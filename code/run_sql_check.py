# -*- coding: utf-8 -*-
"""
run_sql_check.py —— SQL 与 Python 双实现一致性校验

流程：
    1. 建 SQLite 库，把 data/raw 下的两个 CSV 原样载入 ODS 表（全部按 TEXT 载入）
    2. 依次执行 sql/sqlite/01_dwd_clean.sql → 02_feature.sql → 03_rule_preview.sql
    3. 用 SQL 产出的特征表 / 线索集合，与 Python 规则引擎的产出逐行比对

这样做的意义：SQL 只负责清洗与特征，Python 只负责规则，两者独立实现；
若清洗口径或聚合口径有偏差，特征表先对不上，能定位到具体合同。
"""

import csv
import os
import sqlite3
from decimal import Decimal as D

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(BASE, "data", "raw")
OUT = os.path.join(BASE, "out")
SQL = os.path.join(BASE, "sql", "sqlite")
DB = os.path.join(OUT, "check.db")

LINES = []


def log(s=""):
    print(s)
    LINES.append(s)


def q(name):
    return '"' + name.replace('"', '""') + '"'


def load_csv(con, table, path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    head, body = rows[0], rows[1:]
    cols = ", ".join(f"{q(h)} TEXT" for h in head)
    con.execute(f"DROP TABLE IF EXISTS {table}")
    con.execute(f"CREATE TABLE {table} ({cols})")
    ph = ", ".join("?" * len(head))
    con.executemany(f"INSERT INTO {table} VALUES ({ph})", body)
    return len(body)


def num(x):
    return D(str(x)) if x is not None else None


def main():
    os.makedirs(OUT, exist_ok=True)
    if os.path.exists(DB):
        os.remove(DB)
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row

    log("=" * 74)
    log("SQL 与 Python 双实现一致性校验")
    log("=" * 74)

    n1 = load_csv(con, "ods_siku_contract", os.path.join(RAW, "ods_siku_contract.csv"))
    n2 = load_csv(con, "ods_nonsiku_invoice_pay", os.path.join(RAW, "ods_nonsiku_invoice_pay.csv"))
    log(f"\n[1] 载入 ODS：合同表 {n1} 行，发票及支付表 {n2} 行")

    for fn in ["01_dwd_clean.sql", "02_feature.sql", "03_rule_preview.sql"]:
        with open(os.path.join(SQL, fn), encoding="utf-8") as f:
            con.executescript(f.read())
        log(f"[2] 执行 {fn} —— 通过")

    # ---------- 特征表比对 ----------
    sql_feat = {r["contract_code"]: r for r in con.execute("SELECT * FROM dws_contract_feature")}
    py_feat = {}
    with open(os.path.join(OUT, "feature_table.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            py_feat[r["contract_code"]] = r

    log(f"\n[3] 特征表：SQL {len(sql_feat)} 行，Python {len(py_feat)} 行")
    bad = []
    for code in sorted(set(sql_feat) | set(py_feat)):
        s, p = sql_feat.get(code), py_feat.get(code)
        if s is None or p is None:
            bad.append((code, "仅一侧存在"))
            continue
        for col, key in (("c1", "c1"), ("e1", "e1"), ("b2", "b2"), ("s2", "s2")):
            sv = num(s[col])
            pv = num(p[key].replace(",", "")) if p[key] else None
            if (sv is None) != (pv is None) or (sv is not None and abs(sv - pv) > D("0.01")):
                bad.append((code, f"{col} SQL={s[col]} PY={p[key]}"))
    if bad:
        for code, msg in bad[:10]:
            log(f"    ✗ {code}  {msg}")
        log(f"    ✗ 特征表不一致 {len(bad)} 处")
    else:
        log("    ✓ 特征表逐行一致（c1 / e1 / b2 / s2 全部相同）")

    # ---------- 线索集合比对 ----------
    sql_clue = {(r["rule_code"], r["contract_code"]): r for r in con.execute("SELECT * FROM v_rule_result")}
    py_clue = {}
    with open(os.path.join(OUT, "clue_list.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            py_clue[(r["规则编号"], r["合同编号"])] = r

    log(f"\n[4] 线索集合：SQL {len(sql_clue)} 条，Python {len(py_clue)} 条")
    only_sql = set(sql_clue) - set(py_clue)
    only_py = set(py_clue) - set(sql_clue)
    diff_bad = []
    for k in set(sql_clue) & set(py_clue):
        sv = num(sql_clue[k]["diff_amount"])
        pv = num(py_clue[k]["差异金额"].replace(",", ""))
        if abs(sv - pv) > D("0.01"):
            diff_bad.append((k, sv, pv))

    log(f"    SQL 独有 {len(only_sql)} 条，Python 独有 {len(only_py)} 条，金额不一致 {len(diff_bad)} 条")
    for k in sorted(only_sql):
        log(f"      SQL 独有：{k[0]} {k[1]}")
    for k in sorted(only_py):
        log(f"      Python 独有：{k[0]} {k[1]}")
    for k, sv, pv in diff_bad:
        log(f"      金额不一致：{k[0]} {k[1]} SQL={sv} PY={pv}")

    ok = not bad and not only_sql and not only_py and not diff_bad

    # ---------- 数据质量表 ----------
    dq1 = con.execute("SELECT COUNT(*) c FROM dwq_contract_without_invoice").fetchone()["c"]
    dq2 = con.execute("SELECT COUNT(*) c FROM dwq_orphan_invoice").fetchone()["c"]
    log(f"\n[5] 数据质量：司库有合同无报送 {dq1} 条（DQ1），非司库有记录无合同 {dq2} 条（DQ2）")

    log("\n" + "-" * 74)
    log(f"结论：{'通过 —— SQL 与 Python 的线索集合、差异金额完全一致' if ok else '不通过 —— 见上方差异明细'}")
    log("-" * 74)

    with open(os.path.join(OUT, "sql_check_report.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(LINES) + "\n")
    con.close()

    if not ok:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
