# -*- coding: utf-8 -*-
"""
run_model.py —— R1 票款背离 / R4 累计结算超合同额 规则引擎（零第三方依赖）

职责划分：
    SQL（sql/*.sql）   负责数据清洗与特征计算，产出「一份合同一行」的特征表
    Python（本文件）    负责规则计算与结果输出，产出线索表与证据项

两套实现独立编写、结果互校（见 code/run_sql_check.py）。
"""

import csv
import json
import os
from decimal import Decimal as D, ROUND_HALF_UP

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(BASE, "data", "raw")
OUT = os.path.join(BASE, "out")

LEVEL_HIGH, LEVEL_MID = 80.0, 40.0
LOG = []


def log(s=""):
    print(s)
    LOG.append(s)


def money(x):
    return f"{D(x):,.2f}"


def pct(x, n=2):
    return f"{(D(x) * 100):.{n}f}%"


def q2(x):
    return D(x).quantize(D("0.01"), rounding=ROUND_HALF_UP)


def read_csv(name):
    path = os.path.join(RAW, name)
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


# --------------------------------------------------------------------------
# 一、清洗：与 sql/01_dwd_clean.sql 一一对应
# --------------------------------------------------------------------------
def clean_contract(rows, p):
    kept, dropped = {}, []
    for r in rows:
        if r["DAYIDX"] != p["snapshot_date"]:
            dropped.append((r["CONTRACT_CODE"], "数据日期不在快照日"))
            continue
        if r["FRAMEWORK_CONTRACT"] != p["framework_filter"]:
            dropped.append((r["CONTRACT_CODE"], "框架合同（FRAMEWORK_CONTRACT≠T00）"))
            continue
        if r["CURRENCTCODE"] != p["currency_filter"]:
            dropped.append((r["CONTRACT_CODE"], f"币种为 {r['CURRENCTCODE']}，非 CNY"))
            continue
        if D(r["CONTRACT_AMOUNT"]) <= 0:
            dropped.append((r["CONTRACT_CODE"], "合同总金额≤0，无判定基准"))
            continue
        # 同一合同多条时取签订日期最新一条（本合同表一合同一行，防御性保留）
        old = kept.get(r["CONTRACT_CODE"])
        if old is None or r["SIGN_DATE"] >= old["SIGN_DATE"]:
            kept[r["CONTRACT_CODE"]] = r
    return kept, dropped


def clean_invoice(rows, p):
    """状态=3 剔除；同合同取更新时间最新一条。返回 (最新记录字典, 剔除明细)"""
    latest, dropped = {}, []
    for r in rows:
        if r["数据日期"] != p["snapshot_date"]:
            dropped.append((r["合同编号"], f"数据日期={r['数据日期']}，非快照日"))
            continue
        if r["状态"] == "3":
            dropped.append((r["合同编号"], f"状态=3 删除记录（更新时间 {r['更新时间']}）"))
            continue
        cur = latest.get(r["合同编号"])
        if cur is None or r["更新时间"] > cur["更新时间"]:
            if cur is not None:
                dropped.append((r["合同编号"], f"被更新版本覆盖（{cur['更新时间']}）"))
            latest[r["合同编号"]] = r
        else:
            dropped.append((r["合同编号"], f"被更新版本覆盖（{r['更新时间']}）"))
    return latest, dropped


# --------------------------------------------------------------------------
# 二、特征：一份合同一行（C1 / E1 / B2 / S2）
# --------------------------------------------------------------------------
def build_feature(contracts, invoices, dropped_contract, dropped_invoice, raw_codes):
    feats, dq = [], []
    for code, c in contracts.items():
        inv = invoices.get(code)
        if inv is None:
            dq.append(dict(contract_code=code, contract_name=c["CONTRACT_NAME"],
                           level="数据质量", code="DQ1",
                           reason="司库合同表有该合同，非司库发票及支付信息表无对应记录，R1 无法判定"))
            b2 = s2 = None
        else:
            b2, s2 = q2(inv["累计发票金额"]), q2(inv["累计支付金额"])
        feats.append(dict(
            contract_code=code,
            contract_name=c["CONTRACT_NAME"],
            member_code=c["ACNTMCODE"],
            member_name=c["ACNTMNAME"],
            vendor_name=c["OPPOSITE_COMPNAY"],
            contract_type=c["CONTRACT_TYPE"],
            billing_type=c["BILLING_TYPE"],
            sign_date=c["SIGN_DATE"],
            c1=q2(c["CONTRACT_AMOUNT"]),
            e1=q2(c["SETTLEMENT_AMOUNT"]),
            b2=b2, s2=s2,
            has_nonsiku=inv is not None,
        ))
    for code in invoices:
        if code not in raw_codes:
            dq.append(dict(contract_code=code, contract_name=invoices[code]["合同名称"],
                           level="数据质量", code="DQ2",
                           reason="非司库有该合同编号，司库合同表无对应合同，无法参与任一规则判定"))
    for code, why in dropped_contract:
        dq.append(dict(contract_code=code, contract_name="", level="范围排除",
                       code="EX1", reason=f"16 合同表：{why}"))
    for code, why in dropped_invoice:
        dq.append(dict(contract_code=code, contract_name="", level="数据质量",
                       code="DQ3", reason=f"11 发票及支付表：{why}"))
    feats.sort(key=lambda x: x["contract_code"])
    return feats, dq


# --------------------------------------------------------------------------
# 三、规则：R1 / R4
# --------------------------------------------------------------------------
def ratio_of(diff, base):
    return (diff / base) if base and base > 0 else D("0")


def score_of(ratio):
    """评分 = 差异率对 25% 的归一化，封顶 100。等级由评分切分，二者恒不矛盾。"""
    s = float(ratio / D("0.25") * 100)
    return round(min(100.0, s), 1)


def level_of(score):
    return "高" if score >= LEVEL_HIGH else ("中" if score >= LEVEL_MID else "低")


def rule_r4(f, p):
    c1, e1 = f["c1"], f["e1"]
    over = e1 - c1
    if over <= 0:
        return None
    ratio = ratio_of(over, c1)
    if over <= q2(p["tolerance_amount"]) or ratio <= D(str(p["tolerance_ratio"])):
        return None
    sc = score_of(ratio)
    return dict(rule="R4", rule_name="累计结算超合同额", diff=over, ratio=ratio, score=sc,
                level=level_of(sc), direction="超合同额",
                reason=(f"合同 {f['contract_code']} 合同总金额 {money(c1)} 元，"
                        f"已结算金额 {money(e1)} 元，超合同额 {money(over)} 元（{pct(ratio)}）。"
                        f"已结算金额突破合同约定上限，可能存在未登记合同变更、超权限结算或重复结算。"),
                advice="核查合同变更与补充协议是否已登记入账；核对结算审批权限与结算单金额；"
                       "确认是否存在跨合同串用或重复结算。")


def rule_r1(f, p):
    b2, s2 = f["b2"], f["s2"]
    if b2 is None or s2 is None:
        return None
    diff = abs(b2 - s2)
    base = max(b2, s2)
    ratio = ratio_of(diff, base)
    if diff <= q2(p["tolerance_amount"]) or ratio <= D(str(p["tolerance_ratio"])):
        return None
    sc = score_of(ratio)
    if s2 > b2:
        direction = "付款大于开票"
        reason = (f"合同 {f['contract_code']} 累计支付 {money(s2)} 元，累计发票 {money(b2)} 元，"
                  f"付款大于开票 {money(diff)} 元（{pct(ratio)}）。可能存在已付款未取得发票、"
                  f"预付款未冲销或供应商延迟开票。")
        advice = "核查开票申请与实际到票情况；确认是否存在预付款条款；" \
                 "比对付款审批单与发票签收记录。"
    else:
        direction = "开票大于付款"
        reason = (f"合同 {f['contract_code']} 累计发票 {money(b2)} 元，累计支付 {money(s2)} 元，"
                  f"开票大于付款 {money(diff)} 元（{pct(ratio)}）。可能存在先票后款、预开票或红冲未剔除。")
        advice = "核对合同约定账期与付款节点；确认是否存在预开票；" \
                 "核实累计发票金额是否已剔除红冲与作废发票。"
    return dict(rule="R1", rule_name="票款背离", diff=diff, ratio=ratio, score=sc,
                level=level_of(sc), direction=direction, reason=reason, advice=advice)


def build_evidence(f, c_row, i_row):
    ev = []
    for fld, why in [("CONTRACT_CODE", "合同唯一编码，聚合主键"),
                     ("CONTRACT_NAME", "合同名称"),
                     ("CONTRACT_AMOUNT", "合同总金额，R4 分母"),
                     ("SETTLEMENT_AMOUNT", "已结算金额，R4 分子"),
                     ("FRAMEWORK_CONTRACT", "框架标记，T00 参与判定"),
                     ("CURRENCTCODE", "币种"),
                     ("BILLING_TYPE", "结算类型，N02 为付款"),
                     ("DAYIDX", "数据日期")]:
        ev.append(("司库第16张 合同", fld, c_row[fld], why))
    if i_row:
        for fld, why in [("合同编号", "与 CONTRACT_CODE 对齐"),
                         ("累计发票金额", "R1 一侧 B2"),
                         ("累计支付金额", "R1 另一侧 S2"),
                         ("数据日期", "快照日"),
                         ("更新时间", "多版本取最新依据"),
                         ("状态", "1 新增 / 2 修改 / 3 删除")]:
            ev.append(("非司库第11张 发票及支付信息表", fld, i_row[fld], why))
    return ev


# --------------------------------------------------------------------------
# 四、主流程
# --------------------------------------------------------------------------
def main():
    os.makedirs(OUT, exist_ok=True)
    p = json.load(open(os.path.join(BASE, "data", "params.json"), encoding="utf-8"))

    log("=" * 76)
    log("合同台账与票款执行一致性校验模型 · 规则引擎运行日志")
    log(f"模型编码 {p['model_code']} ｜ 版本 {p['model_version']} ｜ 快照日 {p['snapshot_date']}")
    log(f"容差：绝对 {money(p['tolerance_amount'])} 元  且  相对 {pct(p['tolerance_ratio'])}（两门槛同时满足）")
    log("=" * 76)

    raw_c = read_csv("ods_siku_contract.csv")
    raw_i = read_csv("ods_nonsiku_invoice_pay.csv")
    log(f"\n[1] 读入原始数据：16 合同表 {len(raw_c)} 行，11 发票及支付表 {len(raw_i)} 行")

    contracts, drop_c = clean_contract(raw_c, p)
    invoices, drop_i = clean_invoice(raw_i, p)
    log(f"[2] 清洗后：合同表 {len(contracts)} 份（剔除 {len(drop_c)} 条），"
        f"发票表 {len(invoices)} 份（剔除 {len(drop_i)} 条）")

    raw_codes = {r["CONTRACT_CODE"] for r in raw_c}
    feats, dq = build_feature(contracts, invoices, drop_c, drop_i, raw_codes)
    log(f"[3] 特征表：{len(feats)} 行（一份合同一行），数据质量与排除记录 {len(dq)} 条")

    clues, evid = [], []
    for f in feats:
        c_row = contracts[f["contract_code"]]
        i_row = invoices.get(f["contract_code"])
        for fn in (rule_r4, rule_r1):
            hit = fn(f, p)
            if hit:
                hit.update(f)
                clues.append(hit)

    lv_order = {"高": 0, "中": 1, "低": 2}
    clues.sort(key=lambda x: (x["rule"], lv_order[x["level"]], -x["score"], -x["diff"]))
    for i, c in enumerate(clues, 1):
        c["clue_id"] = f"CLUE-{p['snapshot_date'].replace('-', '')}-{i:04d}"
        for k, (tbl, fld, val, why) in enumerate(
                build_evidence(c, contracts[c["contract_code"]],
                               invoices.get(c["contract_code"])), 1):
            evid.append([c["clue_id"], k, tbl, fld, val, why])

    log(f"[4] 规则计算：命中 {len(clues)} 条线索（R1 {sum(1 for c in clues if c['rule']=='R1')} 条，"
        f"R4 {sum(1 for c in clues if c['rule']=='R4')} 条）")
    for c in clues:
        log(f"    {c['clue_id']}  {c['rule']}  {c['contract_code']:<12} {c['level']}  "
            f"评分 {c['score']:>5}  差异 {money(c['diff']):>15} 元  {pct(c['ratio']):>7}  {c['direction']}")

    # ---------- 输出 ----------
    head = ["线索ID", "规则编号", "规则名称", "合同编号", "合同名称", "成员单位",
            "合同类型", "合同金额C1", "已结算金额E1", "累计发票金额B2", "累计支付金额S2",
            "差异金额", "差异率", "风险评分", "风险等级", "风险理由", "排查建议",
            "快照日", "数据来源"]
    rows = []
    for c in clues:
        rows.append([c["clue_id"], c["rule"], c["rule_name"], c["contract_code"], c["contract_name"],
                     c["member_name"], c["contract_type"], money(c["c1"]), money(c["e1"]),
                     "" if c["b2"] is None else money(c["b2"]),
                     "" if c["s2"] is None else money(c["s2"]),
                     money(c["diff"]), pct(c["ratio"]), f"{c['score']:.1f}", c["level"],
                     c["reason"], c["advice"], p["snapshot_date"],
                     "司库第16张合同表" + ("；非司库第11张发票及支付信息表" if c["b2"] is not None else "")])
    with open(os.path.join(OUT, "clue_list.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(head)
        w.writerows(rows)

    with open(os.path.join(OUT, "evidence.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["线索ID", "序号", "来源表", "来源字段", "取值", "说明"])
        w.writerows(evid)

    with open(os.path.join(OUT, "excluded_list.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["合同编号", "合同名称", "类别", "编号", "说明"])
        for d in dq:
            w.writerow([d["contract_code"], d["contract_name"], d["level"], d["code"], d["reason"]])

    feat_head = ["contract_code", "contract_name", "member_name", "c1", "e1", "b2", "s2", "has_nonsiku"]
    with open(os.path.join(OUT, "feature_table.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(feat_head)
        for x in feats:
            w.writerow([x["contract_code"], x["contract_name"], x["member_name"],
                        money(x["c1"]), money(x["e1"]),
                        "" if x["b2"] is None else money(x["b2"]),
                        "" if x["s2"] is None else money(x["s2"]), x["has_nonsiku"]])

    # ---------- 自检 ----------
    log("\n[5] 自检")
    checks = []

    lv_ok = all((c["level"] == "高") == (c["score"] >= LEVEL_HIGH) and
                (c["level"] == "中") == (LEVEL_MID <= c["score"] < LEVEL_HIGH)
                for c in clues)
    checks.append(("评分与等级一致", lv_ok, "所有线索的等级由评分切分得到，无矛盾"))

    tol_ok = all(c["diff"] > q2(p["tolerance_amount"]) and
                 c["ratio"] > D(str(p["tolerance_ratio"])) for c in clues)
    checks.append(("双门槛同时满足", tol_ok, "无一例仅过单一门槛"))

    ev_ok = all(sum(1 for e in evid if e[0] == c["clue_id"]) >= 6 for c in clues)
    checks.append(("每条线索证据≥6项", ev_ok, f"证据项合计 {len(evid)} 条"))

    dir_ok = all(c["direction"] for c in clues)
    checks.append(("差异方向已标注", dir_ok, "付款大于开票 / 开票大于付款 / 超合同额"))

    uniq = len({(c["rule"], c["contract_code"]) for c in clues}) == len(clues)
    checks.append(("线索唯一", uniq, "按「规则编号+合同编号」唯一，一合同可出多条但规则不重复"))

    zero_ok = all(c["diff"] > 0 for c in clues)
    checks.append(("无零差异线索", zero_ok, "容差过滤生效"))

    for name, ok, note in checks:
        log(f"    [{'PASS' if ok else 'FAIL'}] {name} —— {note}")
    if not all(x[1] for x in checks):
        raise SystemExit("自检未通过，终止输出")

    summary = {
        "contract_total": len(raw_c),
        "contract_judged": len(feats),
        "invoice_versions": len(raw_i),
        "invoice_latest": len(invoices),
        "clue_total": len(clues),
        "r1_total": sum(1 for c in clues if c["rule"] == "R1"),
        "r4_total": sum(1 for c in clues if c["rule"] == "R4"),
        "lv_high": sum(1 for c in clues if c["level"] == "高"),
        "lv_mid": sum(1 for c in clues if c["level"] == "中"),
        "lv_low": sum(1 for c in clues if c["level"] == "低"),
        "dq_total": len(dq),
        "evidence_total": len(evid),
        "hit_rate": round(len(clues) / max(1, len(feats)) * 100, 1),
    }
    log(f"\n[6] 汇总：{summary['contract_judged']} 份合同参与判定，命中 {summary['clue_total']} 条线索"
        f"（高 {summary['lv_high']} / 中 {summary['lv_mid']} / 低 {summary['lv_low']}），"
        f"命中率 {summary['hit_rate']}%，证据项 {summary['evidence_total']} 条，"
        f"数据质量与排除 {summary['dq_total']} 条")
    log("=" * 76)

    with open(os.path.join(OUT, "run_log.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(LOG) + "\n")

    payload = dict(params=p, summary=summary,
                   clues=[{k: (str(v) if isinstance(v, D) else v) for k, v in c.items()}
                          for c in clues],
                   features=[{k: (str(v) if isinstance(v, D) else v) for k, v in x.items()}
                             for x in feats],
                   evidence=evid, dq=dq)
    with open(os.path.join(OUT, "clues.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)

    print(f"\n✓ 已输出 out/clue_list.csv（{len(rows)} 条）、out/evidence.csv（{len(evid)} 条）、"
          f"out/excluded_list.csv、out/feature_table.csv、out/run_log.txt")


if __name__ == "__main__":
    main()
