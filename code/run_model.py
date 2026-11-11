# -*- coding: utf-8 -*-
"""
run_model.py —— R1 票款背离 / R2 累计结算超合同额 规则引擎（零第三方依赖）

职责划分：
    SQL（sql/*.sql）   负责数据清洗与特征计算，产出「一份合同一行」的特征表
    Python（本文件）    负责规则计算与结果输出，产出线索表与证据项

两套实现独立编写、结果互校（见 code/run_sql_check.py）。
"""

import csv
import json
import os
from decimal import Decimal as D, ROUND_HALF_UP
from datetime import date as _date, timedelta as _td
from collections import Counter as _Counter

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


def rule_r2(f, p):
    c1, e1 = f["c1"], f["e1"]
    over = e1 - c1
    if over <= 0:
        return None
    ratio = ratio_of(over, c1)
    if over <= q2(p["tolerance_amount"]) or ratio <= D(str(p["tolerance_ratio"])):
        return None
    sc = score_of(ratio)
    return dict(rule="R2", rule_name="累计结算超合同额", diff=over, ratio=ratio, score=sc,
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

# --------------------------------------------------------------------------
# 五、V2.1 输出对象：Candidate / Reason / Fact / Evidence / 关系 / 执行元数据
#     依据《御风穿透式监管平台·预置模型输入输出数据规范》第四章
# --------------------------------------------------------------------------
RISK_TYPE = {"R1": ("PAYMENT_INVOICE_MISMATCH", "票款背离"),
             "R2": ("CONTRACT_OVER_SETTLEMENT", "累计结算超合同额")}
REASON_DEF = {
    "R1_less": ("INVOICE_LESS_THAN_PAYMENT", "关联发票合计小于累计支付"),
    "R1_more": ("INVOICE_MORE_THAN_PAYMENT", "关联发票合计大于累计支付"),
    "R2": ("SETTLEMENT_OVER_CONTRACT", "已结算金额超过合同总金额"),
}
LV = {"高": "HIGH", "中": "MEDIUM", "低": "LOW"}

EV_HEAD = ["evidence_id", "detail_id", "source_data_context_id", "source_record_id",
           "contract_id", "contract_name", "buyer_id", "seller_id",
           "contract_amount", "settle_amount", "currency", "sign_date", "framework_flag",
           "invoice_amount", "payment_amount", "data_date", "record_status", "update_time", "ext"]


def _w(path, head, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(head)
        w.writerows(rows)




NUM_COLS = {"numeric_value", "risk_score", "risk_amount", "record_count", "display_order"}
def _vis_len(v):
    """按东亚字符占两格估算显示宽度"""
    if v is None:
        return 0
    return sum(2 if ord(ch) > 127 else 1 for ch in str(v))


NUM_FMT = {"risk_score": "0.0", "risk_amount": "#,##0.00",
           "record_count": "0", "display_order": "0"}


def write_xlsx(path, sheets):
    """把多张子表写入一个工作簿，每类对象一张工作表；列与 V2.1 字段一一对应，不额外加列"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        log("    [SKIP] 未安装 openpyxl，跳过 XLSX 总表（其余输出不受影响）")
        return False

    wb = Workbook()
    wb.remove(wb.active)
    head_font = Font(bold=True, color="FFFFFF", size=10)
    head_fill = PatternFill("solid", fgColor="1C3557")
    for name, head, rows in sheets:
        ws = wb.create_sheet(name)
        ws.append(head)
        for r in rows:
            vals = []
            for h, v in zip(head, r):
                if h in NUM_COLS and v not in ("", None):
                    try:
                        vals.append(float(v))
                        continue
                    except (TypeError, ValueError):
                        pass
                vals.append(v)
            ws.append(vals)
        for c in ws[1]:
            c.font = head_font
            c.fill = head_fill
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for h, fmt in NUM_FMT.items():
            if h in head:
                ci = head.index(h) + 1
                for r in range(2, ws.max_row + 1):
                    ws.cell(row=r, column=ci).number_format = fmt
        if "numeric_value" in head and "unit" in head:
            ui, ni = head.index("unit") + 1, head.index("numeric_value") + 1
            for r in range(2, ws.max_row + 1):
                u = ws.cell(row=r, column=ui).value
                ws.cell(row=r, column=ni).number_format = (
                    "#,##0.00" if u == "CNY" else "0.0000000000" if u == "ratio" else "General")
        from openpyxl.worksheet.properties import PageSetupProperties
        for i, h in enumerate(head, 1):
            est = _vis_len(h)
            for rr in range(2, min(ws.max_row, 30) + 1):
                est = max(est, _vis_len(ws.cell(row=rr, column=i).value))
            ws.column_dimensions[get_column_letter(i)].width = min(46, max(11, est + 2))
        ws.page_setup.orientation = "landscape"
        ws.print_title_rows = "1:1"
    wb.save(path)
    return True


def emit_v21(clues, contracts, inv_by_code, p):
    """按 V2.1 第四章生成输出对象，返回各类计数"""
    import shutil
    snap = p["snapshot_date"]
    snapc = snap.replace("-", "")
    exec_id = f"EXEC{snapc}01"
    obs_s = f"{p['observation_start']}T00:00:00+08:00"
    obs_e = f"{(_date.fromisoformat(p['observation_end']) + _td(days=1)).isoformat()}T00:00:00+08:00"
    det_ts = f"{snap}T18:00:00+08:00"

    ev_dir = os.path.join(OUT, "evidence")
    if os.path.isdir(ev_dir):
        shutil.rmtree(ev_dir)
    os.makedirs(ev_dir, exist_ok=True)

    cands, reasons, facts, sets, details, rels = [], [], [], [], [], []

    for i, c in enumerate(clues, 1):
        cid = c["candidate_id"]
        rt, rtn = RISK_TYPE[c["rule"]]
        amt = q2(c["diff"])
        desc = (f"按差异率分级：≥20% 为高风险、10%–20% 为中风险、5%–10% 为低风险。"
                f"本线索差异率 {pct(c['ratio'])}，涉及金额 {money(amt)} 元，评定为{c['level']}风险。")
        if c["rule"] == "R1":
            summary = (f"合同 {c['contract_code']} 累计发票 {money(c['b2'])} 元、累计支付 "
                       f"{money(c['s2'])} 元，差额 {money(amt)} 元（{pct(c['ratio'])}），票款执行不同步。")
            rkey = "R1_less" if c["direction"] == "付款大于开票" else "R1_more"
        else:
            summary = (f"合同 {c['contract_code']} 合同总金额 {money(c['c1'])} 元、已结算 "
                       f"{money(c['e1'])} 元，超合同额 {money(amt)} 元（{pct(c['ratio'])}）。")
            rkey = "R2"

        cands.append({
            "candidate_id": cid, "execution_id": exec_id,
            "model_code": p["model_code"], "model_version": p["model_version"],
            "model_config_version": "MC-002",
            "subject_type": "ENTERPRISE",
            "subject_id": c["member_code"], "subject_name": c["member_name"],
            "risk_type": rt, "risk_level": LV[c["level"]],
            "risk_level_description": desc,
            "risk_score": f"{c['score']:.1f}", "risk_amount": f"{amt}",
            "confidence": "",
            "risk_summary": summary, "investigation_advice": c["advice"],
            "observation_start": obs_s, "observation_end": obs_e, "detected_at": det_ts,
            "ext": json.dumps({"rule_code": c["rule"], "rule_name": rtn,
                               "contract_code": c["contract_code"],
                               "direction": c["direction"],
                               "business_object_type": "CONTRACT"}, ensure_ascii=False),
        })

        rcode, rtitle = REASON_DEF[rkey]
        rid = f"R-{cid}-01"
        reasons.append({
            "reason_id": rid, "candidate_id": cid, "reason_code": rcode,
            "reason_title": rtitle, "reason_description": c["reason"],
            "importance": LV[c["level"]], "investigation_advice": c["advice"],
            "display_order": 1,
            "ext": json.dumps({"rule_code": c["rule"]}, ensure_ascii=False),
        })

        # ---------- Business Fact ----------
        def add_fact(code, name, vtype, val, unit, cmp_="NONE", bench="", disp=""):
            fid = f"F-{cid}-{len(facts) + 1:02d}"
            row = {"fact_id": fid, "candidate_id": cid, "fact_code": code, "fact_name": name,
                   "value_type": vtype, "numeric_value": "", "string_value": "",
                   "boolean_value": "", "date_value": "", "json_value": "",
                   "unit": unit, "comparison_type": cmp_, "benchmark_value": bench,
                   "period_start": obs_s, "period_end": obs_e,
                   "display_text": disp, "source_execution_id": exec_id, "ext": ""}
            if vtype == "NUMBER":
                row["numeric_value"] = val
            elif vtype == "DATE":
                row["date_value"] = val
            facts.append(row)

        if c["rule"] == "R1":
            add_fact("CONTRACT_AMOUNT", "合同总金额", "NUMBER", f"{c['c1']}", "CNY",
                     disp=f"{money(c['c1'])} 元")
            add_fact("INVOICE_AMOUNT_CUM", "累计发票金额", "NUMBER", f"{c['b2']}", "CNY",
                     disp=f"{money(c['b2'])} 元")
            add_fact("PAYMENT_AMOUNT_CUM", "累计支付金额", "NUMBER", f"{c['s2']}", "CNY",
                     disp=f"{money(c['s2'])} 元")
            add_fact("INVOICE_GAP_AMOUNT", "票款差额", "NUMBER", f"{amt}", "CNY",
                     "GT", f"{q2(p['tolerance_amount'])}", f"{money(amt)} 元")
            add_fact("INVOICE_GAP_RATIO", "票款差异率", "NUMBER", f"{c['ratio']:.10f}", "ratio",
                     "GT", f"{D(str(p['tolerance_ratio']))}", pct(c["ratio"]))
        else:
            add_fact("CONTRACT_AMOUNT", "合同总金额", "NUMBER", f"{c['c1']}", "CNY",
                     disp=f"{money(c['c1'])} 元")
            add_fact("SETTLEMENT_AMOUNT", "已结算金额", "NUMBER", f"{c['e1']}", "CNY",
                     disp=f"{money(c['e1'])} 元")
            add_fact("OVER_CONTRACT_AMOUNT", "超合同额", "NUMBER", f"{amt}", "CNY",
                     "GT", f"{q2(p['tolerance_amount'])}", f"{money(amt)} 元")
            add_fact("OVER_CONTRACT_RATIO", "超合同率", "NUMBER", f"{c['ratio']:.10f}", "ratio",
                     "GT", f"{D(str(p['tolerance_ratio']))}", pct(c["ratio"]))
        add_fact("SNAPSHOT_DATE", "数据快照日", "DATE", snap, "", disp=f"快照日 {snap}")

        # ---------- Evidence Set / Detail ----------
        c_row = contracts[c["contract_code"]]
        plan = []
        if c["rule"] == "R1":
            plan.append(("DOCUMENT_DETAIL", "发票及支付信息报送记录", "DC-NONS-11",
                         inv_by_code.get(c["contract_code"], [])))
        plan.append(("CONTRACT_DETAIL", "合同台账记录", "DC-SIKU-16", [c_row]))

        for seq, (etype, ename, dc, src_rows) in enumerate(plan, 1):
            if not src_rows:
                continue
            eid = f"EV-{cid}-{seq:02d}"
            brows = []
            for k, r in enumerate(src_rows, 1):
                if etype == "CONTRACT_DETAIL":
                    brows.append([eid, f"D-{k:03d}", dc, r["CONTRACT_CODE"],
                                  r["CONTRACT_CODE"], r["CONTRACT_NAME"], r["ACNTMCODE"],
                                  r["OPPOSITE_COMPNAYCODE"], f"{q2(r['CONTRACT_AMOUNT'])}",
                                  f"{q2(r['SETTLEMENT_AMOUNT'])}", r["CURRENCTCODE"], r["SIGN_DATE"],
                                  r["FRAMEWORK_CONTRACT"], "", "", r["DAYIDX"], "", "",
                                  json.dumps({"schema_code": "CONTRACT_DETAIL"}, ensure_ascii=False)])
                else:
                    brows.append([eid, f"D-{k:03d}", dc, f"{r['合同编号']}|{r['更新时间']}",
                                  r["合同编号"], r["合同名称"], r["采购人代码"], "",
                                  "", "", "CNY", "", "", f"{q2(r['累计发票金额'])}",
                                  f"{q2(r['累计支付金额'])}", r["数据日期"], r["状态"], r["更新时间"],
                                  json.dumps({"schema_code": "INVOICE_PAY_SUMMARY_DETAIL"}, ensure_ascii=False)])
            _w(os.path.join(ev_dir, f"{eid}.csv"), EV_HEAD, brows)
            details.extend(brows)
            sets.append({
                "evidence_id": eid, "candidate_id": cid, "evidence_type": etype,
                "evidence_name": f"合同 {c['contract_code']} {ename}",
                "detail_dataset_ref": f"out/evidence/{eid}.csv",
                "record_count": len(brows), "source_execution_id": exec_id,
                "generated_at": det_ts,
                "ext": json.dumps({"schema_code": brows[0][18] and json.loads(brows[0][18])["schema_code"]},
                                  ensure_ascii=False),
            })
            rels.append({
                "reason_id": rid, "evidence_id": eid, "relation_type": "SUPPORT",
                "relation_strength": "1.0", "created_at": det_ts,
                "source_execution_id": exec_id, "ext": "",
            })

    _w(os.path.join(OUT, "candidate.csv"), list(cands[0].keys()), [list(x.values()) for x in cands])
    _w(os.path.join(OUT, "reason.csv"), list(reasons[0].keys()), [list(x.values()) for x in reasons])
    _w(os.path.join(OUT, "fact.csv"), list(facts[0].keys()), [list(x.values()) for x in facts])
    _w(os.path.join(OUT, "evidence_set.csv"), list(sets[0].keys()), [list(x.values()) for x in sets])
    _w(os.path.join(OUT, "evidence_detail.csv"), EV_HEAD, details)
    _w(os.path.join(OUT, "explanation_evidence_rel.csv"), list(rels[0].keys()),
       [list(x.values()) for x in rels])

    execution = [{
        "execution_id": exec_id, "model_code": p["model_code"], "model_version": p["model_version"],
        "contract_version": "2.0", "model_config_version": "MC-002",
        "observation_start": obs_s, "observation_end": obs_e,
        "started_at": det_ts, "finished_at": det_ts, "execution_status": "SUCCESS",
        "candidate_count": len(cands), "engine_version": "python-stdlib-3",
        "triggered_by": "用户 / 定时调度",
        "request_key": f"REQ-{snapc}-01",
        "runtime_snapshot_ref": "out/runtime_snapshot.json",
        "result_manifest_ref": "out/manifest.json",
        "replay_of_execution_id": "", "ext": "",
    }]
    _w(os.path.join(OUT, "execution.csv"), list(execution[0].keys()), [list(x.values()) for x in execution])

    with open(os.path.join(OUT, "runtime_snapshot.json"), "w", encoding="utf-8") as f:
        json.dump(p, f, ensure_ascii=False, indent=2)
    with open(os.path.join(OUT, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump({"execution_id": exec_id, "files": [
            "candidate.csv", "reason.csv", "fact.csv", "evidence_set.csv",
            "evidence_detail.csv", "explanation_evidence_rel.csv", "execution.csv"],
            "evidence_dir": "evidence/"}, f, ensure_ascii=False, indent=2)

    cnt = _Counter(x["candidate_id"] for x in sets)
    return dict(candidates=len(cands), reasons=len(reasons), facts=len(facts),
                evidence_sets=len(sets), evidence_details=len(details), rels=len(rels),
                min_sets=min(cnt.values()) if cnt else 0,
                _cand=cands, _reason=reasons, _fact=facts, _set=sets, _detail=details)


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

    inv_by_code = {}
    for r in raw_i:
        inv_by_code.setdefault(r["合同编号"], []).append(r)
    for v in inv_by_code.values():
        v.sort(key=lambda x: x["更新时间"], reverse=True)

    raw_codes = {r["CONTRACT_CODE"] for r in raw_c}
    feats, dq = build_feature(contracts, invoices, drop_c, drop_i, raw_codes)
    log(f"[3] 特征表：{len(feats)} 行（一份合同一行），数据质量与排除记录 {len(dq)} 条")

    clues, evid = [], []
    for f in feats:
        c_row = contracts[f["contract_code"]]
        i_row = invoices.get(f["contract_code"])
        for fn in (rule_r2, rule_r1):
            hit = fn(f, p)
            if hit:
                hit.update(f)
                clues.append(hit)

    lv_order = {"高": 0, "中": 1, "低": 2}
    clues.sort(key=lambda x: (x["rule"], lv_order[x["level"]], -x["score"], -x["diff"]))
    for i, c in enumerate(clues, 1):
        c["clue_id"] = f"CLUE-{p['snapshot_date'].replace('-', '')}-{i:04d}"
        c["candidate_id"] = f"C{p['snapshot_date'].replace('-', '')}{i:04d}"
        for k, (tbl, fld, val, why) in enumerate(
                build_evidence(c, contracts[c["contract_code"]],
                               invoices.get(c["contract_code"])), 1):
            evid.append([c["clue_id"], k, tbl, fld, val, why])

    log(f"[4] 规则计算：命中 {len(clues)} 条线索（R1 {sum(1 for c in clues if c['rule']=='R1')} 条，"
        f"R2 {sum(1 for c in clues if c['rule']=='R2')} 条）")
    for c in clues:
        log(f"    {c['candidate_id']}  {c['rule']}  {c['contract_code']:<12} {c['level']}  "
            f"评分 {c['score']:>5}  差异 {money(c['diff']):>15} 元  {pct(c['ratio']):>7}  {c['direction']}")

    # ---------- 输出 ----------
    v21 = emit_v21(clues, contracts, inv_by_code, p)

    # 四类对象合并为一个工作簿：每类对象一张子表，列严格照 V2.1
    ok_xlsx = write_xlsx(os.path.join(OUT, "输出总表.xlsx"), [
        ("Risk Candidate", list(v21["_cand"][0].keys()), [list(x.values()) for x in v21["_cand"]]),
        ("Reason", list(v21["_reason"][0].keys()), [list(x.values()) for x in v21["_reason"]]),
        ("Business Fact", list(v21["_fact"][0].keys()), [list(x.values()) for x in v21["_fact"]]),
        ("Evidence Set", list(v21["_set"][0].keys()), [list(x.values()) for x in v21["_set"]]),
    ])

    # 人读视图（由 Candidate + Fact 派生，非 V2.1 规范交付物）
    head = ["线索ID", "规则编号", "规则名称", "合同编号", "合同名称", "成员单位",
            "合同类型", "合同金额C1", "已结算金额E1", "累计发票金额B2", "累计支付金额S2",
            "差异金额", "差异率", "风险评分", "风险等级", "风险理由", "排查建议",
            "快照日", "数据来源"]
    rows = []
    for c in clues:
        rows.append([c["candidate_id"], c["rule"], c["rule_name"], c["contract_code"],
                     c["contract_name"], c["member_name"], c["contract_type"],
                     money(c["c1"]), money(c["e1"]),
                     "" if c["b2"] is None else money(c["b2"]),
                     "" if c["s2"] is None else money(c["s2"]),
                     money(c["diff"]), pct(c["ratio"]), f"{c['score']:.1f}", c["level"],
                     c["reason"], c["advice"], p["snapshot_date"],
                     "司库第16张合同表" + ("；非司库第11张发票及支付信息表" if c["b2"] is not None else "")])
    _w(os.path.join(OUT, "线索汇总_人读视图.csv"), head, rows)

    _w(os.path.join(OUT, "excluded_list.csv"),
       ["合同编号", "合同名称", "类别", "编号", "说明"],
       [[d["contract_code"], d["contract_name"], d["level"], d["code"], d["reason"]] for d in dq])

    feat_head = ["contract_code", "contract_name", "member_name", "c1", "e1", "b2", "s2", "has_nonsiku"]
    _w(os.path.join(OUT, "feature_table.csv"), feat_head,
       [[x["contract_code"], x["contract_name"], x["member_name"],
         money(x["c1"]), money(x["e1"]),
         "" if x["b2"] is None else money(x["b2"]),
         "" if x["s2"] is None else money(x["s2"]), x["has_nonsiku"]] for x in feats])

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

    ev_ok = (v21["min_sets"] >= 1 and v21["evidence_details"] >= v21["candidates"]
             and v21["rels"] == v21["evidence_sets"])
    checks.append(("V2.1 证据对象完整", ev_ok,
                   f"证据集 {v21['evidence_sets']} 组、明细 {v21['evidence_details']} 行、关联 {v21['rels']} 条"))

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
        "r2_total": sum(1 for c in clues if c["rule"] == "R2"),
        "lv_high": sum(1 for c in clues if c["level"] == "高"),
        "lv_mid": sum(1 for c in clues if c["level"] == "中"),
        "lv_low": sum(1 for c in clues if c["level"] == "低"),
        "dq_total": len(dq),
        "evidence_total": v21["evidence_details"],
        "hit_rate": round(len(clues) / max(1, len(feats)) * 100, 1),
    }
    log(f"\n[6] 汇总：{summary['contract_judged']} 份合同参与判定，命中 {summary['clue_total']} 条线索"
        f"（高 {summary['lv_high']} / 中 {summary['lv_mid']} / 低 {summary['lv_low']}），"
        f"命中率 {summary['hit_rate']}%，证据明细 {summary['evidence_total']} 行，"
        f"数据质量与排除 {summary['dq_total']} 条")
    log("=" * 76)

    with open(os.path.join(OUT, "run_log.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(LOG) + "\n")

    payload = dict(params=p, summary=summary,
                   clues=[{k: (str(v) if isinstance(v, D) else v) for k, v in c.items()}
                          for c in clues],
                   features=[{k: (str(v) if isinstance(v, D) else v) for k, v in x.items()}
                             for x in feats],
                   evidence=evid, dq=dq,
                   v21={k: v21[k] for k in ("_cand", "_reason", "_fact", "_set", "_detail")})
    with open(os.path.join(OUT, "clues.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)

    print(f"\n✓ V2.1 输出：Candidate {v21['candidates']} ｜ Reason {v21['reasons']} ｜ "
          f"Fact {v21['facts']} ｜ Evidence Set {v21['evidence_sets']} ｜ "
          f"Detail {v21['evidence_details']} ｜ 关联 {v21['rels']}")
    print("  文件：out/candidate.csv、reason.csv、fact.csv、evidence_set.csv、"
          "evidence_detail.csv、explanation_evidence_rel.csv、execution.csv、evidence/")
    if ok_xlsx:
        print("  总表：out/输出总表.xlsx（4 张子表：Risk Candidate / Reason / Business Fact / Evidence Set）")


if __name__ == "__main__":
    main()
