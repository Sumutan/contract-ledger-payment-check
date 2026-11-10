# -*- coding: utf-8 -*-
"""
build_doc.py —— 生成《合同台账与票款执行一致性校验模型·模型交付文档》

章节结构对齐考核要求（内容要求 1）：
    业务背景与政策依据 → 模型目标与适用范围（含纳入/排除）→ 输入数据与数据标准
    → 特征清单与计算口径 → 策略规则 → 阈值与参数配置
    → 输出字段（共性/溯源/个性）→ 效果评估与验收口径 → 风险与局限 → 复现步骤
"""

import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEST = os.path.join(BASE, "docs", "合同台账与票款执行一致性校验模型-模型说明.docx")

NAVY = RGBColor(0x1C, 0x35, 0x57)
INK = RGBColor(0x16, 0x20, 0x2E)
GREY = RGBColor(0x5A, 0x66, 0x78)
NOTE = RGBColor(0x3D, 0x35, 0x27)
USABLE = 16.6
LATIN = {"宋体": "Times New Roman", "黑体": "Arial"}


def set_font(run, name="宋体", size=10.5, bold=False, color=INK):
    latin = LATIN.get(name, "Times New Roman")
    run.font.name = latin
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    rf.set(qn("w:eastAsia"), name)
    rf.set(qn("w:ascii"), latin)
    rf.set(qn("w:hAnsi"), latin)


def para(doc, text="", name="宋体", size=10.5, bold=False, color=INK,
         before=0, after=6, line=1.55, indent=0, align=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.left_indent = Cm(indent)
    if align is not None:
        p.alignment = align
    if text:
        set_font(p.add_run(text), name, size, bold, color)
    return p


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def cell_margins(t, top=40, bottom=40, left=80, right=80):
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{k}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    t._tbl.tblPr.append(mar)


def add_table(doc, head, rows, widths, size=8.8):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    cell_margins(t)
    s = sum(widths)
    widths = [w / s * USABLE for w in widths]
    for i, w in enumerate(widths):
        t.columns[i].width = Cm(w)
    hdr = t.rows[0]
    trPr = hdr._tr.get_or_add_trPr()
    e = OxmlElement("w:tblHeader")
    e.set(qn("w:val"), "true")
    trPr.append(e)
    for i, h in enumerate(head):
        c = hdr.cells[i]
        c.width = Cm(widths[i])
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        set_font(p.add_run(h), "黑体", size, True, NAVY)
        shade(c, "EFF3F8")
    for r in rows:
        row = t.add_row()
        trPr2 = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        cs.set(qn("w:val"), "true")
        trPr2.append(cs)
        for i, v in enumerate(r):
            c = row.cells[i]
            c.width = Cm(widths[i])
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.2
            set_font(p.add_run(str(v)), "宋体", size, False, INK)
    if len(rows) <= 10:
        for r in t.rows[:-1]:
            for c in r.cells:
                for cp in c.paragraphs:
                    cp.paragraph_format.keep_with_next = True
    para(doc, "", size=5, after=0)
    return t


def h1(doc, text, newpage=False):
    p = doc.add_paragraph()
    if newpage:
        p.paragraph_format.page_break_before = True
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), "黑体", 14, True, NAVY)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), "6")
    b.set(qn("w:color"), "C9D3E0")
    bdr.append(b)
    pPr.append(bdr)


def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    set_font(p.add_run(text), "黑体", 11.5, True, NAVY)


def bullets(doc, items):
    for it in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.5
        set_font(p.add_run("· "), "宋体", 10.5, True, NAVY)
        set_font(p.add_run(it), "宋体", 10.5)


def formula(doc, lines, size=9):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.keep_together = True
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), "F5F8FC")
    pPr.append(shd)
    for i, ln in enumerate(lines):
        r = p.add_run(("" if i == 0 else "\n") + ln)
        set_font(r, "宋体", size, False, RGBColor(0x2C, 0x3E, 0x55))


def footer_pagenum(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    set_font(run, "宋体", 9, False, GREY)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    run._r.addnext(fld)


def main():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.0)
    footer_pagenum(sec)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    para(doc, "模型交付文档", "宋体", 10, False, GREY, after=3)
    para(doc, "合同台账与票款执行一致性校验模型", "黑体", 21, True, NAVY, after=4)
    para(doc, "作业自建模型 · 不编 CL 编号 · 两表零 join 落地版 · v1.0.0",
         "宋体", 9.5, False, GREY, after=12)

    # ==================== 一 ====================
    h1(doc, "一、业务背景与政策依据")
    para(doc, "企业同时向国资委报送两套数：司库 19 张表、非司库数据表。这两套数是对外承诺的数字，"
              "企业内部还有一套账本上的数字。数字对不对得上，过去只能靠人工抽查：抽十份合同，"
              "能对上的放过去，对不上的再问。抽查的覆盖率取决于人，不取决于数。")
    para(doc, "本模型把三个口径按合同编号拉到一行，只查两件最基础、也最容易验证的事："
              "同一份合同，报的累计发票额与累计支付额对不对得上（R1）；"
              "同一份合同，报的已结算额有没有突破合同额（R2）。"
              "两条规则各读一张表、各算一个差、各出一批线索，不合并评分、不互相依赖。")
    para(doc, "监管动机链条：国资委要求央企按统一标准报送数据（数据底座）→ 报送数据必须反映真实业务，"
              "不得以虚假资料核算 → 但企业同时存在「报送口径」与「业务台账」两套数 → 两套数之间的差异，"
              "可能掩盖超付、虚开发票、资金体外循环 → 以报送数据为基准做自洽校验，差异即为待核线索。")
    add_table(doc, ["政策文件", "发布主体", "对应条款 / 要求", "约束点"], [
        ["《关于推动中央企业加快财务数智化转型升级的指导意见》（1 号文）", "国务院国资委",
         "建立问题识别机制，构建穿透监督模型", "授权层：回答为什么建这个模型"],
        ["《关于加强中央企业穿透式监管的指导意见（试行）》（2 号文）", "国务院国资委",
         "数据采集、智能分析、风险预警、核查派单、整改跟踪形成闭环", "闭环层：线索如何被使用"],
        ["《中华人民共和国会计法》第九条", "全国人大常委会",
         "各单位必须根据实际发生的经济业务事项进行会计核算；任何单位不得以虚假的经济业务事项或者资料进行会计核算",
         "R1：同一笔业务在两套报送口径下必须能对上"],
        ["《中华人民共和国发票管理办法》", "国务院",
         "收款方应当向付款方开具发票；不得虚开发票", "R1：票款背离背后的涉税风险"],
        ["《企业内部控制应用指引第 16 号——合同管理》", "财政部等五部委",
         "合同履行应实行监控，结算与合同约定保持一致；合同变更应履行审批", "R2：已结算不得突破合同额"],
        ["《中央企业违规经营投资责任追究实施办法》", "国务院国资委",
         "未按规定履行合同、超付合同款造成资产损失的，应追究责任", "处置依据，非判定依据"],
    ], [4.4, 2.5, 6.2, 3.5])
    para(doc, "说明：两条规则是内控逻辑在既有字段上的翻译，不是发文条款的提炼。政策只提供授权与方向，"
              "不提供具体阈值，因此文中不出现「依据某号文第 X 条判定」的写法。", "宋体", 9, False, GREY, before=2)

    # ==================== 二 ====================
    h1(doc, "二、模型目标与适用范围", newpage=True)
    add_table(doc, ["项目", "内容"], [
        ["模型名称", "合同台账与票款执行一致性校验模型"],
        ["模型编码 / 版本", "CONTRACT_LEDGER_PAYMENT_CHECK / v1.0.0（2026-11）"],
        ["所属领域", "合同 / 财务"],
        ["模型目标", "在不引入企业内部系统数据的前提下，仅用两套报送数据，自动发现"
                     "「票与款不同步」和「结算突破合同额」两类待核线索"],
        ["判定输入", "2 张表：司库第 16 张 合同表、非司库第 11 张 发票及支付信息表。"
                     "一份合同一行，两条规则各读一张表，互不 join"],
        ["覆盖规则", "R1 票款背离、R2 累计结算超合同额"],
        ["不覆盖", "R2 台账割裂（需企业 IT 四表）、框架分组合计（R3）、逐笔付款对逐张发票"],
        ["输出粒度", "合同编号。一条线索一行，同一合同可同时出两条线索"],
    ], [2.6, 14.0])

    h2(doc, "2.1 纳入条件")
    add_table(doc, ["条件", "取值", "依据"], [
        ["观察期", "2026-01-01 ~ 2026-06-30", "按报送周期取数，两条规则共用"],
        ["快照日", "2026-06-30", "11 表按「数据日期」、16 表按 DAYIDX 对齐"],
        ["合同类型", "FRAMEWORK_CONTRACT = 'T00'", "规范对合同额、已结算额均注明「非框架合同填写」"],
        ["币种", "CURRENCTCODE = 'CNY'", "外币需折算，折算口径未定，本版直接排除"],
        ["金额基准", "CONTRACT_AMOUNT > 0", "无合同额则 R2 无分母，无法判定"],
    ], [2.4, 6.0, 8.2])

    h2(doc, "2.2 排除条件")
    add_table(doc, ["类别", "具体情形", "处理", "演示条数"], [
        ["范围排除", "框架合同（T01）", "不进特征表，不出线索", "2"],
        ["范围排除", "外币合同（非 CNY）", "不进特征表，不出线索", "1"],
        ["版本淘汰", "11 表同一合同多条记录", "按「更新时间」取最新一条；状态=3 直接剔除", "3"],
        ["数据质量", "司库有合同、非司库无报送记录", "R1 不判定，落 DQ1 清单", "1"],
        ["数据质量", "非司库有记录、司库无对应合同", "不参与任一规则，落 DQ2 清单", "1"],
    ], [2.2, 6.4, 6.4, 1.6])

    h2(doc, "2.3 与在册模型的边界")
    para(doc, "在册「合同金额异常风险评估模型（合同金额异常-CL0038）」已含「结算超合同情形」，与本模型 R2 口径重叠。"
              "区别在于：CL0038 是「合同金额与工商数据比对」的综合评估；本模型 R2 只做「已结算 vs 合同额」的单点校验，"
              "但纳入了报送口径的框架过滤与快照对齐。对客说明时建议标注「R2 与 CL0038 互为交叉校验」，"
              "不宣称发现新风险类型。本模型为作业自建，不在已建模型名录内，不编 CL 编号。")

    # ==================== 三 ====================
    h1(doc, "三、输入数据与数据标准")
    para(doc, "两张表共 30 个字段，字段名与类型全部照报送规范原文，未自定义任何未登记字段名。"
              "判定实际只用到 6 个（两个合同编号、两个金额、合同额、已结算额），其余为过滤与对齐字段。")

    h2(doc, "3.1 司库第 16 张 合同表（19 个字段）")
    add_table(doc, ["序号", "数据标记", "数据名称", "类型", "必填性", "在本模型中的用途"], [
        ["1", "COMCODE", "所属集团编码", "文本(18)", "必填", "集团归属"],
        ["2", "COMNAME", "所属集团名称", "文本(100)", "必填", "集团名称"],
        ["3", "ACNTMCODE", "本方单位编码", "文本(18)", "必填", "输出「成员单位」来源"],
        ["4", "ACNTMNAME", "本方单位名称", "文本(128)", "必填", "输出「成员单位」来源"],
        ["5", "OPPOSITE_COMPNAY", "对方单位名称", "文本(512)", "必填", "线索展示"],
        ["6", "OPPOSITE_COMPNAYCODE", "对方单位税号", "文本(18)", "必填", "对方识别"],
        ["7", "CONTRACT_CODE", "合同编号", "文本(128)", "必填", "聚合主键 k，与 11 表对齐"],
        ["8", "CONTRACT_NAME", "合同名称", "文本(512)", "必填", "线索展示"],
        ["9", "CONTRACT_TYPE", "合同类型", "文本(3)", "必填", "N01 采购 / N02 销售 / N03 服务 / N04 工程"],
        ["10", "CONTRACT_SUBJECT", "合同标的物", "文本(512)", "必填", "标的识别"],
        ["11", "BILLING_TYPE", "结算类型", "文本(3)", "必填", "N01 收款 / N02 付款 / N03 无收无支"],
        ["12", "SIGN_DATE", "签订日期", "日期(10)", "必填", "多版本取最新依据"],
        ["13", "START_DATE", "起始日期", "日期(10)", "必填", "合同生效起算"],
        ["14", "END_DATE", "终止日期", "日期(10)", "必填", "无限期填 9999-12-31"],
        ["15", "FRAMEWORK_CONTRACT", "是否框架合同", "文本(3)", "必填", "过滤：仅取 T00"],
        ["16", "CONTRACT_AMOUNT", "合同总金额", "数值(18,2)", "必填", "R2 分母 C1"],
        ["17", "SETTLEMENT_AMOUNT", "已结算金额", "数值(18,2)", "必填", "R2 分子 E1"],
        ["18", "CURRENCTCODE", "币种", "文本(3)", "必填", "过滤：仅取 CNY"],
        ["19", "DAYIDX", "数据日期", "日期(10)", "必填", "快照日对齐"],
    ], [1.0, 4.4, 3.2, 2.0, 1.6, 5.0])

    h2(doc, "3.2 非司库第 11 张 发票及支付信息表（11 个字段）")
    add_table(doc, ["序号", "字段名称", "类型", "必填性", "业务说明", "在本模型中的用途"], [
        ["1", "所属集团统一社会信用代码", "文本(18)", "必填", "集团统一社会信用代码", "集团归属"],
        ["2", "所属集团名称", "文本", "必填", "集团名称", "集团名称"],
        ["3", "更新时间", "时间", "必填", "记录最后更新时间", "多版本取最新的依据"],
        ["4", "状态", "文本", "必填", "1 新增 / 2 修改 / 3 删除", "剔除状态=3"],
        ["5", "合同编号", "文本", "必填", "合同唯一编号", "聚合主键 k"],
        ["6", "合同名称", "文本", "非必填", "合同名称", "线索展示"],
        ["7", "累计发票金额", "数值（元）", "必填", "发票合计金额，保留两位小数", "R1 一侧 B2"],
        ["8", "累计支付金额", "数值（元）", "必填", "支付合计金额，保留两位小数", "R1 另一侧 S2"],
        ["9", "采购人名称", "文本", "非必填", "采购企业名称", "说明该表为采购口径"],
        ["10", "采购人代码", "文本(18)", "必填", "采购企业统一社会信用代码", "采购主体识别"],
        ["11", "数据日期", "文本", "必填", "数据日期 YYYY-MM-DD", "快照日对齐"],
    ], [0.9, 4.4, 2.0, 1.4, 4.8, 3.4])

    h2(doc, "3.3 取数与口径")
    add_table(doc, ["项", "取值", "说明"], [
        ["观察期", "2026-01-01 ~ 2026-06-30", "两条规则共用同一观察期"],
        ["快照日 T", "2026-06-30", "11 表按「数据日期」取 T 日；16 表按 DAYIDX"],
        ["框架合同", "FRAMEWORK_CONTRACT = 'T00'", "不排除会让 R2 出现大量 0 分母与失真值"],
        ["币种", "仅取 CURRENCTCODE = 'CNY'", "外币需折算，折算口径未定"],
        ["金额单位", "统一到元", "司库为币种最小单位保留两位；非司库为元，两边数值同为元"],
        ["红冲作废", "11 表累计发票金额须为剔除红冲后的净额", "规范未明确该列是否含红冲，须填报方确认（P4）"],
        ["多版本", "状态=3 剔除；同合同取「更新时间」最新一条", "11 表以状态位表达增删改"],
    ], [2.0, 5.6, 9.0])

    h2(doc, "3.4 未纳入的表及理由")
    add_table(doc, ["表", "为什么本期不用", "什么条件下再启用"], [
        ["18 应付款项", "R1 已改为「票 vs 款」两方比对，不再需要应付口径；且其合同编号为「非必填、弱校验」"
                       "（仅科目为合同负债时必填），归集不可靠", "需查「应付挂账 / 超额开票」，且合同编号填充率 ≥ 80%"],
        ["17 应收款项", "结构与 18 完全对称，但两条规则都在付款侧；11 表本身是采购口径"
                       "（带采购人名称 / 代码），没有销售侧的累计表", "新增销售合同规则时按 BILLING_TYPE 分支"],
        ["15 增值税发票", "无合同编号，无法归集到合同；只能按对方单位 + 金额 + 时间窗软匹配", "补合同编号或发票号↔合同映射"],
        ["2 资金结算", "无合同编号。交易流水号是笔级主键，详见下段", "需要逐笔取证或去重时"],
        ["5 / 6 票据", "与资金结算重复计入；「仅计已结清」口径未定", "P5 定案且明确去重规则后"],
        ["IT 四表", "仅 R2 使用，本期不交付", "R2 立项时"],
    ], [2.6, 8.0, 6.0])
    h2(doc, "3.5 关于「交易流水号」")
    para(doc, "它是 2 资金结算表的 BANKNUMBER，定义为「此笔交易在业务系统的唯一编码」，"
              "粒度是一笔。整张表 18 个字段没有任何合同编号，要归集到合同只能靠「对方单位 + 金额 + 时间窗」软匹配，"
              "匹配误差会直接变成假线索。因此它在本模型中只作证据键（去重、线索下钻、银行对账），不进判定式。",
         "宋体", 9, False, NOTE, before=2)

    # ==================== 四 ====================
    h1(doc, "四、特征清单与计算口径")
    para(doc, "特征全部为规则类计算量，不涉及模型训练与权重。特征名是业务口径名称，不是新增数据表字段。")
    add_table(doc, ["编号", "特征名", "计算口径", "来源", "用于"], [
        ["T01", "合同金额 C1", "16.CONTRACT_AMOUNT", "16 表", "R2 分母"],
        ["T02", "已结算金额 E1", "16.SETTLEMENT_AMOUNT", "16 表", "R2 分子"],
        ["T03", "累计发票金额 B2", "11.累计发票金额（T 日快照，状态≠3，取最新版本）", "11 表", "R1 一侧"],
        ["T04", "累计支付金额 S2", "11.累计支付金额（同 T03 口径）", "11 表", "R1 另一侧"],
        ["T05", "超合同额 D4", "E1 − C1（保留正负号）", "派生", "R2 判定"],
        ["T06", "超合同率", "D4 ÷ C1", "派生", "R2 判定与分级"],
        ["T07", "票款差额 D1", "ABS(B2 − S2)", "派生", "R1 判定"],
        ["T08", "票款差异率", "D1 ÷ MAX(B2, S2)", "派生", "R1 判定与分级"],
        ["T09", "差异方向", "S2 > B2 为「付款大于开票」，否则「开票大于付款」", "派生", "线索解释"],
        ["T10", "是否有非司库记录", "11 表是否存在该合同编号", "派生", "数据质量 DQ1"],
        ["T11", "报送版本数", "该合同在 11 表的记录条数（去重前）", "11 表", "取数取舍说明"],
        ["T12", "风险评分", "MIN(100, 差异率 ÷ 25% × 100)，保留一位小数", "派生", "排序"],
    ], [1.0, 3.4, 6.6, 1.8, 3.6])

    h2(doc, "4.1 计算示例")
    para(doc, "以 CG-202613 为例（R1）：11 表该合同累计发票 0.00 元、累计支付 450,000.00 元，"
              "T07 = 450,000.00，T08 = 450,000 ÷ 450,000 = 100%，T09 判为「付款大于开票」，"
              "T12 = MIN(100, 100% ÷ 25% × 100) = 100.0，定高级。")
    para(doc, "以 CG-202628 为例（双命中）：合同额 2,000,000.00 元、已结算 2,200,000.00 元，"
              "T05 = 200,000.00，T06 = 10.00%，命中 R2；同一合同 B2 = 1,800,000.00、S2 = 2,000,000.00，"
              "T07 = 200,000.00，T08 = 10.00%，命中 R1。两条线索独立成行，共用同一合同编号。")

    # ==================== 五 ====================
    h1(doc, "五、策略规则")
    h2(doc, "5.1 总览：两条规则，各读一张表")
    formula(doc, [
        "R1 票款背离（只读 11 表）",
        "  B2(k) = 11.累计发票金额   WHERE 合同编号 = k AND 状态 ≠ 3 AND 数据日期 = T",
        "  S2(k) = 11.累计支付金额   WHERE 同一条件",
        "  D1(k) = ABS( B2(k) − S2(k) )        // 无方向的绝对差",
        "",
        "R2 累计结算超合同额（只读 16 表）",
        "  C1(k) = 16.CONTRACT_AMOUNT    WHERE CONTRACT_CODE = k AND 框架 = T00 AND 币种 = CNY",
        "  E1(k) = 16.SETTLEMENT_AMOUNT  WHERE 同一条件",
        "  D4(k) = E1(k) − C1(k)               // 有方向：负数表示尚未结算完，正常",
    ])
    para(doc, "两张表不需要 join：R1 输出「合同编号 + 差额」，R2 输出「合同编号 + 超额」，"
              "在下游按合同编号并排展示即可。这也是本版能落到「两条独立 SQL」的原因。",
         "宋体", 9, False, GREY, before=0)

    h2(doc, "5.2 R1 票款背离")
    formula(doc, ["命中R1 = D1(k) > 10,000 元   且   D1(k) ÷ MAX(B2(k), S2(k)) > 5%"])
    add_table(doc, ["差额方向", "业务含义", "常见成因", "排查方向"], [
        ["B2 < S2\n付款大于开票", "钱付出去了，票没跟上", "已付款未取得发票、预付款、供应商延迟开票",
         "查开票申请与实际到票，核对合同付款条件"],
        ["B2 > S2\n开票大于付款", "票开了，钱没付或没付完", "先票后款（账期内正常）、预开票、红冲未剔除",
         "先排除账期，再看是否长期挂账"],
    ], [3.0, 4.4, 5.4, 3.8])
    para(doc, "算例：CG-202611，B2 = 3,000,000.00，S2 = 3,500,000.00，D1 = 500,000.00 > 1 万，"
              "比率 500,000 ÷ 3,500,000 = 14.29% > 5%，命中，评分 57.1，定中级。")

    h2(doc, "5.3 R2 累计结算超合同额")
    formula(doc, [
        "命中R2 = D4(k) > 10,000 元   且   D4(k) ÷ C1(k) > 5%",
        "说明：E1 < C1 说明合同尚未结算完，属正常状态，不产生线索。",
    ])
    para(doc, "算例：CG-202608，C1 = 1,200,000.00，E1 = 2,400,000.00，D4 = 1,200,000.00，"
              "比率 100%，评分 100.0，定高级。")

    h2(doc, "5.4 可执行 SQL")
    para(doc, "SQL 负责清洗与特征（sql/01_dwd_clean.sql、02_feature.sql），Python 负责规则计算与输出"
              "（code/run_model.py）。以下两段为规则预览，可独立运行复核。",
         "宋体", 9.5, False, GREY)
    formula(doc, [
        "-- R1（非司库单表，无 join）",
        "SELECT 合同编号, ABS(b2 - s2) AS d1,",
        "       ROUND(ABS(b2 - s2) * 1.0 / MAX(b2, s2), 6) AS ratio",
        "FROM   dws_contract_feature",
        "WHERE  b2 IS NOT NULL AND s2 IS NOT NULL AND MAX(b2, s2) > 0",
        "  AND  ABS(b2 - s2) > 10000",
        "  AND  ABS(b2 - s2) * 1.0 / MAX(b2, s2) > 0.05;",
        "",
        "-- R2（司库合同表单表）",
        "SELECT contract_code, e1 - c1 AS d4,",
        "       ROUND((e1 - c1) * 1.0 / c1, 6) AS ratio",
        "FROM   dws_contract_feature",
        "WHERE  c1 > 0 AND (e1 - c1) > 10000",
        "  AND  (e1 - c1) * 1.0 / c1 > 0.05;",
    ])

    # ==================== 六 ====================
    h1(doc, "六、阈值与参数配置")
    para(doc, "共 6 个参数。前两个是真正影响判定结果的阈值，可由业务方调整；其余为取数口径，"
              "改动会改变数据含义，只允许平台侧维护。")
    add_table(doc, ["parameter_code", "默认值", "可调范围", "作用", "可编辑方"], [
        ["tolerance_amount", "10,000.00 元", "0 ~ 200,000", "绝对门槛，挡尾差与四舍五入", "业务管理员"],
        ["tolerance_ratio", "5.00%", "0 ~ 30%", "相对门槛，挡大合同的小比例异常", "业务管理员"],
        ["snapshot_date", "2026-06-30", "固定（按报送周期）", "快照日，两表对齐基准", "平台侧"],
        ["framework_filter", "T00", "固定", "框架过滤，非 T00 不进判定", "平台侧"],
        ["currency_filter", "CNY", "固定", "币种过滤，外币本期不折算", "平台侧"],
        ["score_scale", "25%", "固定", "评分归一化基准：差异率 25% 得满分", "平台侧"],
    ], [3.4, 2.6, 3.2, 5.4, 2.0])

    h2(doc, "6.1 阈值取值的依据")
    para(doc, "两条规则的阈值不是从政策条文推出来的，规范只给字段定义、不给判定值。本版取值依据分三类：")
    bullets(doc, [
        "绝对门槛 1 万元：覆盖万元级合同的分位尾差。低于 1 万的差额在 84 份演示合同中占 12%，全部为四舍五入或分期尾差。",
        "相对门槛 5%：覆盖「先票后款」的正常账期。企业通常先收票后付款，B2 天然大于 S2；"
        "门槛若取 0.5%，几乎所有正常合同都会被判为背离。",
        "评分基准 25%：使差异率 25% 恰好得 100 分，与「≥80 分定高」形成稳定对应，避免评分与等级互相矛盾。",
    ])
    para(doc, "这两个阈值最终须由 POC 用真实核查结论回算。演示值只保证在构造样本上的判定正确性，"
              "不代表真实数据上的误报率。", "宋体", 9, False, NOTE, before=2)

    h2(doc, "6.2 为什么必须是双门槛")
    para(doc, "单门槛一定会错。只用绝对额，一份 1 亿元合同差 30 万会被报（占 0.3%，属正常波动）；"
              "只用相对率，一份 100 万元合同差 4.9 万会被放过（占 4.9%，但金额不小）。"
              "演示数据里这两类样本都留了对照：CG-202606（差 300 万，占 3.61%，不命中）、"
              "CG-202619（差 1.2 万，占 0.6%，不命中）、CG-202621（差 11 万，占 5.21%，命中）。")

    # ==================== 七 ====================
    h1(doc, "七、输出对象（对齐 V2.1）")
    para(doc, "输出按《御风穿透式监管平台·预置模型输入输出数据规范》第四章组织为 6 类对象："
              "Candidate 为唯一模型结果，Reason / Fact / Evidence 用于解释结果，"
              "另有原因-证据关联表与执行元数据。不再是单张宽表。")
    add_table(doc, ["对象", "基数", "本模型取值"], [
        ["Risk Candidate", "一条线索一个", "15 个，按「规则编号 + 合同编号」唯一"],
        ["Reason", "一个 Candidate 1–2 条", "15 条，按命中方向分支"],
        ["Business Fact", "一个 Candidate 4–5 条", "85 条，含金额、比率、快照日"],
        ["Evidence Set", "一个 Candidate ≥1 组", "25 组（R1 两组、R2 一组）"],
        ["Evidence Detail", "一组证据 ≥1 行", "27 行，按源记录一行一条"],
        ["explanation_evidence_rel", "Reason × Evidence", "25 条，关系均为 SUPPORT"],
    ], [3.4, 4.2, 9.0])

    h2(doc, "7.1 Risk Candidate 字段（20 个）")
    add_table(doc, ["字段", "类型", "必填", "含义与本模型取值"], [
        ["candidate_id", "string(64)", "是", "候选唯一 ID，格式 C + 快照日 + 4 位序号，如 C202606300001"],
        ["execution_id", "string(64)", "是", "来源模型执行，EXEC2026063001"],
        ["model_code", "string(64)", "是", "CONTRACT_LEDGER_PAYMENT_CHECK"],
        ["model_version", "string(32)", "是", "1.0.0"],
        ["model_config_version", "string(32)", "是", "MC-002"],
        ["subject_type", "enum", "是", "ENTERPRISE（风险主体为企业）"],
        ["subject_id", "string(128)", "是", "合同承办单位的统一社会信用代码"],
        ["subject_name", "string(256)", "否", "合同承办单位名称"],
        ["risk_type", "string(64)", "是", "PAYMENT_INVOICE_MISMATCH / CONTRACT_OVER_SETTLEMENT"],
        ["risk_level", "enum", "是", "HIGH / MEDIUM / LOW"],
        ["risk_level_description", "string(2000)", "是", "写明分级规则、本线索差异率与涉及金额"],
        ["risk_score", "decimal", "否", "0–100 排序分，差异率对 25% 归一化"],
        ["risk_amount", "decimal", "否", "差异金额，单位元，两位小数"],
        ["confidence", "decimal", "否", "本模型不输出置信度，留空"],
        ["risk_summary", "string(1000)", "是", "一句话摘要，含合同编号与关键金额"],
        ["investigation_advice", "string(2000)", "否", "按命中规则给出的核查方向"],
        ["observation_start", "datetime", "是", "2026-01-01T00:00:00+08:00"],
        ["observation_end", "datetime", "是", "2026-07-01T00:00:00+08:00（左闭右开）"],
        ["detected_at", "datetime", "是", "2026-06-30T18:00:00+08:00"],
        ["ext", "json", "否", "规则编号、合同编号、差异方向、业务对象类型"],
    ], [3.6, 2.6, 1.2, 9.2])

    h2(doc, "7.2 Reason 与原因-证据关联")
    add_table(doc, ["Reason 字段", "类型", "必填", "本模型取值"], [
        ["reason_id", "string(64)", "是", "R-C202606300001-01"],
        ["candidate_id", "string(64)", "是", "所属候选"],
        ["reason_code", "string(64)", "是", "INVOICE_LESS_THAN_PAYMENT / INVOICE_MORE_THAN_PAYMENT / "
                                          "SETTLEMENT_OVER_CONTRACT"],
        ["reason_title", "string(256)", "是", "关联发票合计小于累计支付 / 大于 / 已结算超过合同额"],
        ["reason_description", "string(2000)", "是", "业务化表述，不暴露公式与权重"],
        ["importance", "enum", "否", "与 Candidate 的风险等级一致"],
        ["investigation_advice", "string(2000)", "否", "针对该原因的排查建议"],
        ["display_order", "integer", "否", "1"],
        ["ext", "json", "否", "所属规则编号"],
    ], [4.0, 2.6, 1.2, 8.8])
    add_table(doc, ["关联表字段", "取值", "说明"], [
        ["reason_id / evidence_id", "Reason 与 Evidence Set 各一条", "联合主键"],
        ["relation_type", "SUPPORT", "证据支持该原因成立"],
        ["relation_strength", "1.0", "无权重体系，统一取 1"],
        ["created_at / source_execution_id", "2026-06-30T18:00:00+08:00 / EXEC2026063001", "生成时间与执行归属"],
    ], [5.4, 5.6, 5.6])

    h2(doc, "7.3 Business Fact 字段（18 个）")
    para(doc, "Fact 采用「值类型 + 五个值列」的互斥结构：一条 Fact 只有一个值列非空，且与 value_type 匹配。"
              "金额用 decimal 存储、展示两位小数；比例以 0–1 小数存储，单位标 ratio。")
    add_table(doc, ["字段", "类型", "必填", "本模型取值"], [
        ["fact_id", "string(64)", "是", "F-C202606300001-01"],
        ["candidate_id", "string(64)", "是", "所属候选"],
        ["fact_code", "string(64)", "是", "CONTRACT_AMOUNT / SETTLEMENT_AMOUNT / INVOICE_AMOUNT_CUM / "
                                        "PAYMENT_AMOUNT_CUM / INVOICE_GAP_AMOUNT / INVOICE_GAP_RATIO / "
                                        "OVER_CONTRACT_AMOUNT / OVER_CONTRACT_RATIO / SNAPSHOT_DATE"],
        ["fact_name", "string(256)", "是", "与 fact_code 一一对应的中文名"],
        ["value_type", "enum", "是", "NUMBER / DATE（本模型只用这两类）"],
        ["numeric_value", "decimal(38,10)", "条件", "value_type = NUMBER 时填写"],
        ["string_value / boolean_value / json_value", "—", "条件", "本模型不使用，留空"],
        ["date_value", "date", "条件", "value_type = DATE 时填写，ISO 格式"],
        ["unit", "string(32)", "条件", "金额为 CNY，比率为 ratio"],
        ["comparison_type", "enum", "否", "门槛类事实用 GT，其余 NONE"],
        ["benchmark_value", "typed scalar", "否", "阈值：10,000.00 元 或 0.05"],
        ["period_start / period_end", "datetime", "否", "观察期起止，与 Candidate 一致"],
        ["display_text", "string(1000)", "是", "展示文本，如「3,500,000.00 元」"],
        ["source_execution_id", "string(64)", "否", "所属执行"],
        ["ext", "json", "否", "本模型留空"],
    ], [5.0, 2.4, 1.2, 8.0])

    h2(doc, "7.4 Evidence Set 与 Evidence Detail")
    add_table(doc, ["Evidence Set 字段", "类型", "必填", "本模型取值"], [
        ["evidence_id", "string(64)", "是", "EV-C202606300001-01"],
        ["candidate_id", "string(64)", "是", "所属候选"],
        ["evidence_type", "enum", "是", "DOCUMENT_DETAIL（11 表）/ CONTRACT_DETAIL（16 表）"],
        ["evidence_name", "string(256)", "是", "合同 CG-202611 发票及支付信息报送记录"],
        ["detail_dataset_ref", "string(512)", "是", "out/evidence/EV-xxx.csv，可解析回明细"],
        ["record_count", "bigint", "是", "该组明细行数"],
        ["source_execution_id", "string(64)", "否", "所属执行"],
        ["generated_at", "datetime", "是", "生成时间"],
        ["ext", "json", "否", "schema_code"],
    ], [4.6, 2.4, 1.2, 8.4])
    add_table(doc, ["Detail 字段", "取值说明"], [
        ["evidence_id + detail_id", "联合主键，组内从 D-001 递增"],
        ["source_data_context_id", "DC-SIKU-16 / DC-NONS-11，指回取数上下文"],
        ["source_record_id", "16 表用 CONTRACT_CODE；11 表用「合同编号|更新时间」复合主键"],
        ["业务字段", "CONTRACT_DETAIL 保留合同编号、名称、双方、合同额、已结算、币种、签订日期；"
                    "DOCUMENT_DETAIL 保留合同编号、累计发票、累计支付、数据日期、状态、更新时间"],
        ["ext", "schema_code，标明该行属于哪种明细结构"],
    ], [4.6, 12.0])

    h2(doc, "7.5 执行元数据与本轮产出")
    para(doc, "执行记录见 out/execution.csv，本模型实际填充的字段为：execution_id、model_code、model_version、"
              "contract_version、model_config_version、观察期起止、started_at / finished_at、execution_status、"
              "candidate_count、engine_version、triggered_by、request_key、runtime_snapshot_ref、"
              "result_manifest_ref。ext 与 replay_of_execution_id 留空。")
    add_table(doc, ["对象", "数量", "输出文件"], [
        ["Risk Candidate", "15", "out/candidate.csv"],
        ["Reason", "15", "out/reason.csv"],
        ["Business Fact", "85", "out/fact.csv"],
        ["Evidence Set", "25", "out/evidence_set.csv"],
        ["Evidence Detail", "27", "out/evidence_detail.csv"],
        ["原因-证据关联", "25", "out/explanation_evidence_rel.csv"],
        ["执行元数据", "1", "out/execution.csv"],
        ["明细数据集", "25 个文件", "out/evidence/EV-*.csv"],
    ], [4.0, 3.0, 9.6])
    para(doc, "out/线索汇总_人读视图.csv 由 Candidate 与 Fact 派生，仅供人工翻阅，不作为规范交付物。",
         "宋体", 9, False, GREY, before=2)

    h1(doc, "八、效果评估标准与验收口径")
    h2(doc, "8.1 指标定义")
    add_table(doc, ["指标", "口径", "演示集取值"], [
        ["覆盖率", "参与判定的合同数 ÷ 纳入范围的合同数", "84 / 84 = 100%"],
        ["命中率", "线索数 ÷ 参与判定合同数", "15 / 84 = 17.9%"],
        ["精确率", "设计为真风险的命中数 ÷ 全部命中数", "15 / 15 = 100%"],
        ["查全率", "命中真风险数 ÷ 应命中真风险数", "15 / 15 = 100%"],
        ["误报率", "核查后排除数 ÷ 命中数", "需真实核查结论，POC 后回填"],
        ["证据完备率", "至少 1 组 Evidence Set 的 Candidate 数 ÷ 全部 Candidate 数",
         "15 / 15 = 100%（25 组 / 27 行）"],
    ], [2.6, 8.4, 5.6])

    h2(doc, "8.2 演示集结果")
    add_table(doc, ["规则", "应命中", "实测命中", "误报", "漏报", "精确率", "查全率"], [
        ["R1 票款背离", "10", "10", "0", "0", "100%", "100%"],
        ["R2 累计结算超合同额", "5", "5", "0", "0", "100%", "100%"],
        ["合计", "15", "15", "0", "0", "100%", "100%"],
    ], [3.6, 1.8, 2.0, 1.6, 1.6, 1.8, 1.8])
    para(doc, "这个 100% 只说明「按设计构造的样本集」上的逻辑正确性，不代表真实数据表现。"
              "真实场景里，跨期开票、预开票、代扣手续费、汇率折算都会造成合理差额，"
              "误报率必须在 POC 阶段用业务核查结论回算。", "宋体", 9, False, NOTE, before=2)

    h2(doc, "8.3 验收口径")
    add_table(doc, ["验收项", "标准", "核对方式", "结果"], [
        ["全链路无报错", "gen → run_model → run_sql_check → build_report 四步全部退出码 0",
         "out/run_log.txt 与终端输出", "通过"],
        ["双实现一致", "SQL 特征表与 Python 特征表逐行一致；线索集合与差异金额完全相同",
         "code/run_sql_check.py", "通过"],
        ["评分与等级一致", "等级由评分切分得到，不存在「61 分却高风险」的矛盾展示",
         "run_model.py 自检项 1", "通过"],
        ["双门槛生效", "不存在仅过单一门槛的线索", "run_model.py 自检项 2", "通过"],
        ["证据可追溯", "每条 Candidate ≥1 组 Evidence Set，明细可经 detail_dataset_ref 解析",
         "out/evidence_set.csv 与 out/evidence/", "通过"],
        ["输出契约合规", "Candidate / Reason / Fact / Evidence 字段与 V2.1 第四章逐项对齐",
         "out/candidate.csv 等 6 张表", "通过"],
        ["字段合规", "30 个输入字段全部对齐报送规范，未自定义字段名", "第七章与 data/字段字典.csv", "通过"],
        ["阈值可调", "改动参数无需改代码即可重算", "report/任务报告.html 阈值实验室", "通过"],
        ["异常样本覆盖", "28 条设计样本全部符合预期，含 8 条不命中对照", "文档第九章与报告第五节", "通过"],
    ], [3.0, 6.2, 4.6, 1.4])

    # ==================== 九 ====================
    h1(doc, "九、风险与局限")
    add_table(doc, ["级别", "事项", "影响", "处置"], [
        ["P0", "R1 只看票与款是否同步，看不到「应付」这一会计确认口径",
         "应付已确认 500 万而票、款都只走 300 万时，D1 = 0，本模型不报", "明确写入「不覆盖」；如需覆盖，接 18 表回到三方比对"],
        ["P0", "11 表累计发票金额是否已剔除红冲未确认（P4）",
         "若含红冲，B2 虚高，会把红冲误判成「多开票」", "与填报方确认口径；必要时按发票号关联蓝红票单独校准"],
        ["P1", "两表快照日可能不同天（P7）",
         "时点不一致会让差异系统性偏移", "强制按 T 日对齐；不一致时先对齐再比"],
        ["P1", "评分只由差异率决定，未纳入金额规模",
         "两份差异率相同但金额相差十倍的线索评分相同，排序不够精细", "下游可按「等级 + 差异金额」二级排序"],
        ["P2", "未覆盖框架分组合计（R3）与逐笔票款核对",
         "框架协议下多份子合同合计超预算的情况看不到", "作为后续扩展"],
        ["P2", "未做发票标的与合同标的一致性校验",
         "可能漏掉「票是真的、但不是这份合同的」", "接入发票商品明细后扩展"],
    ], [1.2, 5.4, 5.2, 4.8])

    # ==================== 十 ====================
    h1(doc, "十、复现步骤与运行记录")
    h2(doc, "10.1 运行环境与步骤")
    para(doc, "Python 3.9+，零第三方依赖（标准库 csv / json / decimal / sqlite3）。"
              "SQLite 版 SQL 由标准库 sqlite3 直接执行，无需安装数据库。IDE 用 PyCharm 或 Jupyter 均可，"
              "四个脚本按顺序执行。")
    add_table(doc, ["步骤", "命令", "应核对的结果"], [
        ["1", "python3 code/gen_mock_data.py", "生成 2 张表 CSV：16 表 87 行、11 表 90 行；生成 data/params.json"],
        ["2", "python3 code/run_model.py",
         "输出 15 个 Candidate、85 条 Fact、25 组证据；6 项自检全部 PASS"],
        ["3", "python3 code/run_sql_check.py", "SQL 与 Python 特征表逐行一致、线索集合完全一致，结论「通过」"],
        ["4", "python3 code/build_report.py", "生成 report/任务报告.html，双击即可演示"],
    ], [1.2, 6.6, 8.6])

    h2(doc, "10.2 运行记录（节选）")
    formula(doc, [
        "[1] 读入原始数据：16 合同表 87 行，11 发票及支付表 90 行",
        "[2] 清洗后：合同表 84 份（剔除 3 条），发票表 87 份（剔除 3 条）",
        "[3] 特征表：84 行（一份合同一行），数据质量与排除记录 8 条",
        "[4] 规则计算：命中 15 条线索（R1 10 条，R2 5 条）",
        "[5] 自检",
        "    [PASS] 评分与等级一致    [PASS] 双门槛同时满足    [PASS] V2.1 证据对象完整",
        "    [PASS] 差异方向已标注    [PASS] 线索唯一          [PASS] 无零差异线索",
        "[6] 汇总：84 份合同参与判定，命中 15 条线索（高 4 / 中 7 / 低 4），命中率 17.9%，证据明细 27 行",
        "",
        "✓ V2.1 输出：Candidate 15 ｜ Reason 15 ｜ Fact 85 ｜ Evidence Set 25 ｜ Detail 27 ｜ 关联 25",
        "",
        "SQL 与 Python 双实现一致性校验",
        "    ✓ 特征表逐行一致（c1 / e1 / b2 / s2 全部相同）",
        "    ✓ 线索集合 15 条，SQL 独有 0 条、Python 独有 0 条、金额不一致 0 条",
        "    结论：通过 —— SQL 与 Python 的线索集合、差异金额完全一致",
    ])
    para(doc, "完整日志见 out/run_log.txt 与 out/sql_check_report.txt；规范输出为 out/candidate.csv 等 6 张表与 out/evidence/ 明细，"
              "人读视图为 out/线索汇总_人读视图.csv，排除清单为 out/excluded_list.csv。",
         "宋体", 9, False, GREY, before=2)

    # 表格前的段落设「与下段同页」，避免标题与引语留在页底、表格跳到下一页
    body = doc.element.body
    for tbl in body.iterchildren(qn("w:tbl")):
        prev = tbl.getprevious()
        if prev is not None and prev.tag == qn("w:p"):
            pPr = prev.get_or_add_pPr()
            for e in pPr.findall(qn("w:keepNext")):
                pPr.remove(e)
            pPr.append(OxmlElement("w:keepNext"))

    doc.save(DEST)
    print(f"✓ 已生成：{DEST}")
    print(f"  段落 {len(doc.paragraphs)} ｜ 表格 {len(doc.tables)} ｜ "
          f"{os.path.getsize(DEST) / 1024:.0f} KB")


if __name__ == "__main__":
    main()
