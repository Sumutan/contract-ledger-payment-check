# 合同台账与票款执行一致性校验模型

穿透式监管实战作业 · 作业自建模型 · 不编 CL 编号

| 项 | 内容 |
|---|---|
| 模型编码 | `CONTRACT_LEDGER_PAYMENT_CHECK` |
| 版本 | v1.0.0（2026-11） |
| 判定输入 | **2 张表**：司库第 16 张 合同表、非司库第 11 张 发票及支付信息表 |
| 覆盖规则 | **R1 票款背离**（只读 11 表）、**R2 累计结算超合同额**（只读 16 表） |
| 关键设计 | 两条规则各读一张表、互不 join，可分别开发、验收、上线 |

R2 台账割裂需企业 IT 四表，本期不交付；框架分组合计（R3）不在范围。

---

## 一、四步运行（Python 3.9+，零第三方依赖）

在项目根目录依次执行：

```bash
python3 code/gen_mock_data.py     # 1. 生成 mock 数据（2 张表 CSV + 参数文件）
python3 code/run_model.py         # 2. 规则计算，输出线索与证据，含 6 项自检
python3 code/run_sql_check.py     # 3. SQL 与 Python 双实现一致性校验
python3 code/build_report.py      # 4. 生成可交互 HTML 报告
```

四步全部退出码 0，无报错。IDE 用 PyCharm 或 Jupyter 均可，也可逐段在 Notebook 里跑。

## 二、目录说明

```
├── docs/
│   ├── 合同台账与票款执行一致性校验模型-模型说明.docx   交付文档（11 页 10 章）
│   ├── 合同台账与票款执行一致性校验模型-模型说明.pdf    PDF 版
│   └── build_doc.py                                 文档生成脚本（内容与数据同源）
├── data/
│   ├── raw/ods_siku_contract.csv                     16 合同表，87 行 × 19 列
│   ├── raw/ods_nonsiku_invoice_pay.csv               11 发票及支付表，90 行 × 11 列
│   ├── params.json                                   快照日与 6 个模型参数
│   └── 字段字典.csv                                   30 个字段的名称/类型/必填/说明
├── code/
│   ├── gen_mock_data.py                              mock 数据生成（28 个设计样本）
│   ├── run_model.py                                  规则引擎（Python 侧）
│   ├── run_sql_check.py                              SQL 与 Python 一致性校验
│   └── build_report.py                               HTML 报告生成
├── sql/
│   ├── hive/                                         平台版：01 清洗 / 02 特征 / 03 规则预览
│   └── sqlite/                                       本地可执行版，供 run_sql_check 调用
├── out/
│   ├── 输出总表.xlsx                                  四类对象合并总表（4 张子表，列同 V2.1）
│   ├── candidate.csv                                 Risk Candidate（15 个，V2.1 §4.2 全字段）
│   ├── reason.csv                                    Reason（15 条，§4.3）
│   ├── fact.csv                                      Business Fact（85 条，§4.4）
│   ├── evidence_set.csv                              Evidence Set（25 组，§4.5）
│   ├── evidence_detail.csv                           Evidence Detail（27 行，§4.5.2）
│   ├── explanation_evidence_rel.csv                  原因-证据关联（25 条，§4.6）
│   ├── execution.csv                                 执行元数据（§4.7）
│   ├── evidence/EV-*.csv                             按证据集拆分的可解析明细（25 个）
│   ├── manifest.json / runtime_snapshot.json         结果清单与运行快照
│   ├── 线索汇总_人读视图.csv                          由 Candidate + Fact 派生，非规范交付物
│   ├── excluded_list.csv                             数据质量与排除清单（8 条）
│   ├── feature_table.csv                             特征表（84 行，一份合同一行）
│   ├── run_log.txt                                   运行日志（含 6 项自检）
│   ├── sql_check_report.txt                          双实现一致性校验报告
│   └── clues.json                                    供报告生成使用的中间数据
├── report/任务报告.html                              单文件离线报告，双击即开
└── 演示/                                             投屏演示材料
    ├── 演示.html                                     HTML 幻灯片（17 页，方向键翻页，F 全屏）
    ├── 演示.ipynb                                    Jupyter 逐步演示（4 步 14 单元）
    └── 命令清单.md                                    一页纸命令与排错
```

## 三、职责划分

| 环节 | 承担者 | 产出 |
|---|---|---|
| 数据清洗 | SQL（`01_dwd_clean.sql`） | `dwd_contract_clean`、`dwd_invoice_latest` |
| 特征计算 | SQL（`02_feature.sql`） | `dws_contract_feature`（c1 / e1 / b2 / s2） |
| 规则计算与输出 | Python（`run_model.py`） | Candidate / Reason / Fact / Evidence 四类对象，排除清单 |
| 一致性校验 | Python（`run_sql_check.py`） | 特征表逐行比对 + 线索集合比对 |

Python 侧独立实现了一遍清洗与聚合，不用 SQL 的中间结果，这样两边对不上时能直接定位到具体合同。

## 四、运行结果

| 指标 | 取值 |
|---|---|
| 合同表原始 / 参与判定 | 87 份 / 84 份（框架 2、外币 1 被排除） |
| 发票表报送记录 / 去重后 | 90 条 / 87 份（状态=3 与旧版本共 3 条被剔除） |
| 线索 | 15 条（R1 10、R2 5），高 4 / 中 7 / 低 4 |
| 命中率 | 17.9% |
| 证据 | 25 组 Evidence Set / 27 行明细（每条 Candidate ≥1 组） |
| 数据质量与排除 | 8 条（DQ1 1、DQ2 1、DQ3 3、范围排除 3） |

三方一致：Python ≡ SQLite SQL ≡ 页面内前端复算，线索集合、等级、评分、差异金额逐项相同。

## 五、测试样本设计

共 28 个设计样本，含 9 条**不命中对照**——只放命中样本说明不了模型准不准。

| 类型 | 样本 | 要点 |
|---|---|---|
| 正常对照 | CG-202601 ~ 605 | 差额在容差内、未结算完、尾差 |
| 双门槛对照 | CG-202606 / 619 / 620 / 621 | 差 300 万但占 3.61%（不命中）；差 11 万占 5.21%（命中） |
| R2 命中 | CG-202607 ~ 610 | 超 10% / 100% / 25% / 6% |
| R1 命中 | CG-202611 ~ 617、621 | 款多、票多、无票、低档 |
| 排除 | KJ-2026-001 / 002、CG-202623 | 框架合同、外币合同 |
| 数据质量 | CG-202625、CG-2026-ORPHAN | 有合同无报送、有报送无合同 |
| 多版本 | CG-202627 | 3 条记录，须取最新有效版本 |
| 双命中 | CG-202628 | 同一合同同时触发 R1 与 R2 |

明细见 `report/任务报告.html` 第五节，或文档第九章。

## 六、口径与边界

**R1 只看票与款是否同步，看不到「应付」。** 应付已确认 500 万而票、款都只走 300 万时，差额为 0，本模型不报——这属于应付挂账，不是票款背离。要覆盖需接司库第 18 张应付款项，回到三方比对。

**11 表累计发票金额是否已剔除红冲未确认。** 规范未明确该列口径。若含红冲，B2 会虚高，产生「多开票」误报。

**框架合同全局排除。** 规范对 `CONTRACT_AMOUNT`、`SETTLEMENT_AMOUNT` 均注明「非框架合同填写」，不排除会让 R2 出现大量 0 分母。

**交易流水号不作判定键。** 它在司库第 2 张资金结算表，整张表 18 个字段没有合同编号，归集只能软匹配。本模型保留其作为证据键（去重、下钻、银行对账）。

## 七、已知限制

- Hive 版 SQL（`sql/hive/`）未在真实 Hive 环境验证，本机无集群；一致性校验针对 `sql/sqlite/` 做。
- 评分只由差异率决定，未纳入金额规模，同分不同额的线索需下游按金额二级排序。
- 阈值（1 万元 + 5%）为演示值，须 POC 用真实核查结论回标。
