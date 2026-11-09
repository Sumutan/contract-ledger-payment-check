-- =============================================================================
-- 03_rule_preview.sql ｜ 规则预览（SQLite / 本地可执行版）
-- 与 Python 规则引擎的线索集合应完全一致，由 code/run_sql_check.py 自动比对
-- 参数：容差 绝对 10,000.00 元 且 相对 5%（两门槛同时满足）
-- =============================================================================

DROP TABLE IF EXISTS rs_r4;
CREATE TABLE rs_r4 AS
SELECT
    'R4'                     AS rule_code,
    contract_code,
    c1                       AS contract_amount,
    e1                       AS settlement_amount,
    e1 - c1                  AS diff_amount,
    ROUND((e1 - c1) * 1.0 / c1, 6) AS diff_ratio,
    '超合同额'                AS direction
FROM dws_contract_feature
WHERE c1 > 0
  AND (e1 - c1) > 10000
  AND (e1 - c1) * 1.0 / c1 > 0.05;

DROP TABLE IF EXISTS rs_r1;
CREATE TABLE rs_r1 AS
SELECT
    'R1'                     AS rule_code,
    contract_code,
    b2                       AS invoice_amount,
    s2                       AS payment_amount,
    ABS(b2 - s2)             AS diff_amount,
    ROUND(ABS(b2 - s2) * 1.0 / MAX(b2, s2), 6) AS diff_ratio,
    CASE WHEN s2 > b2 THEN '付款大于开票' ELSE '开票大于付款' END AS direction
FROM dws_contract_feature
WHERE b2 IS NOT NULL
  AND s2 IS NOT NULL
  AND MAX(b2, s2) > 0
  AND ABS(b2 - s2) > 10000
  AND ABS(b2 - s2) * 1.0 / MAX(b2, s2) > 0.05;

-- 合并视图，供导出与比对
DROP VIEW IF EXISTS v_rule_result;
CREATE VIEW v_rule_result AS
SELECT rule_code, contract_code, diff_amount, diff_ratio, direction FROM rs_r4
UNION ALL
SELECT rule_code, contract_code, diff_amount, diff_ratio, direction FROM rs_r1;
