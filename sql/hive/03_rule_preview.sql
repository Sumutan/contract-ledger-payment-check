-- =============================================================================
-- 03_rule_preview.sql ｜ 规则预览（Hive / 平台版）
-- 说明：正式交付中「规则计算与结果输出」由 Python（code/run_model.py）承担。
--       本文件用于平台侧快速预览与人工复算，输出与 Python 线索表一一对应。
-- 参数：容差 绝对 10,000.00 元 且 相对 5%
-- =============================================================================

-- -----------------------------------------------------------------------------
-- 3.1 R4 累计结算超合同额：只读 16 表特征列，方向性判定（只查超不查少）
-- -----------------------------------------------------------------------------
SELECT
    'R4'                                        AS rule_code,
    '累计结算超合同额'                            AS rule_name,
    contract_code,
    contract_name,
    member_name,
    c1                                          AS contract_amount,
    e1                                          AS settlement_amount,
    e1 - c1                                     AS diff_amount,
    ROUND((e1 - c1) / c1, 6)                     AS diff_ratio,
    '超合同额'                                   AS direction
FROM dws_contract_feature
WHERE c1 > 0
  AND (e1 - c1) > 10000
  AND (e1 - c1) / c1 > 0.05
ORDER BY diff_ratio DESC;


-- -----------------------------------------------------------------------------
-- 3.2 R1 票款背离：只读 11 表特征列，无方向差值 + 方向标注
-- -----------------------------------------------------------------------------
SELECT
    'R1'                                        AS rule_code,
    '票款背离'                                   AS rule_name,
    contract_code,
    contract_name,
    member_name,
    b2                                          AS invoice_amount,
    s2                                          AS payment_amount,
    ABS(b2 - s2)                                AS diff_amount,
    ROUND(ABS(b2 - s2) / GREATEST(b2, s2), 6)    AS diff_ratio,
    CASE WHEN s2 > b2 THEN '付款大于开票' ELSE '开票大于付款' END AS direction
FROM dws_contract_feature
WHERE b2 IS NOT NULL
  AND s2 IS NOT NULL
  AND GREATEST(b2, s2) > 0
  AND ABS(b2 - s2) > 10000
  AND ABS(b2 - s2) / GREATEST(b2, s2) > 0.05
ORDER BY diff_ratio DESC;
