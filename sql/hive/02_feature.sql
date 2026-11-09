-- =============================================================================
-- 02_feature.sql ｜ 特征计算层（Hive / 平台版）
-- 输出：dws_contract_feature —— 一份合同一行，供 R1 / R4 两条规则各取所需
-- 设计：两张源表在此左连接一次，此后 R1 与 R4 互不依赖，可分别调度
-- =============================================================================

DROP TABLE IF EXISTS dws_contract_feature;
CREATE TABLE dws_contract_feature STORED AS PARQUET AS
SELECT
    c.CONTRACT_CODE                      AS contract_code,
    c.CONTRACT_NAME                      AS contract_name,
    c.ACNTMCODE                          AS member_code,
    c.ACNTMNAME                          AS member_name,
    c.OPPOSITE_COMPNAY                   AS vendor_name,
    c.CONTRACT_TYPE                      AS contract_type,
    c.BILLING_TYPE                       AS billing_type,
    c.CONTRACT_AMOUNT                    AS c1,
    c.SETTLEMENT_AMOUNT                  AS e1,
    i.累计发票金额                        AS b2,
    i.累计支付金额                        AS s2,
    CASE WHEN i.合同编号 IS NULL THEN 0 ELSE 1 END AS has_nonsiku
FROM dwd_contract_clean c
LEFT JOIN dwd_invoice_latest i
       ON c.CONTRACT_CODE = i.`合同编号`;

-- 数据质量：司库有合同、非司库无报送记录（R1 无法判定）
DROP TABLE IF EXISTS dwq_contract_without_invoice;
CREATE TABLE dwq_contract_without_invoice STORED AS PARQUET AS
SELECT contract_code, contract_name, 'DQ1' AS dq_code,
       '司库合同表有该合同，非司库发票及支付信息表无对应记录' AS dq_reason
FROM dws_contract_feature
WHERE has_nonsiku = 0;

-- 数据质量：非司库有记录、司库无对应合同（无法参与任一规则）
DROP TABLE IF EXISTS dwq_orphan_invoice;
CREATE TABLE dwq_orphan_invoice STORED AS PARQUET AS
SELECT i.`合同编号` AS contract_code, i.`合同名称` AS contract_name,
       'DQ2' AS dq_code,
       '非司库有该合同编号，司库合同表无对应合同' AS dq_reason
FROM dwd_invoice_latest i
LEFT JOIN ods_siku_contract c
       ON c.CONTRACT_CODE = i.`合同编号`
WHERE c.CONTRACT_CODE IS NULL;
