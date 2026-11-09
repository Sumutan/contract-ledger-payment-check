-- =============================================================================
-- 02_feature.sql ｜ 特征计算层（SQLite / 本地可执行版）
-- 与 sql/hive/02_feature.sql 逻辑一一对应
-- =============================================================================

DROP TABLE IF EXISTS dws_contract_feature;
CREATE TABLE dws_contract_feature AS
SELECT
    c.CONTRACT_CODE       AS contract_code,
    c.CONTRACT_NAME       AS contract_name,
    c.ACNTMCODE           AS member_code,
    c.ACNTMNAME           AS member_name,
    c.OPPOSITE_COMPNAY    AS vendor_name,
    c.CONTRACT_TYPE       AS contract_type,
    c.BILLING_TYPE        AS billing_type,
    c.CONTRACT_AMOUNT     AS c1,
    c.SETTLEMENT_AMOUNT   AS e1,
    i."累计发票金额"       AS b2,
    i."累计支付金额"       AS s2,
    CASE WHEN i."合同编号" IS NULL THEN 0 ELSE 1 END AS has_nonsiku
FROM dwd_contract_clean c
LEFT JOIN dwd_invoice_latest i
       ON c.CONTRACT_CODE = i."合同编号";

DROP TABLE IF EXISTS dwq_contract_without_invoice;
CREATE TABLE dwq_contract_without_invoice AS
SELECT contract_code, contract_name, 'DQ1' AS dq_code,
       '司库合同表有该合同，非司库发票及支付信息表无对应记录' AS dq_reason
FROM dws_contract_feature
WHERE has_nonsiku = 0;

DROP TABLE IF EXISTS dwq_orphan_invoice;
CREATE TABLE dwq_orphan_invoice AS
SELECT i."合同编号" AS contract_code, i."合同名称" AS contract_name,
       'DQ2' AS dq_code,
       '非司库有该合同编号，司库合同表无对应合同' AS dq_reason
FROM dwd_invoice_latest i
LEFT JOIN ods_siku_contract c
       ON c.CONTRACT_CODE = i."合同编号"
WHERE c.CONTRACT_CODE IS NULL;
