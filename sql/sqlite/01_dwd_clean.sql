-- =============================================================================
-- 01_dwd_clean.sql ｜ 数据清洗层（SQLite / 本地可执行版）
-- 使用方式：由 code/run_sql_check.py 建表并载入 CSV 后执行本文件
-- 与 sql/hive/01_dwd_clean.sql 逻辑一一对应，仅方言差异
-- 参数：snapshot_date = '2026-06-30'
-- =============================================================================

DROP TABLE IF EXISTS dwd_contract_clean;
CREATE TABLE dwd_contract_clean AS
SELECT
    CONTRACT_CODE,
    CONTRACT_NAME,
    ACNTMCODE,
    ACNTMNAME,
    OPPOSITE_COMPNAY,
    CONTRACT_TYPE,
    BILLING_TYPE,
    FRAMEWORK_CONTRACT,
    CAST(CONTRACT_AMOUNT   AS NUMERIC) AS CONTRACT_AMOUNT,
    CAST(SETTLEMENT_AMOUNT AS NUMERIC) AS SETTLEMENT_AMOUNT,
    CURRENCTCODE,
    DAYIDX
FROM ods_siku_contract
WHERE DAYIDX = '2026-06-30'
  AND FRAMEWORK_CONTRACT = 'T00'
  AND CURRENCTCODE = 'CNY'
  AND CAST(CONTRACT_AMOUNT AS NUMERIC) > 0;

DROP TABLE IF EXISTS dwd_invoice_latest;
CREATE TABLE dwd_invoice_latest AS
SELECT
    "合同编号", "合同名称", "累计发票金额", "累计支付金额", "更新时间", "状态", "数据日期"
FROM (
    SELECT
        "合同编号",
        "合同名称",
        CAST("累计发票金额" AS NUMERIC) AS "累计发票金额",
        CAST("累计支付金额" AS NUMERIC) AS "累计支付金额",
        "更新时间",
        "状态",
        "数据日期",
        ROW_NUMBER() OVER (PARTITION BY "合同编号" ORDER BY "更新时间" DESC) AS rn
    FROM ods_nonsiku_invoice_pay
    WHERE "数据日期" = '2026-06-30'
      AND "状态" <> '3'
) t
WHERE rn = 1;
