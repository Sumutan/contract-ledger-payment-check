-- =============================================================================
-- 01_dwd_clean.sql ｜ 数据清洗层（Hive / 平台版）
-- 模型：合同台账与票款执行一致性校验模型  v1.0.0
-- 输入：ods_siku_contract（司库第 16 张 合同表）
--       ods_nonsiku_invoice_pay（非司库第 11 张 发票及支付信息表）
-- 输出：dwd_contract_clean / dwd_invoice_latest
-- 参数：snapshot_date = '2026-06-30'（调用时以 --hivevar snapshot_date=... 传入）
-- =============================================================================

-- 参数占位，单独运行时把 ${snapshot_date} 替换为 '2026-06-30'
SET hivevar:snapshot_date = 2026-06-30;

-- -----------------------------------------------------------------------------
-- 1.1 合同表清洗
--     ① 快照日对齐          ② 仅非框架合同（合同额与已结算额仅对 T00 填写）
--     ③ 仅人民币            ④ 合同额 > 0（无金额则无判定基准）
-- -----------------------------------------------------------------------------
DROP TABLE IF EXISTS dwd_contract_clean;
CREATE TABLE dwd_contract_clean STORED AS PARQUET AS
SELECT
    CONTRACT_CODE,
    CONTRACT_NAME,
    ACNTMCODE,
    ACNTMNAME,
    OPPOSITE_COMPNAY,
    CONTRACT_TYPE,
    BILLING_TYPE,
    FRAMEWORK_CONTRACT,
    CAST(CONTRACT_AMOUNT   AS DECIMAL(18, 2)) AS CONTRACT_AMOUNT,
    CAST(SETTLEMENT_AMOUNT AS DECIMAL(18, 2)) AS SETTLEMENT_AMOUNT,
    CURRENCTCODE,
    DAYIDX
FROM ods_siku_contract
WHERE DAYIDX = '${hivevar:snapshot_date}'
  AND FRAMEWORK_CONTRACT = 'T00'
  AND CURRENCTCODE = 'CNY'
  AND CAST(CONTRACT_AMOUNT AS DECIMAL(18, 2)) > 0;

-- -----------------------------------------------------------------------------
-- 1.2 发票及支付信息表清洗
--     ① 快照日对齐  ② 剔除状态=3 的删除记录  ③ 同合同多版本按更新时间取最新
-- -----------------------------------------------------------------------------
DROP TABLE IF EXISTS dwd_invoice_latest;
CREATE TABLE dwd_invoice_latest STORED AS PARQUET AS
SELECT
    合同编号, 合同名称, 累计发票金额, 累计支付金额, 更新时间, 状态, 数据日期
FROM (
    SELECT
        `合同编号`,
        `合同名称`,
        CAST(`累计发票金额` AS DECIMAL(18, 2)) AS 累计发票金额,
        CAST(`累计支付金额` AS DECIMAL(18, 2)) AS 累计支付金额,
        `更新时间`,
        `状态`,
        `数据日期`,
        ROW_NUMBER() OVER (PARTITION BY `合同编号` ORDER BY `更新时间` DESC) AS rn
    FROM ods_nonsiku_invoice_pay
    WHERE `数据日期` = '${hivevar:snapshot_date}'
      AND `状态` <> '3'
) t
WHERE rn = 1;
