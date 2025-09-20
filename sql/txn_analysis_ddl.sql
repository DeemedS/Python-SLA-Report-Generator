CREATE TABLE [dbo].[txn_analysis_time_of_day_txn_count] (
	TRANSACTION_DATE DATE NULL,
    WEEK_NUM     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [0-8]        INT NOT NULL DEFAULT 0,
    [9-12]       INT NOT NULL DEFAULT 0,
    [13-16]      INT NOT NULL DEFAULT 0,
    [17-20]      INT NOT NULL DEFAULT 0,
    [21-23]      INT NOT NULL DEFAULT 0,
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_time_of_day_txn_count_date_status
ON [dbo].[txn_analysis_time_of_day_txn_count] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idxtxn_analysis_time_of_day_txn_count_date
ON [dbo].[txn_analysis_time_of_day_txn_count] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_sla] (
	TRANSACTION_DATE DATE NULL,
    WEEK_NUM     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [SLA-<0] INT NOT NULL DEFAULT 0,
    [SLA-=0] INT NOT NULL DEFAULT 0,
    [SLA-1TO30] INT NOT NULL DEFAULT 0,
    [SLA-31TO60] INT NOT NULL DEFAULT 0,
    [SLA-61TO90] INT NOT NULL DEFAULT 0,
	[SLA-91TO120] INT NOT NULL DEFAULT 0,
	[SLA-121TO150] INT NOT NULL DEFAULT 0,
	[SLA->150] INT NOT NULL DEFAULT 0,
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_sla_date_status
ON [dbo].[txn_analysis_sla] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_sla_date
ON [dbo].[txn_analysis_sla] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_sla_wo_cbi] (
	TRANSACTION_DATE DATE NULL,
    WEEK_NUM     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [SLA-<0] INT NOT NULL DEFAULT 0,
    [SLA-=0] INT NOT NULL DEFAULT 0,
    [SLA-1TO30] INT NOT NULL DEFAULT 0,
    [SLA-31TO60] INT NOT NULL DEFAULT 0,
    [SLA-61TO90] INT NOT NULL DEFAULT 0,
	[SLA-91TO120] INT NOT NULL DEFAULT 0,
	[SLA-121TO150] INT NOT NULL DEFAULT 0,
	[SLA->150] INT NOT NULL DEFAULT 0,
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_sla_wo_cbi_date_status
ON [dbo].[txn_analysis_sla_wo_cbi] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_sla_wo_cbi_date
ON [dbo].[txn_analysis_sla_wo_cbi] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_sla_with_cbi] (
	TRANSACTION_DATE DATE NULL,
    WEEK_NUM     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [SLA-<0] INT NOT NULL DEFAULT 0,
    [SLA-=0] INT NOT NULL DEFAULT 0,
    [SLA-1TO30] INT NOT NULL DEFAULT 0,
    [SLA-31TO60] INT NOT NULL DEFAULT 0,
    [SLA-61TO90] INT NOT NULL DEFAULT 0,
	[SLA-91TO120] INT NOT NULL DEFAULT 0,
	[SLA-121TO150] INT NOT NULL DEFAULT 0,
	[SLA->150] INT NOT NULL DEFAULT 0,
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_sla_with_cbi_date_status
ON [dbo].[txn_analysis_sla_with_cbi] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_sla_with_cbi_date
ON [dbo].[txn_analysis_sla_with_cbi] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_transaction_amount] (
	TRANSACTION_DATE DATE NULL,
    [MONTH]     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [< 500] INT NOT NULL DEFAULT 0,
    [500 TO 999] INT NOT NULL DEFAULT 0,
    [1000 TO 2999] INT NOT NULL DEFAULT 0,
    [3000 TO 4999] INT NOT NULL DEFAULT 0,
    [>= 5000] INT NOT NULL DEFAULT 0
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_transaction_amount_date_status
ON [dbo].[txn_analysis_transaction_amount] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_transaction_amount_date
ON [dbo].[txn_analysis_transaction_amount] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_pay_cash_amt_txn] (
	TRANSACTION_DATE DATE NULL,
    [MONTH]     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [PAY_CASH=0] INT NOT NULL DEFAULT 0,
    [PAY_CASH=1TO499] INT NOT NULL DEFAULT 0,
    [PAY_CASH=500T999] INT NOT NULL DEFAULT 0,
    [PAY_CASH=1000TO2999] INT NOT NULL DEFAULT 0,
    [PAY_CASH=3000TO4999] INT NOT NULL DEFAULT 0,
	[PAY_CASH>=5000] INT NOT NULL DEFAULT 0
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_pay_cash_amt_txn_date_status
ON [dbo].[txn_analysis_pay_cash_amt_txn] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_pay_cash_amt_txn_date
ON [dbo].[txn_analysis_pay_cash_amt_txn] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_pay_cash_amt_vol] (
	TRANSACTION_DATE DATE NULL,
    [MONTH]     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [PAY_CASH=0] BIGINT NOT NULL DEFAULT 0,
    [PAY_CASH=1TO499] BIGINT NOT NULL DEFAULT 0,
    [PAY_CASH=500T999] BIGINT NOT NULL DEFAULT 0,
    [PAY_CASH=1000TO2999] BIGINT NOT NULL DEFAULT 0,
    [PAY_CASH=3000TO4999] BIGINT NOT NULL DEFAULT 0,
	[PAY_CASH>=5000] BIGINT NOT NULL DEFAULT 0
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_pay_cash_amt_vol_date_status
ON [dbo].[txn_analysis_pay_cash_amt_vol] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_pay_cash_amt_vol_date
ON [dbo].[txn_analysis_pay_cash_amt_vol] (TRANSACTION_DATE);


CREATE TABLE [dbo].[txn_analysis_total_denom_cbi_txn] (
	TRANSACTION_DATE DATE NULL,
    [MONTH]     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [DENOMINATION=0] INT NOT NULL DEFAULT 0,
    [DENOMINATION=1TO499] INT NOT NULL DEFAULT 0,
    [DENOMINATION=500T999] INT NOT NULL DEFAULT 0,
    [DENOMINATION=1000TO2999] INT NOT NULL DEFAULT 0,
    [DENOMINATION=3000TO4999] INT NOT NULL DEFAULT 0,
	[DENOMINATION>=5000] INT NOT NULL DEFAULT 0
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_denom_cbi_txn_date_status
ON [dbo].[txn_analysis_total_denom_cbi_txn] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_denom_cbi_txn_date
ON [dbo].[txn_analysis_total_denom_cbi_txn] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_total_denom_cbi_vol] (
	TRANSACTION_DATE DATE NULL,
    [MONTH]     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [DENOMINATION=0] BIGINT NOT NULL DEFAULT 0,
    [DENOMINATION=1TO499] BIGINT NOT NULL DEFAULT 0,
    [DENOMINATION=500T999] BIGINT NOT NULL DEFAULT 0,
    [DENOMINATION=1000TO2999] BIGINT NOT NULL DEFAULT 0,
    [DENOMINATION=3000TO4999] BIGINT NOT NULL DEFAULT 0,
	[DENOMINATION>=5000] BIGINT NOT NULL DEFAULT 0
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_denom_cbi_vol_date_status
ON [dbo].[txn_analysis_total_denom_cbi_vol] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_denom_cbi_vol_date
ON [dbo].[txn_analysis_total_denom_cbi_vol] (TRANSACTION_DATE);

CREATE TABLE [dbo].[txn_analysis_total_per_cash_bill_cbi_txn] (
	TRANSACTION_DATE DATE NULL,
    [MONTH]     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [P20_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P50_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P100_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P200_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P500_DENOM] BIGINT NOT NULL DEFAULT 0,
	[P1000_DENOM] BIGINT NOT NULL DEFAULT 0
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_per_cash_bill_cbi_txn_date_status
ON [dbo].[txn_analysis_total_per_cash_bill_cbi_txn] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_per_cash_bill_cbi_txn_date
ON [dbo].[txn_analysis_total_per_cash_bill_cbi_txn] (TRANSACTION_DATE);


CREATE TABLE [dbo].[txn_analysis_total_per_cash_bill_cbi_vol] (
	TRANSACTION_DATE DATE NULL,
    [MONTH]     INT NOT NULL,
    FINAL_STATUS VARCHAR(50) NOT NULL,
    [P20_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P50_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P100_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P200_DENOM] BIGINT NOT NULL DEFAULT 0,
    [P500_DENOM] BIGINT NOT NULL DEFAULT 0,
	[P1000_DENOM] BIGINT NOT NULL DEFAULT 0
);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_per_cash_bill_cbi_vol_date_status
ON [dbo].[txn_analysis_total_per_cash_bill_cbi_vol] (TRANSACTION_DATE, FINAL_STATUS);

CREATE NONCLUSTERED INDEX idx_txn_analysis_total_per_cash_bill_cbi_vol_date
ON [dbo].[txn_analysis_total_per_cash_bill_cbi_vol] (TRANSACTION_DATE);