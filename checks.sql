-- =====================================================================
-- Loan Book Reconciliation & Data Quality Controls - PostgreSQL version
-- Reporting date: 2026-09-30. Synthetic data.
-- Run with:  createdb loanbook && psql -d loanbook -f checks.sql
-- (the CSV files must be in the folder you run psql from)
-- =====================================================================

-- ---------- 0. Load the data ----------
DROP TABLE IF EXISTS loan_system_raw, general_ledger, fx_rates CASCADE;

CREATE TABLE loan_system_raw (
    loan_id             TEXT,
    customer_id         TEXT,
    country             TEXT,
    product             TEXT,
    currency            TEXT,
    original_amount     NUMERIC(14,2),
    outstanding_balance NUMERIC(14,2),
    interest_rate       NUMERIC(8,4),
    origination_date    DATE,
    maturity_date       DATE,
    days_past_due       INT,
    ifrs9_stage         INT
);

CREATE TABLE general_ledger (
    loan_id      TEXT,
    gl_account   TEXT,
    balance_eur  NUMERIC(14,2),
    posting_date DATE
);

CREATE TABLE fx_rates (currency TEXT PRIMARY KEY, eur_per_unit NUMERIC(10,4));
INSERT INTO fx_rates VALUES ('EUR', 1), ('USD', 0.855), ('GBP', 1.148);

\copy loan_system_raw FROM 'loan_system.csv'    WITH (FORMAT csv, HEADER true)
\copy general_ledger  FROM 'general_ledger.csv' WITH (FORMAT csv, HEADER true)

-- ---------- 1. Duplicates ----------
-- Which Loan_IDs appear more than once?
SELECT loan_id, COUNT(*) AS n
FROM loan_system_raw
GROUP BY loan_id
HAVING COUNT(*) > 1;

-- Are they exact copies? (same count of distinct full rows as distinct IDs -> yes)
SELECT COUNT(*) AS raw_rows,
       COUNT(DISTINCT loan_id) AS distinct_ids,
       (SELECT COUNT(*) FROM (SELECT DISTINCT * FROM loan_system_raw) d) AS distinct_rows
FROM loan_system_raw;

-- Clean table: keep one copy of each row, add EUR balance
DROP VIEW IF EXISTS sys_clean;
CREATE VIEW sys_clean AS
SELECT s.*,
       f.eur_per_unit AS rate,
       ROUND(s.outstanding_balance * f.eur_per_unit, 2) AS balance_eur
FROM (SELECT DISTINCT * FROM loan_system_raw) s
LEFT JOIN fx_rates f ON f.currency = s.currency;

-- ---------- 2. Validation rules (each one a WHERE) ----------
SELECT 'V1 Missing Customer_ID' AS rule_name, loan_id FROM sys_clean WHERE customer_id IS NULL OR customer_id = ''
UNION ALL
SELECT 'V2 Negative balance', loan_id FROM sys_clean WHERE outstanding_balance < 0
UNION ALL
SELECT 'V3 Maturity before origination', loan_id FROM sys_clean WHERE maturity_date < origination_date
UNION ALL
SELECT 'V4 Originated after reporting date', loan_id FROM sys_clean WHERE origination_date > DATE '2026-09-30'
UNION ALL
SELECT 'V5 IFRS 9 stage vs DPD', loan_id FROM sys_clean
WHERE (days_past_due > 90 AND ifrs9_stage <> 3) OR (days_past_due > 30 AND ifrs9_stage = 1)
UNION ALL
SELECT 'V6 Interest rate out of range', loan_id FROM sys_clean WHERE interest_rate < 0 OR interest_rate > 0.25
UNION ALL
SELECT 'V7 Balance above original amount', loan_id FROM sys_clean WHERE outstanding_balance > original_amount
ORDER BY 1, 2;

-- Count per rule
SELECT 'V1' AS rule, COUNT(*) FILTER (WHERE customer_id IS NULL OR customer_id = '') AS n FROM sys_clean
UNION ALL SELECT 'V2', COUNT(*) FILTER (WHERE outstanding_balance < 0) FROM sys_clean
UNION ALL SELECT 'V3', COUNT(*) FILTER (WHERE maturity_date < origination_date) FROM sys_clean
UNION ALL SELECT 'V4', COUNT(*) FILTER (WHERE origination_date > DATE '2026-09-30') FROM sys_clean
UNION ALL SELECT 'V5', COUNT(*) FILTER (WHERE (days_past_due > 90 AND ifrs9_stage <> 3)
                                           OR (days_past_due > 30 AND ifrs9_stage = 1)) FROM sys_clean
UNION ALL SELECT 'V6', COUNT(*) FILTER (WHERE interest_rate < 0 OR interest_rate > 0.25) FROM sys_clean
UNION ALL SELECT 'V7', COUNT(*) FILTER (WHERE outstanding_balance > original_amount) FROM sys_clean;

-- ---------- 3. Reconciliation (FULL OUTER JOIN) ----------
DROP VIEW IF EXISTS recon;
CREATE VIEW recon AS
SELECT COALESCE(s.loan_id, g.loan_id) AS loan_id,
       s.currency,
       s.balance_eur            AS system_eur,
       g.balance_eur            AS gl_eur,
       COALESCE(s.balance_eur, 0) - COALESCE(g.balance_eur, 0) AS difference,
       CASE
           WHEN g.loan_id IS NULL THEN 'Missing in GL'
           WHEN s.loan_id IS NULL THEN 'Missing in loan system'
           WHEN s.balance_eur = g.balance_eur THEN 'Matched'
           WHEN ABS(s.balance_eur - g.balance_eur) <= 1 THEN 'Within tolerance'
           ELSE 'Break'
       END AS recon_status
FROM sys_clean s
FULL OUTER JOIN general_ledger g ON g.loan_id = s.loan_id;

-- Totals (must equal the Excel Summary sheet)
SELECT (SELECT SUM(balance_eur) FROM sys_clean)      AS system_total_eur,
       (SELECT SUM(balance_eur) FROM general_ledger) AS gl_total_eur,
       (SELECT SUM(balance_eur) FROM sys_clean)
     - (SELECT SUM(balance_eur) FROM general_ledger) AS total_difference;

-- Reconciliation proof: differences by status add up to the total difference
SELECT recon_status, COUNT(*) AS n, SUM(difference) AS difference_eur
FROM recon
GROUP BY recon_status
ORDER BY recon_status;

-- Detail of the breaks
SELECT * FROM recon WHERE recon_status = 'Break' ORDER BY ABS(difference) DESC;
