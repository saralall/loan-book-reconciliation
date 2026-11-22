# Loan Book Reconciliation & Data Quality Controls

A quarter-end check of the kind banks run before reporting figures: does the **loan system** agree with the **general ledger**, and does the loan data pass basic **quality rules**? Built in **Excel** and verified independently in **PostgreSQL**.

> **Data:** synthetic. A 250-loan Baltic portfolio (LT / LV / EE) containing realistic data-quality issues. No real customers or bank data.

**Reporting date:** 30 Sep 2026 · **Tolerance:** differences up to ±1.00 EUR are treated as rounding

![Summary sheet](screenshots/summary.png)

## Results at a glance

| | |
|---|---|
| Loans reconciled | 250 (after removing 3 duplicate rows) |
| Total difference, system vs. ledger | **140,197.64 EUR**, fully explained, **0.00 unexplained** |
| Exceptions logged | **38** (24 High, 10 Medium, 4 Low) |
| Biggest break | −26,971.89 EUR: a USD loan booked in the ledger without FX conversion |
| Cross-check | Excel and PostgreSQL produce identical totals and counts |

## Repository contents

| File | What it is |
|---|---|
| `loan_book_reconciliation.xlsx` | The workbook: cleaning, validation, reconciliation, exceptions log and summary |
| `checks.sql` | The same checks in PostgreSQL |
| `loan_system.csv`, `general_ledger.csv` | Source data used by `checks.sql` |
| `screenshots/` | Images of the key sheets |

**Workbook sheets:** README → Summary → Exceptions → SYS_Clean → GL_Work → Profile → Data_Dictionary → Loan_System (raw) → General_Ledger (raw) → FX_Rates (raw). Built for Excel 365 (uses XLOOKUP). If any cell looks empty, press **Ctrl+Alt+F9**.

---

## Approach

### 1. Set up
- The raw sheets are never edited. Work happens on copies (`SYS_Clean`, `GL_Work`).
- The data is in Excel Tables (`tSys`, `tGL`, `tFX`, `tExc`, `tDict`), so formulas read as `tSys[Loan_ID]` rather than `$A$2:$A$251`.
- A named cell `RepDate` holds the reporting date. `Data_Dictionary` documents every column and its valid values.

### 2. Profile the raw data
The `Profile` sheet shows rows, blanks, distinct values and min/max per column, before any cleaning. It surfaced the problems that the rules then target:
- 253 rows but only 250 distinct Loan IDs, so there are duplicates
- 4 blank Customer IDs
- a minimum balance of −97,110.43
- a maximum interest rate of 45%
- an origination date after the reporting date, and a maturity date (2015) earlier than any loan was granted (2016)

### 3. Remove duplicates
`Dup_Count = COUNTIFS(tSys[Loan_ID],[@Loan_ID])` found 3 Loan IDs appearing twice: LN-10041, LN-10148 and LN-10196. All three were confirmed as exact copies, and one of each was kept (253 → 250 rows).
**Why first:** leaving them in would overstate the system total by 615,527.45 EUR, and the reconciliation could never tie out.

### 4. Validation rules
There is one TRUE/FALSE column per rule, and `Issue_Count` sums them (`--` turns TRUE into 1). Conditional formatting highlights flagged rows.

| Rule | Check | Found |
|---|---|---|
| V1 | Customer_ID missing | 4 |
| V2 | Negative outstanding balance | 3 |
| V3 | Maturity date before origination date | 3 |
| V4 | Originated after the reporting date | 2 |
| V5 | IFRS 9 stage inconsistent with days past due | 5 |
| V6 | Interest rate outside 0–25% | 2 |
| V7 | Outstanding balance above original amount | 3 |

**V5 is the banking rule.** Under IFRS 9:
- **Stage 1** is a performing loan.
- **Stage 2** means a significant increase in credit risk. The standard backstop is more than 30 days past due.
- **Stage 3** means credit-impaired, with a backstop of more than 90 days past due.

So a loan more than 90 days past due must be in Stage 3, and a loan more than 30 days past due can't be in Stage 1. For example, LN-10229 is 106 days past due but sits in Stage 1, so its loss provision is understated. The rule only flags stages that are *too low*. Stage 2 with 0 days past due is allowed, because loans move to Stage 2 for other reasons too (a rating downgrade or a watch list, for example).

### 5. Convert to EUR
```
Rate        = XLOOKUP([@Currency], tFX[Currency], tFX[EUR_per_unit])
Balance_EUR = ROUND([@Outstanding_Balance]*[@Rate], 2)
```

### 6. Reconcile the loan system with the general ledger
```
GL_Balance   = XLOOKUP([@Loan_ID], tGL[Loan_ID], tGL[Balance_EUR], "MISSING")
Difference   = IF(ISNUMBER([@GL_Balance]), ROUND([@Balance_EUR]-[@GL_Balance],2), "")
Recon_Status = Missing in GL / Matched / Within tolerance (|diff| ≤ 1) / Break
In_System    = ISNUMBER(MATCH([@Loan_ID], tSys[Loan_ID], 0))   -- in GL_Work
```

| Status | Loans | Difference (EUR) |
|---|---|---|
| Matched | 234 | 0.00 |
| Within tolerance | 6 | −1.19 |
| Break | 6 | −33,131.18 |
| Missing in GL | 4 | 217,950.29 |
| Missing in loan system | 3 | −44,620.28 |

**The 6 breaks and their likely causes:**

| Loan | Ccy | System EUR | GL EUR | Difference | Likely cause |
|---|---|---|---|---|---|
| LN-10168 | USD | 159,041.14 | 186,013.03 | −26,971.89 | **FX:** the GL booked the USD amount as if it were EUR (186,013.03 × 0.855 = 159,041.14) |
| LN-10002 | EUR | 4,327.00 | 8,212.06 | −3,885.06 | Amount mismatch on a EUR loan, so not FX; check the GL posting |
| LN-10234 | EUR | 27,634.98 | 31,064.68 | −3,429.70 | Amount mismatch; possibly an unposted repayment or timing difference |
| LN-10116 | EUR | 1,561.72 | −539.96 | +2,101.68 | Amount mismatch; the GL balance is negative, check the posting |
| LN-10015 | USD | 2,113.39 | 3,497.83 | −1,384.44 | Matches neither the converted nor the original USD amount; check FX and posting |
| LN-10163 | EUR | 47,576.61 | 47,138.38 | +438.23 | Small mismatch; a repayment was likely posted in one system only |

**A rounding detail worth knowing:** LN-10043 is a GBP loan where 346,651.25 × 1.148 = 397,955.635 exactly. Excel rounds the half up to .64, but the ledger holds .63, consistent with binary floating-point rounding in the source system. The 0.01 difference falls within tolerance. It's a small, real example of why reconciliations need a tolerance at all.

### 7. Exceptions log
The `Exceptions` sheet has one row per issue: Loan_ID | Check | Description | Amount (EUR) | Severity | Suggested action.

| Severity | Meaning | Count |
|---|---|---|
| High | Affects reported amounts: duplicates, negative balances, loans after the reporting date, balance above original, missing records, breaks | 24 |
| Medium | Classification or plausibility: IFRS 9 stage, impossible dates, out-of-range rates | 10 |
| Low | Reference data: missing Customer_ID | 4 |

Rounding differences within tolerance are **not** exceptions. That's the purpose of the tolerance.

![Exceptions sheet](screenshots/exceptions.png)

### 8. Summary report
Every figure on `Summary` is a live formula (SUM, SUMIFS, COUNTIFS on the tables).

**Reconciliation statement**

| Line | Loans | EUR |
|---|---|---|
| Loan system total | 250 | 20,525,849.06 |
| Less: general ledger total | 249 | 20,385,651.42 |
| **Total difference** | | **140,197.64** |
| Missing in GL | 4 | 217,950.29 |
| Missing in loan system | 3 | −44,620.28 |
| Breaks | 6 | −33,131.18 |
| Rounding (within ±1 EUR) | 6 | −1.19 |
| **Total explained** | | **140,197.64** |
| **Unexplained** | | **0.00** |

**Exposure by country and product (EUR)**, shown as a SUMIFS grid and a PivotTable:

| | Business | Consumer | Credit card | Mortgage | Total |
|---|---|---|---|---|---|
| EE | 2,353,962.96 | 145,816.47 | 5,908.84 | 1,456,088.61 | 3,961,776.88 |
| LT | 5,340,519.80 | 183,067.63 | 33,958.79 | 3,948,860.65 | 9,506,406.87 |
| LV | 4,924,527.13 | 228,094.73 | 14,650.43 | 1,890,393.02 | 7,057,665.31 |
| **Total** | 12,619,009.89 | 556,978.83 | 54,518.06 | 7,295,342.28 | 20,525,849.06 |

**IFRS 9 stage split**

| Stage | Loans | Balance (EUR) | % of balance |
|---|---|---|---|
| 1 | 221 | 17,850,638.40 | 87.0% |
| 2 | 21 | 2,295,191.30 | 11.2% |
| 3 | 8 | 380,019.36 | 1.9% |

The sheet also has exceptions by check and severity, a stacked column chart of exposure, and a sign-off box (Prepared by / Date / Reviewed by / Status).

### 9. Independent check in PostgreSQL
`checks.sql` repeats every step outside Excel:
- `GROUP BY … HAVING COUNT(*) > 1` finds the duplicates, and `SELECT DISTINCT *` builds the clean view.
- Each rule V1–V7 is a `WHERE` condition.
- The reconciliation is a `FULL OUTER JOIN` on `loan_id`, with a `CASE` for the status.

The results match Excel exactly: 20,525,849.06 vs. 20,385,651.42, a difference of 140,197.64, and the same count for every status.

```bash
createdb loanbook
cd loan-book-reconciliation      # the folder with the CSV files
psql -d loanbook -f checks.sql
```
Tested on PostgreSQL 16.

---

## Limitations
- LN-10105 and LN-10146 were originated after the reporting date. They're flagged as exceptions but still included in the reconciliation totals, because both systems hold them. In a real Q3 report, they would be excluded or their dates corrected.
- The exceptions log holds values, not formulas, so it's a snapshot of this run.
- The likely causes of breaks are hypotheses. Confirming them would need the posting history from Finance.

## If this ran every month
Load both sources with Power Query (or SQL), keep the rules and the reconciliation as fixed queries, refresh Summary and Exceptions with one click, and publish to Power BI. The analyst would then only investigate exceptions instead of rebuilding the checks each month.

---

**Skills used:** Excel 365 (Tables, XLOOKUP, SUMIFS/COUNTIFS, PivotTables, conditional formatting, charts) · PostgreSQL (CTEs, FULL OUTER JOIN, conditional aggregation) · data validation · reconciliation · IFRS 9 staging
