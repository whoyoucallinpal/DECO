# DECO Pro Forma - Complete Row Trace & Errors Found

## Purpose
This document traces EXACT row numbers in each sheet to verify all cross-sheet references are correct.

## EXPENSE DETAIL Sheet
**Actual Rows:**
- Row 1: Header "DECO ART CENTER - EXPENSE DETAIL"
- Row 2: (blank)
- Row 3: Column headers
- Row 4: Owner Salaries
- Row 5: Studio Manager
- Row 6: Teachers
- Row 7: Assistants
- Row 8: Payroll Taxes
- Row 9: Rent
- Row 10: Utilities
- Row 11: Insurance
- Row 12: Property Tax
- Row 13: Repairs & Maintenance
- Row 14: Marketing
- Row 15: Professional Services
- Row 16: Liquor License
- Row 17: Supply Replenishment
- Row 18: COGS - Art Supplies
- Row 19: COGS - Pottery
- Row 20: COGS - Beverages
- Row 21: COGS - Retail
- Row 22: TOTAL EXPENSES (NOTE: Does NOT include loan payments to avoid circular reference)

---

## P&L Sheet
**Actual Rows:**
- Row 1: Header
- Row 2: (blank)
- Row 3: Column headers
- Row 4: REVENUE header
- Row 5: Total Revenue
- Row 6: COST OF GOODS SOLD header
- Row 7: Art Supplies COGS
- Row 8: Pottery COGS
- Row 9: Beverage COGS
- Row 10: Retail COGS
- Row 11: Total COGS
- Row 12: GROSS PROFIT
- Row 13: OPERATING EXPENSES header
- Row 14: Owner Salaries
- Row 15: Studio Manager
- Row 16: Teachers
- Row 17: Assistants
- Row 18: Payroll Taxes
- Row 19: Rent
- Row 20: Utilities
- Row 21: Insurance
- Row 22: Property Tax
- Row 23: Repairs & Maintenance
- Row 24: Marketing
- Row 25: Professional Services
- Row 26: Liquor License
- Row 27: Supply Replenishment
- Row 28: Total Operating Expenses
- Row 29: EBITDA
- Row 30: Depreciation
- Row 31: Interest Expense
- Row 32: NET INCOME BEFORE TAX

---

## LOAN AMORTIZATION Sheet
**Actual Rows:**
- Row 1: Header
- Row 2: (blank)
- Row 3: Loan Amount (B3 has formula)
- Row 4: Annual Interest Rate (B4 has formula)
- Row 5: Monthly Interest Rate (B5 has formula)
- Row 6: Loan Term (Years) (B6 has formula)
- Row 7: Loan Term (Months) (B7 has formula)
- Row 8: Grace Period (Months) (B8 has formula)
- Row 9: Monthly Payment (B9 has formula)
- Row 10: (blank)
- Row 11: Amortization table headers
- Row 12: Month -2 (first data row)
- Row 13: Month -1
- Row 14: Month 0
- Row 15: Month 1
- Row 16: Month 2
- ... (continues for 120 months)

---

## CRITICAL ERRORS FOUND:

### ERROR #1: Loan Amortization PMT Formula (Line 1330)
**Current:**
```javascript
sheet.getRange(row, 2).setFormula(`=IF(B4>0,PMT(B${monthlyRateRow},B${loanTermMonthsRow},-B4),0)`)
```

**Problem:** Uses `-B4` (Annual Interest Rate) as the loan amount instead of `-B3` (Loan Amount)

**Fix:** Change to:
```javascript
sheet.getRange(row, 2).setFormula(`=IF(B3>0,PMT(B${monthlyRateRow},B${loanTermMonthsRow},-B3),0)`)
```

This is a SHOW-STOPPER bug - all loan payments will be calculated based on the interest rate percentage (e.g., 6.5%) instead of the actual loan amount (e.g., $200,000)!

---

## Cross-Sheet References To Verify:
- P&L → Expense Detail rows 4-21: ✓ CORRECT
- P&L → Loan Amortization rows (amortRow = 14 + month): ✓ CORRECT
- Cash Flow → Loan Amortization rows (amortRow = 14 + month): ✓ CORRECT
- Expense Detail → Loan Amortization rows (amortRow = 13 + month + 1 = 14 + month): ✓ CORRECT
- Annual Summary → P&L rows: NEED TO VERIFY
- Dashboard → Annual Summary rows: NEED TO VERIFY

