# DECO Pro Forma Testing Checklist

## How to Test

1. **Delete any existing spreadsheet** created by the old version
2. **Run the script:** Execute `createDECOProForma()` in Google Apps Script
3. **Check these key cells:**

---

## Critical Tests (MUST PASS)

### ✓ Loan Amortization Sheet
- **Cell B9 (Monthly Payment):** Should show ~$2,000-$3,000 (NOT $0.50!)
  - If it shows less than $100, the PMT formula bug still exists
  - Expected: Around $2,200 for a $200K loan at 6.5% over 10 years

- **Cell B15 (Month 1 Payment in table):** Should match B9
- **Cell F15 (Month 1 Interest):** Should be reasonable (e.g., $1,083 for $200K at 6.5%)
- **Cell E15 (Month 1 Principal):** Should be reasonable (Payment - Interest)

### ✓ P&L Sheet
- **Row 5 (Total Revenue), Column B (Month 1):** Should have a dollar amount (not #REF!)
- **Row 11 (Total COGS), Column B:** Should have a dollar amount
- **Row 28 (Total Operating Expenses), Column B:** Should have a dollar amount
- **Row 31 (Interest Expense), Column B:** Should match Loan Amortization F15
- **Row 32 (Net Income), Column B:** Should be a calculated value (could be negative in early months)

### ✓ Cash Flow Sheet
- **Row 4 (Beginning Cash Balance), Column B (Month 1):** Should be $0
- **Row 17 (Ending Cash Balance), Column B:** Should be a calculated value
- **Row 12 (Loan Principal Payment), Column B:** Should match Loan Amortization E15
- **Row 13 (Loan Interest Payment), Column B:** Should match Loan Amortization F15

### ✓ Annual Summary Sheet
- **Row 5 (Total Revenue), Column B (Year 1):** Should be sum of 10 months
- **Row 19 (EBITDA), Column B:** Should be a calculated value
- **Row 22 (Ending Cash Balance), Column B:** Should match Cash Flow K17

### ✓ Dashboard Sheet
- **Row 6 (Loan Amount):** Should match Startup Costs B30
- **Row 10 (Monthly Payment):** Should match Loan Amortization B9

---

## What to Look For

**✓ GOOD SIGNS:**
- All cells show dollar amounts (no #REF!, #DIV/0!, #VALUE! errors)
- Loan payment is realistic ($2,000-3,000 range)
- Revenue and expenses flow through all sheets
- Net Income calculations work (even if negative)

**✗ BAD SIGNS:**
- #REF! errors (means a sheet reference is broken)
- #DIV/0! errors (means dividing by zero somewhere)
- Loan payment under $100 (PMT formula still broken)
- Empty cells in P&L, Cash Flow, or Annual Summary

---

## If You See Errors

**Note exactly which cells/sheets have errors and what type:**
- Cell address (e.g., "Dashboard B6")
- Error type (#REF!, #DIV/0!, #VALUE!, or just blank)
- What formula is in that cell (click on it to see)

This will help me fix any remaining issues precisely.

---

## Quick Sanity Check

**Does this make business sense?**
- Month 1 should show startup activity (loan proceeds come in, startup costs go out)
- Early months may show negative Net Income (normal during ramp-up)
- By Year 2-3, Net Income should turn positive
- Cash balance should build over time (not go deeply negative)
