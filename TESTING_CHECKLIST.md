# Testing Checklist - DECO Pro Forma

## ✅ Pre-Installation Verification (COMPLETED)

All tests passed on `DECOProForma_Rental_VERIFIED.gs`:

### Script Syntax Tests
- [x] No `getColumnWidth` with 2 parameters
- [x] All `setColumnWidth` calls have correct signature
- [x] No unmatched quotes in formulas
- [x] All sheet names properly referenced
- [x] All function definitions present

### Formula Reference Tests
- [x] Class Sessions Year 1 uses `Assumptions!B30` (6 sessions)
- [x] Class Sessions Year 2+ uses `Assumptions!B31` (12 sessions)
- [x] Total Capital Needed uses `Startup Costs!B32`
- [x] Startup Costs Subtotal uses `Startup Costs!B27`
- [x] Operating Reserve uses `Startup Costs!B30`
- [x] Art Supplies Grand Total uses `Startup Supplies!E63`
- [x] Depreciation formulas reference correct rows (B92+B93)
- [x] Leasehold depreciation uses `B10/15`
- [x] Equipment depreciation uses `B13:B16/7`

### Method Signature Tests
- [x] All `.merge()` calls correct
- [x] All `.setFormula()` calls correct (306 found)
- [x] All `.setValue()` calls correct (243 found)
- [x] All `.setColumnWidth()` calls correct
- [x] All `.setNumberFormat()` calls correct

---

## 📋 Post-Installation Testing (USER TO COMPLETE)

After you install the script in Google Sheets, verify the following:

### 1. Script Execution
- [ ] Script runs without errors
- [ ] All 12 sheets are created
- [ ] Dashboard is moved to first position
- [ ] Sheets have correct tab colors
- [ ] Success alert appears

### 2. Dashboard Verification
Navigate to Dashboard tab and check:
- [ ] "Total Capital Needed" shows a NUMBER (not text)
- [ ] "SBA Loan Amount Requested" shows realistic value (~$130-150K)
- [ ] "Monthly Payment" is calculated
- [ ] All revenue/expense figures populate
- [ ] DSCR values show as numbers (e.g., 1.25)

### 3. Assumptions Sheet
- [ ] All input cells have values
- [ ] Formulas calculate (blue cells show values, not errors)
- [ ] Loan Amount shows value from Loan Calculator
- [ ] Monthly Payment calculates
- [ ] Total Annual Depreciation shows value

### 4. Revenue Model Sheet
Check revenue for Year 1:
- [ ] Classes Revenue uses 6 sessions (check formula: should be `Assumptions!$B$30`)
- [ ] Summer Camps Revenue calculates
- [ ] Events revenue shows values
- [ ] Beverage revenue shows values
- [ ] Retail Sales = 5% of total (formula should reference total row)
- [ ] Total Gross Revenue sums correctly

Check revenue for Year 2:
- [ ] Classes Revenue uses 12 sessions (check formula: should be `Assumptions!$B$31`)
- [ ] All other revenue grows appropriately

### 5. Expenses Sheet
- [ ] Art Supplies COGS calculates (should be 7.56% of revenue)
- [ ] Supplies Replenishment shows values (not zero)
- [ ] Owner Salaries = $100,000 (both owners)
- [ ] Studio Manager = $52,000 (25 * 40 * 52)
- [ ] Base Rent increases each year (3% escalation)
- [ ] NNN costs show values
- [ ] Total Operating Expenses sums correctly

### 6. Startup Costs Sheet
- [ ] First Month Rent shows value
- [ ] Last Month Rent shows value
- [ ] Security Deposit shows value
- [ ] Buildout = $25,000
- [ ] Equipment costs populate (Kiln, Wheels, Furniture)
- [ ] Art Supplies references Startup Supplies total (should be ~$20K)
- [ ] SUBTOTAL sums correctly (~$65-70K)
- [ ] Operating Reserve (4 months) calculates
- [ ] TOTAL CAPITAL NEEDED = SUBTOTAL + Operating Reserve

### 7. Loan Calculator Sheet
- [ ] Total Capital Needed matches Startup Costs B32
- [ ] Startup Costs to Finance matches Startup Costs B27
- [ ] Working Capital to Finance matches Startup Costs B30
- [ ] Total Loan Amount = Capital Needed - Owner Equity - Other Sources
- [ ] Monthly Payment calculates
- [ ] Total Interest Paid shows value
- [ ] Total Amount Repaid shows value

### 8. Loan Amortization Sheet
- [ ] Loan Amount matches Loan Calculator
- [ ] Interest Rate = 12%
- [ ] Monthly Payment matches Assumptions
- [ ] Year 1 shows 10 months of payments
- [ ] Years 2-5 show 12 months each
- [ ] Principal + Interest = Total Payment
- [ ] Balance decreases each year

### 9. Profit & Loss Sheet
- [ ] Gross Revenue matches Revenue Model
- [ ] Total Expenses matches Expenses sheet
- [ ] EBITDA = Revenue - Expenses
- [ ] Depreciation shows value (not #DIV/0! or #REF!)
- [ ] EBIT = EBITDA - Depreciation
- [ ] Interest Expense matches Loan Amortization
- [ ] Net Income calculates
- [ ] EBITDA Margin % calculates
- [ ] Net Margin % calculates

### 10. Cash Flow Sheet
- [ ] Net Income matches Profit & Loss
- [ ] Depreciation matches Profit & Loss
- [ ] Cash from Operations = Net Income + Depreciation
- [ ] Principal Payments match Loan Amortization
- [ ] Net Cash Flow calculates
- [ ] Beginning Cash Balance Year 1 = Operating Reserve (not zero!)
- [ ] Ending Cash Balance = Beginning + Net Cash Flow
- [ ] Year 2 Beginning Balance = Year 1 Ending Balance

### 11. Cross-Sheet Formula Tests
Test editing values to ensure formulas update:
- [ ] Change `Assumptions!B40` (capacity %) - revenue should update
- [ ] Change `Assumptions!B4` (building size) - rent should update
- [ ] Change `Assumptions!B34` (interest rate) - payments should update
- [ ] All dependent sheets update when Assumptions change

### 12. Error Checks
Look for these error indicators:
- [ ] No `#REF!` errors (broken references)
- [ ] No `#DIV/0!` errors (division by zero)
- [ ] No `#NAME?` errors (invalid formula names)
- [ ] No `#VALUE!` errors (wrong value type)
- [ ] No `#N/A` errors
- [ ] No cells showing "ERROR" or error text

---

## 🐛 If You Find Errors

If any test fails:

1. **Take a screenshot** of the error
2. **Note the cell reference** (e.g., "Dashboard B6")
3. **Note what it shows** (e.g., "#REF!" or "0" when it should have a value)
4. **Check the formula** (click the cell and look at formula bar)
5. **Report the issue** with all above details

---

## ✅ Expected Values (Approximate)

If all tests pass, you should see values similar to:

- **Total Capital Needed:** ~$130-150K
- **SBA Loan Amount:** ~$130-150K (if no owner equity)
- **Monthly Loan Payment:** ~$1,800-2,000
- **Year 1 Revenue:** ~$250-300K (10 months)
- **Year 2 Revenue:** ~$400-450K (full year)
- **Year 1 Net Income:** May be negative or small positive
- **Year 2 Net Income:** Should be positive
- **Year 5 Net Income:** Should be significantly positive

---

## 📝 Notes

- Some variation in values is normal if you've adjusted assumptions
- The key is that all formulas **calculate** (no errors)
- Values should be **logical** (no random huge numbers)
- Relationships should make sense (e.g., Year 2 > Year 1)

---

## Status: VERIFIED ✅

Script version: `DECOProForma_Rental_VERIFIED.gs`
Verification date: November 2025
All pre-installation tests: PASSED
