# DECO Pro Forma V3 - Row Mapping Audit

## ASSUMPTIONS SHEET - Actual Row Numbers

| Row | Cell | Value | Notes |
|-----|------|-------|-------|
| 4 | B4 | 5000 | Square Footage |
| 5 | B5 | 15 | Base Rent ($/sf/year) |
| 6 | B6 | 4 | NNN - Triple Net ($/sf/year) |
| 7 | B7 | =B4*(B5+B6) | Total Annual Rent |
| 8 | B8 | =B7/12 | Monthly Rent |
| 11 | B11 | March 2026 | Opening Date |
| 12 | B12 | 10 | Year 1 Operating Months |
| 13 | B13 | 6 | Class Sessions Year 1 |
| 14 | B14 | 12 | Class Sessions Year 2+ |
| 17 | B17 | ='Startup Costs'!B28 | Startup Costs Total |
| 18 | B18 | 4 | Operating Reserve (months) |
| 19 | B19 | =Expenses!B35/10*B17 | Operating Reserve Amount **WRONG FORMULA** |
| 20 | B20 | =B16+B18 | Total Loan Amount **WRONG FORMULA** |
| 21 | B21 | 10 | Loan Term (years) |
| 22 | B22 | 0.1225 | Interest Rate |
| 23 | B23 | =-PMT(B21/12,B20*12,B19) | Monthly Payment **WRONG FORMULA** |
| 26 | B26 | 0.50 | Year 1 Capacity % |
| 27 | B27 | 0.20 | Annual Growth % |
| 28 | B28 | 0.05 | Drop-in Rate % |
| 31 | B31 | 50000 | Owner 1 Annual Salary |
| 32 | B32 | 50000 | Owner 2 Annual Salary |
| 33 | B33 | 25 | Studio Manager Hourly |
| 34 | B34 | 40 | Studio Manager Hours/Week |
| 35 | B35 | 25 | Teacher Hourly |
| 36 | B36 | 15 | Assistant Hourly |
| 37 | B37 | 0.0765 | Payroll Tax Rate |
| 40 | B40 | 8 | Paint & Sip Events/Month |
| 41 | B41 | 45 | Paint & Sip Price |
| 42 | B42 | 15 | Paint & Sip Attendance |
| 43 | B43 | 1 | Private Events/Month |
| 44 | B44 | 1500 | Private Event Base Fee |
| 45 | B45 | 20 | Private Event Per Person |
| 46 | B46 | 25 | Private Event Attendance |
| 47 | B47 | 1 | Workshops/Month |
| 48 | B48 | 75 | Workshop Price |
| 49 | B49 | 15 | Workshop Attendance |
| 52 | B52 | 2 | Drinks/Person (Paint & Sip) |
| 53 | B53 | 2 | Drinks/Person (Private Event) |
| 54 | B54 | 1.5 | Drinks/Person (Workshop) |
| 55 | B55 | 4 | Coffee Price |
| 56 | B56 | 6 | Beer Price |
| 57 | B57 | 8 | Wine Price |
| 58 | B58 | 0.30 | Beverage COGS % |
| 61 | B61 | 0.0756 | Art Supplies COGS % |
| 62 | B62 | 0.0202 | Pottery COGS % |
| 63 | B63 | 0.30 | Supplies Replenish % (Quarterly) |
| 64 | B64 | 1200 | Electricity/Month |
| 65 | B65 | 400 | Gas/Month |
| 66 | B66 | 300 | Water/Month |
| 67 | B67 | 200 | Internet/Month |
| 68 | B68 | 1500 | Insurance/Month |
| 69 | B69 | 0 | Property Tax/Year |
| 70 | B70 | 5000 | Repairs & Maintenance/Year |
| 71 | B71 | 1500 | Marketing Year 1/Month |
| 72 | B72 | 0.03 | Marketing Year 2+ % |
| 73 | B73 | 3600 | Accounting/Year |
| 74 | B74 | 2400 | Legal/Year |
| 75 | B75 | 1800 | Payroll Service/Year |
| 76 | B76 | 1030 | Liquor License/Year |
| 79 | B79 | =15000/15 | Leasehold Improvements Depreciation |
| 80 | B80 | =17800/7 | Equipment Depreciation |
| 81 | B81 | =B79+B80 | Total Annual Depreciation |

## REVENUE MODEL SHEET - Actual Row Numbers

| Row | Content | Notes |
|-----|---------|-------|
| 5 | Classes Revenue | SUMPRODUCT formula |
| 6 | Drop-in Classes | Based on Classes Revenue |
| 10 | Summer Camps Revenue | SUMPRODUCT formula |
| 14 | Paint & Sip Events | |
| 15 | Private Events | |
| 16 | Visiting Artist Workshops | |
| 20 | Paint & Sip Beverages | |
| 21 | Private Event Beverages | |
| 22 | Workshop Beverages | |
| 26 | Retail Sales (5% of program revenue) | Calculated later |
| 28 | TOTAL GROSS REVENUE | SUM of all revenue |

## EXPENSES SHEET - Actual Row Numbers

| Row | Content | Notes |
|-----|---------|-------|
| 5 | Art Supplies | COGS calculation |
| 6 | Pottery Supplies | COGS calculation |
| 7 | Beverage COGS | |
| 8 | Supplies Replenishment (Quarterly) | |
| 11 | Owner Salaries | |
| 12 | Studio Manager | |
| 13 | Teachers (calculated from enrollment) | |
| 14 | Assistants | |
| 15 | Payroll Taxes | |
| 18 | Rent (Base + NNN) | |
| 19 | Loan Payment | |
| ... | (more rows) | Need to trace |

---

## IDENTIFIED ERRORS

### 1. ASSUMPTIONS SHEET INTERNAL ERRORS

**B19 - Operating Reserve Amount:**
- Current: `=Expenses!B35/10*B17`
- Problem: B17 is Startup Costs Total, should be B18 (Operating Reserve months)
- Should be: `=Expenses!B[total expenses row]/12*B18`

**B20 - Total Loan Amount:**
- Current: `=B16+B18`
- Problem: B16 is section header (blank), B18 is Operating Reserve months (4)
- Should be: `=B17+B19` (Startup Costs + Operating Reserve Amount)

**B23 - Monthly Payment:**
- Current: `=-PMT(B21/12,B20*12,B19)`
- Problem: All three parameters are wrong!
  - B21 = Loan Term (10), not Interest Rate
  - B20 = wrong Total Loan Amount
  - B19 = Operating Reserve Amount, not Loan Amount
- Should be: `=-PMT(B22/12,B21*12,B20)`
  - Rate: B22/12 (Interest Rate / 12)
  - Nper: B21*12 (Loan Term * 12)
  - Pv: B20 (Total Loan Amount)

### 2. REVENUE MODEL ERRORS

**Row 5 - Classes Revenue (Year 1):**
- Current: `Assumptions!$B$8` (Monthly Rent)
- Should be: `Assumptions!$B$13` (Class Sessions Year 1 = 6)

**Row 5 - Classes Revenue (Years 2-5):**
- Current: `Assumptions!$B$9` (blank row)
- Should be: `Assumptions!$B$14` (Class Sessions Year 2+ = 12)

**Row 6 - Drop-in Classes:**
- Current: `Assumptions!$B$23` (Monthly Payment formula)
- Should be: `Assumptions!$B$28` (Drop-in Rate % = 0.05)

**Rows 14-16 - Events & Workshops:**
- Need to check all Assumptions references (likely all off by ~6 rows)

**Rows 20-22 - Beverage Sales:**
- Current formulas reference B34, B35, B36, B37, B38, etc.
- These are in the STAFF COMPENSATION section, should be in EVENTS & WORKSHOPS and BEVERAGE SALES sections
- Need complete recalculation

### 3. EXPENSES SHEET ERRORS

**Row 5 - Art Supplies:**
- Current: `'Revenue Model'!B26*Assumptions!$B$55`
- Problems:
  1. Revenue Model B26 is Retail Sales (wrong revenue source)
  2. Assumptions B55 is Coffee Price ($4), not Art Supplies COGS %
- Should be: `=[appropriate revenue rows]*Assumptions!$B$61` (Art Supplies COGS % = 0.0756)
- **QUESTION FOR USER:** Which revenue should Art Supplies COGS apply to?

**Row 6 - Pottery Supplies:**
- Current: `'Revenue Model'!B26*Assumptions!$B$56`
- Problems:
  1. Revenue Model B26 is Retail Sales (wrong)
  2. Assumptions B56 is Beer Price ($6), not Pottery COGS %
- Should be: `=[appropriate revenue rows]*Assumptions!$B$62` (Pottery COGS % = 0.0202)

**Row 7 - Beverage COGS:**
- Current: `('Revenue Model'!B20+B21+B22)*Assumptions!$B$53`
- Assumptions B53 is Drinks/Person (Private Event) = 2, not Beverage COGS %
- Should be: `('Revenue Model'!B20+B21+B22)*Assumptions!$B$58` (Beverage COGS % = 0.30)

**Row 8 - Supplies Replenishment:**
- Current: `'Startup Supplies'!E66*Assumptions!$B$57*3`
- Assumptions B57 is Wine Price ($8), not Supplies Replenish %
- Should be: `'Startup Supplies'!E66*Assumptions!$B$63*3` (Supplies Replenish % = 0.30)

**Row 11 - Owner Salaries:**
- Current: `Assumptions!$B$25+Assumptions!$B$26`
- B25 is section header, B26 is Year 1 Capacity % (0.50)
- Should be: `Assumptions!$B$31+Assumptions!$B$32` (Owner 1 + Owner 2 = $50K each)

**Row 12 - Studio Manager:**
- Current: `Assumptions!$B$27*Assumptions!$B$28*52`
- B27 is Annual Growth % (0.20), B28 is Drop-in Rate % (0.05)
- Should be: `Assumptions!$B$33*Assumptions!$B$34*52` (Hourly $25 * Hours/Week 40 * 52)

**Row 13 - Teachers:**
- Current: Uses `Assumptions!$B$8` (Monthly Rent) and `Assumptions!$B$29` (blank)
- Should use: `Assumptions!$B$13` (Class Sessions Yr1) and `Assumptions!$B$35` (Teacher Hourly $25)
- Year 2+ should use: `Assumptions!$B$14` (Class Sessions Yr2+ = 12)

**Row 15 - Payroll Taxes:**
- Current: `SUM(B11:B14)*Assumptions!$B$31`
- Assumptions B31 is Owner 1 Salary ($50K), not Payroll Tax Rate
- Should be: `SUM(B11:B14)*Assumptions!$B$37` (Payroll Tax Rate = 0.0765)

**Row 18 - Rent:**
- Current: `Assumptions!$B$8*10` (Monthly Rent * 10) - This one is CORRECT!

**Row 19 - Loan Payment:**
- Current: `Assumptions!$B$22*10`
- B22 is Interest Rate (0.1225), not Monthly Payment
- Should be: `Assumptions!$B$23*10` (Monthly Payment)

---

## PATTERN ANALYSIS

There appears to be a systematic offset of approximately **6 rows** in most Assumptions references throughout the file. This suggests the row tracking variable `r` got out of sync during code development.

### Systematic Offsets Identified:
1. **Revenue Model to Assumptions:** Most references are 6+ rows too low
2. **Expenses to Assumptions:** Most references are 6+ rows too low
3. **Assumptions internal formulas:** Some self-references are off by 1-2 rows

---

## NEXT STEPS

1. ✅ Complete row mapping for all sheets
2. ⏳ Ask user for clarification on COGS calculations (which revenue streams to include)
3. ⏳ Fix all Assumptions sheet internal formulas
4. ⏳ Fix all Revenue Model references to Assumptions
5. ⏳ Fix all Expenses references to Assumptions and Revenue Model
6. ⏳ Check other sheets (Startup Costs, Loan Amortization, Dashboard, etc.)
7. ⏳ Test and verify all formulas
8. ⏳ Commit fixes
