# DECO Art Center Pro Forma - Requirements Document

## Project Overview
Create a Google Apps Script that generates a comprehensive 5-year financial pro forma for DECO Art Center in Tulsa, OK (zip 74104). This model will:
- Project financial performance for SBA loan application
- Allow adjustable assumptions for scenario planning
- Show monthly detail for all 60 months plus annual summaries
- Auto-calculate required loan amount based on startup costs + working capital

## Timeline & Structure

### Key Dates
- **January 2026 (Month -2)**: Loan funded, pre-opening expenses begin
- **February 2026 (Month -1)**: Buildout and setup continues
- **March 2026 (Month 1)**: Business opens for operations
- **Month Labels**: MAR26, APR26, MAY26, JUN26, etc. through FEB31

### Output Structure
- **Monthly Detail**: All 60 months (Mar 2026 - Feb 2031)
- **Annual Summary**: Separate sheet with 5-year overview

---

## Facility Assumptions

| Item | Default Value | Adjustable? |
|------|--------------|-------------|
| Square Footage | 5,000 sf | Yes |
| Base Rent | $16/sf/month | Yes |
| NNN (Triple Net) | $4/sf/month | Yes |
| Total Monthly Rent | $100,000/year | Auto-calc |

---

## Revenue Streams

### 1. Regular Classes (Year-Round)

**Schedule**: All months, 12 sessions per year (1 session = 1 month of classes)

| Class Type | Duration | Price | Max Students | Yr1 Enroll | Yr2 | Yr3 | Yr4 | Yr5 |
|------------|----------|-------|--------------|------------|-----|-----|-----|-----|
| Adults - Painting | 2.0 hrs | $199 | 15 | 85% | 94% | 100% | 100% | 100% |
| Adults - Pottery | 2.5 hrs | $229 | 6 | 85% | 94% | 100% | 100% | 100% |
| Teens 13-18 - Drawing | 2.0 hrs | $199 | 15 | 85% | 94% | 100% | 100% | 100% |
| Teens 13-18 - Pottery | 2.0 hrs | $229 | 6 | 85% | 94% | 100% | 100% | 100% |
| Kids 9-12 - Mixed Media | 1.5 hrs | $159 | 15 | 85% | 94% | 100% | 100% | 100% |
| Kids 9-12 - Pottery | 1.5 hrs | $179 | 6 | 85% | 94% | 100% | 100% | 100% |
| Kids 6-8 - General Art | 1.5 hrs | $154 | 15 | 85% | 94% | 100% | 100% | 100% |
| Kids 4-5 - General Art | 1.5 hrs | $154 | 15 | 85% | 94% | 100% | 100% | 100% |
| Homeschool 6-12 | 1.5 hrs | $154 | 30 | 50% | 60% | 70% | 85% | 100% |
| Homeschool Teens | 1.5 hrs | $154 | 30 | 50% | 60% | 70% | 85% | 100% |

**Revenue Calculation**: Price × Max Students × Enrollment % × 12 sessions/year

**Staffing Requirements**:
- 1 teacher per class (class duration × $25/hr)
- 1 assistant if enrolled students > 10 (class duration × $15/hr)

---

### 2. Summer Camps (June, July, August)

**Schedule**: 12 weeks total (June-Aug), all camp types run simultaneously

| Age Group | Duration | Days | Type | Price | Max Students | Yr1 | Yr2 | Yr3 | Yr4 | Yr5 |
|-----------|----------|------|------|-------|--------------|-----|-----|-----|-----|-----|
| 5-7 | Half Day AM | 5 | General | $159 | 15 | 85% | 94% | 100% | 100% | 100% |
| 5-7 | Half Day AM | 5 | Ceramics/NW | $179 | 15 | 85% | 94% | 100% | 100% | 100% |
| 5-7 | Half Day PM | 4 | General | $135 | 15 | 85% | 94% | 100% | 100% | 100% |
| 5-7 | Full Day | 5 | General | $275 | 15 | 85% | 94% | 100% | 100% | 100% |
| 5-7 | Full Day | 5 | Ceramics/NW | $250 | 15 | 85% | 94% | 100% | 100% | 100% |
| 8-12 | Half Day AM | 5 | General | $159 | 15 | 85% | 94% | 100% | 100% | 100% |
| 8-12 | Half Day AM | 5 | Ceramics/NW | $179 | 15 | 85% | 94% | 100% | 100% | 100% |
| 8-12 | Half Day AM | 4 | General | $135 | 15 | 85% | 94% | 100% | 100% | 100% |
| 8-12 | Full Day | 5 | General | $275 | 15 | 85% | 94% | 100% | 100% | 100% |
| 8-12 | Full Day | 5 | Ceramics/NW | $295 | 15 | 85% | 94% | 100% | 100% | 100% |
| 8-12 | Full Day | 4 | General | $234 | 15 | 85% | 94% | 100% | 100% | 100% |
| 12-16 | Half Day | 5 | General | $159 | 15 | 85% | 94% | 100% | 100% | 100% |
| 12-16 | Half Day PM | 5 | Ceramics/W | $179 | 6 | 85% | 94% | 100% | 100% | 100% |
| 12-16 | Full Day | 5 | General | $275 | 15 | 85% | 94% | 100% | 100% | 100% |
| 12-16 | Full Day | 5 | Ceramics/W | $295 | 6 | 85% | 94% | 100% | 100% | 100% |
| All Ages | Extended Care | 5 | Add-on | $50 | 40 | 50% | 55% | 60% | 66% | 72% |

**Notes**:
- Ceramics/W = Wheel pottery lessons
- Ceramics/NW = Non-wheel (handbuilding, slab)
- Extended Care: 7:30-9am & 3-6pm, $50/week, max 40 students total across all camps
- **Enrollment capped at 100%** (ignore >100% values in source data)

**Revenue Calculation**: Price × Max Students × MIN(Enrollment %, 100%) × 12 weeks

**Market Comparison** (Tulsa area):
- Camp Philbrook: $260-$320/week full-day
- Kravis Arts Camp: $200/week
- Living Arts: $115/week
- **DECO pricing ($275-$295 full-day): Competitive ✓**

---

### 3. School Break Camps (Spring & Fall)

**Schedule**:
- Spring Break: 1 week (4 days) - typically March/April
- Fall Break: 1 week (4 days) - typically October/November

**Structure**: Use summer camp pricing and options, but with **50% enrollment** of summer camp percentages

**Example**:
- Summer camp showing 85% enrollment → School break = 42.5% enrollment
- Still capped at 100% maximum

---

### 4. Paint & Sip Events

| Parameter | Value | Adjustable? |
|-----------|-------|-------------|
| Events per month | 8 | Yes |
| Price per person | $45 | Yes |
| Attendees per event | 15 | Yes |
| Total monthly attendees | 120 | Auto-calc |

**Growth**: Apply same enrollment % growth as Regular Classes
- Year 1: 85% of capacity
- Year 2: 94%
- Year 3+: 100%

**Alcohol Revenue** (calculated separately):
- Drinks per person: 2 (adjustable)
- Participation rate: 75% (adjustable)
- Drink mix: 40% wine ($8), 40% beer ($6), 20% cocktails ($10)
- Weighted avg: $7.60/drink

**Market Research** (Tulsa):
- Pinot's Palette: $35-$45/person (drinks not included)
- **DECO pricing: $45 competitive ✓**

---

### 5. Private Events

| Parameter | Value | Adjustable? |
|-----------|-------|-------------|
| Events per month | 1 | Yes |
| Reservation fee | $250 | Yes |
| Price per person | $25 | Yes |
| Average attendees | 25 | Yes |

**Growth**: Apply same enrollment % growth as Regular Classes

**Alcohol Revenue** (calculated separately):
- Drinks per person: 2 (adjustable)
- Same drink mix as Paint & Sip

---

### 6. Workshops (Guest Artists)

| Parameter | Value | Adjustable? |
|-----------|-------|-------------|
| Workshops per month | 2 | Yes |
| Price per person | $75 | Yes |
| Attendees per workshop | 15 | Yes |
| Revenue split | 60% DECO / 40% Artist | Yes |

**Growth**: Constant (no enrollment % applied)

**Alcohol Revenue**:
- Drinks per person: 1.5 (adjustable)
- Same drink mix

---

### 7. Retail Revenue

**Calculation**: 5% of all other revenue (non-retail)

**COGS**: 50% (adjustable)

**Product Mix**: DECO branded items, local artisan goods (soaps, jewelry, clothing, stickers, art, books)

**Market Research**:
- Typical retail markup: 50-150% over wholesale
- COGS 50-60% for buy-wholesale-sell-retail model ✓

---

### 8. Beverage Revenue Summary

**Drink Pricing**:
- Wine: $8/glass
- Beer: $6/bottle
- Cocktails: $10/drink (average)
- Coffee/Tea/Soft drinks: Negligible

**Mix**: 40% wine, 40% beer, 20% cocktails = $7.60 weighted average

**COGS**: 30% (adjustable)

**Attached to**:
- Paint & Sip: 2 drinks/person, 75% participation
- Private Events: 2 drinks/person
- Workshops: 1.5 drinks/person

---

## Expenses

### Payroll & Labor

#### Owners
- Owner 1: $50,000/year starting Year 1 (adjustable)
- Owner 2: $50,000/year starting Year 1 (adjustable)

#### Studio Manager
- Year 1: 0 hours (owners manage)
- Year 2+: 40 hours/week @ $25/hr (adjustable)
- Annual cost (Year 2+): $52,000

#### Teachers
- Rate: $25/hr (adjustable)
- Hours: Auto-calculated based on:
  - Regular classes: Sum of all class durations
  - Summer camps: Half-day = 4 hrs, Full-day = 8 hrs per camp per week
  - School break camps: Same as summer
  - 1 teacher required per class/camp

#### Assistants
- Rate: $15/hr (adjustable)
- Hours: Auto-calculated for classes/camps with >10 enrolled students
- Same duration as class/camp

#### Payroll Taxes
- Rate: **9.75%** of total payroll (adjustable)
- Breakdown:
  - Federal FICA: 7.65%
  - Federal FUTA: 0.6%
  - Oklahoma SUI: 1.5%

---

### Occupancy Costs

| Expense | Calculation | Annual (5,000 sf) |
|---------|-------------|-------------------|
| Base Rent | $16/sf/month | $960,000 |
| NNN | $4/sf/month | $240,000 |
| **Total Rent** | | **$1,200,000** |
| Electricity | $1,200/month | $14,400 |
| Gas | $400/month | $4,800 |
| Water | $300/month | $3,600 |
| Internet | $200/month | $2,400 |
| **Utilities** | | **$25,200** |
| Insurance | $1,500/month | $18,000 |
| Property Tax | Annual | $19,500 |
| Repairs & Maintenance | Annual | $26,000 |

All values adjustable.

---

### Cost of Goods Sold (COGS)

| Category | Rate | Applied To |
|----------|------|------------|
| Art Supplies | 7.56% | Art class/camp revenue |
| Pottery Supplies | 2.02% | Pottery class/camp revenue |
| Beverages | 30% | Beverage revenue |
| Retail | 50% | Retail revenue |

**Supply Replenishment**:
- Initial supplies: $17,699
- Replenishment: 30% of initial = $5,310
- Frequency: Quarterly (4× per year)
- Annual total: $21,240

---

### Marketing

**Pre-Opening & Launch** (Months -2, -1, 0, 1, 2):
- $3,000/month for 5 months
- Total: $15,000

**Ongoing** (Month 3+):
- 3% of revenue (adjustable)

---

### Professional Services

| Service | Annual Cost | Adjustable? |
|---------|-------------|-------------|
| Accounting | $3,600 | Yes |
| Legal | $2,400 | Yes |
| Payroll Service | $1,800 | Yes |
| **Total** | **$7,800** | |

**Market Research** (Tulsa small business):
- Accounting: $200-400/month → Using $300/month ✓

---

### Licenses & Permits

**Liquor License** (Oklahoma Mixed Beverage):
- Initial (Year 1): $1,005
- Annual renewal (Year 2+): $905
- Model value: $1,030/year average (adjustable)

---

### Depreciation (Non-Cash Expense)

| Asset Category | Annual Amount | Method |
|----------------|---------------|---------|
| Building/Improvements | $33,333 | 39-year straight line |
| Equipment | $17,857 | 7-year straight line |
| **Total** | **$51,190** | |

---

## Startup Costs

### Equipment & Furnishings
- Kiln: $8,000
- Pottery Wheels (6): $4,800
- Classroom Furniture: $2,500
- Lobby Furniture: $2,500
- **Subtotal: $17,800**

### Initial Supplies
- Art Supplies (per detailed inventory): $17,699
- **Subtotal: $17,699**

### Setup & Professional
- Renovation/Build-out: $25,000
- Signage & Branding: $5,000
- Initial Marketing: $3,000
- Legal & Permits: $5,000
- POS System & Software: $3,000
- **Subtotal: $41,000**

### Pre-Opening Costs Subtotal
**$76,499** (all items adjustable)

### Operating Reserve
- **Auto-calculate**: 4 months of operating expenses (adjustable)
- Calculate based on Year 1 average monthly operating expenses
- Add to startup costs for total loan amount

### Total Project Cost
**$76,499 + Operating Reserve**

---

## SBA Loan Structure

| Parameter | Default | Adjustable? |
|-----------|---------|-------------|
| Loan Type | SBA 7(a) | N/A |
| Loan Amount | Auto-calculated | N/A |
| Interest Rate | 12% | Yes |
| Term | 10 years | Yes |
| Grace Period | 0 months | Yes |

**Loan Amount Calculation**:
```
Total Startup Costs ($76,499)
+ Operating Reserve (4 months of expenses)
= Total Loan Amount
```

**Monthly Payment Calculation**:
Standard amortization formula with monthly compounding

**Loan Servicing Start**:
- If grace period = 0: Start Month 1 (March 2026)
- If grace period > 0: Start after grace period

---

## Financial Statements Required

### 1. Profit & Loss (P&L)
**Monthly detail for 60 months** showing:

**Revenue**:
- Regular Classes
- Summer Camps
- School Break Camps
- Paint & Sip Events
- Private Events
- Workshops (DECO share only)
- Retail
- Beverage Sales
- **Total Revenue**

**Cost of Goods Sold**:
- Art Supplies
- Pottery Supplies
- Beverages
- Retail COGS
- Supply Replenishment
- **Total COGS**
- **Gross Profit**

**Operating Expenses**:
- Owner Salaries (2)
- Studio Manager
- Teachers
- Assistants
- Payroll Taxes
- Rent (Base + NNN)
- Utilities (Electricity, Gas, Water, Internet)
- Insurance
- Property Tax
- Repairs & Maintenance
- Marketing
- Accounting
- Legal
- Payroll Service
- Liquor License
- Loan Payment (Principal + Interest)
- **Total Operating Expenses**

**EBITDA**: Gross Profit - Operating Expenses (before depreciation, interest, taxes)

**Depreciation**

**Net Income** (before taxes)

---

### 2. Cash Flow Statement
**Monthly detail for 60 months** showing:

**Beginning Cash Balance**

**Cash Inflows**:
- All revenue (same as P&L)
- Loan proceeds (Month -2)

**Cash Outflows**:
- All operating expenses (same as P&L)
- Startup costs (Months -2, -1, 0)
- Loan principal payments
- **Exclude depreciation** (non-cash)

**Net Cash Flow**

**Ending Cash Balance**

---

### 3. Annual Summary Sheet
**5-year view** showing:
- Total Revenue by year
- Total Expenses by year
- Net Income by year
- EBITDA by year
- Cash Flow by year
- Ending Cash Balance by year

---

## Sheet Structure (To Be Determined)

**Option A**: 12-Sheet Structure (from previous versions)
1. Dashboard
2. Assumptions
3. Class Schedule
4. Summer Camps
5. Startup Supplies
6. Startup Costs
7. Revenue Model
8. Expenses
9. Loan Calculator
10. Loan Amortization
11. Profit & Loss
12. Cash Flow

**Option B**: Consolidated Structure
1. Dashboard / Summary
2. Assumptions (all adjustable inputs)
3. Revenue Detail (monthly, all streams)
4. Expense Detail (monthly, all categories)
5. Startup Costs
6. Loan Amortization
7. P&L (monthly, 60 months)
8. Cash Flow (monthly, 60 months)
9. Annual Summary (5-year)

**Question for User**: Which structure do you prefer, or do you have a different preference?

---

## Key Calculation Rules

### Revenue Recognition
- All revenue recognized in month services are provided
- No deferred revenue for multi-session classes

### Seasonality
- Regular classes: All 12 months
- Summer camps: June, July, August only (12 weeks total)
- School break camps: 1 week Spring (Mar/Apr), 1 week Fall (Oct/Nov)
- Paint & Sip: Year-round
- Private Events: Year-round
- Workshops: Year-round

### Growth Application
- Regular classes: Per class schedule enrollment %
- Camps: Per camp schedule enrollment %, capped at 100%
- Paint & Sip: Same % as regular classes
- Private Events: Same % as regular classes
- Workshops: No growth (constant)
- Retail: 5% of other revenue (grows as revenue grows)

### Staffing Calculations
- Teachers: Required hours based on actual enrolled students and class schedules
- Assistants: Only for classes/camps with >10 enrolled students
- Studio Manager: 0 in Year 1, 40 hrs/week in Year 2+

### Expense Timing
- Startup costs: Months -2, -1, 0
- Operating expenses: Begin Month 1 (except marketing starts Month -2)
- Loan payments: Begin based on grace period setting

---

## Data Validation & Error Prevention

1. **No circular references**: All calculations must be linear
2. **No merged cells**: Avoid formatting issues
3. **Consistent row tracking**: Use variables for row numbers
4. **Formula auditing**: Every cell with a formula should be traceable
5. **Percentage caps**: Enrollment capped at 100% where specified
6. **Negative value checks**: Cash, enrollment, prices cannot be negative
7. **Date continuity**: Month labels must be sequential

---

## Assumptions - All Adjustable Inputs

Every assumption listed in this document should be user-adjustable via an Assumptions sheet, including:

**Facility**:
- Square footage
- Rent $/sf/month
- NNN $/sf/month
- Utility rates

**Pricing**:
- All class prices
- All camp prices
- Event prices
- Drink prices

**Capacity & Enrollment**:
- Max students per class/camp
- Enrollment % by year
- Event attendance

**Labor Rates**:
- Owner salaries
- Manager hourly rate & hours
- Teacher hourly rate
- Assistant hourly rate
- Payroll tax rate

**COGS & Operating Expenses**:
- All COGS percentages
- All operating expense amounts
- Marketing budget and %
- Supply replenishment %

**Loan Terms**:
- Interest rate
- Term length
- Grace period
- Operating reserve months

---

## Success Criteria

✓ All 60 months of financial projections calculate without errors
✓ No #REF!, #VALUE!, #NUM!, or circular reference errors
✓ All assumptions are adjustable
✓ Loan amount auto-calculates based on startup costs + reserve
✓ Monthly loan payments amortize correctly over term
✓ Cash flow tracks accurately month-to-month
✓ Teacher/assistant hours calculate based on enrollment
✓ Revenue grows according to enrollment % by year
✓ All formulas are documented and traceable
✓ Model can be used for scenario planning (change assumptions, see results)

---

## Next Steps

1. **User reviews this requirements document**
2. **User confirms or adjusts any requirements**
3. **User decides on sheet structure (Option A, B, or custom)**
4. **Build Google Apps Script based on approved requirements**
5. **User tests script and reports any issues**
6. **Fix issues systematically based on test results**

---

*Requirements document created: 2025-11-21*
*All data sources: regular_class_schedule.csv, summer_camps_schedule.csv, startup_supplies.csv, assumptions.csv, startup_costs.csv*
