# DECO Art Center - 5-Year Pro Forma (Rental Model)

A comprehensive Google Sheets-based financial projection tool for securing an SBA loan for DECO Art Center in Tulsa, OK.

## 📋 Overview

This pro forma model is specifically designed for a **rental/lease scenario** and automatically calculates the SBA loan amount needed based on startup costs and operating capital requirements.

## 🚀 Installation

1. **Create a new Google Sheet**
   - Go to [Google Sheets](https://sheets.google.com)
   - Create a blank spreadsheet
   - Name it "DECO Art Center Pro Forma"

2. **Open Apps Script**
   - In your Google Sheet: `Extensions` → `Apps Script`
   - Delete any default code

3. **Copy the Script**
   - Open `DECOProForma_Rental.gs`
   - Copy ALL the code
   - Paste into the Apps Script editor

4. **Save & Run**
   - Click the Save icon (💾)
   - Select the function `createDECOProForma` from the dropdown
   - Click "Run" (▶️)
   - **Authorize** the script when prompted (this is safe - you're authorizing your own script)
   - Wait 60-90 seconds for completion

5. **Success!**
   - You'll see a confirmation message
   - Return to your Google Sheet
   - You'll now have 12 tabs with your complete pro forma

## 📊 Key Features

### Rental Model
- **No building purchase** - lease-based model
- **Adjustable square footage** (default: 5,000 sq ft)
- **Adjustable rent** (default: $16/sf/year)
- **NNN breakdown** for Tulsa, OK (property tax, insurance, CAM)
- **5-year lease** with 3% annual escalation
- **Move-in costs** included (first, last, security deposit)

### Automatic Loan Calculator
- **Auto-calculates** total SBA loan needed
- Includes startup costs + 4 months operating reserve
- **Adjustable** interest rate (default: 12%)
- **10-year term** for working capital loan
- Shows monthly payment, total interest, and DSCR

### Revenue Projections
- **Regular classes** (10 different class types)
- **Summer camps** (16 different camp configurations)
- **Events & workshops** (Paint & Sip, private events, artist workshops)
- **Beverage sales** (coffee, beer, wine)
- **Retail sales** (5% of total revenue)
- **Conservative growth** assumptions (50% capacity Year 1, 20% annual growth)

### Comprehensive Financials
- Profit & Loss Statement (5 years)
- Cash Flow Projection (5 years)
- Startup Costs breakdown
- Operating Expenses (adjusted for Tulsa, OK)
- Loan Amortization Schedule
- Executive Dashboard

## 🎯 How to Use

### Step 1: Start with the Dashboard
The **Dashboard** tab gives you an executive summary:
- SBA loan amount needed
- Monthly payment
- 5-year revenue/profit projections
- Key financial ratios (including Debt Service Coverage Ratio)

### Step 2: Review & Adjust Assumptions
The **Assumptions** tab is your control panel. All inputs are editable:

#### Facility
- Building size (sq ft)
- Base rent ($/sf/year)
- NNN breakdown
- Lease terms

#### Operations
- Class capacity and growth rates
- Staff compensation
- Event pricing and frequency
- Operating expenses

#### Financing
- Interest rate
- Loan term
- Owner equity (if any)

### Step 3: Customize Your Offerings
- **Class Schedule**: Edit class types, prices, capacity
- **Summer Camps**: Adjust camp offerings and pricing
- **Startup Supplies**: Modify initial supply needs

### Step 4: Review Financial Statements
- **Revenue Model**: See detailed revenue projections by source
- **Expenses**: Review all operating costs
- **Profit & Loss**: Analyze profitability over 5 years
- **Cash Flow**: Understand cash position
- **Loan Calculator**: See total funding needed

## 💡 Key Assumptions (Baseline)

### Facility
- **Location**: Tulsa, OK 74119
- **Size**: 5,000 sq ft
- **Rent**: $16/sf/year ($6,667/month)
- **NNN**: $4/sf/year ($1,667/month)
- **Total Monthly**: $8,333

### Startup Costs
- Leasehold improvements: $25,000
- Equipment (kiln, wheels, furniture): $17,800
- Art supplies: ~$20,000
- Professional services: $16,000
- Move-in costs: $25,000
- **4-month operating reserve**

### Loan Terms
- **SBA 7(a) Loan**
- Interest rate: 12%
- Term: 10 years
- Amount: **Auto-calculated** based on startup + working capital

### Operations
- **Opening**: March 2026
- **Year 1**: 10 months operation
- **Capacity**: 50% Year 1, growing 20% annually
- **Staffing**: 2 owners + studio manager + teachers/assistants

### Revenue Streams
- Regular classes (10 types)
- Summer camps (16 configurations)
- Paint & Sip events (8/month @ $45)
- Private events (1/month)
- Workshops (1/month @ $150)
- Beverages & retail

## 📈 What Makes This Model Strong for Lenders

1. **Conservative Assumptions**
   - 50% capacity in Year 1
   - Realistic pricing based on Tulsa market
   - 4-month cash reserve

2. **Multiple Revenue Streams**
   - Not dependent on single source
   - Year-round income (classes + camps + events)

3. **Detailed Financial Planning**
   - Every expense category accounted for
   - Tulsa-specific costs (utilities, insurance, taxes)
   - Realistic staffing model

4. **Strong Financial Metrics**
   - Debt Service Coverage Ratio (DSCR) shown
   - EBITDA and Net Income projections
   - Cash flow positive trajectory

5. **Fully Adjustable**
   - Easy to run scenarios
   - All assumptions clearly documented
   - Professional presentation

## 🔧 Customization Guide

### To Change Rent
1. Go to **Assumptions** tab
2. Edit `B5` (Base Rent $/sf/year)
3. Edit `B10-B12` (NNN breakdown)
4. All calculations update automatically

### To Adjust Loan Amount
1. Go to **Loan Calculator** tab
2. Edit `B6` (Owner Equity)
3. Edit `B7` (Other Funding)
4. Loan amount recalculates automatically

### To Modify Class Offerings
1. Go to **Class Schedule** tab
2. Edit prices, capacity, or sessions
3. Add new rows for additional classes
4. Revenue model updates automatically

### To Change Growth Assumptions
1. Go to **Assumptions** tab
2. Edit `B40` (Year 1 Capacity %)
3. Edit `B41` (Annual Growth %)
4. All years update automatically

## 📁 File Structure

```
DECO/
├── DECOProForma_Rental.gs    # Google Apps Script (main file)
└── README.md                  # This file
```

## 🎓 Understanding the Financials

### What is DSCR (Debt Service Coverage Ratio)?
- **DSCR = EBITDA ÷ Annual Debt Payment**
- Lenders want to see **DSCR > 1.25**
- Shows you can cover loan payments with operating income
- Your model shows this ratio for all 5 years

### What is EBITDA?
- **Earnings Before Interest, Taxes, Depreciation, Amortization**
- Shows operating profitability
- Key metric lenders use to assess business health

### Why 4 Months Reserve?
- Covers operating expenses during slow periods
- Shows lenders you've planned for contingencies
- SBA loans typically require 3-6 months working capital

## ⚠️ Important Notes

1. **Run Only Once**: The script creates all tabs. If you need to recreate, you can run again - it will update existing sheets.

2. **Don't Delete Sheets**: All sheets are interconnected with formulas. Deleting one will break others.

3. **Yellow/Green Cells**: In the Assumptions tab, white cells are inputs (edit these), blue cells are formulas (auto-calculate).

4. **Year 1 = 10 Months**: Assumes March 2026 opening, so only 10 months of operation in Year 1.

5. **Adjust for Your Situation**: This is a baseline model. Customize all assumptions to match your actual plans.

## 🆘 Troubleshooting

### Script won't run?
- Make sure you're running `createDECOProForma` function
- Check that you've authorized the script
- Try refreshing the page and running again

### Formulas showing errors?
- Don't rename sheets (breaks cross-sheet references)
- Check that all 12 sheets were created
- Try running the script again to recreate

### Numbers seem wrong?
- Check the **Assumptions** tab - all inputs are there
- Verify you've customized for your specific situation
- Review **Class Schedule** and **Summer Camps** tabs

## 📞 Questions?

This model is designed to be comprehensive yet flexible. Every assumption can be adjusted to match your specific situation.

**Key Tabs to Review with Your Lender:**
1. **Dashboard** - Executive summary
2. **Loan Calculator** - Total funding request
3. **Profit & Loss** - 5-year profitability
4. **Cash Flow** - Cash position over time
5. **Assumptions** - All underlying assumptions

## 📄 License

This is a custom financial model created for DECO Art Center. Feel free to adapt for your use.

---

**Created**: November 2025
**Version**: 1.0 - Rental Model
**For**: SBA Loan Application - DECO Art Center, Tulsa, OK
