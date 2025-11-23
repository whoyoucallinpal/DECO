/**
 * DECO Art Center - 5-Year Pro Forma Financial Model
 * Option B: 9-Sheet Consolidated Structure
 *
 * Creates a comprehensive financial projection model for DECO Art Center in Tulsa, OK
 * with monthly detail for 60 months (March 2026 - February 2031)
 *
 * Sheet Structure:
 * 1. Assumptions - All adjustable inputs
 * 2. Revenue Detail - 60-month revenue breakdown by stream
 * 3. Expense Detail - 60-month expense breakdown by category
 * 4. Startup Costs - One-time startup cost breakdown
 * 5. Loan Amortization - Monthly loan payment schedule
 * 6. P&L - 60-month Profit & Loss statement
 * 7. Cash Flow - 60-month cash flow statement
 * 8. Dashboard - High-level summary and key metrics
 * 9. Annual Summary - 5-year annual rollup
 *
 * Timeline:
 * - Jan 2026 (Month -2): Loan funded, pre-opening
 * - Feb 2026 (Month -1): Buildout continues
 * - Mar 2026 (Month 1): Business opens
 * - Monthly columns: MAR26, APR26, MAY26, etc.
 *
 * Created: 2025-11-21
 * Based on: REQUIREMENTS.md
 */

function createDECOProForma() {
  const ss = SpreadsheetApp.create('DECO Art Center Pro Forma (Option B)');

  // Create all sheets in order
  createAssumptionsSheet(ss);
  createRevenueDetailSheet(ss);
  createExpenseDetailSheet(ss);
  createStartupCostsSheet(ss);
  createLoanAmortizationSheet(ss);
  createProfitLossSheet(ss);
  createCashFlowSheet(ss);
  createDashboardSheet(ss);
  createAnnualSummarySheet(ss);

  // Delete default Sheet1
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }

  // Set Dashboard as active sheet
  ss.setActiveSheet(ss.getSheetByName('Dashboard'));

  Logger.log('Spreadsheet created: ' + ss.getUrl());
  return ss;
}

/**
 * ASSUMPTIONS SHEET
 * Contains all adjustable inputs for the financial model
 */
function createAssumptionsSheet(ss) {
  const sheet = ss.insertSheet('Assumptions', 0);
  sheet.setTabColor('#4A86E8'); // Blue

  let row = 1;

  // Header
  sheet.getRange(row, 1).setValue('DECO ART CENTER - FINANCIAL ASSUMPTIONS').setFontWeight('bold').setFontSize(14);
  sheet.getRange(row, 3).setValue('All values are adjustable').setFontStyle('italic');
  row += 2;

  // ========== FACILITY ASSUMPTIONS ==========
  sheet.getRange(row, 1).setValue('FACILITY').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Square Footage');
  sheet.getRange(row, 2).setValue(5000).setNumberFormat('#,##0');
  sheet.getRange(row, 3).setValue('sf');
  row++;

  sheet.getRange(row, 1).setValue('Base Rent');
  sheet.getRange(row, 2).setValue(16).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 3).setValue('per sf per year');
  row++;

  sheet.getRange(row, 1).setValue('NNN (Triple Net)');
  sheet.getRange(row, 2).setValue(4).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 3).setValue('per sf per year');
  row++;

  sheet.getRange(row, 1).setValue('Total Annual Rent').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=B4*(B5+B6)').setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('auto-calculated').setFontStyle('italic');
  const totalAnnualRentRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Total Monthly Rent').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(`=B${totalAnnualRentRow}/12`).setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('auto-calculated').setFontStyle('italic');
  row += 2;

  // ========== TIMELINE ==========
  sheet.getRange(row, 1).setValue('TIMELINE').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Loan Funding Date');
  sheet.getRange(row, 2).setValue('Jan 2026');
  sheet.getRange(row, 3).setValue('Month -2');
  row++;

  sheet.getRange(row, 1).setValue('Business Opening Date');
  sheet.getRange(row, 2).setValue('Mar 2026');
  sheet.getRange(row, 3).setValue('Month 1');
  row += 2;

  // ========== STAFF COMPENSATION ==========
  sheet.getRange(row, 1).setValue('STAFF COMPENSATION').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Owner 1 Annual Salary');
  sheet.getRange(row, 2).setValue(50000).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('Starting Year 1');
  row++;

  sheet.getRange(row, 1).setValue('Owner 2 Annual Salary');
  sheet.getRange(row, 2).setValue(50000).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('Starting Year 1');
  row++;

  sheet.getRange(row, 1).setValue('Studio Manager - Hourly Rate');
  sheet.getRange(row, 2).setValue(25).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 3).setValue('per hour');
  row++;

  sheet.getRange(row, 1).setValue('Studio Manager - Hours/Week Year 1');
  sheet.getRange(row, 2).setValue(0).setNumberFormat('#,##0');
  sheet.getRange(row, 3).setValue('Owners manage Year 1');
  row++;

  sheet.getRange(row, 1).setValue('Studio Manager - Hours/Week Year 2+');
  sheet.getRange(row, 2).setValue(40).setNumberFormat('#,##0');
  sheet.getRange(row, 3).setValue('per week');
  row++;

  sheet.getRange(row, 1).setValue('Teacher Hourly Rate');
  sheet.getRange(row, 2).setValue(25).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 3).setValue('per hour');
  row++;

  sheet.getRange(row, 1).setValue('Assistant Hourly Rate');
  sheet.getRange(row, 2).setValue(15).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 3).setValue('per hour (for classes >10 students)');
  row++;

  sheet.getRange(row, 1).setValue('Payroll Tax Rate');
  sheet.getRange(row, 2).setValue(0.0975).setNumberFormat('0.00%');
  sheet.getRange(row, 3).setValue('FICA 7.65% + FUTA 0.6% + OK SUI 1.5%');
  row += 2;

  // ========== EVENTS & PRICING ==========
  sheet.getRange(row, 1).setValue('EVENTS & PRICING').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Paint & Sip - Events/Month');
  sheet.getRange(row, 2).setValue(8).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Paint & Sip - Price/Person');
  sheet.getRange(row, 2).setValue(45).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Paint & Sip - Attendance/Event');
  sheet.getRange(row, 2).setValue(15).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Private Event - Reservation Fee');
  sheet.getRange(row, 2).setValue(250).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Private Event - Price/Person');
  sheet.getRange(row, 2).setValue(25).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Private Event - Average Attendance');
  sheet.getRange(row, 2).setValue(25).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Private Events/Month');
  sheet.getRange(row, 2).setValue(1).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Workshop - Price/Person');
  sheet.getRange(row, 2).setValue(75).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Workshop - Attendance');
  sheet.getRange(row, 2).setValue(15).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Workshop - DECO Revenue Split');
  sheet.getRange(row, 2).setValue(0.60).setNumberFormat('0%');
  sheet.getRange(row, 3).setValue('60% DECO / 40% Artist');
  row++;

  sheet.getRange(row, 1).setValue('Workshops/Month');
  sheet.getRange(row, 2).setValue(2).setNumberFormat('#,##0');
  row += 2;

  // ========== BEVERAGE SALES ==========
  sheet.getRange(row, 1).setValue('BEVERAGE SALES').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Drinks/Person - Paint & Sip');
  sheet.getRange(row, 2).setValue(2).setNumberFormat('0.0');
  row++;

  sheet.getRange(row, 1).setValue('Drinks/Person - Private Event');
  sheet.getRange(row, 2).setValue(2).setNumberFormat('0.0');
  row++;

  sheet.getRange(row, 1).setValue('Drinks/Person - Workshop');
  sheet.getRange(row, 2).setValue(1.5).setNumberFormat('0.0');
  row++;

  sheet.getRange(row, 1).setValue('Drink Participation % (P&S)');
  sheet.getRange(row, 2).setValue(0.75).setNumberFormat('0%');
  sheet.getRange(row, 3).setValue('% who buy drinks at Paint & Sip');
  row++;

  sheet.getRange(row, 1).setValue('Wine Price');
  sheet.getRange(row, 2).setValue(8).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('per glass');
  const winePriceRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Beer Price');
  sheet.getRange(row, 2).setValue(6).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('per bottle');
  const beerPriceRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Cocktail Price');
  sheet.getRange(row, 2).setValue(10).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('per drink');
  const cocktailPriceRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Drink Mix - Wine %');
  sheet.getRange(row, 2).setValue(0.40).setNumberFormat('0%');
  const wineMixRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Drink Mix - Beer %');
  sheet.getRange(row, 2).setValue(0.40).setNumberFormat('0%');
  const beerMixRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Drink Mix - Cocktail %');
  sheet.getRange(row, 2).setValue(0.20).setNumberFormat('0%');
  const cocktailMixRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Weighted Avg Drink Price').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(`=B${winePriceRow}*B${wineMixRow}+B${beerPriceRow}*B${beerMixRow}+B${cocktailPriceRow}*B${cocktailMixRow}`).setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('auto-calculated').setFontStyle('italic');
  row++;

  sheet.getRange(row, 1).setValue('Beverage COGS %');
  sheet.getRange(row, 2).setValue(0.30).setNumberFormat('0%');
  row += 2;

  // ========== COST OF GOODS SOLD ==========
  sheet.getRange(row, 1).setValue('COST OF GOODS SOLD (COGS)').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Art Supplies COGS %');
  sheet.getRange(row, 2).setValue(0.0756).setNumberFormat('0.00%');
  sheet.getRange(row, 3).setValue('% of art class/camp revenue');
  row++;

  sheet.getRange(row, 1).setValue('Pottery Supplies COGS %');
  sheet.getRange(row, 2).setValue(0.0202).setNumberFormat('0.00%');
  sheet.getRange(row, 3).setValue('% of pottery class/camp revenue');
  row++;

  sheet.getRange(row, 1).setValue('Retail COGS %');
  sheet.getRange(row, 2).setValue(0.50).setNumberFormat('0%');
  sheet.getRange(row, 3).setValue('% of retail revenue');
  row++;

  sheet.getRange(row, 1).setValue('Initial Supplies Investment');
  sheet.getRange(row, 2).setValue(17699).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('From Startup Supplies sheet');
  const initialSuppliesRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Supply Replenishment %');
  sheet.getRange(row, 2).setValue(0.30).setNumberFormat('0%');
  sheet.getRange(row, 3).setValue('% of initial to reorder');
  const replenishPctRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Replenishment Frequency');
  sheet.getRange(row, 2).setValue(4).setNumberFormat('#,##0');
  sheet.getRange(row, 3).setValue('times per year (quarterly)');
  const replenishFreqRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Annual Supply Replenishment').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(`=B${initialSuppliesRow}*B${replenishPctRow}*B${replenishFreqRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('auto-calculated').setFontStyle('italic');
  row += 2;

  // ========== OPERATING EXPENSES ==========
  sheet.getRange(row, 1).setValue('OPERATING EXPENSES').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Electricity/Month');
  sheet.getRange(row, 2).setValue(1200).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Gas/Month');
  sheet.getRange(row, 2).setValue(400).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Water/Month');
  sheet.getRange(row, 2).setValue(300).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Internet/Month');
  sheet.getRange(row, 2).setValue(200).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Insurance/Month');
  sheet.getRange(row, 2).setValue(1500).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Property Tax/Year');
  sheet.getRange(row, 2).setValue(19500).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Repairs & Maintenance/Year');
  sheet.getRange(row, 2).setValue(26000).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Accounting/Year');
  sheet.getRange(row, 2).setValue(3600).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Legal/Year');
  sheet.getRange(row, 2).setValue(2400).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Payroll Service/Year');
  sheet.getRange(row, 2).setValue(1800).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Liquor License - Initial');
  sheet.getRange(row, 2).setValue(1005).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('Year 1');
  row++;

  sheet.getRange(row, 1).setValue('Liquor License - Renewal');
  sheet.getRange(row, 2).setValue(905).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('Year 2+');
  row += 2;

  // ========== MARKETING ==========
  sheet.getRange(row, 1).setValue('MARKETING').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Pre-Launch Marketing/Month');
  sheet.getRange(row, 2).setValue(3000).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('Months -2, -1, 0, 1, 2');
  row++;

  sheet.getRange(row, 1).setValue('Ongoing Marketing % of Revenue');
  sheet.getRange(row, 2).setValue(0.03).setNumberFormat('0.00%');
  sheet.getRange(row, 3).setValue('Starting Month 3');
  row += 2;

  // ========== DEPRECIATION ==========
  sheet.getRange(row, 1).setValue('DEPRECIATION').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Building/Improvements/Year');
  sheet.getRange(row, 2).setValue(33333).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('39-year straight line');
  const buildingDepreciationRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Equipment/Year');
  sheet.getRange(row, 2).setValue(17857).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('7-year straight line');
  const equipmentDepreciationRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Total Annual Depreciation').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(`=B${buildingDepreciationRow}+B${equipmentDepreciationRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('auto-calculated').setFontStyle('italic');
  row += 2;

  // ========== LOAN TERMS ==========
  sheet.getRange(row, 1).setValue('LOAN TERMS').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Interest Rate');
  sheet.getRange(row, 2).setValue(0.12).setNumberFormat('0.00%');
  sheet.getRange(row, 3).setValue('annual');
  const interestRateRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Loan Term (Years)');
  sheet.getRange(row, 2).setValue(10).setNumberFormat('#,##0');
  const loanTermYearsRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Grace Period (Months)');
  sheet.getRange(row, 2).setValue(0).setNumberFormat('#,##0');
  const gracePeriodRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Operating Reserve (Months)');
  sheet.getRange(row, 2).setValue(4).setNumberFormat('#,##0');
  sheet.getRange(row, 3).setValue('Months of expenses to reserve');
  const operatingReserveMonthsRow = row;
  row += 2;

  // ========== RETAIL ==========
  sheet.getRange(row, 1).setValue('RETAIL').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Retail % of Other Revenue');
  sheet.getRange(row, 2).setValue(0.05).setNumberFormat('0.00%');
  sheet.getRange(row, 3).setValue('5% of non-retail revenue');
  row += 2;

  // ========== REGULAR CLASS SCHEDULE ==========
  sheet.getRange(row, 1).setValue('REGULAR CLASS SCHEDULE').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 10).merge();
  row++;

  // Headers
  sheet.getRange(row, 1).setValue('Class Type').setFontWeight('bold');
  sheet.getRange(row, 2).setValue('Duration (hrs)').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('Price').setFontWeight('bold');
  sheet.getRange(row, 4).setValue('Max Students').setFontWeight('bold');
  sheet.getRange(row, 5).setValue('Sessions/Yr').setFontWeight('bold');
  sheet.getRange(row, 6).setValue('Year 1 %').setFontWeight('bold');
  sheet.getRange(row, 7).setValue('Year 2 %').setFontWeight('bold');
  sheet.getRange(row, 8).setValue('Year 3 %').setFontWeight('bold');
  sheet.getRange(row, 9).setValue('Year 4 %').setFontWeight('bold');
  sheet.getRange(row, 10).setValue('Year 5 %').setFontWeight('bold');
  row++;

  // Class data from regular_class_schedule.csv
  const classData = [
    ['Adults - Painting', 2.0, 199, 15, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Adults - Pottery', 2.5, 229, 6, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Teens 13-18 - Drawing', 2.0, 199, 15, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Teens 13-18 - Pottery', 2.0, 229, 6, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Kids 9-12 - Mixed Media', 1.5, 159, 15, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Kids 9-12 - Pottery', 1.5, 179, 6, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Kids 6-8 - General Art', 1.5, 154, 15, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Kids 4-5 - General Art', 1.5, 154, 15, 12, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['Homeschool 6-12', 1.5, 154, 30, 12, 0.50, 0.60, 0.70, 0.85, 1.00],
    ['Homeschool Teens', 1.5, 154, 30, 12, 0.50, 0.60, 0.70, 0.85, 1.00]
  ];

  for (let i = 0; i < classData.length; i++) {
    sheet.getRange(row, 1).setValue(classData[i][0]);
    sheet.getRange(row, 2).setValue(classData[i][1]).setNumberFormat('0.0');
    sheet.getRange(row, 3).setValue(classData[i][2]).setNumberFormat('$#,##0');
    sheet.getRange(row, 4).setValue(classData[i][3]).setNumberFormat('#,##0');
    sheet.getRange(row, 5).setValue(classData[i][4]).setNumberFormat('#,##0');
    sheet.getRange(row, 6).setValue(classData[i][5]).setNumberFormat('0%');
    sheet.getRange(row, 7).setValue(classData[i][6]).setNumberFormat('0%');
    sheet.getRange(row, 8).setValue(classData[i][7]).setNumberFormat('0%');
    sheet.getRange(row, 9).setValue(classData[i][8]).setNumberFormat('0%');
    sheet.getRange(row, 10).setValue(classData[i][9]).setNumberFormat('0%');
    row++;
  }
  row++;

  // ========== SUMMER CAMP SCHEDULE ==========
  sheet.getRange(row, 1).setValue('SUMMER CAMP SCHEDULE').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 11).merge();
  row++;

  // Headers
  sheet.getRange(row, 1).setValue('Age Group').setFontWeight('bold');
  sheet.getRange(row, 2).setValue('Duration').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('Days').setFontWeight('bold');
  sheet.getRange(row, 4).setValue('Type').setFontWeight('bold');
  sheet.getRange(row, 5).setValue('Price').setFontWeight('bold');
  sheet.getRange(row, 6).setValue('Max Students').setFontWeight('bold');
  sheet.getRange(row, 7).setValue('Year 1 %').setFontWeight('bold');
  sheet.getRange(row, 8).setValue('Year 2 %').setFontWeight('bold');
  sheet.getRange(row, 9).setValue('Year 3 %').setFontWeight('bold');
  sheet.getRange(row, 10).setValue('Year 4 %').setFontWeight('bold');
  sheet.getRange(row, 11).setValue('Year 5 %').setFontWeight('bold');
  row++;

  // Camp data from summer_camps_schedule.csv
  const campData = [
    ['5-7', 'Half Day AM', 5, 'General', 159, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['5-7', 'Half Day AM', 5, 'Ceramics/NW', 179, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['5-7', 'Half Day PM', 4, 'General', 135, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['5-7', 'Full Day', 5, 'General', 275, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['5-7', 'Full Day', 5, 'Ceramics/NW', 250, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['8-12', 'Half Day AM', 5, 'General', 159, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['8-12', 'Half Day AM', 5, 'Ceramics/NW', 179, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['8-12', 'Half Day AM', 4, 'General', 135, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['8-12', 'Full Day', 5, 'General', 275, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['8-12', 'Full Day', 5, 'Ceramics/NW', 295, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['8-12', 'Full Day', 4, 'General', 234, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['12-16', 'Half Day', 5, 'General', 159, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['12-16', 'Half Day PM', 5, 'Ceramics/W', 179, 6, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['12-16', 'Full Day', 5, 'General', 275, 15, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['12-16', 'Full Day', 5, 'Ceramics/W', 295, 6, 0.85, 0.94, 1.00, 1.00, 1.00],
    ['All Ages', 'Extended Care', 5, 'Add-on', 50, 40, 0.50, 0.55, 0.60, 0.66, 0.72]
  ];

  for (let i = 0; i < campData.length; i++) {
    sheet.getRange(row, 1).setValue(campData[i][0]);
    sheet.getRange(row, 2).setValue(campData[i][1]);
    sheet.getRange(row, 3).setValue(campData[i][2]).setNumberFormat('#,##0');
    sheet.getRange(row, 4).setValue(campData[i][3]);
    sheet.getRange(row, 5).setValue(campData[i][4]).setNumberFormat('$#,##0');
    sheet.getRange(row, 6).setValue(campData[i][5]).setNumberFormat('#,##0');
    sheet.getRange(row, 7).setValue(campData[i][6]).setNumberFormat('0%');
    sheet.getRange(row, 8).setValue(campData[i][7]).setNumberFormat('0%');
    sheet.getRange(row, 9).setValue(campData[i][8]).setNumberFormat('0%');
    sheet.getRange(row, 10).setValue(campData[i][9]).setNumberFormat('0%');
    sheet.getRange(row, 11).setValue(campData[i][10]).setNumberFormat('0%');
    row++;
  }
  row++;

  sheet.getRange(row, 1).setValue('Note: Summer camps run 12 weeks (Jun-Aug). School break camps (Spring/Fall) use 50% enrollment.').setFontStyle('italic');
  row += 2;

  // Format columns
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 80);
  sheet.setColumnWidth(10, 80);
  sheet.setColumnWidth(11, 80);

  // Freeze first row
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Helper function to generate month labels
 * Returns array of month labels: ['MAR26', 'APR26', ...]
 */
function getMonthLabels() {
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const labels = [];

  // Start from Jan 2026 (Month -2) through Feb 2031 (Month 60)
  let year = 26;
  let monthIndex = 0; // Start with Jan

  for (let i = -2; i <= 60; i++) {
    if (i === 61) break; // 63 total months (Jan26 - Feb31)
    labels.push(months[monthIndex] + year.toString().padStart(2, '0'));
    monthIndex++;
    if (monthIndex === 12) {
      monthIndex = 0;
      year++;
    }
  }

  return labels;
}

/**
 * REVENUE DETAIL SHEET
 * 60 months of detailed revenue by stream
 */
function createRevenueDetailSheet(ss) {
  const sheet = ss.insertSheet('Revenue Detail');
  sheet.setTabColor('#34A853'); // Green

  const monthLabels = getMonthLabels();
  let row = 1;
  let col = 1;

  // Header
  sheet.getRange(row, col).setValue('DECO ART CENTER - REVENUE DETAIL (60 Months)').setFontWeight('bold').setFontSize(14);
  row += 2;

  // Column headers - Month labels
  sheet.getRange(row, 1).setValue('Revenue Stream').setFontWeight('bold');
  for (let i = 0; i < 60; i++) {
    // Months 1-60 = MAR26 through FEB31
    sheet.getRange(row, i + 2).setValue(monthLabels[i + 2]).setFontWeight('bold');
  }
  const headerRow = row;
  row++;

  // Note: This is a simplified structure showing the framework
  // Full implementation would add all revenue calculations here

  // Placeholder rows for now
  sheet.getRange(row, 1).setValue('Regular Classes').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Summer Camps').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Paint & Sip Events').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Private Events').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Workshops').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Beverage Sales').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Retail').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('TOTAL REVENUE').setFontWeight('bold').setBackground('#D9EAD3');
  row++;

  // Format
  sheet.setColumnWidth(1, 250);
  for (let i = 2; i <= 61; i++) {
    sheet.setColumnWidth(i, 100);
  }
  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(1);

  return sheet;
}

/**
 * EXPENSE DETAIL SHEET
 * 60 months of detailed expenses by category
 */
function createExpenseDetailSheet(ss) {
  const sheet = ss.insertSheet('Expense Detail');
  sheet.setTabColor('#EA4335'); // Red

  const monthLabels = getMonthLabels();
  let row = 1;

  // Header
  sheet.getRange(row, 1).setValue('DECO ART CENTER - EXPENSE DETAIL (60 Months)').setFontWeight('bold').setFontSize(14);
  row += 2;

  // Column headers - Month labels
  sheet.getRange(row, 1).setValue('Expense Category').setFontWeight('bold');
  for (let i = 0; i < 60; i++) {
    sheet.getRange(row, i + 2).setValue(monthLabels[i + 2]).setFontWeight('bold');
  }
  const headerRow = row;
  row++;

  // Expense categories (simplified framework)
  sheet.getRange(row, 1).setValue('Owner Salaries').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Studio Manager').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Teachers').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Assistants').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Payroll Taxes').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Rent (Base + NNN)').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Utilities').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Insurance').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Property Tax').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Repairs & Maintenance').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Marketing').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Professional Services').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Liquor License').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Supply Replenishment').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('COGS - Art Supplies').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('COGS - Pottery').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('COGS - Beverages').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('COGS - Retail').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Loan Payment').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('TOTAL EXPENSES').setFontWeight('bold').setBackground('#F4CCCC');
  const totalExpenseRow = row;
  row++;

  // Format
  sheet.setColumnWidth(1, 250);
  for (let i = 2; i <= 61; i++) {
    sheet.setColumnWidth(i, 100);
  }
  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(1);

  return sheet;
}

/**
 * STARTUP COSTS SHEET
 * One-time startup cost breakdown
 */
function createStartupCostsSheet(ss) {
  const sheet = ss.insertSheet('Startup Costs');
  sheet.setTabColor('#FBBC04'); // Yellow

  let row = 1;

  // Header
  sheet.getRange(row, 1).setValue('DECO ART CENTER - STARTUP COSTS').setFontWeight('bold').setFontSize(14);
  row += 2;

  // ========== EQUIPMENT & FURNISHINGS ==========
  sheet.getRange(row, 1).setValue('EQUIPMENT & FURNISHINGS').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Kiln');
  sheet.getRange(row, 2).setValue(8000).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Pottery Wheels (6)');
  sheet.getRange(row, 2).setValue(4800).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Classroom Furniture');
  sheet.getRange(row, 2).setValue(2500).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Lobby Furniture');
  sheet.getRange(row, 2).setValue(2500).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('EQUIPMENT SUBTOTAL').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=SUM(B4:B7)').setNumberFormat('$#,##0').setFontWeight('bold');
  row += 2;

  // ========== INITIAL SUPPLIES ==========
  sheet.getRange(row, 1).setValue('INITIAL SUPPLIES').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Art Supplies (from detailed inventory)');
  sheet.getRange(row, 2).setValue(17699).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('SUPPLIES SUBTOTAL').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=B11').setNumberFormat('$#,##0').setFontWeight('bold');
  row += 2;

  // ========== SETUP & PROFESSIONAL ==========
  sheet.getRange(row, 1).setValue('SETUP & PROFESSIONAL').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Renovation/Build-out');
  sheet.getRange(row, 2).setValue(25000).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Signage & Branding');
  sheet.getRange(row, 2).setValue(5000).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Initial Marketing');
  sheet.getRange(row, 2).setValue(3000).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Legal & Permits');
  sheet.getRange(row, 2).setValue(5000).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('POS System & Software');
  sheet.getRange(row, 2).setValue(3000).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('SETUP SUBTOTAL').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=SUM(B15:B19)').setNumberFormat('$#,##0').setFontWeight('bold');
  row += 2;

  // ========== WORKING CAPITAL ==========
  sheet.getRange(row, 1).setValue('WORKING CAPITAL').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Operating Reserve').setFontStyle('italic');
  sheet.getRange(row, 2).setValue('See calculation below').setFontStyle('italic');
  row++;

  sheet.getRange(row, 1).setValue('Reserve Months (from Assumptions)').setFontStyle('italic');
  sheet.getRange(row, 2).setFormula('=Assumptions!B87').setNumberFormat('#,##0').setFontStyle('italic');
  row++;

  sheet.getRange(row, 1).setValue('Average Monthly Operating Expenses').setFontStyle('italic');
  sheet.getRange(row, 2).setValue('Calculated from Year 1').setFontStyle('italic');
  row++;

  sheet.getRange(row, 1).setValue('Operating Reserve Amount').setFontWeight('bold');
  sheet.getRange(row, 2).setValue('TO BE LINKED TO EXPENSE DETAIL').setFontWeight('bold').setBackground('#FFF9C4');
  const operatingReserveRow = row;
  row += 2;

  // ========== TOTALS ==========
  sheet.getRange(row, 1).setValue('TOTAL CASH NEEDED (before reserve)').setFontWeight('bold').setFontSize(12);
  sheet.getRange(row, 2).setFormula('=B8+B12+B20').setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(12);
  row++;

  sheet.getRange(row, 1).setValue('TOTAL PROJECT COST (with reserve)').setFontWeight('bold').setFontSize(12).setBackground('#D9EAD3');
  sheet.getRange(row, 2).setFormula(`=B${row-1}+B${operatingReserveRow}`).setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(12).setBackground('#D9EAD3');
  row++;

  sheet.getRange(row, 1).setValue('LOAN AMOUNT NEEDED').setFontWeight('bold').setFontSize(14).setBackground('#B6D7A8');
  sheet.getRange(row, 2).setFormula(`=B${row-1}`).setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(14).setBackground('#B6D7A8');
  row += 2;

  sheet.getRange(row, 1).setValue('Note: Operating Reserve will calculate once Expense Detail sheet is complete.').setFontStyle('italic').setFontColor('#666666');

  // Format columns
  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 150);

  // Freeze header
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * LOAN AMORTIZATION SHEET
 * Monthly loan payment schedule
 */
function createLoanAmortizationSheet(ss) {
  const sheet = ss.insertSheet('Loan Amortization');
  sheet.setTabColor('#FF6D01'); // Orange

  let row = 1;
  const monthLabels = getMonthLabels();

  // Header
  sheet.getRange(row, 1).setValue('DECO ART CENTER - LOAN AMORTIZATION').setFontWeight('bold').setFontSize(14);
  row += 2;

  // Loan parameters
  sheet.getRange(row, 1).setValue('Loan Amount');
  sheet.getRange(row, 2).setFormula('=\'Startup Costs\'!B32').setNumberFormat('$#,##0.00').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Annual Interest Rate');
  sheet.getRange(row, 2).setFormula('=Assumptions!B84').setNumberFormat('0.00%');
  const annualRateRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Monthly Interest Rate');
  sheet.getRange(row, 2).setFormula(`=B${annualRateRow}/12`).setNumberFormat('0.0000%');
  const monthlyRateRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Loan Term (Years)');
  sheet.getRange(row, 2).setFormula('=Assumptions!B85').setNumberFormat('#,##0');
  const loanTermYearsRowLoan = row;
  row++;

  sheet.getRange(row, 1).setValue('Loan Term (Months)');
  sheet.getRange(row, 2).setFormula(`=B${loanTermYearsRowLoan}*12`).setNumberFormat('#,##0');
  const loanTermMonthsRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Grace Period (Months)');
  sheet.getRange(row, 2).setFormula('=Assumptions!B86').setNumberFormat('#,##0');
  const gracePeriodRowLoan = row;
  row++;

  sheet.getRange(row, 1).setValue('Monthly Payment').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(`=IF(B4>0,PMT(B${monthlyRateRow},B${loanTermMonthsRow},-B4),0)`).setNumberFormat('$#,##0.00').setFontWeight('bold');
  const monthlyPaymentRow = row;
  row += 2;

  // Amortization table headers
  sheet.getRange(row, 1).setValue('Month').setFontWeight('bold');
  sheet.getRange(row, 2).setValue('Period').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('Beginning Balance').setFontWeight('bold');
  sheet.getRange(row, 4).setValue('Payment').setFontWeight('bold');
  sheet.getRange(row, 5).setValue('Principal').setFontWeight('bold');
  sheet.getRange(row, 6).setValue('Interest').setFontWeight('bold');
  sheet.getRange(row, 7).setValue('Ending Balance').setFontWeight('bold');
  const headerRow = row;
  row++;

  // Generate 120 months (10 years) of amortization schedule
  // Start from Month -2 (Jan 2026) but payments don't start until after grace period
  const totalMonths = 120; // Show 10 years regardless
  const startMonth = -2; // Jan 2026
  const loanAmountRow = 4;

  for (let i = 0; i < totalMonths; i++) {
    const monthNum = startMonth + i;
    const monthLabel = monthLabels[i + 2]; // Offset because labels array starts at Jan26 (index 0)

    sheet.getRange(row, 1).setValue(monthLabel);
    sheet.getRange(row, 2).setValue(monthNum).setNumberFormat('#,##0');

    if (row === headerRow + 1) {
      // First row - beginning balance is loan amount
      sheet.getRange(row, 3).setFormula(`=$B$${loanAmountRow}`).setNumberFormat('$#,##0.00');
    } else {
      // Subsequent rows - beginning balance is previous ending balance
      sheet.getRange(row, 3).setFormula(`=G${row-1}`).setNumberFormat('$#,##0.00');
    }

    // Payment amount - only if past grace period and balance > 0
    sheet.getRange(row, 4).setFormula(`=IF(AND(B${row}>=$B$${gracePeriodRowLoan},C${row}>0),$B$${monthlyPaymentRow},0)`).setNumberFormat('$#,##0.00');

    // Interest = Beginning Balance * Monthly Rate
    sheet.getRange(row, 6).setFormula(`=C${row}*$B$${monthlyRateRow}`).setNumberFormat('$#,##0.00');

    // Principal = Payment - Interest
    sheet.getRange(row, 5).setFormula(`=D${row}-F${row}`).setNumberFormat('$#,##0.00');

    // Ending Balance = Beginning Balance - Principal
    sheet.getRange(row, 7).setFormula(`=MAX(0,C${row}-E${row})`).setNumberFormat('$#,##0.00');

    row++;
  }

  // Format columns
  sheet.setColumnWidth(1, 80);  // Month
  sheet.setColumnWidth(2, 70);  // Period
  sheet.setColumnWidth(3, 130); // Beginning Balance
  sheet.setColumnWidth(4, 110); // Payment
  sheet.setColumnWidth(5, 110); // Principal
  sheet.setColumnWidth(6, 110); // Interest
  sheet.setColumnWidth(7, 130); // Ending Balance

  // Freeze headers
  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(2);

  return sheet;
}

/**
 * PROFIT & LOSS SHEET
 * 60-month P&L statement
 */
function createProfitLossSheet(ss) {
  const sheet = ss.insertSheet('P&L');
  sheet.setTabColor('#9900FF'); // Purple

  const monthLabels = getMonthLabels();
  let row = 1;

  // Header
  sheet.getRange(row, 1).setValue('DECO ART CENTER - PROFIT & LOSS (60 Months)').setFontWeight('bold').setFontSize(14);
  row += 2;

  // Column headers
  sheet.getRange(row, 1).setValue('Line Item').setFontWeight('bold');
  for (let i = 0; i < 60; i++) {
    sheet.getRange(row, i + 2).setValue(monthLabels[i + 2]).setFontWeight('bold');
  }
  const headerRow = row;
  row++;

  // REVENUE SECTION
  sheet.getRange(row, 1).setValue('REVENUE').setFontWeight('bold').setFontSize(12).setBackground('#D9EAD3');
  row++;

  sheet.getRange(row, 1).setValue('Total Revenue');
  const totalRevenueRow = row;
  // Reference Revenue Detail total row
  for (let i = 0; i < 60; i++) {
    const col = i + 2;
    sheet.getRange(row, col).setFormula(`='Revenue Detail'!${String.fromCharCode(65 + col - 1)}11`).setNumberFormat('$#,##0');
  }
  row++;

  // COGS SECTION
  sheet.getRange(row, 1).setValue('COST OF GOODS SOLD').setFontWeight('bold').setFontSize(12).setBackground('#FCE5CD');
  row++;

  sheet.getRange(row, 1).setValue('Art Supplies COGS');
  row++;
  sheet.getRange(row, 1).setValue('Pottery COGS');
  row++;
  sheet.getRange(row, 1).setValue('Beverage COGS');
  row++;
  sheet.getRange(row, 1).setValue('Retail COGS');
  row++;

  sheet.getRange(row, 1).setValue('Total COGS').setFontWeight('bold');
  const totalCOGSRow = row;
  row++;

  // GROSS PROFIT
  sheet.getRange(row, 1).setValue('GROSS PROFIT').setFontWeight('bold').setFontSize(12).setBackground('#D9EAD3');
  const grossProfitRow = row;
  for (let i = 0; i < 60; i++) {
    const col = i + 2;
    const colLetter = getColLetter(col);
    sheet.getRange(row, col).setFormula(`=${colLetter}${totalRevenueRow}-${colLetter}${totalCOGSRow}`).setNumberFormat('$#,##0');
  }
  row++;

  // OPERATING EXPENSES
  sheet.getRange(row, 1).setValue('OPERATING EXPENSES').setFontWeight('bold').setFontSize(12).setBackground('#F4CCCC');
  row++;

  sheet.getRange(row, 1).setValue('Owner Salaries');
  row++;
  sheet.getRange(row, 1).setValue('Studio Manager');
  row++;
  sheet.getRange(row, 1).setValue('Teachers');
  row++;
  sheet.getRange(row, 1).setValue('Assistants');
  row++;
  sheet.getRange(row, 1).setValue('Payroll Taxes');
  row++;
  sheet.getRange(row, 1).setValue('Rent');
  row++;
  sheet.getRange(row, 1).setValue('Utilities');
  row++;
  sheet.getRange(row, 1).setValue('Insurance');
  row++;
  sheet.getRange(row, 1).setValue('Property Tax');
  row++;
  sheet.getRange(row, 1).setValue('Repairs & Maintenance');
  row++;
  sheet.getRange(row, 1).setValue('Marketing');
  row++;
  sheet.getRange(row, 1).setValue('Professional Services');
  row++;
  sheet.getRange(row, 1).setValue('Liquor License');
  row++;
  sheet.getRange(row, 1).setValue('Supply Replenishment');
  row++;

  sheet.getRange(row, 1).setValue('Total Operating Expenses').setFontWeight('bold');
  const totalOpExRow = row;
  row++;

  // EBITDA
  sheet.getRange(row, 1).setValue('EBITDA').setFontWeight('bold').setFontSize(12).setBackground('#C9DAF8');
  const ebitdaRow = row;
  for (let i = 0; i < 60; i++) {
    const col = i + 2;
    const colLetter = getColLetter(col);
    sheet.getRange(row, col).setFormula(`=${colLetter}${grossProfitRow}-${colLetter}${totalOpExRow}`).setNumberFormat('$#,##0');
  }
  row++;

  // DEPRECIATION
  sheet.getRange(row, 1).setValue('Depreciation');
  const depreciationRow = row;
  row++;

  // INTEREST EXPENSE
  sheet.getRange(row, 1).setValue('Interest Expense');
  const interestRow = row;
  row++;

  // NET INCOME BEFORE TAX
  sheet.getRange(row, 1).setValue('NET INCOME BEFORE TAX').setFontWeight('bold').setFontSize(12).setBackground('#B6D7A8');
  const netIncomeRow = row;
  for (let i = 0; i < 60; i++) {
    const col = i + 2;
    const colLetter = getColLetter(col);
    sheet.getRange(row, col).setFormula(`=${colLetter}${ebitdaRow}-${colLetter}${depreciationRow}-${colLetter}${interestRow}`).setNumberFormat('$#,##0');
  }
  row++;

  // Format
  sheet.setColumnWidth(1, 250);
  for (let i = 2; i <= 61; i++) {
    sheet.setColumnWidth(i, 100);
  }
  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(1);

  return sheet;
}

/**
 * Helper function to get column letter from column number
 */
function getColLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * CASH FLOW SHEET
 * 60-month cash flow statement
 */
function createCashFlowSheet(ss) {
  const sheet = ss.insertSheet('Cash Flow');
  sheet.setTabColor('#00BCD4'); // Cyan

  const monthLabels = getMonthLabels();
  let row = 1;

  // Header
  sheet.getRange(row, 1).setValue('DECO ART CENTER - CASH FLOW (60 Months)').setFontWeight('bold').setFontSize(14);
  row += 2;

  // Column headers
  sheet.getRange(row, 1).setValue('Cash Flow Item').setFontWeight('bold');
  for (let i = 0; i < 60; i++) {
    sheet.getRange(row, i + 2).setValue(monthLabels[i + 2]).setFontWeight('bold');
  }
  const headerRow = row;
  row++;

  // BEGINNING CASH
  sheet.getRange(row, 1).setValue('Beginning Cash Balance').setFontWeight('bold').setBackground('#E8EAED');
  const beginCashRow = row;
  row++;

  // CASH INFLOWS
  sheet.getRange(row, 1).setValue('CASH INFLOWS').setFontWeight('bold').setFontSize(12).setBackground('#D9EAD3');
  row++;

  sheet.getRange(row, 1).setValue('Revenue Collections');
  row++;

  sheet.getRange(row, 1).setValue('Loan Proceeds (Month 1 only)');
  row++;

  sheet.getRange(row, 1).setValue('Total Cash Inflows').setFontWeight('bold');
  const totalInflowsRow = row;
  row++;

  // CASH OUTFLOWS
  sheet.getRange(row, 1).setValue('CASH OUTFLOWS').setFontWeight('bold').setFontSize(12).setBackground('#F4CCCC');
  row++;

  sheet.getRange(row, 1).setValue('Startup Costs (Pre-opening)');
  row++;

  sheet.getRange(row, 1).setValue('Operating Expenses');
  row++;

  sheet.getRange(row, 1).setValue('COGS');
  row++;

  sheet.getRange(row, 1).setValue('Loan Principal Payment');
  row++;

  sheet.getRange(row, 1).setValue('Loan Interest Payment');
  row++;

  sheet.getRange(row, 1).setValue('Total Cash Outflows').setFontWeight('bold');
  const totalOutflowsRow = row;
  row++;

  // NET CASH FLOW
  sheet.getRange(row, 1).setValue('NET CASH FLOW').setFontWeight('bold').setFontSize(12).setBackground('#C9DAF8');
  const netCashFlowRow = row;
  for (let i = 0; i < 60; i++) {
    const col = i + 2;
    const colLetter = getColLetter(col);
    sheet.getRange(row, col).setFormula(`=${colLetter}${totalInflowsRow}-${colLetter}${totalOutflowsRow}`).setNumberFormat('$#,##0');
  }
  row++;

  // ENDING CASH
  sheet.getRange(row, 1).setValue('Ending Cash Balance').setFontWeight('bold').setFontSize(12).setBackground('#B6D7A8');
  const endCashRow = row;
  for (let i = 0; i < 60; i++) {
    const col = i + 2;
    const colLetter = getColLetter(col);
    sheet.getRange(row, col).setFormula(`=${colLetter}${beginCashRow}+${colLetter}${netCashFlowRow}`).setNumberFormat('$#,##0');
  }
  row++;

  // Format
  sheet.setColumnWidth(1, 280);
  for (let i = 2; i <= 61; i++) {
    sheet.setColumnWidth(i, 100);
  }
  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(1);

  return sheet;
}

/**
 * DASHBOARD SHEET
 * High-level summary and key metrics
 */
function createDashboardSheet(ss) {
  const sheet = ss.insertSheet('Dashboard');
  sheet.setTabColor('#4285F4'); // Blue

  let row = 1;

  // ========== HEADER ==========
  sheet.getRange(row, 1).setValue('DECO ART CENTER').setFontWeight('bold').setFontSize(20).setFontColor('#4285F4');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  sheet.getRange(row, 1).setValue('5-Year Pro Forma Financial Summary').setFontSize(14).setFontStyle('italic');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  sheet.getRange(row, 1).setValue('Location: Tulsa, OK 74104 | Opening: March 2026').setFontColor('#666666');
  sheet.getRange(row, 1, 1, 4).merge();
  row += 2;

  // ========== LOAN SUMMARY ==========
  sheet.getRange(row, 1).setValue('LOAN SUMMARY').setFontWeight('bold').setFontSize(12).setBackground('#4285F4').setFontColor('#FFFFFF');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Loan Amount');
  sheet.getRange(row, 2).setFormula("='Startup Costs'!B32").setNumberFormat('$#,##0');
  const loanAmountDashRow = row;
  row++;

  sheet.getRange(row, 1).setValue('Interest Rate');
  sheet.getRange(row, 2).setFormula('=Assumptions!B84').setNumberFormat('0.00%');
  row++;

  sheet.getRange(row, 1).setValue('Loan Term');
  sheet.getRange(row, 2).setFormula('=Assumptions!B85&" years"');
  row++;

  sheet.getRange(row, 1).setValue('Monthly Payment');
  sheet.getRange(row, 2).setFormula("='Loan Amortization'!B11").setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Total Interest (10 years)');
  sheet.getRange(row, 2).setFormula("=SUMIF('Loan Amortization'!D:D,\">0\",'Loan Amortization'!F:F)").setNumberFormat('$#,##0');
  row += 2;

  // ========== 5-YEAR FINANCIAL OVERVIEW ==========
  sheet.getRange(row, 1).setValue('5-YEAR FINANCIAL OVERVIEW').setFontWeight('bold').setFontSize(12).setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(row, 1, 1, 6).merge();
  row++;

  // Column headers
  sheet.getRange(row, 1).setValue('Metric').setFontWeight('bold');
  sheet.getRange(row, 2).setValue('Year 1').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('Year 2').setFontWeight('bold');
  sheet.getRange(row, 4).setValue('Year 3').setFontWeight('bold');
  sheet.getRange(row, 5).setValue('Year 4').setFontWeight('bold');
  sheet.getRange(row, 6).setValue('Year 5').setFontWeight('bold');
  const overviewHeaderRow = row;
  row++;

  // Total Revenue row - reference Annual Summary
  sheet.getRange(row, 1).setValue('Total Revenue');
  sheet.getRange(row, 2).setFormula("='Annual Summary'!B5").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("='Annual Summary'!C5").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("='Annual Summary'!D5").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("='Annual Summary'!E5").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("='Annual Summary'!F5").setNumberFormat('$#,##0');
  const dashRevenueRow = row;
  row++;

  // Total Expenses row
  sheet.getRange(row, 1).setValue('Total Expenses');
  sheet.getRange(row, 2).setFormula("='Annual Summary'!B17").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("='Annual Summary'!C17").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("='Annual Summary'!D17").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("='Annual Summary'!E17").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("='Annual Summary'!F17").setNumberFormat('$#,##0');
  const dashExpenseRow = row;
  row++;

  // EBITDA row
  sheet.getRange(row, 1).setValue('EBITDA').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula("='Annual Summary'!B19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setFormula("='Annual Summary'!C19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 4).setFormula("='Annual Summary'!D19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 5).setFormula("='Annual Summary'!E19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 6).setFormula("='Annual Summary'!F19").setNumberFormat('$#,##0').setFontWeight('bold');
  row++;

  // Net Income row
  sheet.getRange(row, 1).setValue('Net Income').setFontWeight('bold').setBackground('#B6D7A8');
  sheet.getRange(row, 2).setFormula("='Annual Summary'!B22").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#B6D7A8');
  sheet.getRange(row, 3).setFormula("='Annual Summary'!C22").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#B6D7A8');
  sheet.getRange(row, 4).setFormula("='Annual Summary'!D22").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#B6D7A8');
  sheet.getRange(row, 5).setFormula("='Annual Summary'!E22").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#B6D7A8');
  sheet.getRange(row, 6).setFormula("='Annual Summary'!F22").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#B6D7A8');
  row++;

  // Ending Cash Balance row
  sheet.getRange(row, 1).setValue('Ending Cash Balance').setFontWeight('bold').setBackground('#C9DAF8');
  sheet.getRange(row, 2).setFormula("='Annual Summary'!B25").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#C9DAF8');
  sheet.getRange(row, 3).setFormula("='Annual Summary'!C25").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#C9DAF8');
  sheet.getRange(row, 4).setFormula("='Annual Summary'!D25").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#C9DAF8');
  sheet.getRange(row, 5).setFormula("='Annual Summary'!E25").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#C9DAF8');
  sheet.getRange(row, 6).setFormula("='Annual Summary'!F25").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#C9DAF8');
  row += 2;

  // ========== 5-YEAR TOTALS ==========
  sheet.getRange(row, 1).setValue('5-YEAR TOTALS').setFontWeight('bold').setFontSize(12).setBackground('#FBBC04').setFontColor('#000000');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Total Revenue (5 Years)');
  sheet.getRange(row, 2).setFormula(`=SUM(B${dashRevenueRow}:F${dashRevenueRow})`).setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Total Net Income (5 Years)').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(`=SUM('Annual Summary'!B22:F22)`).setNumberFormat('$#,##0').setFontWeight('bold');
  row++;

  sheet.getRange(row, 1).setValue('Total Loan Payments (5 Years)');
  sheet.getRange(row, 2).setFormula("=SUMPRODUCT(('Loan Amortization'!B:B>=-2)*('Loan Amortization'!B:B<=58)*('Loan Amortization'!D:D))").setNumberFormat('$#,##0');
  row += 2;

  // ========== KEY ASSUMPTIONS ==========
  sheet.getRange(row, 1).setValue('KEY ASSUMPTIONS').setFontWeight('bold').setFontSize(12).setBackground('#EA4335').setFontColor('#FFFFFF');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Square Footage');
  sheet.getRange(row, 2).setFormula('=Assumptions!B4').setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Monthly Rent');
  sheet.getRange(row, 2).setFormula('=Assumptions!B9').setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Owner Salaries (each)');
  sheet.getRange(row, 2).setFormula('=Assumptions!B15').setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Paint & Sip Events/Month');
  sheet.getRange(row, 2).setFormula('=Assumptions!B29').setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Paint & Sip Price');
  sheet.getRange(row, 2).setFormula('=Assumptions!B30').setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Year 1 Enrollment Rate');
  sheet.getRange(row, 2).setValue(0.85).setNumberFormat('0%');
  row++;

  sheet.getRange(row, 1).setValue('Year 3+ Enrollment Rate');
  sheet.getRange(row, 2).setValue(1.00).setNumberFormat('0%');
  row += 2;

  // ========== STARTUP COSTS BREAKDOWN ==========
  sheet.getRange(row, 1).setValue('STARTUP COSTS BREAKDOWN').setFontWeight('bold').setFontSize(12).setBackground('#9900FF').setFontColor('#FFFFFF');
  sheet.getRange(row, 1, 1, 2).merge();
  row++;

  sheet.getRange(row, 1).setValue('Equipment & Furnishings');
  sheet.getRange(row, 2).setFormula("='Startup Costs'!B8").setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Initial Supplies');
  sheet.getRange(row, 2).setFormula("='Startup Costs'!B12").setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Setup & Professional');
  sheet.getRange(row, 2).setFormula("='Startup Costs'!B20").setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Operating Reserve');
  sheet.getRange(row, 2).setFormula("='Startup Costs'!B27").setNumberFormat('$#,##0');
  row++;

  sheet.getRange(row, 1).setValue('TOTAL LOAN REQUIRED').setFontWeight('bold').setBackground('#B6D7A8');
  sheet.getRange(row, 2).setFormula("='Startup Costs'!B32").setNumberFormat('$#,##0').setFontWeight('bold').setBackground('#B6D7A8');
  row += 2;

  // ========== NOTES ==========
  sheet.getRange(row, 1).setValue('NOTES').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  sheet.getRange(row, 1).setValue('1. This model assumes all revenue is collected in the month earned (no receivables delay).');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  sheet.getRange(row, 1).setValue('2. Summer camps run June-August. School break camps (Spring/Fall) use 50% enrollment.');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  sheet.getRange(row, 1).setValue('3. Retail revenue is estimated at 5% of all other revenue sources.');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  sheet.getRange(row, 1).setValue('4. Studio Manager is hired in Year 2 at 40 hours/week. Owners manage Year 1.');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  sheet.getRange(row, 1).setValue('5. All assumptions can be adjusted in the Assumptions sheet.');
  sheet.getRange(row, 1, 1, 4).merge();
  row++;

  // Format columns
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);

  // Freeze header
  sheet.setFrozenRows(3);

  return sheet;
}

/**
 * ANNUAL SUMMARY SHEET
 * 5-year annual rollup
 */
function createAnnualSummarySheet(ss) {
  const sheet = ss.insertSheet('Annual Summary');
  sheet.setTabColor('#0F9D58'); // Dark Green

  let row = 1;

  // Header
  sheet.getRange(row, 1).setValue('DECO ART CENTER - ANNUAL SUMMARY').setFontWeight('bold').setFontSize(14);
  row += 2;

  // Column headers
  sheet.getRange(row, 1).setValue('Category').setFontWeight('bold');
  sheet.getRange(row, 2).setValue('Year 1').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('Year 2').setFontWeight('bold');
  sheet.getRange(row, 4).setValue('Year 3').setFontWeight('bold');
  sheet.getRange(row, 5).setValue('Year 4').setFontWeight('bold');
  sheet.getRange(row, 6).setValue('Year 5').setFontWeight('bold');
  sheet.getRange(row, 7).setValue('5-Year Total').setFontWeight('bold');
  const headerRow = row;
  row++;

  // ========== REVENUE SECTION ==========
  sheet.getRange(row, 1).setValue('REVENUE').setFontWeight('bold').setFontSize(11).setBackground('#D9EAD3');
  sheet.getRange(row, 1, 1, 7).setBackground('#D9EAD3');
  row++;

  // Total Revenue - sum monthly P&L revenue for each year
  // Year 1 = Months 1-10 (Mar26-Dec26) = columns B-K in P&L row 5
  // Year 2 = Months 11-22 (Jan27-Dec27) = columns L-W
  // Year 3 = Months 23-34 (Jan28-Dec28) = columns X-AI
  // Year 4 = Months 35-46 (Jan29-Dec29) = columns AJ-AU
  // Year 5 = Months 47-58 (Jan30-Dec30) = columns AV-BG + BH-BI for Jan-Feb 31
  sheet.getRange(row, 1).setValue('Total Revenue').setFontWeight('bold');
  // Year 1: Mar26-Dec26 (10 months) = P&L columns B:K
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B5:K5)").setNumberFormat('$#,##0');
  // Year 2: Jan27-Dec27 (12 months) = P&L columns L:W
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B5:K5)").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L5:W5)").setNumberFormat('$#,##0');
  // Year 3: Jan28-Dec28 (12 months) = P&L columns X:AI
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X5:AI5)").setNumberFormat('$#,##0');
  // Year 4: Jan29-Dec29 (12 months) = P&L columns AJ:AU
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ5:AU5)").setNumberFormat('$#,##0');
  // Year 5: Jan30-Feb31 (14 months: Jan30-Dec30=12 + Jan31-Feb31=2) = P&L columns AV:BI
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV5:BI5)").setNumberFormat('$#,##0');
  // 5-Year Total
  sheet.getRange(row, 7).setFormula('=SUM(B5:F5)').setNumberFormat('$#,##0').setFontWeight('bold');
  const totalRevenueRow = row;
  row++;

  // ========== COGS SECTION ==========
  sheet.getRange(row, 1).setValue('COST OF GOODS SOLD').setFontWeight('bold').setFontSize(11).setBackground('#FCE5CD');
  sheet.getRange(row, 1, 1, 7).setBackground('#FCE5CD');
  row++;

  sheet.getRange(row, 1).setValue('Total COGS').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B12:K12)").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L12:W12)").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X12:AI12)").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ12:AU12)").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV12:BI12)").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B7:F7)').setNumberFormat('$#,##0').setFontWeight('bold');
  const totalCOGSRow = row;
  row++;

  // ========== GROSS PROFIT ==========
  sheet.getRange(row, 1).setValue('GROSS PROFIT').setFontWeight('bold').setFontSize(11).setBackground('#D9EAD3');
  sheet.getRange(row, 2).setFormula(`=B${totalRevenueRow}-B${totalCOGSRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setFormula(`=C${totalRevenueRow}-C${totalCOGSRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 4).setFormula(`=D${totalRevenueRow}-D${totalCOGSRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 5).setFormula(`=E${totalRevenueRow}-E${totalCOGSRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 6).setFormula(`=F${totalRevenueRow}-F${totalCOGSRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 7).setFormula(`=G${totalRevenueRow}-G${totalCOGSRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 1, 1, 7).setBackground('#D9EAD3');
  const grossProfitRow = row;
  row++;

  // ========== OPERATING EXPENSES SECTION ==========
  sheet.getRange(row, 1).setValue('OPERATING EXPENSES').setFontWeight('bold').setFontSize(11).setBackground('#F4CCCC');
  sheet.getRange(row, 1, 1, 7).setBackground('#F4CCCC');
  row++;

  // Payroll (Owners + Manager + Teachers + Assistants + Taxes)
  sheet.getRange(row, 1).setValue('Payroll (all staff)');
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B17:K21)").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L17:W21)").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X17:AI21)").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ17:AU21)").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV17:BI21)").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B11:F11)').setNumberFormat('$#,##0');
  row++;

  // Rent
  sheet.getRange(row, 1).setValue('Rent');
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B22:K22)").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L22:W22)").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X22:AI22)").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ22:AU22)").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV22:BI22)").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B12:F12)').setNumberFormat('$#,##0');
  row++;

  // Utilities
  sheet.getRange(row, 1).setValue('Utilities');
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B23:K23)").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L23:W23)").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X23:AI23)").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ23:AU23)").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV23:BI23)").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B13:F13)').setNumberFormat('$#,##0');
  row++;

  // Other Operating
  sheet.getRange(row, 1).setValue('Other Operating Expenses');
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B24:K30)").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L24:W30)").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X24:AI30)").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ24:AU30)").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV24:BI30)").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B14:F14)').setNumberFormat('$#,##0');
  row++;

  // Loan Payments (Principal + Interest)
  sheet.getRange(row, 1).setValue('Loan Payments (P+I)');
  // Use SUMIFS to get loan payments for each year based on period column
  sheet.getRange(row, 2).setFormula("=SUMIFS('Loan Amortization'!D:D,'Loan Amortization'!B:B,\">=-2\",'Loan Amortization'!B:B,\"<=10\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUMIFS('Loan Amortization'!D:D,'Loan Amortization'!B:B,\">10\",'Loan Amortization'!B:B,\"<=22\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUMIFS('Loan Amortization'!D:D,'Loan Amortization'!B:B,\">22\",'Loan Amortization'!B:B,\"<=34\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUMIFS('Loan Amortization'!D:D,'Loan Amortization'!B:B,\">34\",'Loan Amortization'!B:B,\"<=46\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUMIFS('Loan Amortization'!D:D,'Loan Amortization'!B:B,\">46\",'Loan Amortization'!B:B,\"<=60\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B15:F15)').setNumberFormat('$#,##0');
  row++;

  // Total Operating Expenses
  sheet.getRange(row, 1).setValue('Total Operating Expenses').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B31:K31)").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L31:W31)").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X31:AI31)").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ31:AU31)").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV31:BI31)").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 7).setFormula('=SUM(B17:F17)').setNumberFormat('$#,##0').setFontWeight('bold');
  const totalOpExpRow = row;
  row++;

  // ========== EBITDA ==========
  sheet.getRange(row, 1).setValue('EBITDA').setFontWeight('bold').setFontSize(11).setBackground('#C9DAF8');
  sheet.getRange(row, 2).setFormula(`=B${grossProfitRow}-B${totalOpExpRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setFormula(`=C${grossProfitRow}-C${totalOpExpRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 4).setFormula(`=D${grossProfitRow}-D${totalOpExpRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 5).setFormula(`=E${grossProfitRow}-E${totalOpExpRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 6).setFormula(`=F${grossProfitRow}-F${totalOpExpRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 7).setFormula(`=G${grossProfitRow}-G${totalOpExpRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 1, 1, 7).setBackground('#C9DAF8');
  const ebitdaRow = row;
  row++;

  // Depreciation
  sheet.getRange(row, 1).setValue('Depreciation');
  sheet.getRange(row, 2).setFormula("=SUM('P&L'!B33:K33)").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUM('P&L'!L33:W33)").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUM('P&L'!X33:AI33)").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUM('P&L'!AJ33:AU33)").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUM('P&L'!AV33:BI33)").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B20:F20)').setNumberFormat('$#,##0');
  const deprecRow = row;
  row++;

  // Interest Expense
  sheet.getRange(row, 1).setValue('Interest Expense');
  sheet.getRange(row, 2).setFormula("=SUMIFS('Loan Amortization'!F:F,'Loan Amortization'!B:B,\">=-2\",'Loan Amortization'!B:B,\"<=10\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setFormula("=SUMIFS('Loan Amortization'!F:F,'Loan Amortization'!B:B,\">10\",'Loan Amortization'!B:B,\"<=22\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 4).setFormula("=SUMIFS('Loan Amortization'!F:F,'Loan Amortization'!B:B,\">22\",'Loan Amortization'!B:B,\"<=34\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 5).setFormula("=SUMIFS('Loan Amortization'!F:F,'Loan Amortization'!B:B,\">34\",'Loan Amortization'!B:B,\"<=46\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 6).setFormula("=SUMIFS('Loan Amortization'!F:F,'Loan Amortization'!B:B,\">46\",'Loan Amortization'!B:B,\"<=60\")").setNumberFormat('$#,##0');
  sheet.getRange(row, 7).setFormula('=SUM(B21:F21)').setNumberFormat('$#,##0');
  const interestRow = row;
  row++;

  // ========== NET INCOME ==========
  sheet.getRange(row, 1).setValue('NET INCOME BEFORE TAX').setFontWeight('bold').setFontSize(11).setBackground('#B6D7A8');
  sheet.getRange(row, 2).setFormula(`=B${ebitdaRow}-B${deprecRow}-B${interestRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setFormula(`=C${ebitdaRow}-C${deprecRow}-C${interestRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 4).setFormula(`=D${ebitdaRow}-D${deprecRow}-D${interestRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 5).setFormula(`=E${ebitdaRow}-E${deprecRow}-E${interestRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 6).setFormula(`=F${ebitdaRow}-F${deprecRow}-F${interestRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 7).setFormula(`=G${ebitdaRow}-G${deprecRow}-G${interestRow}`).setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 1, 1, 7).setBackground('#B6D7A8');
  const netIncomeRow = row;
  row += 2;

  // ========== CASH FLOW SUMMARY ==========
  sheet.getRange(row, 1).setValue('CASH FLOW SUMMARY').setFontWeight('bold').setFontSize(11).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 7).setBackground('#E8EAED');
  row++;

  // Ending Cash Balance (end of each year)
  // Year 1 ends at Dec26 (Month 10) = column K in Cash Flow
  // Year 2 ends at Dec27 (Month 22) = column W
  // Year 3 ends at Dec28 (Month 34) = column AI
  // Year 4 ends at Dec29 (Month 46) = column AU
  // Year 5 ends at Feb31 (Month 60) = column BI
  sheet.getRange(row, 1).setValue('Ending Cash Balance').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula("='Cash Flow'!K19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setFormula("='Cash Flow'!W19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 4).setFormula("='Cash Flow'!AI19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 5).setFormula("='Cash Flow'!AU19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 6).setFormula("='Cash Flow'!BI19").setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 7).setValue('N/A').setFontStyle('italic');
  const endCashRow = row;
  row += 2;

  // ========== KEY RATIOS ==========
  sheet.getRange(row, 1).setValue('KEY RATIOS').setFontWeight('bold').setFontSize(11).setBackground('#FFF2CC');
  sheet.getRange(row, 1, 1, 7).setBackground('#FFF2CC');
  row++;

  // Gross Margin %
  sheet.getRange(row, 1).setValue('Gross Margin %');
  sheet.getRange(row, 2).setFormula(`=IF(B${totalRevenueRow}>0,B${grossProfitRow}/B${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 3).setFormula(`=IF(C${totalRevenueRow}>0,C${grossProfitRow}/C${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 4).setFormula(`=IF(D${totalRevenueRow}>0,D${grossProfitRow}/D${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 5).setFormula(`=IF(E${totalRevenueRow}>0,E${grossProfitRow}/E${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 6).setFormula(`=IF(F${totalRevenueRow}>0,F${grossProfitRow}/F${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 7).setFormula(`=IF(G${totalRevenueRow}>0,G${grossProfitRow}/G${totalRevenueRow},0)`).setNumberFormat('0.0%');
  row++;

  // EBITDA Margin %
  sheet.getRange(row, 1).setValue('EBITDA Margin %');
  sheet.getRange(row, 2).setFormula(`=IF(B${totalRevenueRow}>0,B${ebitdaRow}/B${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 3).setFormula(`=IF(C${totalRevenueRow}>0,C${ebitdaRow}/C${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 4).setFormula(`=IF(D${totalRevenueRow}>0,D${ebitdaRow}/D${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 5).setFormula(`=IF(E${totalRevenueRow}>0,E${ebitdaRow}/E${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 6).setFormula(`=IF(F${totalRevenueRow}>0,F${ebitdaRow}/F${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 7).setFormula(`=IF(G${totalRevenueRow}>0,G${ebitdaRow}/G${totalRevenueRow},0)`).setNumberFormat('0.0%');
  row++;

  // Net Profit Margin %
  sheet.getRange(row, 1).setValue('Net Profit Margin %');
  sheet.getRange(row, 2).setFormula(`=IF(B${totalRevenueRow}>0,B${netIncomeRow}/B${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 3).setFormula(`=IF(C${totalRevenueRow}>0,C${netIncomeRow}/C${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 4).setFormula(`=IF(D${totalRevenueRow}>0,D${netIncomeRow}/D${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 5).setFormula(`=IF(E${totalRevenueRow}>0,E${netIncomeRow}/E${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 6).setFormula(`=IF(F${totalRevenueRow}>0,F${netIncomeRow}/F${totalRevenueRow},0)`).setNumberFormat('0.0%');
  sheet.getRange(row, 7).setFormula(`=IF(G${totalRevenueRow}>0,G${netIncomeRow}/G${totalRevenueRow},0)`).setNumberFormat('0.0%');
  row += 2;

  // Notes
  sheet.getRange(row, 1).setValue('Note: Year 1 is 10 months (Mar-Dec 2026). Year 5 is 14 months (Jan 2030-Feb 2031).').setFontStyle('italic').setFontColor('#666666');
  sheet.getRange(row, 1, 1, 7).merge();
  row++;

  sheet.getRange(row, 1).setValue('All values reference P&L and Cash Flow monthly sheets for aggregation.').setFontStyle('italic').setFontColor('#666666');
  sheet.getRange(row, 1, 1, 7).merge();

  // Format columns
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 120);

  // Freeze header row only (no column freeze due to merged note cells)
  sheet.setFrozenRows(headerRow);

  return sheet;
}
