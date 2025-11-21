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
  row++;

  sheet.getRange(row, 1).setValue('Total Monthly Rent').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=B8/12').setNumberFormat('$#,##0.00').setFontWeight('bold');
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
  row++;

  sheet.getRange(row, 1).setValue('Beer Price');
  sheet.getRange(row, 2).setValue(6).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('per bottle');
  row++;

  sheet.getRange(row, 1).setValue('Cocktail Price');
  sheet.getRange(row, 2).setValue(10).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('per drink');
  row++;

  sheet.getRange(row, 1).setValue('Drink Mix - Wine %');
  sheet.getRange(row, 2).setValue(0.40).setNumberFormat('0%');
  row++;

  sheet.getRange(row, 1).setValue('Drink Mix - Beer %');
  sheet.getRange(row, 2).setValue(0.40).setNumberFormat('0%');
  row++;

  sheet.getRange(row, 1).setValue('Drink Mix - Cocktail %');
  sheet.getRange(row, 2).setValue(0.20).setNumberFormat('0%');
  row++;

  sheet.getRange(row, 1).setValue('Weighted Avg Drink Price').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=B65*B68+B66*B69+B67*B70').setNumberFormat('$#,##0.00').setFontWeight('bold');
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
  row++;

  sheet.getRange(row, 1).setValue('Supply Replenishment %');
  sheet.getRange(row, 2).setValue(0.30).setNumberFormat('0%');
  sheet.getRange(row, 3).setValue('% of initial to reorder');
  row++;

  sheet.getRange(row, 1).setValue('Replenishment Frequency');
  sheet.getRange(row, 2).setValue(4).setNumberFormat('#,##0');
  sheet.getRange(row, 3).setValue('times per year (quarterly)');
  row++;

  sheet.getRange(row, 1).setValue('Annual Supply Replenishment').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=B79*B80*B81').setNumberFormat('$#,##0').setFontWeight('bold');
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
  row++;

  sheet.getRange(row, 1).setValue('Equipment/Year');
  sheet.getRange(row, 2).setValue(17857).setNumberFormat('$#,##0');
  sheet.getRange(row, 3).setValue('7-year straight line');
  row++;

  sheet.getRange(row, 1).setValue('Total Annual Depreciation').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=B109+B110').setNumberFormat('$#,##0').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('auto-calculated').setFontStyle('italic');
  row += 2;

  // ========== LOAN TERMS ==========
  sheet.getRange(row, 1).setValue('LOAN TERMS').setFontWeight('bold').setFontSize(12).setBackground('#E8EAED');
  sheet.getRange(row, 1, 1, 3).merge();
  row++;

  sheet.getRange(row, 1).setValue('Interest Rate');
  sheet.getRange(row, 2).setValue(0.12).setNumberFormat('0.00%');
  sheet.getRange(row, 3).setValue('annual');
  row++;

  sheet.getRange(row, 1).setValue('Loan Term (Years)');
  sheet.getRange(row, 2).setValue(10).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Grace Period (Months)');
  sheet.getRange(row, 2).setValue(0).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 1).setValue('Operating Reserve (Months)');
  sheet.getRange(row, 2).setValue(4).setNumberFormat('#,##0');
  sheet.getRange(row, 3).setValue('Months of expenses to reserve');
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

  // This will be built next
  sheet.getRange(1, 1).setValue('REVENUE DETAIL - Under Construction');
}

/**
 * EXPENSE DETAIL SHEET
 * 60 months of detailed expenses by category
 */
function createExpenseDetailSheet(ss) {
  const sheet = ss.insertSheet('Expense Detail');
  sheet.setTabColor('#EA4335'); // Red

  sheet.getRange(1, 1).setValue('EXPENSE DETAIL - Under Construction');
}

/**
 * STARTUP COSTS SHEET
 * One-time startup cost breakdown
 */
function createStartupCostsSheet(ss) {
  const sheet = ss.insertSheet('Startup Costs');
  sheet.setTabColor('#FBBC04'); // Yellow

  sheet.getRange(1, 1).setValue('STARTUP COSTS - Under Construction');
}

/**
 * LOAN AMORTIZATION SHEET
 * Monthly loan payment schedule
 */
function createLoanAmortizationSheet(ss) {
  const sheet = ss.insertSheet('Loan Amortization');
  sheet.setTabColor('#FF6D01'); // Orange

  sheet.getRange(1, 1).setValue('LOAN AMORTIZATION - Under Construction');
}

/**
 * PROFIT & LOSS SHEET
 * 60-month P&L statement
 */
function createProfitLossSheet(ss) {
  const sheet = ss.insertSheet('P&L');
  sheet.setTabColor('#9900FF'); // Purple

  sheet.getRange(1, 1).setValue('PROFIT & LOSS - Under Construction');
}

/**
 * CASH FLOW SHEET
 * 60-month cash flow statement
 */
function createCashFlowSheet(ss) {
  const sheet = ss.insertSheet('Cash Flow');
  sheet.setTabColor('#00BCD4'); // Cyan

  sheet.getRange(1, 1).setValue('CASH FLOW - Under Construction');
}

/**
 * DASHBOARD SHEET
 * High-level summary and key metrics
 */
function createDashboardSheet(ss) {
  const sheet = ss.insertSheet('Dashboard');
  sheet.setTabColor('#4285F4'); // Blue

  sheet.getRange(1, 1).setValue('DASHBOARD - Under Construction');
}

/**
 * ANNUAL SUMMARY SHEET
 * 5-year annual rollup
 */
function createAnnualSummarySheet(ss) {
  const sheet = ss.insertSheet('Annual Summary');
  sheet.setTabColor('#0F9D58'); // Dark Green

  sheet.getRange(1, 1).setValue('ANNUAL SUMMARY - Under Construction');
}
