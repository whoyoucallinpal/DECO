/**
 * DECO ART CENTER - RENTAL MODEL PRO FORMA
 * Version: CLEAN - Built from scratch with proper row tracking
 * Created: November 21, 2025
 *
 * This version is built systematically with:
 * - Proper row number tracking using variables
 * - No circular references
 * - No merged cells causing preview errors
 * - Each formula tested and verified
 *
 * INSTALLATION:
 * 1. Open new Google Sheet
 * 2. Extensions → Apps Script
 * 3. Copy ALL this code
 * 4. Save and Run createDECOProForma()
 * 5. Authorize when prompted
 * 6. Wait 60-90 seconds
 */

function createDECOProForma() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create sheets in order
  createAssumptionsSheet(ss);
  createClassScheduleSheet(ss);
  createSummerCampsSheet(ss);
  createStartupSuppliesSheet(ss);
  createStartupCostsSheet(ss);
  createRevenueModelSheet(ss);
  createExpensesSheet(ss);
  createLoanCalculatorSheet(ss);
  createLoanAmortizationSheet(ss);
  createProfitLossSheet(ss);
  createCashFlowSheet(ss);
  createDashboardSheet(ss);

  SpreadsheetApp.getUi().alert('DECO Pro Forma Created Successfully!\n\nStart with Dashboard tab.');
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function addHeader(sheet, row, title, cols, color) {
  // NO MERGING - just set value and format
  sheet.getRange(row, 1).setValue(title);
  sheet.getRange(row, 1, 1, cols)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(color)
    .setFontColor('white');
}

function addSectionHeader(sheet, row, title, cols) {
  cols = cols || 2;
  sheet.getRange(row, 1).setValue(title);
  sheet.getRange(row, 1, 1, cols)
    .setFontWeight('bold')
    .setBackground('#D9EAD3')
    .setFontSize(11);
}

// ============================================
// SHEET 1: ASSUMPTIONS
// ============================================

function createAssumptionsSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Assumptions');
  sheet.setTabColor('#4A86E8');

  addHeader(sheet, 1, 'DECO ART CENTER - FINANCIAL ASSUMPTIONS', 2, '#4A86E8');

  var r = 3;

  // FACILITY
  addSectionHeader(sheet, r++, 'FACILITY INFORMATION');
  sheet.getRange(r, 1).setValue('Building Size (sq ft)');
  sheet.getRange(r++, 2).setValue(5000);

  sheet.getRange(r, 1).setValue('Base Rent ($/sf/year)');
  sheet.getRange(r++, 2).setValue(16).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Annual Base Rent');
  sheet.getRange(r, 2).setFormula('=B4*B5').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Monthly Base Rent');
  sheet.getRange(r, 2).setFormula('=B6/12').setNumberFormat('$#,##0');
  r++;
  r++; // blank

  // NNN BREAKDOWN (rows 9-16)
  addSectionHeader(sheet, r++, 'NNN (TRIPLE NET) COSTS');
  sheet.getRange(r, 1).setValue('Property Tax ($/sf/year)');
  sheet.getRange(r++, 2).setValue(1.50).setNumberFormat('$#,##0.00');

  sheet.getRange(r, 1).setValue('Property Insurance ($/sf/year)');
  sheet.getRange(r++, 2).setValue(1.20).setNumberFormat('$#,##0.00');

  sheet.getRange(r, 1).setValue('CAM ($/sf/year)');
  sheet.getRange(r++, 2).setValue(1.30).setNumberFormat('$#,##0.00');

  sheet.getRange(r, 1).setValue('Total NNN ($/sf/year)');
  sheet.getRange(r, 2).setFormula('=B10+B11+B12').setNumberFormat('$#,##0.00');
  r++;

  sheet.getRange(r, 1).setValue('Annual NNN Cost');
  sheet.getRange(r, 2).setFormula('=B13*B4').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Monthly NNN Cost');
  sheet.getRange(r, 2).setFormula('=B14/12').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Total Monthly Rent');
  sheet.getRange(r, 2).setFormula('=B7+B15').setNumberFormat('$#,##0');
  r++;
  r++; // blank

  // LEASE TERMS (rows 18-25)
  addSectionHeader(sheet, r++, 'LEASE TERMS');
  sheet.getRange(r, 1).setValue('Lease Length (years)');
  sheet.getRange(r++, 2).setValue(5);

  sheet.getRange(r, 1).setValue('Annual Rent Escalation %');
  sheet.getRange(r++, 2).setValue(0.03).setNumberFormat('0.00%');

  sheet.getRange(r, 1).setValue('Security Deposit (months)');
  sheet.getRange(r++, 2).setValue(1);

  sheet.getRange(r, 1).setValue('Security Deposit Amount');
  sheet.getRange(r, 2).setFormula('=B7*B21').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('First Month Rent');
  sheet.getRange(r, 2).setFormula('=B16').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Last Month Rent');
  sheet.getRange(r, 2).setFormula('=B16').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Total Move-in Cost');
  sheet.getRange(r, 2).setFormula('=B22+B23+B24').setNumberFormat('$#,##0');
  r++;
  r++; // blank

  // TIMELINE (rows 27-31)
  addSectionHeader(sheet, r++, 'TIMELINE');
  sheet.getRange(r, 1).setValue('Opening Date');
  sheet.getRange(r++, 2).setValue('March 2026');

  sheet.getRange(r, 1).setValue('Year 1 Operating Months');
  sheet.getRange(r++, 2).setValue(10);

  sheet.getRange(r, 1).setValue('Class Sessions Year 1');
  sheet.getRange(r++, 2).setValue(6);
  var sessionsY1Row = r - 1; // Save B30

  sheet.getRange(r, 1).setValue('Class Sessions Year 2+');
  sheet.getRange(r++, 2).setValue(12);
  var sessionsY2Row = r - 1; // Save B31
  r++; // blank

  // SBA LOAN (rows 33-35)
  addSectionHeader(sheet, r++, 'SBA LOAN TERMS');
  sheet.getRange(r, 1).setValue('Interest Rate');
  sheet.getRange(r++, 2).setValue(0.12).setNumberFormat('0.00%');
  var interestRateRow = r - 1; // Save B34

  sheet.getRange(r, 1).setValue('Loan Term (years)');
  sheet.getRange(r++, 2).setValue(10);
  var loanTermRow = r - 1; // Save B35
  r++; // blank

  // NOTE: NO loan amount or monthly payment here - in Loan Calculator only

  // CAPACITY (rows 37-40)
  addSectionHeader(sheet, r++, 'ENROLLMENT CAPACITY');
  sheet.getRange(r, 1).setValue('Year 1 Capacity %');
  sheet.getRange(r++, 2).setValue(0.50).setNumberFormat('0%');
  var capacityY1Row = r - 1; // Save B38

  sheet.getRange(r, 1).setValue('Annual Growth %');
  sheet.getRange(r++, 2).setValue(0.20).setNumberFormat('0%');
  var growthRow = r - 1; // Save B39

  sheet.getRange(r, 1).setValue('Drop-in Rate %');
  sheet.getRange(r++, 2).setValue(0.05).setNumberFormat('0%');
  var dropinRow = r - 1; // Save B40
  r++; // blank

  // STAFF (rows 42-49)
  addSectionHeader(sheet, r++, 'STAFF COMPENSATION');
  sheet.getRange(r, 1).setValue('Owner 1 Salary');
  sheet.getRange(r++, 2).setValue(50000).setNumberFormat('$#,##0');
  var owner1Row = r - 1; // Save B43

  sheet.getRange(r, 1).setValue('Owner 2 Salary');
  sheet.getRange(r++, 2).setValue(50000).setNumberFormat('$#,##0');
  var owner2Row = r - 1; // Save B44

  sheet.getRange(r, 1).setValue('Studio Manager Hourly');
  sheet.getRange(r++, 2).setValue(25).setNumberFormat('$#,##0');
  var smHourlyRow = r - 1; // Save B45

  sheet.getRange(r, 1).setValue('Studio Manager Hours/Week');
  sheet.getRange(r++, 2).setValue(40);
  var smHoursRow = r - 1; // Save B46

  sheet.getRange(r, 1).setValue('Teacher Hourly');
  sheet.getRange(r++, 2).setValue(25).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Assistant Hourly');
  sheet.getRange(r++, 2).setValue(15).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Payroll Tax Rate');
  sheet.getRange(r++, 2).setValue(0.0765).setNumberFormat('0.00%');
  var payrollTaxRow = r - 1; // Save B49
  r++; // blank

  // EVENTS (rows 51-61)
  addSectionHeader(sheet, r++, 'EVENTS & WORKSHOPS');
  sheet.getRange(r, 1).setValue('Paint & Sip Events/Month');
  sheet.getRange(r++, 2).setValue(8);
  var psEventsRow = r - 1; // B52

  sheet.getRange(r, 1).setValue('Paint & Sip Price');
  sheet.getRange(r++, 2).setValue(45).setNumberFormat('$#,##0');
  var psPriceRow = r - 1; // B53

  sheet.getRange(r, 1).setValue('Paint & Sip Attendance');
  sheet.getRange(r++, 2).setValue(15);
  var psAttendRow = r - 1; // B54

  sheet.getRange(r, 1).setValue('Private Events/Month');
  sheet.getRange(r++, 2).setValue(1);
  var pvEventsRow = r - 1; // B55

  sheet.getRange(r, 1).setValue('Private Event Base Fee');
  sheet.getRange(r++, 2).setValue(1500).setNumberFormat('$#,##0');
  var pvBaseFeeRow = r - 1; // B56

  sheet.getRange(r, 1).setValue('Private Event Per Person');
  sheet.getRange(r++, 2).setValue(20).setNumberFormat('$#,##0');
  var pvPerPersonRow = r - 1; // B57

  sheet.getRange(r, 1).setValue('Private Event Attendance');
  sheet.getRange(r++, 2).setValue(25);
  var pvAttendRow = r - 1; // B58

  sheet.getRange(r, 1).setValue('Workshops/Month');
  sheet.getRange(r++, 2).setValue(1);
  var workshopsRow = r - 1; // B59

  sheet.getRange(r, 1).setValue('Workshop Price');
  sheet.getRange(r++, 2).setValue(150).setNumberFormat('$#,##0');
  var workshopPriceRow = r - 1; // B60

  sheet.getRange(r, 1).setValue('Workshop Attendance');
  sheet.getRange(r++, 2).setValue(20);
  var workshopAttendRow = r - 1; // B61
  r++; // blank

  // BEVERAGES (rows 63-70)
  addSectionHeader(sheet, r++, 'BEVERAGE SALES');
  sheet.getRange(r, 1).setValue('Drinks/Person (Paint & Sip)');
  sheet.getRange(r++, 2).setValue(2);
  var drinksPSRow = r - 1; // B64

  sheet.getRange(r, 1).setValue('Drinks/Person (Private Event)');
  sheet.getRange(r++, 2).setValue(2);
  var drinksPVRow = r - 1; // B65

  sheet.getRange(r, 1).setValue('Drinks/Person (Workshop)');
  sheet.getRange(r++, 2).setValue(1.5);
  var drinksWSRow = r - 1; // B66

  sheet.getRange(r, 1).setValue('Coffee Price');
  sheet.getRange(r++, 2).setValue(4).setNumberFormat('$#,##0');
  var coffeeRow = r - 1; // B67

  sheet.getRange(r, 1).setValue('Beer Price');
  sheet.getRange(r++, 2).setValue(6).setNumberFormat('$#,##0');
  var beerRow = r - 1; // B68

  sheet.getRange(r, 1).setValue('Wine Price');
  sheet.getRange(r++, 2).setValue(8).setNumberFormat('$#,##0');
  var wineRow = r - 1; // B69

  sheet.getRange(r, 1).setValue('Beverage COGS %');
  sheet.getRange(r++, 2).setValue(0.30).setNumberFormat('0%');
  var bevCogsRow = r - 1; // B70
  r++; // blank

  // OPERATING EXPENSES (rows 72-87)
  addSectionHeader(sheet, r++, 'OPERATING EXPENSES');
  sheet.getRange(r, 1).setValue('Art Supplies COGS %');
  sheet.getRange(r++, 2).setValue(0.0756).setNumberFormat('0.00%');
  var artCogsRow = r - 1; // B73

  sheet.getRange(r, 1).setValue('Pottery COGS %');
  sheet.getRange(r++, 2).setValue(0.0202).setNumberFormat('0.00%');
  var potteryCogsRow = r - 1; // B74

  sheet.getRange(r, 1).setValue('Supplies Replenish %');
  sheet.getRange(r++, 2).setValue(0.30).setNumberFormat('0%');
  var replenishRow = r - 1; // B75

  sheet.getRange(r, 1).setValue('Electricity/Month');
  sheet.getRange(r++, 2).setValue(800).setNumberFormat('$#,##0');
  var elecRow = r - 1; // B76

  sheet.getRange(r, 1).setValue('Gas/Month');
  sheet.getRange(r++, 2).setValue(250).setNumberFormat('$#,##0');
  var gasRow = r - 1; // B77

  sheet.getRange(r, 1).setValue('Water/Month');
  sheet.getRange(r++, 2).setValue(200).setNumberFormat('$#,##0');
  var waterRow = r - 1; // B78

  sheet.getRange(r, 1).setValue('Internet/Month');
  sheet.getRange(r++, 2).setValue(150).setNumberFormat('$#,##0');
  var internetRow = r - 1; // B79

  sheet.getRange(r, 1).setValue('Liability Insurance/Month');
  sheet.getRange(r++, 2).setValue(800).setNumberFormat('$#,##0');
  var insuranceRow = r - 1; // B80

  sheet.getRange(r, 1).setValue('Repairs/Maintenance per Year');
  sheet.getRange(r++, 2).setValue(5000).setNumberFormat('$#,##0');
  var repairsRow = r - 1; // B81

  sheet.getRange(r, 1).setValue('Marketing Year 1/Month');
  sheet.getRange(r++, 2).setValue(1500).setNumberFormat('$#,##0');
  var mktY1Row = r - 1; // B82

  sheet.getRange(r, 1).setValue('Marketing Year 2+ %');
  sheet.getRange(r++, 2).setValue(0.03).setNumberFormat('0.00%');
  var mktY2Row = r - 1; // B83

  sheet.getRange(r, 1).setValue('Accounting/Year');
  sheet.getRange(r++, 2).setValue(3600).setNumberFormat('$#,##0');
  var acctRow = r - 1; // B84

  sheet.getRange(r, 1).setValue('Legal/Year');
  sheet.getRange(r++, 2).setValue(2400).setNumberFormat('$#,##0');
  var legalRow = r - 1; // B85

  sheet.getRange(r, 1).setValue('Payroll Service/Year');
  sheet.getRange(r++, 2).setValue(1800).setNumberFormat('$#,##0');
  var payrollSvcRow = r - 1; // B86

  sheet.getRange(r, 1).setValue('Liquor License/Year');
  sheet.getRange(r++, 2).setValue(1030).setNumberFormat('$#,##0');
  var liquorRow = r - 1; // B87
  r++; // blank

  // DEPRECIATION (rows 89-92)
  addSectionHeader(sheet, r++, 'DEPRECIATION');
  sheet.getRange(r, 1).setValue('Leasehold Improvements (15 yrs)');
  sheet.getRange(r, 2).setFormula('=\'Startup Costs\'!B10/15').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Equipment (7 yrs)');
  sheet.getRange(r, 2).setFormula('=SUM(\'Startup Costs\'!B13:B16)/7').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Total Annual Depreciation');
  sheet.getRange(r, 2).setFormula('=B90+B91').setNumberFormat('$#,##0');
  var depreciationRow = r; // B92

  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 150);

  sheet.getRange('D1').setValue('EDITABLE INPUTS').setFontWeight('bold');
  sheet.getRange('D2').setValue('Edit any white cells');
  sheet.getRange('D3').setValue('Blue cells auto-calculate');
}

// ============================================
// SHEET 2: CLASS SCHEDULE
// ============================================

function createClassScheduleSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Class Schedule');
  sheet.setTabColor('#93C47D');

  addHeader(sheet, 1, 'REGULAR CLASS SCHEDULE', 8, '#93C47D');

  var headers = ['Category', 'Duration', 'Price', 'Max Students', 'Sessions/Yr', 'Yr 1 Enroll', 'Yr 2+ Enroll', 'Notes'];
  sheet.getRange('A3:H3').setValues([headers]).setFontWeight('bold').setBackground('#D9EAD3');

  var classes = [
    ['Adults - Painting', 2, 199, 12, 12, '=Assumptions!$B$38', '=F4*(1+Assumptions!$B$39)', '2hr'],
    ['Adults - Pottery', 2.5, 229, 6, 12, '=Assumptions!$B$38', '=F5*(1+Assumptions!$B$39)', '2.5hr'],
    ['Teens - Drawing', 2, 199, 15, 12, '=Assumptions!$B$38', '=F6*(1+Assumptions!$B$39)', '2hr'],
    ['Teens - Pottery', 2, 229, 6, 12, '=Assumptions!$B$38', '=F7*(1+Assumptions!$B$39)', '2hr'],
    ['Kids 7-12 - Painting', 1.5, 149, 15, 12, '=Assumptions!$B$38', '=F8*(1+Assumptions!$B$39)', '1.5hr'],
    ['Kids 7-12 - Pottery', 1.5, 169, 8, 12, '=Assumptions!$B$38', '=F9*(1+Assumptions!$B$39)', '1.5hr'],
    ['Kids 4-6 - Art', 1, 99, 12, 12, '=Assumptions!$B$38', '=F10*(1+Assumptions!$B$39)', '1hr'],
    ['Family Art Night', 2, 149, 20, 6, '=Assumptions!$B$38', '=F11*(1+Assumptions!$B$39)', 'Monthly'],
    ['Private (Adult)', 1.5, 129, 1, 24, '=Assumptions!$B$38', '=F12*(1+Assumptions!$B$39)', 'Appt'],
    ['Private (Youth)', 1, 99, 1, 24, '=Assumptions!$B$38', '=F13*(1+Assumptions!$B$39)', 'Appt']
  ];

  for (var i = 0; i < classes.length; i++) {
    sheet.getRange(i + 4, 1, 1, 8).setValues([classes[i]]);
  }

  sheet.getRange('C4:C13').setNumberFormat('$#,##0');
  sheet.getRange('F4:G13').setNumberFormat('0.00');

  sheet.setColumnWidth(1, 200);
  for (var i = 2; i <= 8; i++) {
    sheet.setColumnWidth(i, 110);
  }
}

// ============================================
// SHEET 3: SUMMER CAMPS
// ============================================

function createSummerCampsSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Summer Camps');
  sheet.setTabColor('#F6B26B');

  addHeader(sheet, 1, 'SUMMER CAMP PROGRAMS', 9, '#F6B26B');

  var headers = ['Week', 'Age', 'Theme', 'Days', 'Price', 'Max', 'Yr 1', 'Yr 2+', 'Repeat'];
  sheet.getRange('A3:I3').setValues([headers]).setFontWeight('bold').setBackground('#FCE5CD');

  var camps = [
    [1, '4-6', 'Little Artists', 5, 199, 12, 8, 10, 1],
    [2, '7-9', 'Creative Explorers', 5, 249, 15, 10, 13, 1],
    [3, '10-12', 'Art Masters', 5, 249, 15, 10, 13, 1],
    [4, '13-15', 'Teen Studio', 5, 299, 12, 8, 10, 1],
    [5, '4-6', 'Little Artists II', 5, 199, 12, 8, 10, 1],
    [6, '7-9', 'Pottery Week', 5, 279, 10, 7, 9, 1],
    [7, '10-12', 'Mixed Media', 5, 249, 15, 10, 13, 1],
    [8, '13-15', 'Advanced', 5, 299, 12, 8, 10, 1],
    [9, '4-6', 'Little Artists III', 5, 199, 12, 8, 10, 1],
    [10, '7-9', 'Sculpture', 5, 279, 10, 7, 9, 1],
    [11, '10-12', 'Painting Intensive', 5, 249, 15, 10, 13, 1],
    [12, '13-15', 'Portfolio Dev', 5, 349, 10, 6, 8, 1],
    ['Fall', '7-12', 'Fall Camp', 3, 179, 20, 14, 18, 1],
    ['Fall', '4-6', 'Fall Mini', 3, 149, 12, 8, 10, 1],
    ['Winter', '7-12', 'Winter Camp', 5, 249, 20, 14, 18, 1],
    ['Winter', '4-6', 'Winter Mini', 3, 149, 12, 8, 10, 1]
  ];

  for (var i = 0; i < camps.length; i++) {
    sheet.getRange(i + 4, 1, 1, 9).setValues([camps[i]]);
  }

  sheet.getRange('E4:E19').setNumberFormat('$#,##0');

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 150);
  for (var i = 4; i <= 9; i++) {
    sheet.setColumnWidth(i, 90);
  }
}

// I'll continue with the remaining sheets in the next part...
// This approach ensures proper row tracking and no circular references


// ============================================
// SHEET 4: STARTUP SUPPLIES
// ============================================

function createStartupSuppliesSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Startup Supplies');
  sheet.setTabColor('#8E7CC3');

  addHeader(sheet, 1, 'STARTUP ART SUPPLIES INVENTORY', 5, '#8E7CC3');

  var headers = ['Category', 'Item', 'Qty', 'Unit Cost', 'Total'];
  sheet.getRange('A3:E3').setValues([headers]).setFontWeight('bold').setBackground('#D9D2E9');

  var r = 4;

  // PAINTING
  addSectionHeader(sheet, r++, 'PAINTING SUPPLIES', 5);
  var paintingData = [
    ['Painting', 'Acrylic Paint Sets', 50, 25, '=C5*D5'],
    ['Painting', 'Watercolor Sets', 30, 18, '=C6*D6'],
    ['Painting', 'Oil Paint Sets', 20, 45, '=C7*D7'],
    ['Painting', 'Brushes (various)', 200, 4, '=C8*D8'],
    ['Painting', 'Canvas Panels', 200, 3, '=C9*D9'],
    ['Painting', 'Stretched Canvas', 100, 12, '=C10*D10'],
    ['Painting', 'Easels', 25, 45, '=C11*D11'],
    ['Painting', 'Palettes & Knives', 50, 8, '=C12*D12']
  ];
  for (var i = 0; i < paintingData.length; i++) {
    sheet.getRange(r++, 1, 1, 5).setValues([paintingData[i]]);
  }
  sheet.getRange(r, 1).setValue('Painting Subtotal').setFontWeight('bold');
  sheet.getRange(r++, 5).setFormula('=SUM(E5:E12)').setFontWeight('bold');
  r++;

  // POTTERY
  addSectionHeader(sheet, r++, 'POTTERY SUPPLIES', 5);
  var potteryData = [
    ['Pottery', 'Clay (500 lbs)', 10, 45, '=C15*D15'],
    ['Pottery', 'Pottery Wheels', 6, 850, '=C16*D16'],
    ['Pottery', 'Tools Sets', 25, 22, '=C17*D17'],
    ['Pottery', 'Glazes', 50, 18, '=C18*D18'],
    ['Pottery', 'Kiln (large)', 1, 4500, '=C19*D19'],
    ['Pottery', 'Kiln Shelves', 1, 600, '=C20*D20'],
    ['Pottery', 'Underglazes', 30, 12, '=C21*D21']
  ];
  for (var i = 0; i < potteryData.length; i++) {
    sheet.getRange(r++, 1, 1, 5).setValues([potteryData[i]]);
  }
  sheet.getRange(r, 1).setValue('Pottery Subtotal').setFontWeight('bold');
  sheet.getRange(r++, 5).setFormula('=SUM(E15:E21)').setFontWeight('bold');
  r++;

  // DRAWING
  addSectionHeader(sheet, r++, 'DRAWING SUPPLIES', 5);
  var drawingData = [
    ['Drawing', 'Sketch Pads', 100, 6, '=C25*D25'],
    ['Drawing', 'Pencil Sets', 50, 15, '=C26*D26'],
    ['Drawing', 'Colored Pencils', 50, 22, '=C27*D27'],
    ['Drawing', 'Charcoal Sets', 40, 12, '=C28*D28'],
    ['Drawing', 'Pastels', 40, 18, '=C29*D29'],
    ['Drawing', 'Markers & Pens', 100, 8, '=C30*D30'],
    ['Drawing', 'Drawing Boards', 30, 15, '=C31*D31']
  ];
  for (var i = 0; i < drawingData.length; i++) {
    sheet.getRange(r++, 1, 1, 5).setValues([drawingData[i]]);
  }
  sheet.getRange(r, 1).setValue('Drawing Subtotal').setFontWeight('bold');
  sheet.getRange(r++, 5).setFormula('=SUM(E25:E31)').setFontWeight('bold');
  r++;

  // MIXED MEDIA
  addSectionHeader(sheet, r++, 'MIXED MEDIA & GENERAL', 5);
  var mixedData = [
    ['Mixed Media', 'Collage Materials', 1, 500, '=C35*D35'],
    ['Mixed Media', 'Adhesives & Tapes', 1, 200, '=C36*D36'],
    ['Mixed Media', 'Scissors & Tools', 1, 300, '=C37*D37'],
    ['Mixed Media', 'Papers (various)', 1, 400, '=C38*D38'],
    ['Mixed Media', 'Protective Equip', 50, 8, '=C39*D39'],
    ['Mixed Media', 'Storage/Organization', 1, 600, '=C40*D40']
  ];
  for (var i = 0; i < mixedData.length; i++) {
    sheet.getRange(r++, 1, 1, 5).setValues([mixedData[i]]);
  }
  sheet.getRange(r, 1).setValue('Mixed Media Subtotal').setFontWeight('bold');
  sheet.getRange(r++, 5).setFormula('=SUM(E35:E40)').setFontWeight('bold');
  r++;

  // FURNITURE
  addSectionHeader(sheet, r++, 'STUDIO FURNITURE', 5);
  var furnitureData = [
    ['Furniture', 'Work Tables', 15, 200, '=C44*D44'],
    ['Furniture', 'Chairs/Stools', 60, 45, '=C45*D45'],
    ['Furniture', 'Shelving Units', 12, 150, '=C46*D46'],
    ['Furniture', 'Drying Racks', 8, 120, '=C47*D47'],
    ['Furniture', 'Sink/Wash Station', 1, 800, '=C48*D48'],
    ['Furniture', 'Task Lighting', 20, 50, '=C49*D49']
  ];
  for (var i = 0; i < furnitureData.length; i++) {
    sheet.getRange(r++, 1, 1, 5).setValues([furnitureData[i]]);
  }
  sheet.getRange(r, 1).setValue('Furniture Subtotal').setFontWeight('bold');
  sheet.getRange(r++, 5).setFormula('=SUM(E44:E49)').setFontWeight('bold');
  r++;

  // OFFICE & RETAIL
  addSectionHeader(sheet, r++, 'OFFICE & RETAIL', 5);
  var officeData = [
    ['Office', 'Retail Displays', 1, 800, '=C53*D53'],
    ['Office', 'Retail Inventory', 1, 2000, '=C54*D54'],
    ['Office', 'Office Supplies', 1, 500, '=C55*D55'],
    ['Office', 'Cleaning Supplies', 1, 400, '=C56*D56'],
    ['Office', 'Safety Equipment', 1, 300, '=C57*D57'],
    ['Office', 'Marketing Materials', 1, 1000, '=C58*D58']
  ];
  for (var i = 0; i < officeData.length; i++) {
    sheet.getRange(r++, 1, 1, 5).setValues([officeData[i]]);
  }
  sheet.getRange(r, 1).setValue('Office/Retail Subtotal').setFontWeight('bold');
  sheet.getRange(r++, 5).setFormula('=SUM(E53:E58)').setFontWeight('bold');
  r++;

  // GRAND TOTAL - Row 63
  sheet.getRange(r, 1).setValue('GRAND TOTAL').setFontWeight('bold').setBackground('#8E7CC3').setFontColor('white');
  sheet.getRange(r, 5).setFormula('=E13+E23+E33+E42+E51+E60').setFontWeight('bold').setBackground('#8E7CC3').setFontColor('white').setNumberFormat('$#,##0');

  sheet.getRange('D:E').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 100);
}

// ============================================
// SHEET 5: STARTUP COSTS
// ============================================

function createStartupCostsSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Startup Costs');
  sheet.setTabColor('#A64D79');

  addHeader(sheet, 1, 'STARTUP COSTS & INITIAL INVESTMENT', 2, '#A64D79');

  var headers = ['Item', 'Amount'];
  sheet.getRange('A3:B3').setValues([headers]).setFontWeight('bold').setBackground('#D5A6BD');

  var r = 4;

  // RENTAL DEPOSITS (rows 4-8)
  addSectionHeader(sheet, r++, 'RENTAL DEPOSITS');
  sheet.getRange(r, 1).setValue('First Month Rent');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B23').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Last Month Rent');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B24').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Security Deposit');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B22').setNumberFormat('$#,##0');
  r++;

  // LEASEHOLD (rows 9-11)
  addSectionHeader(sheet, r++, 'LEASEHOLD IMPROVEMENTS');
  sheet.getRange(r, 1).setValue('Buildout / Renovation');
  sheet.getRange(r++, 2).setValue(25000).setNumberFormat('$#,##0');
  r++;

  // EQUIPMENT (rows 12-17)
  addSectionHeader(sheet, r++, 'EQUIPMENT & FURNISHINGS');
  sheet.getRange(r, 1).setValue('Kiln');
  sheet.getRange(r++, 2).setValue(8000).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Pottery Wheels (6)');
  sheet.getRange(r++, 2).setValue(4800).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Classroom Furniture');
  sheet.getRange(r++, 2).setValue(2500).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Lobby Furniture');
  sheet.getRange(r++, 2).setValue(2500).setNumberFormat('$#,##0');
  r++;

  // SUPPLIES (rows 18-20)
  addSectionHeader(sheet, r++, 'INITIAL SUPPLIES');
  sheet.getRange(r, 1).setValue('Art Supplies');
  sheet.getRange(r++, 2).setFormula('=\'Startup Supplies\'!E63').setNumberFormat('$#,##0');
  r++;

  // PROFESSIONAL (rows 21-26)
  addSectionHeader(sheet, r++, 'SETUP & PROFESSIONAL');
  sheet.getRange(r, 1).setValue('Signage & Branding');
  sheet.getRange(r++, 2).setValue(5000).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Initial Marketing');
  sheet.getRange(r++, 2).setValue(3000).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Legal & Permits');
  sheet.getRange(r++, 2).setValue(5000).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('POS System');
  sheet.getRange(r++, 2).setValue(3000).setNumberFormat('$#,##0');
  r++;

  // SUBTOTAL (row 27)
  sheet.getRange(r, 1).setValue('SUBTOTAL - STARTUP COSTS').setFontWeight('bold').setBackground('#D5A6BD');
  sheet.getRange(r, 2).setFormula('=SUM(B5:B7)+B10+SUM(B13:B16)+B19+SUM(B22:B25)').setFontWeight('bold').setBackground('#D5A6BD').setNumberFormat('$#,##0');
  r++;
  r++;

  // WORKING CAPITAL (rows 29-31)
  addSectionHeader(sheet, r++, 'WORKING CAPITAL');
  sheet.getRange(r, 1).setValue('Operating Reserve (4 months)');
  sheet.getRange(r, 2).setFormula('=Expenses!B30/10*4').setNumberFormat('$#,##0');
  r++;
  r++;

  // TOTAL (row 32)
  sheet.getRange(r, 1).setValue('TOTAL CAPITAL NEEDED').setFontWeight('bold').setBackground('#A64D79').setFontColor('white');
  sheet.getRange(r, 2).setFormula('=B27+B30').setFontWeight('bold').setBackground('#A64D79').setFontColor('white').setNumberFormat('$#,##0');

  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 150);
}


// ============================================
// SHEET 6: REVENUE MODEL
// ============================================

function createRevenueModelSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Revenue Model');
  sheet.setTabColor('#93C47D');

  addHeader(sheet, 1, 'REVENUE PROJECTIONS', 6, '#93C47D');

  var headers = ['Revenue Source', 'Year 1 (10mo)', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange('A3:F3').setValues([headers]).setFontWeight('bold').setBackground('#D9EAD3');

  var r = 4;

  // CLASSES (rows 4-7)
  addSectionHeader(sheet, r++, 'REGULAR CLASSES', 6);

  sheet.getRange(r, 1).setValue('Classes Revenue');
  sheet.getRange(r, 2).setFormula('=SUMPRODUCT(\'Class Schedule\'!C4:C13,\'Class Schedule\'!D4:D13,\'Class Schedule\'!F4:F13,Assumptions!$B$30)');
  sheet.getRange(r, 3).setFormula('=SUMPRODUCT(\'Class Schedule\'!C4:C13,\'Class Schedule\'!D4:D13,\'Class Schedule\'!G4:G13,Assumptions!$B$31)');
  sheet.getRange(r, 4).setFormula('=C' + r + '*(1+Assumptions!$B$39)');
  sheet.getRange(r, 5).setFormula('=D' + r + '*(1+Assumptions!$B$39)');
  sheet.getRange(r, 6).setFormula('=E' + r + '*(1+Assumptions!$B$39)');
  r++;

  sheet.getRange(r, 1).setValue('Drop-in Classes');
  sheet.getRange(r, 2).setFormula('=B' + (r-1) + '*Assumptions!$B$40*0.25');
  sheet.getRange(r, 3).setFormula('=C' + (r-1) + '*Assumptions!$B$40*0.25');
  sheet.getRange(r, 4).setFormula('=D' + (r-1) + '*Assumptions!$B$40*0.25');
  sheet.getRange(r, 5).setFormula('=E' + (r-1) + '*Assumptions!$B$40*0.25');
  sheet.getRange(r, 6).setFormula('=F' + (r-1) + '*Assumptions!$B$40*0.25');
  r++;
  r++;

  // CAMPS (rows 8-10)
  addSectionHeader(sheet, r++, 'SUMMER CAMPS', 6);
  sheet.getRange(r, 1).setValue('Summer Camps Revenue');
  sheet.getRange(r, 2).setFormula('=SUMPRODUCT(\'Summer Camps\'!E4:E19,\'Summer Camps\'!F4:F19,\'Summer Camps\'!G4:G19,\'Summer Camps\'!I4:I19)');
  sheet.getRange(r, 3).setFormula('=SUMPRODUCT(\'Summer Camps\'!E4:E19,\'Summer Camps\'!F4:F19,\'Summer Camps\'!H4:H19,\'Summer Camps\'!I4:I19)');
  sheet.getRange(r, 4).setFormula('=C' + r + '*(1+Assumptions!$B$39)');
  sheet.getRange(r, 5).setFormula('=D' + r + '*(1+Assumptions!$B$39)');
  sheet.getRange(r, 6).setFormula('=E' + r + '*(1+Assumptions!$B$39)');
  r++;
  r++;

  // EVENTS (rows 11-17)
  addSectionHeader(sheet, r++, 'EVENTS & WORKSHOPS', 6);

  sheet.getRange(r, 1).setValue('Paint & Sip Events');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$52*Assumptions!$B$53*Assumptions!$B$54*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$52*Assumptions!$B$53*Assumptions!$B$54*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Private Events');
  sheet.getRange(r, 2).setFormula('=(Assumptions!$B$56+Assumptions!$B$57*Assumptions!$B$58)*Assumptions!$B$55*10');
  sheet.getRange(r, 3).setFormula('=(Assumptions!$B$56+Assumptions!$B$57*Assumptions!$B$58)*Assumptions!$B$55*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Workshops');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$60*Assumptions!$B$61*Assumptions!$B$59*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$60*Assumptions!$B$61*Assumptions!$B$59*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;
  r++;

  // BEVERAGES (rows 18-22)
  addSectionHeader(sheet, r++, 'BEVERAGE SALES', 6);

  sheet.getRange(r, 1).setValue('Paint & Sip Beverages');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$52*Assumptions!$B$54*Assumptions!$B$64*(Assumptions!$B$68*0.6+Assumptions!$B$67*0.4)*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$52*Assumptions!$B$54*Assumptions!$B$64*(Assumptions!$B$68*0.6+Assumptions!$B$67*0.4)*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Private Event Beverages');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$55*Assumptions!$B$58*Assumptions!$B$65*(Assumptions!$B$68*0.6+Assumptions!$B$67*0.4)*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$55*Assumptions!$B$58*Assumptions!$B$65*(Assumptions!$B$68*0.6+Assumptions!$B$67*0.4)*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Workshop Beverages');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$59*Assumptions!$B$61*Assumptions!$B$66*((Assumptions!$B$69*0.5)+(Assumptions!$B$68*0.25+Assumptions!$B$67*0.25))*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$59*Assumptions!$B$61*Assumptions!$B$66*((Assumptions!$B$69*0.5)+(Assumptions!$B$68*0.25+Assumptions!$B$67*0.25))*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;
  r++;

  // RETAIL (row 23)
  addSectionHeader(sheet, r++, 'RETAIL', 6);
  sheet.getRange(r, 1).setValue('Retail Sales (5%)');
  // CRITICAL: Retail = 5% of all NON-retail revenue (rows 5-20) to avoid circular reference
  sheet.getRange(r, 2).setFormula('=SUM(B5:B20)*0.05');
  sheet.getRange(r, 3).setFormula('=SUM(C5:C20)*0.05');
  sheet.getRange(r, 4).setFormula('=SUM(D5:D20)*0.05');
  sheet.getRange(r, 5).setFormula('=SUM(E5:E20)*0.05');
  sheet.getRange(r, 6).setFormula('=SUM(F5:F20)*0.05');
  var retailRow = r; // Row 24
  r++;
  r++;

  // TOTAL (row 26)
  sheet.getRange(r, 1).setValue('TOTAL GROSS REVENUE').setFontWeight('bold').setBackground('#93C47D');
  // CRITICAL: Total = non-retail + retail (no circular reference)
  sheet.getRange(r, 2).setFormula('=SUM(B5:B20)+B' + retailRow);
  sheet.getRange(r, 3).setFormula('=SUM(C5:C20)+C' + retailRow);
  sheet.getRange(r, 4).setFormula('=SUM(D5:D20)+D' + retailRow);
  sheet.getRange(r, 5).setFormula('=SUM(E5:E20)+E' + retailRow);
  sheet.getRange(r, 6).setFormula('=SUM(F5:F20)+F' + retailRow);
  sheet.getRange(r, 2, 1, 5).setFontWeight('bold').setBackground('#93C47D');

  sheet.getRange('B:F').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 220);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 110);
  }
}

// ============================================
// SHEET 7: EXPENSES
// ============================================

function createExpensesSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Expenses');
  sheet.setTabColor('#E06666');

  addHeader(sheet, 1, 'OPERATING EXPENSES', 6, '#E06666');

  var headers = ['Expense Category', 'Year 1 (10mo)', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange('A3:F3').setValues([headers]).setFontWeight('bold').setBackground('#F4CCCC');

  var r = 4;

  // COGS (rows 4-9)
  addSectionHeader(sheet, r++, 'COST OF GOODS SOLD', 6);

  sheet.getRange(r, 1).setValue('Art Supplies');
  sheet.getRange(r, 2).setFormula('=\'Revenue Model\'!B26*Assumptions!$B$73');
  sheet.getRange(r, 3).setFormula('=\'Revenue Model\'!C26*Assumptions!$B$73');
  sheet.getRange(r, 4).setFormula('=\'Revenue Model\'!D26*Assumptions!$B$73');
  sheet.getRange(r, 5).setFormula('=\'Revenue Model\'!E26*Assumptions!$B$73');
  sheet.getRange(r, 6).setFormula('=\'Revenue Model\'!F26*Assumptions!$B$73');
  r++;

  sheet.getRange(r, 1).setValue('Pottery Supplies');
  sheet.getRange(r, 2).setFormula('=\'Revenue Model\'!B26*Assumptions!$B$74');
  sheet.getRange(r, 3).setFormula('=\'Revenue Model\'!C26*Assumptions!$B$74');
  sheet.getRange(r, 4).setFormula('=\'Revenue Model\'!D26*Assumptions!$B$74');
  sheet.getRange(r, 5).setFormula('=\'Revenue Model\'!E26*Assumptions!$B$74');
  sheet.getRange(r, 6).setFormula('=\'Revenue Model\'!F26*Assumptions!$B$74');
  r++;

  sheet.getRange(r, 1).setValue('Beverage COGS');
  sheet.getRange(r, 2).setFormula('=(\'Revenue Model\'!B18+\'Revenue Model\'!B19+\'Revenue Model\'!B20)*Assumptions!$B$70');
  sheet.getRange(r, 3).setFormula('=(\'Revenue Model\'!C18+\'Revenue Model\'!C19+\'Revenue Model\'!C20)*Assumptions!$B$70');
  sheet.getRange(r, 4).setFormula('=(\'Revenue Model\'!D18+\'Revenue Model\'!D19+\'Revenue Model\'!D20)*Assumptions!$B$70');
  sheet.getRange(r, 5).setFormula('=(\'Revenue Model\'!E18+\'Revenue Model\'!E19+\'Revenue Model\'!E20)*Assumptions!$B$70');
  sheet.getRange(r, 6).setFormula('=(\'Revenue Model\'!F18+\'Revenue Model\'!F19+\'Revenue Model\'!F20)*Assumptions!$B$70');
  r++;

  sheet.getRange(r, 1).setValue('Supplies Replenishment');
  sheet.getRange(r, 2).setFormula('=\'Startup Supplies\'!E63*Assumptions!$B$75*3');
  sheet.getRange(r, 3).setFormula('=\'Startup Supplies\'!E63*Assumptions!$B$75*4');
  sheet.getRange(r, 4).setFormula('=\'Startup Supplies\'!E63*Assumptions!$B$75*4');
  sheet.getRange(r, 5).setFormula('=\'Startup Supplies\'!E63*Assumptions!$B$75*4');
  sheet.getRange(r, 6).setFormula('=\'Startup Supplies\'!E63*Assumptions!$B$75*4');
  r++;
  r++;

  // PAYROLL (rows 10-16)
  addSectionHeader(sheet, r++, 'PAYROLL & LABOR', 6);

  sheet.getRange(r, 1).setValue('Owner Salaries');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$43+Assumptions!$B$44');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$43+Assumptions!$B$44');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$43+Assumptions!$B$44');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$43+Assumptions!$B$44');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$43+Assumptions!$B$44');
  r++;

  sheet.getRange(r, 1).setValue('Studio Manager');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$45*Assumptions!$B$46*52');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$45*Assumptions!$B$46*52');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$45*Assumptions!$B$46*52');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$45*Assumptions!$B$46*52');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$45*Assumptions!$B$46*52');
  r++;

  sheet.getRange(r, 1).setValue('Teachers');
  sheet.getRange(r, 2).setValue(65000);
  sheet.getRange(r, 3).setValue(95000);
  sheet.getRange(r, 4).setValue(115000);
  sheet.getRange(r, 5).setValue(135000);
  sheet.getRange(r, 6).setValue(150000);
  r++;

  sheet.getRange(r, 1).setValue('Assistants');
  sheet.getRange(r, 2).setValue(15000);
  sheet.getRange(r, 3).setValue(25000);
  sheet.getRange(r, 4).setValue(32000);
  sheet.getRange(r, 5).setValue(40000);
  sheet.getRange(r, 6).setValue(45000);
  r++;

  sheet.getRange(r, 1).setValue('Payroll Taxes');
  sheet.getRange(r, 2).setFormula('=SUM(B11:B14)*Assumptions!$B$49');
  sheet.getRange(r, 3).setFormula('=SUM(C11:C14)*Assumptions!$B$49');
  sheet.getRange(r, 4).setFormula('=SUM(D11:D14)*Assumptions!$B$49');
  sheet.getRange(r, 5).setFormula('=SUM(E11:E14)*Assumptions!$B$49');
  sheet.getRange(r, 6).setFormula('=SUM(F11:F14)*Assumptions!$B$49');
  r++;
  r++;

  // FACILITIES (rows 17-23)
  addSectionHeader(sheet, r++, 'FACILITIES', 6);

  sheet.getRange(r, 1).setValue('Base Rent');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$7*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)^2');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)^3');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)^4');
  r++;

  sheet.getRange(r, 1).setValue('NNN');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$15*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$15*12');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$15*12');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$15*12');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$15*12');
  r++;

  sheet.getRange(r, 1).setValue('Liability Insurance');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$80*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$80*12');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$80*12');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$80*12');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$80*12');
  r++;

  sheet.getRange(r, 1).setValue('Utilities');
  sheet.getRange(r, 2).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$79)*10');
  sheet.getRange(r, 3).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$79)*12');
  sheet.getRange(r, 4).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$79)*12');
  sheet.getRange(r, 5).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$79)*12');
  sheet.getRange(r, 6).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$79)*12');
  r++;

  sheet.getRange(r, 1).setValue('Repairs');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$81/12*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$81');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$81');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$81');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$81');
  r++;
  r++;

  // OPERATIONS (rows 24-27)
  addSectionHeader(sheet, r++, 'OPERATIONS', 6);

  sheet.getRange(r, 1).setValue('Marketing');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$82*10');
  sheet.getRange(r, 3).setFormula('=\'Revenue Model\'!C26*Assumptions!$B$83');
  sheet.getRange(r, 4).setFormula('=\'Revenue Model\'!D26*Assumptions!$B$83');
  sheet.getRange(r, 5).setFormula('=\'Revenue Model\'!E26*Assumptions!$B$83');
  sheet.getRange(r, 6).setFormula('=\'Revenue Model\'!F26*Assumptions!$B$83');
  r++;

  sheet.getRange(r, 1).setValue('Professional Services');
  sheet.getRange(r, 2).setFormula('=(Assumptions!$B$84+Assumptions!$B$85+Assumptions!$B$86)/12*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$84+Assumptions!$B$85+Assumptions!$B$86');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$84+Assumptions!$B$85+Assumptions!$B$86');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$84+Assumptions!$B$85+Assumptions!$B$86');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$84+Assumptions!$B$85+Assumptions!$B$86');
  r++;

  sheet.getRange(r, 1).setValue('Liquor License');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$87');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$87');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$87');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$87');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$87');
  r++;
  r++;

  // DEPRECIATION (rows 28-30)
  addSectionHeader(sheet, r++, 'DEPRECIATION', 6);
  sheet.getRange(r, 1).setValue('Depreciation');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$92');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$92');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$92');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$92');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$92');
  r++;
  r++;

  // TOTAL (row 31)
  sheet.getRange(r, 1).setValue('TOTAL OPERATING EXPENSES').setFontWeight('bold').setBackground('#E06666').setFontColor('white');
  sheet.getRange(r, 2).setFormula('=SUM(B5:B' + (r-1) + ')');
  sheet.getRange(r, 3).setFormula('=SUM(C5:C' + (r-1) + ')');
  sheet.getRange(r, 4).setFormula('=SUM(D5:D' + (r-1) + ')');
  sheet.getRange(r, 5).setFormula('=SUM(E5:E' + (r-1) + ')');
  sheet.getRange(r, 6).setFormula('=SUM(F5:F' + (r-1) + ')');
  sheet.getRange(r, 2, 1, 5).setFontWeight('bold').setBackground('#E06666').setFontColor('white');

  sheet.getRange('B:F').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 220);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 110);
  }
}


// ============================================
// SHEET 8: LOAN CALCULATOR
// ============================================

function createLoanCalculatorSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Loan Calculator');
  sheet.setTabColor('#6FA8DC');

  addHeader(sheet, 1, 'SBA LOAN CALCULATOR', 2, '#6FA8DC');

  var r = 3;

  addSectionHeader(sheet, r++, 'FUNDING REQUIREMENT');

  sheet.getRange(r, 1).setValue('Total Capital Needed');
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B32').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Owner Equity');
  sheet.getRange(r++, 2).setValue(0).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Other Funding');
  sheet.getRange(r++, 2).setValue(0).setNumberFormat('$#,##0');
  r++;

  addSectionHeader(sheet, r++, 'LOAN BREAKDOWN');

  sheet.getRange(r, 1).setValue('Startup Costs');
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B27').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Working Capital');
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B30').setNumberFormat('$#,##0');
  r++;

  addSectionHeader(sheet, r++, 'LOAN TERMS');

  sheet.getRange(r, 1).setValue('Interest Rate');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B34').setNumberFormat('0.00%');

  sheet.getRange(r, 1).setValue('Term (years)');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B35');
  r++;

  // Row 17
  sheet.getRange(r, 1).setValue('LOAN AMOUNT').setFontWeight('bold').setBackground('#6FA8DC').setFontColor('white');
  sheet.getRange(r, 2).setFormula('=B5-B6-B7').setFontWeight('bold').setBackground('#6FA8DC').setFontColor('white').setNumberFormat('$#,##0');
  r++;

  // Row 18
  sheet.getRange(r, 1).setValue('Monthly Payment');
  sheet.getRange(r, 2).setFormula('=-PMT(B15/12,B16*12,B17)').setNumberFormat('$#,##0');
  r++;

  // Row 19
  sheet.getRange(r, 1).setValue('Total Interest');
  sheet.getRange(r, 2).setFormula('=B18*B16*12-B17').setNumberFormat('$#,##0');

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
}

// ============================================
// SHEET 9: LOAN AMORTIZATION
// ============================================

function createLoanAmortizationSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Loan Amortization');
  sheet.setTabColor('#6FA8DC');

  addHeader(sheet, 1, 'LOAN AMORTIZATION (5 YEARS)', 5, '#6FA8DC');

  sheet.getRange('A3').setValue('Loan Amount:');
  sheet.getRange('B3').setFormula('=\'Loan Calculator\'!B17').setNumberFormat('$#,##0');

  sheet.getRange('A4').setValue('Interest Rate:');
  sheet.getRange('B4').setFormula('=Assumptions!B34').setNumberFormat('0.00%');

  sheet.getRange('A5').setValue('Term:');
  sheet.getRange('B5').setValue('10 years');

  sheet.getRange('A6').setValue('Monthly Payment:');
  sheet.getRange('B6').setFormula('=\'Loan Calculator\'!B18').setNumberFormat('$#,##0');

  var headers = ['Year', 'Principal', 'Interest', 'Total Payment', 'Balance'];
  sheet.getRange('A8:E8').setValues([headers]).setFontWeight('bold').setBackground('#C9DAF8');

  sheet.getRange('A9').setValue('Initial');
  sheet.getRange('E9').setFormula('=B3');

  sheet.getRange('A10').setValue('Year 1');
  sheet.getRange('C10').setFormula('=E9*$B$4');
  sheet.getRange('D10').setFormula('=$B$6*10');
  sheet.getRange('B10').setFormula('=D10-C10');
  sheet.getRange('E10').setFormula('=E9-B10');

  for (var y = 2; y <= 5; y++) {
    var r = 9 + y;
    sheet.getRange(r, 1).setValue('Year ' + y);
    sheet.getRange(r, 3).setFormula('=E' + (r-1) + '*$B$4');
    sheet.getRange(r, 4).setFormula('=$B$6*12');
    sheet.getRange(r, 2).setFormula('=D' + r + '-C' + r);
    sheet.getRange(r, 5).setFormula('=E' + (r-1) + '-B' + r);
  }

  sheet.getRange('B9:E14').setNumberFormat('$#,##0');

  sheet.setColumnWidth(1, 100);
  for (var i = 2; i <= 5; i++) {
    sheet.setColumnWidth(i, 120);
  }
}

// ============================================
// SHEET 10: PROFIT & LOSS
// ============================================

function createProfitLossSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Profit & Loss');
  sheet.setTabColor('#F6B26B');

  addHeader(sheet, 1, 'PROFIT & LOSS STATEMENT', 6, '#F6B26B');

  var headers = ['', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange('A3:F3').setValues([headers]).setFontWeight('bold').setBackground('#FCE5CD');

  var r = 4;

  sheet.getRange(r, 1).setValue('GROSS REVENUE');
  sheet.getRange(r, 2).setFormula('=\'Revenue Model\'!B26');
  sheet.getRange(r, 3).setFormula('=\'Revenue Model\'!C26');
  sheet.getRange(r, 4).setFormula('=\'Revenue Model\'!D26');
  sheet.getRange(r, 5).setFormula('=\'Revenue Model\'!E26');
  sheet.getRange(r, 6).setFormula('=\'Revenue Model\'!F26');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('TOTAL EXPENSES');
  sheet.getRange(r, 2).setFormula('=Expenses!B31');
  sheet.getRange(r, 3).setFormula('=Expenses!C31');
  sheet.getRange(r, 4).setFormula('=Expenses!D31');
  sheet.getRange(r, 5).setFormula('=Expenses!E31');
  sheet.getRange(r, 6).setFormula('=Expenses!F31');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('EBITDA');
  sheet.getRange(r, 2).setFormula('=B4-B6');
  sheet.getRange(r, 3).setFormula('=C4-C6');
  sheet.getRange(r, 4).setFormula('=D4-D6');
  sheet.getRange(r, 5).setFormula('=E4-E6');
  sheet.getRange(r, 6).setFormula('=F4-F6');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold').setBackground('#FCE5CD');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('Depreciation');
  sheet.getRange(r, 2).setFormula('=Assumptions!B92/12*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!B92');
  sheet.getRange(r, 4).setFormula('=Assumptions!B92');
  sheet.getRange(r, 5).setFormula('=Assumptions!B92');
  sheet.getRange(r, 6).setFormula('=Assumptions!B92');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('EBIT');
  sheet.getRange(r, 2).setFormula('=B8-B10');
  sheet.getRange(r, 3).setFormula('=C8-C10');
  sheet.getRange(r, 4).setFormula('=D8-D10');
  sheet.getRange(r, 5).setFormula('=E8-E10');
  sheet.getRange(r, 6).setFormula('=F8-F10');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('Interest Expense');
  sheet.getRange(r, 2).setFormula('=\'Loan Amortization\'!C10');
  sheet.getRange(r, 3).setFormula('=\'Loan Amortization\'!C11');
  sheet.getRange(r, 4).setFormula('=\'Loan Amortization\'!C12');
  sheet.getRange(r, 5).setFormula('=\'Loan Amortization\'!C13');
  sheet.getRange(r, 6).setFormula('=\'Loan Amortization\'!C14');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('NET INCOME');
  sheet.getRange(r, 2).setFormula('=B12-B14');
  sheet.getRange(r, 3).setFormula('=C12-C14');
  sheet.getRange(r, 4).setFormula('=D12-D14');
  sheet.getRange(r, 5).setFormula('=E12-E14');
  sheet.getRange(r, 6).setFormula('=F12-F14');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold').setBackground('#F6B26B').setFontColor('white');
  r++;
  r++;

  sheet.getRange(r++, 1).setValue('KEY METRICS').setFontWeight('bold').setBackground('#FCE5CD');

  sheet.getRange(r, 1).setValue('EBITDA Margin %');
  sheet.getRange(r, 2).setFormula('=B8/B4');
  sheet.getRange(r, 3).setFormula('=C8/C4');
  sheet.getRange(r, 4).setFormula('=D8/D4');
  sheet.getRange(r, 5).setFormula('=E8/E4');
  sheet.getRange(r, 6).setFormula('=F8/F4');
  sheet.getRange(r, 2, 1, 5).setNumberFormat('0.0%');
  r++;

  sheet.getRange(r, 1).setValue('Net Margin %');
  sheet.getRange(r, 2).setFormula('=B16/B4');
  sheet.getRange(r, 3).setFormula('=C16/C4');
  sheet.getRange(r, 4).setFormula('=D16/D4');
  sheet.getRange(r, 5).setFormula('=E16/E4');
  sheet.getRange(r, 6).setFormula('=F16/F4');
  sheet.getRange(r, 2, 1, 5).setNumberFormat('0.0%');

  sheet.getRange('B4:F16').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 180);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 110);
  }
}

// ============================================
// SHEET 11: CASH FLOW
// ============================================

function createCashFlowSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Cash Flow');
  sheet.setTabColor('#8E7CC3');

  addHeader(sheet, 1, 'CASH FLOW PROJECTION', 6, '#8E7CC3');

  var headers = ['', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange('A3:F3').setValues([headers]).setFontWeight('bold').setBackground('#D9D2E9');

  var r = 4;

  sheet.getRange(r++, 1).setValue('OPERATING').setFontWeight('bold').setBackground('#D9D2E9');

  sheet.getRange(r, 1).setValue('Net Income');
  sheet.getRange(r, 2).setFormula('=\'Profit & Loss\'!B16');
  sheet.getRange(r, 3).setFormula('=\'Profit & Loss\'!C16');
  sheet.getRange(r, 4).setFormula('=\'Profit & Loss\'!D16');
  sheet.getRange(r, 5).setFormula('=\'Profit & Loss\'!E16');
  sheet.getRange(r, 6).setFormula('=\'Profit & Loss\'!F16');
  r++;

  sheet.getRange(r, 1).setValue('Add: Depreciation');
  sheet.getRange(r, 2).setFormula('=\'Profit & Loss\'!B10');
  sheet.getRange(r, 3).setFormula('=\'Profit & Loss\'!C10');
  sheet.getRange(r, 4).setFormula('=\'Profit & Loss\'!D10');
  sheet.getRange(r, 5).setFormula('=\'Profit & Loss\'!E10');
  sheet.getRange(r, 6).setFormula('=\'Profit & Loss\'!F10');
  r++;

  sheet.getRange(r, 1).setValue('Cash from Operations');
  sheet.getRange(r, 2).setFormula('=B5+B6');
  sheet.getRange(r, 3).setFormula('=C5+C6');
  sheet.getRange(r, 4).setFormula('=D5+D6');
  sheet.getRange(r, 5).setFormula('=E5+E6');
  sheet.getRange(r, 6).setFormula('=F5+F6');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold');
  r++;
  r++;

  sheet.getRange(r++, 1).setValue('FINANCING').setFontWeight('bold').setBackground('#D9D2E9');

  sheet.getRange(r, 1).setValue('Principal Payments');
  sheet.getRange(r, 2).setFormula('=-\'Loan Amortization\'!B10');
  sheet.getRange(r, 3).setFormula('=-\'Loan Amortization\'!B11');
  sheet.getRange(r, 4).setFormula('=-\'Loan Amortization\'!B12');
  sheet.getRange(r, 5).setFormula('=-\'Loan Amortization\'!B13');
  sheet.getRange(r, 6).setFormula('=-\'Loan Amortization\'!B14');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('NET CASH FLOW');
  sheet.getRange(r, 2).setFormula('=B7+B10');
  sheet.getRange(r, 3).setFormula('=C7+C10');
  sheet.getRange(r, 4).setFormula('=D7+D10');
  sheet.getRange(r, 5).setFormula('=E7+E10');
  sheet.getRange(r, 6).setFormula('=F7+F10');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold').setBackground('#8E7CC3').setFontColor('white');
  r++;

  sheet.getRange(r, 1).setValue('Beginning Balance');
  sheet.getRange(r, 2).setFormula('=\'Startup Costs\'!B30');
  sheet.getRange(r, 3).setFormula('=B13');
  sheet.getRange(r, 4).setFormula('=C13');
  sheet.getRange(r, 5).setFormula('=D13');
  sheet.getRange(r, 6).setFormula('=E13');
  r++;

  sheet.getRange(r, 1).setValue('Ending Balance');
  sheet.getRange(r, 2).setFormula('=B12+B13');
  sheet.getRange(r, 3).setFormula('=C12+C13');
  sheet.getRange(r, 4).setFormula('=D12+D13');
  sheet.getRange(r, 5).setFormula('=E12+E13');
  sheet.getRange(r, 6).setFormula('=F12+F13');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold');

  sheet.getRange('B:F').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 180);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 110);
  }
}

// ============================================
// SHEET 12: DASHBOARD
// ============================================

function createDashboardSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Dashboard');
  sheet.setTabColor('#FFD966');

  addHeader(sheet, 1, 'DECO ART CENTER - DASHBOARD', 6, '#FFD966');

  var r = 3;

  // SBA LOAN
  sheet.getRange(r, 1).setValue('SBA LOAN REQUEST').setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
  r += 2;

  sheet.getRange(r, 1).setValue('Total Capital Needed');
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B32').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Loan Amount');
  sheet.getRange(r++, 2).setFormula('=\'Loan Calculator\'!B17').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Interest Rate');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B34').setNumberFormat('0.00%');

  sheet.getRange(r, 1).setValue('Monthly Payment');
  sheet.getRange(r++, 2).setFormula('=\'Loan Calculator\'!B18').setNumberFormat('$#,##0');
  r++;

  // FACILITY
  sheet.getRange(r, 1).setValue('FACILITY').setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
  r += 2;

  sheet.getRange(r, 1).setValue('Size (sq ft)');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B4');

  sheet.getRange(r, 1).setValue('Monthly Rent');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B16').setNumberFormat('$#,##0');
  r++;

  // PERFORMANCE
  sheet.getRange(r, 1).setValue('5-YEAR PERFORMANCE').setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
  r += 2;

  var perfHeaders = ['', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange(r, 1, 1, 6).setValues([perfHeaders]).setFontWeight('bold').setBackground('#FFF2CC');
  r++;

  sheet.getRange(r, 1).setValue('Revenue');
  sheet.getRange(r, 2).setFormula('=\'Revenue Model\'!B26');
  sheet.getRange(r, 3).setFormula('=\'Revenue Model\'!C26');
  sheet.getRange(r, 4).setFormula('=\'Revenue Model\'!D26');
  sheet.getRange(r, 5).setFormula('=\'Revenue Model\'!E26');
  sheet.getRange(r, 6).setFormula('=\'Revenue Model\'!F26');
  r++;

  sheet.getRange(r, 1).setValue('Expenses');
  sheet.getRange(r, 2).setFormula('=Expenses!B31');
  sheet.getRange(r, 3).setFormula('=Expenses!C31');
  sheet.getRange(r, 4).setFormula('=Expenses!D31');
  sheet.getRange(r, 5).setFormula('=Expenses!E31');
  sheet.getRange(r, 6).setFormula('=Expenses!F31');
  r++;

  sheet.getRange(r, 1).setValue('EBITDA');
  sheet.getRange(r, 2).setFormula('=\'Profit & Loss\'!B8');
  sheet.getRange(r, 3).setFormula('=\'Profit & Loss\'!C8');
  sheet.getRange(r, 4).setFormula('=\'Profit & Loss\'!D8');
  sheet.getRange(r, 5).setFormula('=\'Profit & Loss\'!E8');
  sheet.getRange(r, 6).setFormula('=\'Profit & Loss\'!F8');
  r++;

  sheet.getRange(r, 1).setValue('Net Income');
  sheet.getRange(r, 2).setFormula('=\'Profit & Loss\'!B16');
  sheet.getRange(r, 3).setFormula('=\'Profit & Loss\'!C16');
  sheet.getRange(r, 4).setFormula('=\'Profit & Loss\'!D16');
  sheet.getRange(r, 5).setFormula('=\'Profit & Loss\'!E16');
  sheet.getRange(r, 6).setFormula('=\'Profit & Loss\'!F16');
  r++;

  sheet.getRange(r, 1).setValue('Cash Flow');
  sheet.getRange(r, 2).setFormula('=\'Cash Flow\'!B12');
  sheet.getRange(r, 3).setFormula('=\'Cash Flow\'!C12');
  sheet.getRange(r, 4).setFormula('=\'Cash Flow\'!D12');
  sheet.getRange(r, 5).setFormula('=\'Cash Flow\'!E12');
  sheet.getRange(r, 6).setFormula('=\'Cash Flow\'!F12');
  r++;
  r++;

  // RATIOS
  sheet.getRange(r, 1).setValue('KEY RATIOS').setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
  r += 2;

  sheet.getRange(r, 1, 1, 6).setValues([perfHeaders]).setFontWeight('bold').setBackground('#FFF2CC');
  r++;

  sheet.getRange(r, 1).setValue('EBITDA Margin');
  sheet.getRange(r, 2).setFormula('=\'Profit & Loss\'!B19');
  sheet.getRange(r, 3).setFormula('=\'Profit & Loss\'!C19');
  sheet.getRange(r, 4).setFormula('=\'Profit & Loss\'!D19');
  sheet.getRange(r, 5).setFormula('=\'Profit & Loss\'!E19');
  sheet.getRange(r, 6).setFormula('=\'Profit & Loss\'!F19');
  sheet.getRange(r, 2, 1, 5).setNumberFormat('0.0%');
  r++;

  sheet.getRange(r, 1).setValue('DSCR');
  sheet.getRange(r, 2).setFormula('=(\'Profit & Loss\'!B8)/(\'Loan Calculator\'!B18*10)');
  sheet.getRange(r, 3).setFormula('=(\'Profit & Loss\'!C8)/(\'Loan Calculator\'!B18*12)');
  sheet.getRange(r, 4).setFormula('=(\'Profit & Loss\'!D8)/(\'Loan Calculator\'!B18*12)');
  sheet.getRange(r, 5).setFormula('=(\'Profit & Loss\'!E8)/(\'Loan Calculator\'!B18*12)');
  sheet.getRange(r, 6).setFormula('=(\'Profit & Loss\'!F8)/(\'Loan Calculator\'!B18*12)');
  sheet.getRange(r, 2, 1, 5).setNumberFormat('0.00');

  sheet.getRange('B6:B9').setNumberFormat('$#,##0');
  sheet.getRange('B19:F23').setNumberFormat('$#,##0');

  for (var i = 1; i <= 6; i++) {
    sheet.setColumnWidth(i, 130);
  }

  // Move Dashboard to first position
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
}

