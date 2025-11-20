/**
 * DECO ART CENTER - RENTAL MODEL PRO FORMA
 * Version: Rental Model with SBA Loan Calculator
 * Created: November 2025
 *
 * INSTALLATION:
 * 1. Open new Google Sheet (or use existing)
 * 2. Extensions → Apps Script
 * 3. Copy ALL this code
 * 4. Save and Run createDECOProForma()
 * 5. Authorize when prompted
 * 6. Wait 60-90 seconds for completion
 *
 * Can be run multiple times safely - will update existing sheets
 */

function createDECOProForma() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create sheets in dependency order (critical!)
  createAssumptionsSheet(ss);
  createClassScheduleSheet(ss);
  createSummerCampsSheet(ss);
  createStartupSuppliesSheet(ss);
  createRevenueModelSheet(ss);
  createExpensesSheet(ss);
  createStartupCostsSheet(ss);
  createLoanCalculatorSheet(ss);
  createLoanAmortizationSheet(ss);
  createProfitLossSheet(ss);
  createCashFlowSheet(ss);
  createDashboardSheet(ss);

  SpreadsheetApp.getUi().alert('DECO Pro Forma Created Successfully!\n\nStart with Dashboard, then review Assumptions.\nEdit Class Schedule and Summer Camps for your specific offerings.');
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    // Sheet exists - clear it
    sheet.clear();
  } else {
    // Sheet doesn't exist - create it
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function addHeader(sheet, row, title, cols, color) {
  sheet.getRange(row, 1, 1, cols).merge()
    .setValue(title)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(color).setFontColor('white');
}

function addSectionHeader(sheet, row, title, cols) {
  cols = cols || 2;
  sheet.getRange(row, 1, 1, cols).merge()
    .setValue(title)
    .setFontWeight('bold')
    .setBackground('#D9EAD3')
    .setFontSize(11);
}

function addRow(sheet, row, label, values) {
  sheet.getRange(row, 1).setValue(label);
  for (var i = 0; i < values.length; i++) {
    var cell = sheet.getRange(row, i + 2);
    if (typeof values[i] === 'string' && values[i].startsWith('=')) {
      cell.setFormula(values[i]);
    } else {
      cell.setValue(values[i]);
    }
  }
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
  addSectionHeader(sheet, r++, 'FACILITY INFORMATION (RENTAL)');
  sheet.getRange(r, 1).setValue('Building Size (sq ft)'); sheet.getRange(r++, 2).setValue(5000);
  sheet.getRange(r, 1).setValue('Base Rent ($/sf/year)'); sheet.getRange(r++, 2).setValue(16).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Annual Base Rent'); sheet.getRange(r, 2).setFormula('=B4*B5').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Monthly Base Rent'); sheet.getRange(r, 2).setFormula('=B6/12').setNumberFormat('$#,##0'); r++;
  r++;

  // NNN BREAKDOWN
  addSectionHeader(sheet, r++, 'NNN (TRIPLE NET) COSTS - TULSA, OK 74119');
  sheet.getRange(r, 1).setValue('Property Tax ($/sf/year)'); sheet.getRange(r++, 2).setValue(1.50).setNumberFormat('$#,##0.00');
  sheet.getRange(r, 1).setValue('Property Insurance ($/sf/year)'); sheet.getRange(r++, 2).setValue(1.20).setNumberFormat('$#,##0.00');
  sheet.getRange(r, 1).setValue('CAM - Common Area Maint ($/sf/year)'); sheet.getRange(r++, 2).setValue(1.30).setNumberFormat('$#,##0.00');
  sheet.getRange(r, 1).setValue('Total NNN ($/sf/year)'); sheet.getRange(r, 2).setFormula('=B10+B11+B12').setNumberFormat('$#,##0.00'); r++;
  sheet.getRange(r, 1).setValue('Annual NNN Cost'); sheet.getRange(r, 2).setFormula('=B13*B4').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Monthly NNN Cost'); sheet.getRange(r, 2).setFormula('=B14/12').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Total Monthly Rent (Base + NNN)'); sheet.getRange(r, 2).setFormula('=B7+B15').setNumberFormat('$#,##0'); r++;
  r++;

  // LEASE TERMS
  addSectionHeader(sheet, r++, 'LEASE TERMS');
  sheet.getRange(r, 1).setValue('Lease Length (years)'); sheet.getRange(r++, 2).setValue(5);
  sheet.getRange(r, 1).setValue('Annual Rent Escalation %'); sheet.getRange(r++, 2).setValue(0.03).setNumberFormat('0.00%');
  sheet.getRange(r, 1).setValue('Security Deposit (months)'); sheet.getRange(r++, 2).setValue(1);
  sheet.getRange(r, 1).setValue('Security Deposit Amount'); sheet.getRange(r, 2).setFormula('=B7*B21').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('First Month Rent'); sheet.getRange(r, 2).setFormula('=B16').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Last Month Rent'); sheet.getRange(r, 2).setFormula('=B16').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Total Move-in Cost'); sheet.getRange(r, 2).setFormula('=B22+B23+B24').setNumberFormat('$#,##0'); r++;
  r++;

  // TIMELINE
  addSectionHeader(sheet, r++, 'TIMELINE');
  sheet.getRange(r, 1).setValue('Opening Date'); sheet.getRange(r++, 2).setValue('March 2026');
  sheet.getRange(r, 1).setValue('Year 1 Operating Months'); sheet.getRange(r++, 2).setValue(10);
  sheet.getRange(r, 1).setValue('Class Sessions Year 1'); sheet.getRange(r++, 2).setValue(6);
  sheet.getRange(r, 1).setValue('Class Sessions Year 2+'); sheet.getRange(r++, 2).setValue(12);
  r++;

  // SBA LOAN
  addSectionHeader(sheet, r++, 'SBA LOAN TERMS');
  sheet.getRange(r, 1).setValue('Interest Rate'); sheet.getRange(r++, 2).setValue(0.12).setNumberFormat('0.00%');
  sheet.getRange(r, 1).setValue('Loan Term (years)'); sheet.getRange(r++, 2).setValue(10);
  sheet.getRange(r, 1).setValue('Loan Amount (auto-calculated)'); sheet.getRange(r, 2).setFormula('=\'Loan Calculator\'!B18').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Monthly Payment'); sheet.getRange(r, 2).setFormula('=-PMT(B34/12,B35*12,B36)').setNumberFormat('$#,##0'); r++;
  r++;

  // CAPACITY
  addSectionHeader(sheet, r++, 'ENROLLMENT CAPACITY');
  sheet.getRange(r, 1).setValue('Year 1 Capacity %'); sheet.getRange(r++, 2).setValue(0.50).setNumberFormat('0%');
  sheet.getRange(r, 1).setValue('Annual Growth %'); sheet.getRange(r++, 2).setValue(0.20).setNumberFormat('0%');
  sheet.getRange(r, 1).setValue('Drop-in Rate %'); sheet.getRange(r++, 2).setValue(0.05).setNumberFormat('0%');
  r++;

  // STAFF
  addSectionHeader(sheet, r++, 'STAFF COMPENSATION');
  sheet.getRange(r, 1).setValue('Owner 1 (Grant) Annual Salary'); sheet.getRange(r++, 2).setValue(50000).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Owner 2 (Andrea) Annual Salary'); sheet.getRange(r++, 2).setValue(50000).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Studio Manager Hourly'); sheet.getRange(r++, 2).setValue(25).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Studio Manager Hours/Week'); sheet.getRange(r++, 2).setValue(40);
  sheet.getRange(r, 1).setValue('Teacher Hourly'); sheet.getRange(r++, 2).setValue(25).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Assistant Hourly'); sheet.getRange(r++, 2).setValue(15).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Payroll Tax Rate'); sheet.getRange(r++, 2).setValue(0.0765).setNumberFormat('0.00%');
  r++;

  // EVENTS
  addSectionHeader(sheet, r++, 'EVENTS & WORKSHOPS');
  sheet.getRange(r, 1).setValue('Paint & Sip Events/Month'); sheet.getRange(r++, 2).setValue(8);
  sheet.getRange(r, 1).setValue('Paint & Sip Price'); sheet.getRange(r++, 2).setValue(45).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Paint & Sip Attendance'); sheet.getRange(r++, 2).setValue(15);
  sheet.getRange(r, 1).setValue('Private Events/Month'); sheet.getRange(r++, 2).setValue(1);
  sheet.getRange(r, 1).setValue('Private Event Base Fee'); sheet.getRange(r++, 2).setValue(1500).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Private Event Per Person'); sheet.getRange(r++, 2).setValue(20).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Private Event Attendance'); sheet.getRange(r++, 2).setValue(25);
  sheet.getRange(r, 1).setValue('Workshops/Month'); sheet.getRange(r++, 2).setValue(1);
  sheet.getRange(r, 1).setValue('Workshop Price'); sheet.getRange(r++, 2).setValue(150).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Workshop Attendance'); sheet.getRange(r++, 2).setValue(20);
  r++;

  // BEVERAGES
  addSectionHeader(sheet, r++, 'BEVERAGE SALES');
  sheet.getRange(r, 1).setValue('Drinks/Person (Paint & Sip)'); sheet.getRange(r++, 2).setValue(2);
  sheet.getRange(r, 1).setValue('Drinks/Person (Private Event)'); sheet.getRange(r++, 2).setValue(2);
  sheet.getRange(r, 1).setValue('Drinks/Person (Workshop)'); sheet.getRange(r++, 2).setValue(1.5);
  sheet.getRange(r, 1).setValue('Coffee Price'); sheet.getRange(r++, 2).setValue(4).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Beer Price'); sheet.getRange(r++, 2).setValue(6).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Wine Price'); sheet.getRange(r++, 2).setValue(8).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Beverage COGS %'); sheet.getRange(r++, 2).setValue(0.30).setNumberFormat('0%');
  r++;

  // EXPENSES
  addSectionHeader(sheet, r++, 'OPERATING EXPENSES');
  sheet.getRange(r, 1).setValue('Art Supplies COGS %'); sheet.getRange(r++, 2).setValue(0.0756).setNumberFormat('0.00%');
  sheet.getRange(r, 1).setValue('Pottery COGS %'); sheet.getRange(r++, 2).setValue(0.0202).setNumberFormat('0.00%');
  sheet.getRange(r, 1).setValue('Supplies Replenish % (Quarterly)'); sheet.getRange(r++, 2).setValue(0.30).setNumberFormat('0%');
  sheet.getRange(r, 1).setValue('Electricity/Month'); sheet.getRange(r++, 2).setValue(800).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Gas/Month'); sheet.getRange(r++, 2).setValue(250).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Water/Month'); sheet.getRange(r++, 2).setValue(200).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Internet/Month'); sheet.getRange(r++, 2).setValue(150).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Liability Insurance/Month'); sheet.getRange(r++, 2).setValue(800).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Repairs/Maintenance per Year'); sheet.getRange(r++, 2).setValue(5000).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Marketing Year 1/Month'); sheet.getRange(r++, 2).setValue(1500).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Marketing Year 2+ % of Revenue'); sheet.getRange(r++, 2).setValue(0.03).setNumberFormat('0.00%');
  sheet.getRange(r, 1).setValue('Accounting/Year'); sheet.getRange(r++, 2).setValue(3600).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Legal/Year'); sheet.getRange(r++, 2).setValue(2400).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Payroll Service/Year'); sheet.getRange(r++, 2).setValue(1800).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Liquor License/Year'); sheet.getRange(r++, 2).setValue(1030).setNumberFormat('$#,##0');
  r++;

  // DEPRECIATION
  addSectionHeader(sheet, r++, 'DEPRECIATION');
  sheet.getRange(r, 1).setValue('Leasehold Improvements (15 years)'); sheet.getRange(r, 2).setFormula('=\'Startup Costs\'!B9/15').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Equipment (7 years)'); sheet.getRange(r, 2).setFormula('=SUM(\'Startup Costs\'!B10:B13)/7').setNumberFormat('$#,##0'); r++;
  sheet.getRange(r, 1).setValue('Total Annual Depreciation'); sheet.getRange(r, 2).setFormula('=B96+B97').setNumberFormat('$#,##0'); r++;

  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 150);

  // Notes
  sheet.getRange('D1').setValue('EDITABLE INPUTS').setFontWeight('bold').setFontSize(12);
  sheet.getRange('D2').setValue('All values can be edited');
  sheet.getRange('D3').setValue('Blue cells = formulas (auto-calculate)');
  sheet.getRange('D4').setValue('White cells = your inputs');
  sheet.getRange('D5').setValue('Loan amount auto-calculates from Loan Calculator');
}

// ============================================
// SHEET 2: CLASS SCHEDULE
// ============================================

function createClassScheduleSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Class Schedule');
  sheet.setTabColor('#93C47D');

  addHeader(sheet, 1, 'REGULAR CLASS SCHEDULE', 8, '#93C47D');

  var headers = ['Category', 'Duration (hrs)', 'Price', 'Max Students', 'Sessions/Year', 'Year 1 Enroll %', 'Year 2+', 'Notes'];
  sheet.getRange('A3:H3').setValues([headers]).setFontWeight('bold').setBackground('#D9EAD3');

  var classes = [
    ['Adults - Painting', 2, 199, 12, 12, '=Assumptions!$B$40', '=F4*(1+Assumptions!$B$41)', '2hr session'],
    ['Adults - Pottery', 2.5, 229, 6, 12, '=Assumptions!$B$40', '=F5*(1+Assumptions!$B$41)', '2.5hr session'],
    ['Teens 13-18 - Drawing', 2, 199, 15, 12, '=Assumptions!$B$40', '=F6*(1+Assumptions!$B$41)', '2hr session'],
    ['Teens 13-18 - Pottery', 2, 229, 6, 12, '=Assumptions!$B$40', '=F7*(1+Assumptions!$B$41)', '2hr session'],
    ['Kids 9-12 - Mixed Media', 1.5, 159, 12, 12, '=Assumptions!$B$40', '=F8*(1+Assumptions!$B$41)', '1.5hr session'],
    ['Kids 9-12 - Pottery', 1.5, 179, 6, 12, '=Assumptions!$B$40', '=F9*(1+Assumptions!$B$41)', '1.5hr session'],
    ['Kids 6-8 - General Art', 1, 129, 12, 12, '=Assumptions!$B$40', '=F10*(1+Assumptions!$B$41)', '1hr session'],
    ['Kids 4-5 - General Art', 1, 129, 12, 12, '=Assumptions!$B$40', '=F11*(1+Assumptions!$B$41)', '1hr session'],
    ['Homeschool 6-12', 1.5, 154, 12, 12, '=Assumptions!$B$40', '=F12*(1+Assumptions!$B$41)', '1.5hr session'],
    ['Homeschool Teens', 1.5, 154, 12, 12, '=Assumptions!$B$40', '=F13*(1+Assumptions!$B$41)', '1.5hr session']
  ];

  for (var i = 0; i < classes.length; i++) {
    sheet.getRange(4 + i, 1, 1, 8).setValues([classes[i]]);
  }

  sheet.getRange('C4:C13').setNumberFormat('$#,##0');
  sheet.getRange('F4:G13').setNumberFormat('0%');

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 150);

  sheet.getRange('A15').setValue('INSTRUCTIONS:').setFontWeight('bold');
  sheet.getRange('A16').setValue('• Edit Max Students, Price, or Sessions as needed');
  sheet.getRange('A17').setValue('• Enrollment % auto-calculates from Assumptions');
  sheet.getRange('A18').setValue('• Add rows if you need more class types');
  sheet.getRange('A19').setValue('• Year 1 = 6 sessions (Mar-Apr, Sep-Dec)');
  sheet.getRange('A20').setValue('• Year 2+ = 12 sessions (full year)');
}

// ============================================
// SHEET 3: SUMMER CAMPS
// ============================================

function createSummerCampsSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Summer Camps');
  sheet.setTabColor('#F6B26B');

  addHeader(sheet, 1, 'SUMMER CAMP SCHEDULE', 9, '#F6B26B');

  var headers = ['Age Group', 'Duration', 'Days', 'Type', 'Price', 'Max Students', 'Year 1 Enroll %', 'Year 2+', 'Weeks Offered'];
  sheet.getRange('A3:I3').setValues([headers]).setFontWeight('bold').setBackground('#FCE5CD');

  var camps = [
    ['5-7', 'Half Day', 5, 'General', 159, 12, '=Assumptions!$B$40', '=G4*(1+Assumptions!$B$41)', 4],
    ['5-7', 'Half Day', 5, 'Ceramics', 179, 8, '=Assumptions!$B$40', '=G5*(1+Assumptions!$B$41)', 3],
    ['5-7', 'Half Day', 4, 'General', '=159*0.85', 12, '=Assumptions!$B$40', '=G6*(1+Assumptions!$B$41)', 2],
    ['5-7', 'Full Day', 5, 'General', 275, 12, '=Assumptions!$B$40', '=G7*(1+Assumptions!$B$41)', 4],
    ['5-7', 'Full Day', 5, 'Ceramics', 295, 8, '=Assumptions!$B$40', '=G8*(1+Assumptions!$B$41)', 2],
    ['8-12', 'Half Day', 5, 'General', 159, 12, '=Assumptions!$B$40', '=G9*(1+Assumptions!$B$41)', 6],
    ['8-12', 'Half Day', 5, 'Ceramics', 179, 8, '=Assumptions!$B$40', '=G10*(1+Assumptions!$B$41)', 4],
    ['8-12', 'Half Day', 4, 'General', '=159*0.85', 12, '=Assumptions!$B$40', '=G11*(1+Assumptions!$B$41)', 3],
    ['8-12', 'Full Day', 5, 'General', 275, 12, '=Assumptions!$B$40', '=G12*(1+Assumptions!$B$41)', 6],
    ['8-12', 'Full Day', 5, 'Ceramics', 295, 8, '=Assumptions!$B$40', '=G13*(1+Assumptions!$B$41)', 4],
    ['8-12', 'Full Day', 4, 'General', '=275*0.85', 12, '=Assumptions!$B$40', '=G14*(1+Assumptions!$B$41)', 2],
    ['12-16', 'Half Day', 5, 'General', 159, 15, '=Assumptions!$B$40', '=G15*(1+Assumptions!$B$41)', 4],
    ['12-16', 'Half Day', 5, 'Ceramics', 179, 10, '=Assumptions!$B$40', '=G16*(1+Assumptions!$B$41)', 3],
    ['12-16', 'Full Day', 5, 'General', 275, 15, '=Assumptions!$B$40', '=G17*(1+Assumptions!$B$41)', 4],
    ['12-16', 'Full Day', 5, 'Ceramics', 295, 10, '=Assumptions!$B$40', '=G18*(1+Assumptions!$B$41)', 3],
    ['All Ages', 'Extended Care', 5, 'Add-on', 50, 20, '=Assumptions!$B$40', '=G19*(1+Assumptions!$B$41)', 12]
  ];

  for (var i = 0; i < camps.length; i++) {
    sheet.getRange(4 + i, 1, 1, 9).setValues([camps[i]]);
  }

  sheet.getRange('E4:E19').setNumberFormat('$#,##0');
  sheet.getRange('G4:H19').setNumberFormat('0%');

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 110);

  sheet.getRange('A21').setValue('INSTRUCTIONS:').setFontWeight('bold');
  sheet.getRange('A22').setValue('• Edit prices, max students, or weeks offered');
  sheet.getRange('A23').setValue('• Enrollment % auto-updates from Assumptions');
  sheet.getRange('A24').setValue('• Add rows for additional camp types');
  sheet.getRange('A25').setValue('• Up to 6 camps can run simultaneously (one per classroom)');
}

// ============================================
// SHEET 4: STARTUP SUPPLIES
// ============================================

function createStartupSuppliesSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Startup Supplies');
  sheet.setTabColor('#A64D79');

  addHeader(sheet, 1, 'STARTUP ART SUPPLIES', 5, '#A64D79');

  var headers = ['Category', 'Item', 'Quantity', 'Unit Price', 'Total'];
  sheet.getRange('A3:E3').setValues([headers]).setFontWeight('bold').setBackground('#D5A6BD');

  var supplies = [
    ['CANVAS', '', '', '', ''],
    ['', 'Canvas 8x10', 100, 3.50, '=C5*D5'],
    ['', 'Canvas 11x14', 100, 5.00, '=C6*D6'],
    ['', 'Canvas 16x20', 75, 8.50, '=C7*D7'],
    ['', 'Canvas 18x24', 50, 11.00, '=C8*D8'],
    ['', 'CANVAS SUBTOTAL', '', '', '=SUM(E5:E8)'],
    ['', '', '', '', ''],
    ['PAPER', '', '', '', ''],
    ['', 'Drawing Paper 9x12 (500 sheets)', 10, 25.00, '=C12*D12'],
    ['', 'Watercolor Paper 9x12 (100)', 20, 18.00, '=C13*D13'],
    ['', 'Mixed Media Paper 11x14 (100)', 15, 22.00, '=C14*D14'],
    ['', 'Cardstock Variety Pack', 25, 15.00, '=C15*D15'],
    ['', 'PAPER SUBTOTAL', '', '', '=SUM(E12:E15)'],
    ['', '', '', '', ''],
    ['PAINT', '', '', '', ''],
    ['', 'Acrylic Paint Set (24 colors)', 50, 28.00, '=C19*D19'],
    ['', 'Tempera Paint Gallon (assorted)', 40, 12.00, '=C20*D20'],
    ['', 'Watercolor Sets (24 pan)', 50, 15.00, '=C21*D21'],
    ['', 'Oil Paint Set (12 tubes)', 20, 45.00, '=C22*D22'],
    ['', 'PAINT SUBTOTAL', '', '', '=SUM(E19:E22)'],
    ['', '', '', '', ''],
    ['BRUSHES', '', '', '', ''],
    ['', 'Brush Sets (variety 10pc)', 75, 12.00, '=C26*D26'],
    ['', 'Detail Brush Sets', 40, 8.00, '=C27*D27'],
    ['', 'Foam Brushes (100 pack)', 10, 15.00, '=C28*D28'],
    ['', 'BRUSH SUBTOTAL', '', '', '=SUM(E26:E28)'],
    ['', '', '', '', ''],
    ['DRAWING SUPPLIES', '', '', '', ''],
    ['', 'Pencil Sets (12pc graphite)', 100, 8.00, '=C32*D32'],
    ['', 'Colored Pencil Sets (24)', 75, 12.00, '=C33*D33'],
    ['', 'Charcoal Sets', 50, 10.00, '=C34*D34'],
    ['', 'Pastels (oil & soft)', 50, 18.00, '=C35*D35'],
    ['', 'Erasers (variety)', 200, 1.50, '=C36*D36'],
    ['', 'DRAWING SUBTOTAL', '', '', '=SUM(E32:E36)'],
    ['', '', '', '', ''],
    ['MARKERS & INK', '', '', '', ''],
    ['', 'Marker Sets (washable, 24)', 100, 8.00, '=C40*D40'],
    ['', 'Fine Point Markers (12)', 75, 6.00, '=C41*D41'],
    ['', 'India Ink (bottles)', 40, 8.00, '=C42*D42'],
    ['', 'MARKER SUBTOTAL', '', '', '=SUM(E40:E42)'],
    ['', '', '', '', ''],
    ['CLAY & POTTERY', '', '', '', ''],
    ['', 'Clay (25lb bags)', 150, 18.00, '=C46*D46'],
    ['', 'Glazes (assorted pints)', 100, 12.00, '=C47*D47'],
    ['', 'Pottery Tools Sets', 50, 15.00, '=C48*D48'],
    ['', 'Underglazes (set)', 25, 35.00, '=C49*D49'],
    ['', 'Wire & Cutting Tools', 50, 8.00, '=C50*D50'],
    ['', 'CLAY SUBTOTAL', '', '', '=SUM(E46:E50)'],
    ['', '', '', '', ''],
    ['MISCELLANEOUS', '', '', '', ''],
    ['', 'Glue (bottles & sticks)', 200, 2.50, '=C54*D54'],
    ['', 'Scissors (classroom)', 100, 3.00, '=C55*D55'],
    ['', 'Rulers & Templates', 150, 2.00, '=C56*D56'],
    ['', 'Aprons (child/adult)', 150, 8.00, '=C57*D57'],
    ['', 'Storage Containers', 100, 5.00, '=C58*D58'],
    ['', 'Palettes (disposable 100pk)', 20, 10.00, '=C59*D59'],
    ['', 'Paper Towels (cases)', 25, 35.00, '=C60*D60'],
    ['', 'MISC SUBTOTAL', '', '', '=SUM(E54:E60)'],
    ['', '', '', '', ''],
    ['', 'GRAND TOTAL', '', '', '=E9+E16+E23+E29+E37+E43+E51+E61']
  ];

  for (var i = 0; i < supplies.length; i++) {
    var row = 4 + i;
    sheet.getRange(row, 1, 1, 5).setValues([supplies[i]]);

    if (supplies[i][0] !== '' || supplies[i][1].indexOf('SUBTOTAL') > -1 || supplies[i][1].indexOf('GRAND TOTAL') > -1) {
      sheet.getRange(row, 1, 1, 5).setFontWeight('bold');
    }

    if (supplies[i][1].indexOf('SUBTOTAL') > -1) {
      sheet.getRange(row, 1, 1, 5).setBackground('#EAD1DC');
    }

    if (supplies[i][1].indexOf('GRAND TOTAL') > -1) {
      sheet.getRange(row, 1, 1, 5).setBackground('#A64D79').setFontColor('white');
    }
  }

  sheet.getRange('D5:E' + (4 + supplies.length)).setNumberFormat('$#,##0.00');

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);

  sheet.getRange('A' + (4 + supplies.length + 2)).setValue('NOTES:').setFontWeight('bold');
  sheet.getRange('A' + (4 + supplies.length + 3)).setValue('• Add/edit items as needed');
  sheet.getRange('A' + (4 + supplies.length + 4)).setValue('• Subtotals auto-calculate');
  sheet.getRange('A' + (4 + supplies.length + 5)).setValue('• Grand Total feeds Startup Costs');
  sheet.getRange('A' + (4 + supplies.length + 6)).setValue('• Quarterly replenishment = 30% of this total');
}

// ============================================
// SHEET 5: REVENUE MODEL
// ============================================

function createRevenueModelSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Revenue Model');
  sheet.setTabColor('#93C47D');

  addHeader(sheet, 1, 'DECO ART CENTER - REVENUE PROJECTIONS', 6, '#93C47D');

  var headers = ['Revenue Source', 'Year 1 (10mo)', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange('A3:F3').setValues([headers]).setFontWeight('bold').setBackground('#D9EAD3');

  var r = 4;

  addSectionHeader(sheet, r++, 'REGULAR CLASSES', 6);

  sheet.getRange(r, 1).setValue('Classes Revenue');
  sheet.getRange(r, 2).setFormula('=SUMPRODUCT(\'Class Schedule\'!C4:C13,\'Class Schedule\'!D4:D13,\'Class Schedule\'!F4:F13,Assumptions!$B$29)');
  sheet.getRange(r, 3).setFormula('=SUMPRODUCT(\'Class Schedule\'!C4:C13,\'Class Schedule\'!D4:D13,\'Class Schedule\'!G4:G13,Assumptions!$B$30)');
  sheet.getRange(r, 4).setFormula('=C' + r + '*(1+Assumptions!$B$41)');
  sheet.getRange(r, 5).setFormula('=D' + r + '*(1+Assumptions!$B$41)');
  sheet.getRange(r, 6).setFormula('=E' + r + '*(1+Assumptions!$B$41)');
  r++;

  sheet.getRange(r, 1).setValue('Drop-in Classes');
  sheet.getRange(r, 2).setFormula('=B' + (r-1) + '*Assumptions!$B$42*0.25');
  sheet.getRange(r, 3).setFormula('=C' + (r-1) + '*Assumptions!$B$42*0.25');
  sheet.getRange(r, 4).setFormula('=D' + (r-1) + '*Assumptions!$B$42*0.25');
  sheet.getRange(r, 5).setFormula('=E' + (r-1) + '*Assumptions!$B$42*0.25');
  sheet.getRange(r, 6).setFormula('=F' + (r-1) + '*Assumptions!$B$42*0.25');
  r++;
  r++;

  addSectionHeader(sheet, r++, 'SUMMER CAMPS', 6);
  sheet.getRange(r, 1).setValue('Summer Camps Revenue');
  sheet.getRange(r, 2).setFormula('=SUMPRODUCT(\'Summer Camps\'!E4:E19,\'Summer Camps\'!F4:F19,\'Summer Camps\'!G4:G19,\'Summer Camps\'!I4:I19)');
  sheet.getRange(r, 3).setFormula('=SUMPRODUCT(\'Summer Camps\'!E4:E19,\'Summer Camps\'!F4:F19,\'Summer Camps\'!H4:H19,\'Summer Camps\'!I4:I19)');
  sheet.getRange(r, 4).setFormula('=C' + r + '*(1+Assumptions!$B$41)');
  sheet.getRange(r, 5).setFormula('=D' + r + '*(1+Assumptions!$B$41)');
  sheet.getRange(r, 6).setFormula('=E' + r + '*(1+Assumptions!$B$41)');
  r++;
  r++;

  addSectionHeader(sheet, r++, 'EVENTS & WORKSHOPS', 6);

  sheet.getRange(r, 1).setValue('Paint & Sip Events');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$54*Assumptions!$B$55*Assumptions!$B$56*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$54*Assumptions!$B$55*Assumptions!$B$56*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Private Events');
  sheet.getRange(r, 2).setFormula('=(Assumptions!$B$58+Assumptions!$B$59*Assumptions!$B$60)*Assumptions!$B$57*10');
  sheet.getRange(r, 3).setFormula('=(Assumptions!$B$58+Assumptions!$B$59*Assumptions!$B$60)*Assumptions!$B$57*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Visiting Artist Workshops');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$62*Assumptions!$B$63*Assumptions!$B$61*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$62*Assumptions!$B$63*Assumptions!$B$61*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;
  r++;

  addSectionHeader(sheet, r++, 'BEVERAGE SALES', 6);

  sheet.getRange(r, 1).setValue('Paint & Sip Beverages');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$54*Assumptions!$B$56*Assumptions!$B$66*(Assumptions!$B$70*0.6+Assumptions!$B$69*0.4)*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$54*Assumptions!$B$56*Assumptions!$B$66*(Assumptions!$B$70*0.6+Assumptions!$B$69*0.4)*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Private Event Beverages');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$57*Assumptions!$B$60*Assumptions!$B$67*(Assumptions!$B$70*0.6+Assumptions!$B$69*0.4)*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$57*Assumptions!$B$60*Assumptions!$B$67*(Assumptions!$B$70*0.6+Assumptions!$B$69*0.4)*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;

  sheet.getRange(r, 1).setValue('Workshop Beverages');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$61*Assumptions!$B$63*Assumptions!$B$68*((Assumptions!$B$71*0.5)+(Assumptions!$B$70*0.25+Assumptions!$B$69*0.25))*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$61*Assumptions!$B$63*Assumptions!$B$68*((Assumptions!$B$71*0.5)+(Assumptions!$B$70*0.25+Assumptions!$B$69*0.25))*12');
  sheet.getRange(r, 4).setFormula('=C' + r);
  sheet.getRange(r, 5).setFormula('=C' + r);
  sheet.getRange(r, 6).setFormula('=C' + r);
  r++;
  r++;

  addSectionHeader(sheet, r++, 'RETAIL', 6);
  sheet.getRange(r, 1).setValue('Retail Sales (5% of total revenue)');
  var retailRow = r;
  r++;
  r++;

  sheet.getRange(r, 1).setValue('TOTAL GROSS REVENUE').setFontWeight('bold').setBackground('#93C47D');
  sheet.getRange(r, 2).setFormula('=SUM(B5:B' + (retailRow) + ')');
  sheet.getRange(r, 3).setFormula('=SUM(C5:C' + (retailRow) + ')');
  sheet.getRange(r, 4).setFormula('=SUM(D5:D' + (retailRow) + ')');
  sheet.getRange(r, 5).setFormula('=SUM(E5:E' + (retailRow) + ')');
  sheet.getRange(r, 6).setFormula('=SUM(F5:F' + (retailRow) + ')');
  sheet.getRange(r, 2, 1, 5).setFontWeight('bold').setBackground('#93C47D');

  sheet.getRange(retailRow, 2).setFormula('=(B' + r + '-B' + retailRow + ')*0.05');
  sheet.getRange(retailRow, 3).setFormula('=(C' + r + '-C' + retailRow + ')*0.05');
  sheet.getRange(retailRow, 4).setFormula('=(D' + r + '-D' + retailRow + ')*0.05');
  sheet.getRange(retailRow, 5).setFormula('=(E' + r + '-E' + retailRow + ')*0.05');
  sheet.getRange(retailRow, 6).setFormula('=(F' + r + '-F' + retailRow + ')*0.05');

  sheet.getRange('B:F').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 250);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 120);
  }
}

// ============================================
// SHEET 6: EXPENSES
// ============================================

function createExpensesSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Expenses');
  sheet.setTabColor('#E06666');

  addHeader(sheet, 1, 'DECO ART CENTER - OPERATING EXPENSES', 6, '#E06666');

  var headers = ['Expense Category', 'Year 1 (10mo)', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange('A3:F3').setValues([headers]).setFontWeight('bold').setBackground('#F4CCCC');

  var r = 4;

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
  sheet.getRange(r, 2).setFormula('=(\'Revenue Model\'!B20+\'Revenue Model\'!B21+\'Revenue Model\'!B22)*Assumptions!$B$72');
  sheet.getRange(r, 3).setFormula('=(\'Revenue Model\'!C20+\'Revenue Model\'!C21+\'Revenue Model\'!C22)*Assumptions!$B$72');
  sheet.getRange(r, 4).setFormula('=(\'Revenue Model\'!D20+\'Revenue Model\'!D21+\'Revenue Model\'!D22)*Assumptions!$B$72');
  sheet.getRange(r, 5).setFormula('=(\'Revenue Model\'!E20+\'Revenue Model\'!E21+\'Revenue Model\'!E22)*Assumptions!$B$72');
  sheet.getRange(r, 6).setFormula('=(\'Revenue Model\'!F20+\'Revenue Model\'!F21+\'Revenue Model\'!F22)*Assumptions!$B$72');
  r++;

  sheet.getRange(r, 1).setValue('Supplies Replenishment (Quarterly)');
  sheet.getRange(r, 2).setFormula('=\'Startup Supplies\'!E66*Assumptions!$B$75*3');
  sheet.getRange(r, 3).setFormula('=\'Startup Supplies\'!E66*Assumptions!$B$75*4');
  sheet.getRange(r, 4).setFormula('=\'Startup Supplies\'!E66*Assumptions!$B$75*4');
  sheet.getRange(r, 5).setFormula('=\'Startup Supplies\'!E66*Assumptions!$B$75*4');
  sheet.getRange(r, 6).setFormula('=\'Startup Supplies\'!E66*Assumptions!$B$75*4');
  r++;
  r++;

  addSectionHeader(sheet, r++, 'PAYROLL & LABOR', 6);

  sheet.getRange(r, 1).setValue('Owner Salaries');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$44+Assumptions!$B$45');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$44+Assumptions!$B$45');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$44+Assumptions!$B$45');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$44+Assumptions!$B$45');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$44+Assumptions!$B$45');
  r++;

  sheet.getRange(r, 1).setValue('Studio Manager');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$46*Assumptions!$B$47*52');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$46*Assumptions!$B$47*52');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$46*Assumptions!$B$47*52');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$46*Assumptions!$B$47*52');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$46*Assumptions!$B$47*52');
  r++;

  sheet.getRange(r, 1).setValue('Teachers (estimated from schedule)');
  sheet.getRange(r, 2).setValue(65000);
  sheet.getRange(r, 3).setValue(95000);
  sheet.getRange(r, 4).setValue(115000);
  sheet.getRange(r, 5).setValue(135000);
  sheet.getRange(r, 6).setValue(150000);
  r++;

  sheet.getRange(r, 1).setValue('Assistants (estimated)');
  sheet.getRange(r, 2).setValue(15000);
  sheet.getRange(r, 3).setValue(25000);
  sheet.getRange(r, 4).setValue(32000);
  sheet.getRange(r, 5).setValue(40000);
  sheet.getRange(r, 6).setValue(45000);
  r++;

  sheet.getRange(r, 1).setValue('Payroll Taxes');
  sheet.getRange(r, 2).setFormula('=SUM(B11:B14)*Assumptions!$B$50');
  sheet.getRange(r, 3).setFormula('=SUM(C11:C14)*Assumptions!$B$50');
  sheet.getRange(r, 4).setFormula('=SUM(D11:D14)*Assumptions!$B$50');
  sheet.getRange(r, 5).setFormula('=SUM(E11:E14)*Assumptions!$B$50');
  sheet.getRange(r, 6).setFormula('=SUM(F11:F14)*Assumptions!$B$50');
  r++;
  r++;

  addSectionHeader(sheet, r++, 'FACILITIES & OCCUPANCY', 6);

  sheet.getRange(r, 1).setValue('Base Rent');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$7*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)^2');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)^3');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$7*12*(1+Assumptions!$B$20)^4');
  r++;

  sheet.getRange(r, 1).setValue('NNN (Triple Net)');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$15*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$15*12');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$15*12');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$15*12');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$15*12');
  r++;

  sheet.getRange(r, 1).setValue('Liability Insurance');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$79*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$79*12');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$79*12');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$79*12');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$79*12');
  r++;

  sheet.getRange(r, 1).setValue('Utilities (Elec+Gas+Water+Internet)');
  sheet.getRange(r, 2).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$80)*10');
  sheet.getRange(r, 3).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$80)*12');
  sheet.getRange(r, 4).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$80)*12');
  sheet.getRange(r, 5).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$80)*12');
  sheet.getRange(r, 6).setFormula('=(Assumptions!$B$76+Assumptions!$B$77+Assumptions!$B$78+Assumptions!$B$80)*12');
  r++;

  sheet.getRange(r, 1).setValue('Repairs & Maintenance');
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$81/12*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$81');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$81');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$81');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$81');
  r++;
  r++;

  addSectionHeader(sheet, r++, 'OPERATIONS', 6);

  sheet.getRange(r, 1).setValue('Marketing & Advertising');
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
  sheet.getRange(r, 2).setFormula('=Assumptions!$B$87/12*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!$B$87');
  sheet.getRange(r, 4).setFormula('=Assumptions!$B$87');
  sheet.getRange(r, 5).setFormula('=Assumptions!$B$87');
  sheet.getRange(r, 6).setFormula('=Assumptions!$B$87');
  r++;

  sheet.getRange(r, 1).setValue('Miscellaneous');
  sheet.getRange(r, 2).setValue(2000);
  sheet.getRange(r, 3).setValue(3000);
  sheet.getRange(r, 4).setValue(3500);
  sheet.getRange(r, 5).setValue(4000);
  sheet.getRange(r, 6).setValue(4500);
  r++;
  r++;

  sheet.getRange(r, 1).setValue('TOTAL OPERATING EXPENSES').setFontWeight('bold').setBackground('#E06666').setFontColor('white');
  sheet.getRange(r, 2).setFormula('=SUM(B5:B' + (r-1) + ')');
  sheet.getRange(r, 3).setFormula('=SUM(C5:C' + (r-1) + ')');
  sheet.getRange(r, 4).setFormula('=SUM(D5:D' + (r-1) + ')');
  sheet.getRange(r, 5).setFormula('=SUM(E5:E' + (r-1) + ')');
  sheet.getRange(r, 6).setFormula('=SUM(F5:F' + (r-1) + ')');
  sheet.getRange(r, 2, 1, 5).setFontWeight('bold').setBackground('#E06666').setFontColor('white');

  sheet.getRange('B:F').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 280);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 120);
  }
}

// ============================================
// SHEET 7: STARTUP COSTS
// ============================================

function createStartupCostsSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Startup Costs');
  sheet.setTabColor('#A64D79');

  addHeader(sheet, 1, 'STARTUP COSTS & INITIAL INVESTMENT', 2, '#A64D79');

  var headers = ['Item', 'Amount'];
  sheet.getRange('A3:B3').setValues([headers]).setFontWeight('bold').setBackground('#D5A6BD');

  var r = 4;

  addSectionHeader(sheet, r++, 'RENTAL DEPOSITS');
  sheet.getRange(r, 1).setValue('First Month Rent');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B23').setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Last Month Rent');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B24').setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Security Deposit');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B22').setNumberFormat('$#,##0');
  r++;

  addSectionHeader(sheet, r++, 'LEASEHOLD IMPROVEMENTS');
  sheet.getRange(r, 1).setValue('Buildout / Renovation');
  sheet.getRange(r++, 2).setValue(25000).setNumberFormat('$#,##0');
  r++;

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

  addSectionHeader(sheet, r++, 'INITIAL SUPPLIES');
  sheet.getRange(r, 1).setValue('Art Supplies (from Startup Supplies sheet)');
  sheet.getRange(r++, 2).setFormula('=\'Startup Supplies\'!E66').setNumberFormat('$#,##0');
  r++;

  addSectionHeader(sheet, r++, 'SETUP & PROFESSIONAL');
  sheet.getRange(r, 1).setValue('Signage & Branding');
  sheet.getRange(r++, 2).setValue(5000).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Initial Marketing');
  sheet.getRange(r++, 2).setValue(3000).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('Legal & Permits');
  sheet.getRange(r++, 2).setValue(5000).setNumberFormat('$#,##0');
  sheet.getRange(r, 1).setValue('POS System & Software');
  sheet.getRange(r++, 2).setValue(3000).setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('SUBTOTAL - STARTUP COSTS').setFontWeight('bold').setBackground('#D5A6BD');
  sheet.getRange(r, 2).setFormula('=SUM(B5:B7)+B9+SUM(B11:B14)+B16+SUM(B19:B22)').setFontWeight('bold').setBackground('#D5A6BD').setNumberFormat('$#,##0');
  var subtotalRow = r;
  r++;
  r++;

  addSectionHeader(sheet, r++, 'WORKING CAPITAL');
  sheet.getRange(r, 1).setValue('Operating Reserve (4 months)');
  sheet.getRange(r++, 2).setFormula('=Expenses!B35/10*4').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('TOTAL CAPITAL NEEDED').setFontWeight('bold').setBackground('#A64D79').setFontColor('white');
  sheet.getRange(r, 2).setFormula('=B' + subtotalRow + '+B' + (r-1)).setFontWeight('bold').setBackground('#A64D79').setFontColor('white').setNumberFormat('$#,##0');

  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 150);

  sheet.getRange('D3').setValue('NOTE:').setFontWeight('bold');
  sheet.getRange('D4').setValue('This total feeds the Loan Calculator');
  sheet.getRange('D5').setValue('to determine your SBA loan amount');
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
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B' + 29).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Owner Equity / Cash Investment');
  sheet.getRange(r++, 2).setValue(0).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Other Funding Sources');
  sheet.getRange(r++, 2).setValue(0).setNumberFormat('$#,##0');
  r++;

  addSectionHeader(sheet, r++, 'LOAN BREAKDOWN');

  sheet.getRange(r, 1).setValue('Startup Costs to Finance');
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B' + 24).setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Working Capital to Finance');
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B' + 28).setNumberFormat('$#,##0');
  r++;

  addSectionHeader(sheet, r++, 'SBA LOAN TERMS');

  sheet.getRange(r, 1).setValue('Interest Rate');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B34').setNumberFormat('0.00%');

  sheet.getRange(r, 1).setValue('Loan Term (years)');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B35');
  r++;

  sheet.getRange(r, 1).setValue('TOTAL LOAN AMOUNT NEEDED').setFontWeight('bold').setBackground('#6FA8DC').setFontColor('white');
  sheet.getRange(r, 2).setFormula('=B5-B6-B7').setFontWeight('bold').setBackground('#6FA8DC').setFontColor('white').setNumberFormat('$#,##0');
  var loanRow = r;
  r++;

  sheet.getRange(r, 1).setValue('Monthly Payment');
  sheet.getRange(r, 2).setFormula('=-PMT(B15/12,B16*12,B' + loanRow + ')').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Total Interest Paid (10 years)');
  sheet.getRange(r, 2).setFormula('=B19*B16*12-B' + loanRow).setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1).setValue('Total Amount Repaid');
  sheet.getRange(r, 2).setFormula('=B19*B16*12').setNumberFormat('$#,##0');

  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 150);

  sheet.getRange('D3').setValue('INSTRUCTIONS:').setFontWeight('bold');
  sheet.getRange('D4').setValue('1. Enter owner equity/cash (if any) in B6');
  sheet.getRange('D5').setValue('2. Enter other funding sources (if any) in B7');
  sheet.getRange('D6').setValue('3. Loan amount auto-calculates');
  sheet.getRange('D7').setValue('4. This feeds the Assumptions sheet');
}

// ============================================
// SHEET 9: LOAN AMORTIZATION
// ============================================

function createLoanAmortizationSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Loan Amortization');
  sheet.setTabColor('#6FA8DC');

  addHeader(sheet, 1, 'LOAN AMORTIZATION SCHEDULE (5 YEARS)', 5, '#6FA8DC');

  sheet.getRange('A3').setValue('Loan Amount:');
  sheet.getRange('B3').setFormula('=\'Loan Calculator\'!B18').setNumberFormat('$#,##0');

  sheet.getRange('A4').setValue('Interest Rate:');
  sheet.getRange('B4').setFormula('=Assumptions!B34').setNumberFormat('0.00%');

  sheet.getRange('A5').setValue('Term:');
  sheet.getRange('B5').setValue('10 years');

  sheet.getRange('A6').setValue('Monthly Payment:');
  sheet.getRange('B6').setFormula('=Assumptions!B37').setNumberFormat('$#,##0');

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
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
}

// ============================================
// SHEET 10: PROFIT & LOSS
// ============================================

function createProfitLossSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Profit & Loss');
  sheet.setTabColor('#F6B26B');

  addHeader(sheet, 1, 'PRO FORMA PROFIT & LOSS STATEMENT', 6, '#F6B26B');

  var headers = ['', 'Year 1 (10mo)', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
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
  sheet.getRange(r, 2).setFormula('=Expenses!B35');
  sheet.getRange(r, 3).setFormula('=Expenses!C35');
  sheet.getRange(r, 4).setFormula('=Expenses!D35');
  sheet.getRange(r, 5).setFormula('=Expenses!E35');
  sheet.getRange(r, 6).setFormula('=Expenses!F35');
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
  sheet.getRange(r, 2).setFormula('=Assumptions!B98/12*10');
  sheet.getRange(r, 3).setFormula('=Assumptions!B98');
  sheet.getRange(r, 4).setFormula('=Assumptions!B98');
  sheet.getRange(r, 5).setFormula('=Assumptions!B98');
  sheet.getRange(r, 6).setFormula('=Assumptions!B98');
  r++;
  r++;

  sheet.getRange(r, 1).setValue('EBIT (Operating Income)');
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

  sheet.getRange(r, 1).setValue('NET INCOME (Before Tax)');
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
  sheet.setColumnWidth(1, 220);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 120);
  }
}

// ============================================
// SHEET 11: CASH FLOW
// ============================================

function createCashFlowSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Cash Flow');
  sheet.setTabColor('#8E7CC3');

  addHeader(sheet, 1, 'CASH FLOW PROJECTION', 6, '#8E7CC3');

  var headers = ['', 'Year 1 (10mo)', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  sheet.getRange('A3:F3').setValues([headers]).setFontWeight('bold').setBackground('#D9D2E9');

  var r = 4;

  sheet.getRange(r++, 1).setValue('OPERATING ACTIVITIES').setFontWeight('bold').setBackground('#D9D2E9');

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

  sheet.getRange(r++, 1).setValue('FINANCING ACTIVITIES').setFontWeight('bold').setBackground('#D9D2E9');

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

  sheet.getRange(r, 1).setValue('Beginning Cash Balance');
  sheet.getRange(r, 2).setFormula('=\'Startup Costs\'!B28');
  sheet.getRange(r, 3).setFormula('=B13');
  sheet.getRange(r, 4).setFormula('=C13');
  sheet.getRange(r, 5).setFormula('=D13');
  sheet.getRange(r, 6).setFormula('=E13');
  r++;

  sheet.getRange(r, 1).setValue('Ending Cash Balance');
  sheet.getRange(r, 2).setFormula('=B12+B13');
  sheet.getRange(r, 3).setFormula('=C12+C13');
  sheet.getRange(r, 4).setFormula('=D12+D13');
  sheet.getRange(r, 5).setFormula('=E12+E13');
  sheet.getRange(r, 6).setFormula('=F12+F13');
  sheet.getRange(r, 1, 1, 6).setFontWeight('bold');

  sheet.getRange('B:F').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 220);
  for (var i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 120);
  }
}

// ============================================
// SHEET 12: DASHBOARD
// ============================================

function createDashboardSheet(ss) {
  var sheet = getOrCreateSheet(ss, 'Dashboard');
  sheet.setTabColor('#FFD966');

  addHeader(sheet, 1, 'DECO ART CENTER - FINANCIAL DASHBOARD', 6, '#FFD966');

  var r = 3;

  sheet.getRange(r, 1, 1, 6).merge().setValue('SBA LOAN REQUEST')
    .setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
  r += 2;

  sheet.getRange(r, 1).setValue('Total Capital Needed').setFontWeight('bold');
  sheet.getRange(r++, 2).setFormula('=\'Startup Costs\'!B29').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('SBA Loan Amount Requested').setFontWeight('bold');
  sheet.getRange(r++, 2).setFormula('=\'Loan Calculator\'!B18').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Loan Term').setFontWeight('bold');
  sheet.getRange(r++, 2).setValue('10 years');

  sheet.getRange(r, 1).setValue('Interest Rate').setFontWeight('bold');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B34').setNumberFormat('0.00%');

  sheet.getRange(r, 1).setValue('Monthly Payment').setFontWeight('bold');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B37').setNumberFormat('$#,##0');
  r++;

  sheet.getRange(r, 1, 1, 6).merge().setValue('FACILITY DETAILS')
    .setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
  r += 2;

  sheet.getRange(r, 1).setValue('Building Size');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B4&" sq ft"');

  sheet.getRange(r, 1).setValue('Monthly Rent (Base + NNN)');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B16').setNumberFormat('$#,##0');

  sheet.getRange(r, 1).setValue('Lease Term');
  sheet.getRange(r++, 2).setFormula('=Assumptions!B19&" years"');
  r++;

  sheet.getRange(r, 1, 1, 6).merge().setValue('5-YEAR FINANCIAL PERFORMANCE')
    .setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
  r += 2;

  var perfHeaders = ['Metric', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
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
  sheet.getRange(r, 2).setFormula('=Expenses!B35');
  sheet.getRange(r, 3).setFormula('=Expenses!C35');
  sheet.getRange(r, 4).setFormula('=Expenses!D35');
  sheet.getRange(r, 5).setFormula('=Expenses!E35');
  sheet.getRange(r, 6).setFormula('=Expenses!F35');
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

  sheet.getRange(r, 1, 1, 6).merge().setValue('KEY RATIOS')
    .setFontWeight('bold').setFontSize(12).setBackground('#FFE599');
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

  sheet.getRange(r, 1).setValue('Net Margin');
  sheet.getRange(r, 2).setFormula('=\'Profit & Loss\'!B20');
  sheet.getRange(r, 3).setFormula('=\'Profit & Loss\'!C20');
  sheet.getRange(r, 4).setFormula('=\'Profit & Loss\'!D20');
  sheet.getRange(r, 5).setFormula('=\'Profit & Loss\'!E20');
  sheet.getRange(r, 6).setFormula('=\'Profit & Loss\'!F20');
  sheet.getRange(r, 2, 1, 5).setNumberFormat('0.0%');
  r++;

  sheet.getRange(r, 1).setValue('Debt Service Coverage Ratio');
  sheet.getRange(r, 2).setFormula('=(\'Profit & Loss\'!B8)/(Assumptions!$B$37*10)');
  sheet.getRange(r, 3).setFormula('=(\'Profit & Loss\'!C8)/(Assumptions!$B$37*12)');
  sheet.getRange(r, 4).setFormula('=(\'Profit & Loss\'!D8)/(Assumptions!$B$37*12)');
  sheet.getRange(r, 5).setFormula('=(\'Profit & Loss\'!E8)/(Assumptions!$B$37*12)');
  sheet.getRange(r, 6).setFormula('=(\'Profit & Loss\'!F8)/(Assumptions!$B$37*12)');
  sheet.getRange(r, 2, 1, 5).setNumberFormat('0.00');
  r++;
  r++;

  sheet.getRange(r++, 1).setValue('NOTES FOR LENDERS:').setFontWeight('bold');
  sheet.getRange(r++, 1).setValue('• Rental model - 5 year lease at $16/sf + $4 NNN');
  sheet.getRange(r++, 1).setValue('• Conservative 50% capacity Year 1, growing 20% annually');
  sheet.getRange(r++, 1).setValue('• Multiple revenue streams (classes, camps, events, beverages, retail)');
  sheet.getRange(r++, 1).setValue('• All pricing based on Tulsa, OK market (zip 74119)');
  sheet.getRange(r++, 1).setValue('• Year 1 = 10 months operation (March-December 2026)');
  sheet.getRange(r++, 1).setValue('• 4-month operating reserve included in loan request');
  sheet.getRange(r++, 1).setValue('• Strong DSCR (>1.25 indicates ability to service debt)');

  sheet.getRange('B6:B10').setNumberFormat('$#,##0');
  sheet.getRange('B23:F27').setNumberFormat('$#,##0');

  for (var i = 1; i <= 6; i++) {
    sheet.setColumnWidth(i, 150);
  }

  // Move Dashboard to first position
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
}
