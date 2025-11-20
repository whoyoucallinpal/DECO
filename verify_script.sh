#!/bin/bash
echo "=== COMPREHENSIVE SCRIPT VERIFICATION ==="
echo ""

FILE="DECOProForma_Rental_VERIFIED.gs"

echo "1. Checking for syntax errors..."
# Check for common typos
grep -n "getColumnWidth.*," "$FILE" && echo "ERROR: Found getColumnWidth with 2 params" || echo "✓ No getColumnWidth errors"

# Check for mismatched quotes in formulas
grep -n "setFormula.*[^']='[^']*$" "$FILE" && echo "WARNING: Possible unmatched quotes" || echo "✓ No unmatched quote issues"

# Check sheet name consistency
echo ""
echo "2. Verifying sheet names..."
SHEETS=("Assumptions" "Class Schedule" "Summer Camps" "Startup Supplies" "Revenue Model" "Expenses" "Startup Costs" "Loan Calculator" "Loan Amortization" "Profit & Loss" "Cash Flow" "Dashboard")

for sheet in "${SHEETS[@]}"; do
    count=$(grep -c "'$sheet'" "$FILE")
    if [ $count -gt 0 ]; then
        echo "✓ $sheet: $count references"
    else
        echo "⚠ $sheet: No references found"
    fi
done

echo ""
echo "3. Checking for common method signature errors..."
# Check for incorrect method calls
grep -n "\.merge()" "$FILE" | head -3 && echo "✓ merge() calls found" || echo "⚠ No merge calls"
grep -n "\.setFormula(" "$FILE" | wc -l | xargs echo "✓ setFormula calls:"
grep -n "\.setValue(" "$FILE" | wc -l | xargs echo "✓ setValue calls:"

echo ""
echo "4. Verifying critical cell references..."
# Check that fixed references are present
grep -q "Assumptions!\$B\$30" "$FILE" && echo "✓ Class Sessions Year 1 (B30)" || echo "❌ Missing B30 reference"
grep -q "Assumptions!\$B\$31" "$FILE" && echo "✓ Class Sessions Year 2+ (B31)" || echo "❌ Missing B31 reference"
grep -q "'Startup Costs'!B32" "$FILE" && echo "✓ Total Capital Needed (B32)" || echo "❌ Missing B32 reference"
grep -q "'Startup Costs'!B27" "$FILE" && echo "✓ Startup Costs Subtotal (B27)" || echo "❌ Missing B27 reference"
grep -q "'Startup Costs'!B30" "$FILE" && echo "✓ Operating Reserve (B30)" || echo "❌ Missing B30 reference"
grep -q "'Startup Supplies'!E63" "$FILE" && echo "✓ Supplies Grand Total (E63)" || echo "❌ Missing E63 reference"

echo ""
echo "5. Checking function definitions..."
FUNCTIONS=("createDECOProForma" "getOrCreateSheet" "addHeader" "addSectionHeader" "createAssumptionsSheet" "createClassScheduleSheet" "createSummerCampsSheet" "createStartupSuppliesSheet" "createRevenueModelSheet" "createExpensesSheet" "createStartupCostsSheet" "createLoanCalculatorSheet" "createLoanAmortizationSheet" "createProfitLossSheet" "createCashFlowSheet" "createDashboardSheet")

for func in "${FUNCTIONS[@]}"; do
    grep -q "function $func" "$FILE" && echo "✓ $func" || echo "❌ Missing $func"
done

echo ""
echo "=== VERIFICATION COMPLETE ==="
