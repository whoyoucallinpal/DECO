#!/bin/bash

# Fix all incorrect Assumptions references
# Working backwards from high to low row numbers to avoid double-substitution

cd /home/user/DECO

# Create working copy
cp DECOProForma_Rental_FIXED_REFS.gs temp.gs

# Depreciation section (B88-B92 → B90-B92, but B90 stays)
sed -i 's/Assumptions!B94/Assumptions!B92/g' temp.gs
sed -i 's/Assumptions!\$B\$94/Assumptions!\$B\$92/g' temp.gs

# Operating expenses section - work backwards from high to low
sed -i 's/Assumptions!\$B\$87/Assumptions!\$B\$87/g' temp.gs  # Liquor License - correct
sed -i 's/Assumptions!\$B\$86/Assumptions!\$B\$86/g' temp.gs  # Payroll Service - correct
sed -i 's/Assumptions!\$B\$85/Assumptions!\$B\$85/g' temp.gs  # Legal - correct
sed -i 's/Assumptions!\$B\$84/Assumptions!\$B\$84/g' temp.gs  # Accounting - correct
sed -i 's/Assumptions!\$B\$83/Assumptions!\$B\$83/g' temp.gs  # Marketing Y2+ - correct
sed -i 's/Assumptions!\$B\$82/Assumptions!\$B\$82/g' temp.gs  # Marketing Y1 - correct
sed -i 's/Assumptions!\$B\$81/Assumptions!\$B\$81/g' temp.gs  # Repairs - correct
sed -i 's/Assumptions!\$B\$80/Assumptions!\$B\$80/g' temp.gs  # Liability Insurance - correct
sed -i 's/Assumptions!\$B\$79/Assumptions!\$B\$79/g' temp.gs  # Internet - correct
sed -i 's/Assumptions!\$B\$78/Assumptions!\$B\$78/g' temp.gs  # Water - correct
sed -i 's/Assumptions!\$B\$77/Assumptions!\$B\$77/g' temp.gs  # Gas - correct
sed -i 's/Assumptions!\$B\$76/Assumptions!\$B\$76/g' temp.gs  # Electricity - correct
sed -i 's/Assumptions!\$B\$75/Assumptions!\$B\$75/g' temp.gs  # Supplies Replenish - correct
sed -i 's/Assumptions!\$B\$74/Assumptions!\$B\$74/g' temp.gs  # Pottery COGS - correct
sed -i 's/Assumptions!\$B\$73/Assumptions!\$B\$73/g' temp.gs  # Art COGS - correct

# Beverages section (B66-B72 → B64-B70)
sed -i 's/Assumptions!\$B\$72/Assumptions!\$B\$70/g' temp.gs  # Beverage COGS %
sed -i 's/Assumptions!\$B\$71/Assumptions!\$B\$69/g' temp.gs  # Wine Price
sed -i 's/Assumptions!\$B\$70/Assumptions!\$B\$68/g' temp.gs  # Beer Price
sed -i 's/Assumptions!\$B\$69/Assumptions!\$B\$67/g' temp.gs  # Coffee Price
sed -i 's/Assumptions!\$B\$68/Assumptions!\$B\$66/g' temp.gs  # Drinks/Person Workshop
sed -i 's/Assumptions!\$B\$67/Assumptions!\$B\$65/g' temp.gs  # Drinks/Person Private
sed -i 's/Assumptions!\$B\$66/Assumptions!\$B\$64/g' temp.gs  # Drinks/Person Paint & Sip

# Events section (B54-B63 → B52-B61)
sed -i 's/Assumptions!\$B\$63/Assumptions!\$B\$61/g' temp.gs  # Workshop Attendance
sed -i 's/Assumptions!\$B\$62/Assumptions!\$B\$60/g' temp.gs  # Workshop Price
sed -i 's/Assumptions!\$B\$61/Assumptions!\$B\$59/g' temp.gs  # Workshops/Month
sed -i 's/Assumptions!\$B\$60/Assumptions!\$B\$58/g' temp.gs  # Private Event Attendance
sed -i 's/Assumptions!\$B\$59/Assumptions!\$B\$57/g' temp.gs  # Private Event Per Person
sed -i 's/Assumptions!\$B\$58/Assumptions!\$B\$56/g' temp.gs  # Private Event Base Fee
sed -i 's/Assumptions!\$B\$57/Assumptions!\$B\$55/g' temp.gs  # Private Events/Month
sed -i 's/Assumptions!\$B\$56/Assumptions!\$B\$54/g' temp.gs  # Paint & Sip Attendance
sed -i 's/Assumptions!\$B\$55/Assumptions!\$B\$53/g' temp.gs  # Paint & Sip Price
sed -i 's/Assumptions!\$B\$54/Assumptions!\$B\$52/g' temp.gs  # Paint & Sip Events/Month

# Staff section (B42-B50 → B43-B49 for used fields)
sed -i 's/Assumptions!\$B\$50/Assumptions!\$B\$49/g' temp.gs  # Payroll Tax Rate
sed -i 's/Assumptions!\$B\$48/Assumptions!\$B\$49/g' temp.gs  # Payroll Tax Rate (another ref)
sed -i 's/Assumptions!\$B\$47/Assumptions!\$B\$48/g' temp.gs  # Assistant Hourly
sed -i 's/Assumptions!\$B\$46/Assumptions!\$B\$47/g' temp.gs  # Teacher Hourly
sed -i 's/Assumptions!\$B\$45/Assumptions!\$B\$46/g' temp.gs  # Studio Manager Hours/Week
sed -i 's/Assumptions!\$B\$44/Assumptions!\$B\$45/g' temp.gs  # Studio Manager Hourly
sed -i 's/Assumptions!\$B\$43/Assumptions!\$B\$44/g' temp.gs  # Owner 2 Salary
sed -i 's/Assumptions!\$B\$42/Assumptions!\$B\$43/g' temp.gs  # Owner 1 Salary

# Capacity section (already correct: B38-B40)
# Timeline section (already correct: B30-B31)
# SBA section (already correct: B34-B35)

mv temp.gs DECOProForma_Rental_FIXED_REFS.gs

echo "All references fixed!"
