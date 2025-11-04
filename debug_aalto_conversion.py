import pandas as pd
import re

# Load Excel
df = pd.read_excel('Conversion_month.xlsx')

# Check 33.00 CHF entries
print("=== All 33 CHF entries ===")
matches_33 = df[df['Unit Price'] == 33.0]
print(matches_33[['Unit Price', 'Unit Price (EUR)', 'Wine Name', 'Size', 'Minimum Quantity', 'Campaign Sub-Type', 'Producer Name']].to_string())

print("\n=== Aalto wines with 33 CHF (Min Qty = 36) ===")
aalto_33 = df[(df['Unit Price'] == 33.0) & (df['Wine Name'].str.contains('Aalto', case=False, na=False)) & (df['Minimum Quantity'] == 36)]
print(aalto_33[['Unit Price', 'Unit Price (EUR)', 'Wine Name', 'Size', 'Minimum Quantity', 'Campaign Sub-Type']].to_string())

# Check 34.00 CHF entries for Aalto
print("\n=== Aalto wines with 34 CHF (Min Qty = 0) ===")
aalto_34 = df[(df['Unit Price'] == 34.0) & (df['Wine Name'].str.contains('Aalto', case=False, na=False)) & (df['Minimum Quantity'] == 0)]
print(aalto_34[['Unit Price', 'Unit Price (EUR)', 'Wine Name', 'Size', 'Minimum Quantity', 'Campaign Sub-Type']].to_string())
