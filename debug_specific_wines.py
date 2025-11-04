import pandas as pd
from docx import Document

# Load Excel
df = pd.read_excel('Conversion_month.xlsx')

# Load Word doc
doc = Document('month recap.docx')

print("="*80)
print("DEBUGGING SPECIFIC WINE MATCHES")
print("="*80)

# Check Champagne Extra Brut Grand Cru Le Mesnil
print("\n1. Champagne Extra Brut Grand Cru Le Mesnil 2022")
print("   Document CHF: 165.00 | Actual CHF should be: 150.00")
print("\n   Excel entries with CHF 165.00:")
chf_165 = df[df['Unit Price'] == 165.0]
print(chf_165[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())

print("\n   Excel entries with CHF 150.00 and matching wine:")
chf_150 = df[(df['Unit Price'] == 150.0) & (df['Wine Name'].str.contains('Mesnil', case=False, na=False))]
print(chf_150[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())

# Check Il Pino di Biserno
print("\n2. Il Pino di Biserno 2022")
print("   Document CHF: 37.00 | Actual CHF should be: 34.00")
print("\n   Excel entries with CHF 37.00:")
chf_37 = df[df['Unit Price'] == 37.0]
print(chf_37[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].head(10).to_string())

print("\n   Excel entries with CHF 34.00 and 'Pino' or 'Biserno':")
chf_34 = df[(df['Unit Price'] == 34.0) & (df['Wine Name'].str.contains('Pino|Biserno', case=False, na=False))]
print(chf_34[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())

# Check Clos l'Église
print("\n3. Clos l'Église 2009")
print("   Document CHF: 105.00 | Actual CHF should be: 99.00")
print("\n   Excel entries with CHF 105.00:")
chf_105 = df[df['Unit Price'] == 105.0]
print(chf_105[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())

print("\n   Excel entries with CHF 99.00:")
chf_99 = df[df['Unit Price'] == 99.0]
print(chf_99[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.', 'Campaign Sub-Type']].to_string())

# Check Montrose issue (Qty=36)
print("\n4. Montrose 2021 (Qty=36)")
print("   Document CHF: 80.00 | Actual CHF should be: 75.00")
print("\n   Excel entries with CHF 80.00 and Montrose:")
chf_80_montrose = df[(df['Unit Price'] == 80.0) & (df['Wine Name'].str.contains('Montrose', case=False, na=False))]
print(chf_80_montrose[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())

print("\n   Excel entries with CHF 75.00 and Montrose:")
chf_75_montrose = df[(df['Unit Price'] == 75.0) & (df['Wine Name'].str.contains('Montrose', case=False, na=False))]
print(chf_75_montrose[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())

# Check in document what prices are actually written
print("\n" + "="*80)
print("CHECKING ACTUAL PRICES IN DOCUMENT")
print("="*80)

for i, para in enumerate(doc.paragraphs):
    text = para.text
    if 'Mesnil' in text and ('165' in text or '150' in text):
        print(f"\nMesnil paragraph {i}:")
        print(text[:200])
    if 'Pino' in text and 'Biserno' in text and ('37' in text or '34' in text):
        print(f"\nPino di Biserno paragraph {i}:")
        print(text[:200])
    if 'Clos' in text and ('glise' in text or 'Eglise' in text) and ('105' in text or '99' in text):
        print(f"\nClos l'Église paragraph {i}:")
        print(text[:300])
    if 'Montrose' in text and '2021' in text and ('80' in text or '75' in text):
        print(f"\nMontrose paragraph {i}:")
        print(text[:200])
