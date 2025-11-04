import pandas as pd
import sys
import io

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Load Lines.xlsx (our output) and Conversion_month.xlsx (source)
lines_df = pd.read_excel('Lines.xlsx')
excel_df = pd.read_excel('Conversion_month.xlsx')

print("="*80)
print("VERIFICATION: Checking if all conversions are bulletproof")
print("="*80)

# Check matches that are NOT bulletproof (fuzzy_filtered, fallback, market_price)
problematic = lines_df[lines_df['Match Type'].isin(['fuzzy_filtered', 'fallback_1.08', 'market_price_1.08'])]

print(f"\n⚠️  Found {len(problematic)} non-bulletproof matches:")
print(f"   - fuzzy_filtered: {(lines_df['Match Type'] == 'fuzzy_filtered').sum()}")
print(f"   - fallback_1.08: {(lines_df['Match Type'] == 'fallback_1.08').sum()}")
print(f"   - market_price_1.08: {(lines_df['Match Type'] == 'market_price_1.08').sum()}")

print("\n" + "="*80)
print("FUZZY_FILTERED MATCHES (Need to become bulletproof):")
print("="*80)

fuzzy = lines_df[lines_df['Match Type'] == 'fuzzy_filtered']
for idx, row in fuzzy.iterrows():
    wine_name = row['Wine Name']
    vintage = row['Vintage Code']
    size = row['Size']
    min_qty = row['Minimum Quantity']
    chf = row['Unit Price']
    eur = row['Unit Price (EUR)']
    item_no = row['Item No.']

    # Check if this wine exists in Excel with matching Item No.
    if pd.notna(item_no):
        excel_match = excel_df[
            (excel_df['Item No.'] == item_no) &
            (excel_df['Minimum Quantity'] == min_qty) &
            (excel_df['Unit Price'] == chf)
        ]

        if len(excel_match) > 0:
            print(f"\n✓ {wine_name} {vintage} (Qty={min_qty})")
            print(f"  CHF {chf} → EUR {eur} | Item No. {int(item_no)}")
            print(f"  Status: Item No. exists in Excel - SHOULD be item_no_match!")
        else:
            print(f"\n✗ {wine_name} {vintage} (Qty={min_qty})")
            print(f"  CHF {chf} → EUR {eur} | Item No. {int(item_no) if pd.notna(item_no) else 'N/A'}")
            print(f"  Status: No exact Excel match found")
    else:
        print(f"\n✗ {wine_name} {vintage} (Qty={min_qty})")
        print(f"  CHF {chf} → EUR {eur} | Item No. MISSING")
        print(f"  Status: Calculated/Fallback conversion")

print("\n" + "="*80)
print("FALLBACK MATCHES (Need to be fixed):")
print("="*80)

fallback = lines_df[lines_df['Match Type'].isin(['fallback_1.08', 'market_price_1.08'])]
for idx, row in fallback.iterrows():
    print(f"\n✗ {row['Wine Name']} {row['Vintage Code']} (Qty={row['Minimum Quantity']})")
    print(f"  CHF {row['Unit Price']} → EUR {row['Unit Price (EUR)']} [CALCULATED]")
    print(f"  Match Type: {row['Match Type']}")

print("\n" + "="*80)
print("RECOMMENDATION:")
print("="*80)
print("\n❌ Conversions are NOT bulletproof yet.")
print(f"   {len(problematic)} out of {len(lines_df)} matches need improvement.")
print("\nNext steps:")
print("1. Fix Item No. matching logic to work for ALL wines with Item No.")
print("2. Verify CHF prices in Word document match Excel prices")
print("3. Re-run converter to get 100% bulletproof matches")
