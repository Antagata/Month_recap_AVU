import pandas as pd
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Load Excel
df = pd.read_excel('Conversion_month.xlsx')

# Test case: Rieussec 2019, CHF 42.00, Qty 0, Size 75
print("="*80)
print("DEBUG: Why isn't Item No. matching working?")
print("="*80)

# Simulate the matching logic
chf_price = "42.00"
context_vintage = 2019
detected_size = 75.0
detected_quantity = 0

print(f"\nTest: CHF {chf_price}, Vintage {context_vintage}, Size {detected_size}, Qty {detected_quantity}")

# Get wine_options for this CHF
wine_options = df[df['Unit Price'] == 42.0]
print(f"\nFound {len(wine_options)} wine options for CHF 42.00")

# Build item_number_map (same as in converter)
item_number_map = {}
for _, row in df.iterrows():
    item_no = row.get('Item No.')
    if pd.notna(item_no):
        item_key = int(item_no)
        if item_key not in item_number_map:
            item_number_map[item_key] = []

        wine_data = {
            'wine_name': row['Wine Name'],
            'eur_value': f"{row['Unit Price (EUR)']:.2f}",
            'vintage': row['Vintage'],
            'size': row['Size'],
            'min_quantity': row['Minimum Quantity'],
            'item_no': item_no,
            'chf_value': f"{row['Unit Price']:.2f}"
        }
        item_number_map[item_key].append(wine_data)

print(f"\nBuilt item_number_map with {len(item_number_map)} unique items")

# Try to match
print("\n" + "="*80)
print("ATTEMPTING ITEM NO. MATCHING")
print("="*80)

matched = False
for idx, option in wine_options.iterrows():
    item_no = option.get('Item No.')
    print(f"\n--- Checking option: {option['Wine Name']} (Item No: {item_no}) ---")

    if pd.notna(item_no):
        item_key = int(item_no)
        print(f"  Item No. is valid: {item_key}")

        if item_key in item_number_map:
            print(f"  Item No. {item_key} found in map!")
            item_entries = item_number_map[item_key]
            print(f"  Found {len(item_entries)} entries for this Item No.")

            for entry_idx, entry in enumerate(item_entries):
                print(f"\n  Entry {entry_idx + 1}:")
                print(f"    Wine: {entry['wine_name']}")
                print(f"    Vintage: {entry['vintage']} (context: {context_vintage}) -> Match: {entry['vintage'] == context_vintage}")
                print(f"    Size: {entry['size']} (detected: {detected_size}) -> Match: {entry['size'] == detected_size}")
                print(f"    Min Qty: {entry['min_quantity']} (detected: {detected_quantity}) -> Match: {entry['min_quantity'] == detected_quantity}")
                print(f"    CHF: {entry['chf_value']} (looking for: {chf_price}) -> Match: {entry['chf_value'] == chf_price}")

                if (entry['vintage'] == context_vintage and
                    entry['size'] == detected_size and
                    entry['min_quantity'] == detected_quantity and
                    entry['chf_value'] == chf_price):
                    print(f"\n  ✓✓✓ BULLETPROOF MATCH FOUND! ✓✓✓")
                    print(f"  EUR Value: {entry['eur_value']}")
                    matched = True
                    break

            if matched:
                break
        else:
            print(f"  Item No. {item_key} NOT in map")
    else:
        print(f"  Item No. is NaN")

if not matched:
    print("\n✗ NO ITEM NO. MATCH FOUND")
    print("\nPossible reasons:")
    print("1. Vintage not extracted correctly from document")
    print("2. CHF value format mismatch (42.00 vs 42.0)")
    print("3. Size or quantity detection issue")
