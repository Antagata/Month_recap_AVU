import pandas as pd

# Load Excel
df = pd.read_excel('Conversion_month.xlsx')

# Test case: Rieussec 2019, Size 75, CHF 42.00
chf_price = "42.00"
context_vintage = 2019
detected_size = 75.0
detected_quantity = 0

print(f"Looking for: CHF {chf_price}, Vintage {context_vintage}, Size {detected_size}, Qty {detected_quantity}")

# Get all entries for this CHF price
matches = df[df['Unit Price'] == 42.0]
print(f"\nFound {len(matches)} entries with CHF 42.00")

# Filter by vintage and size
filtered = matches[(matches['Vintage'] == context_vintage) & (matches['Size'] == detected_size)]
print(f"\nFiltered by vintage {context_vintage} and size {detected_size}: {len(filtered)} entries")
print(filtered[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())

# Check if Item No. is available
if len(filtered) > 0:
    item_no = filtered.iloc[0]['Item No.']
    print(f"\nItem No. for first match: {item_no}")
    print(f"Type: {type(item_no)}")
    print(f"Is NaN: {pd.isna(item_no)}")

    # Get all entries with this Item No.
    if pd.notna(item_no):
        item_key = int(item_no)
        all_item_entries = df[df['Item No.'] == item_key]
        print(f"\nAll entries with Item No. {item_key}:")
        print(all_item_entries[['Wine Name', 'Vintage', 'Size', 'Minimum Quantity', 'Unit Price', 'Unit Price (EUR)', 'Item No.']].to_string())
