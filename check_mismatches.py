import pandas as pd
import sys
import io

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

df = pd.read_excel(r'Outputs\OMT lines\Matched_OMT Main Offer List_20251124_185222.xlsx')

print('First 20 rows - checking for mismatches:')
print('='*150)

for i in range(20):
    extracted = df.iloc[i]['Extracted Wine Name (from Multi.txt)']
    producer = df.iloc[i]['Producer Name']
    wine_db = df.iloc[i]['Wine Name']
    item_no = df.iloc[i]['Item No. Int']
    min_qty = df.iloc[i]['Minimum Quantity']

    print(f"{i+1}. Extracted: {extracted}")
    print(f"   Producer: {producer} | Wine (DB): {wine_db} | Item: {item_no} | Min Qty: {min_qty}")
    print()
