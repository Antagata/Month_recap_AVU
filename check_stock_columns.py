import pandas as pd

# Check Stock Lines for item with 290 CHF price
stock = pd.read_excel(r'C:\Users\Marco.Africani\OneDrive - AVU SA\AVU CPI Campaign\Puzzle_control_Reports\SOURCE_FILES\Stock Lines.xlsx')

print("Searching Stock Lines for items with 290 CHF price:")
matches = stock[stock['OMT Last Private Offer Price'] == 290.0]
print(f"Found {len(matches)} matches in Stock Lines")
if len(matches) > 0:
    print(matches[['No.', 'Wine Name', 'Vintage Code', 'OMT Last Private Offer Price']].to_string(index=False))

print("\n" + "="*100)

# Now check OMT for those item numbers with min qty 36
omt = pd.read_excel(r'C:\Users\Marco.Africani\OneDrive - AVU SA\AVU CPI Campaign\Puzzle_control_Reports\SOURCE_FILES\OMT Main Offer List.xlsx')
omt['Item No. Int'] = pd.to_numeric(omt['Item No.'], errors='coerce')

if len(matches) > 0:
    item_no = matches.iloc[0]['No.']
    print(f"\nSearching OMT for Item {item_no} with 290 CHF and min_qty 36:")
    omt_matches = omt[
        (omt['Item No. Int'] == item_no) &
        (omt['Unit Price'] == 290.0) &
        (omt['Minimum Quantity'] == 36) &
        (omt['Campaign Type'] == 'PRIVATE') &
        (omt['Campaign Sub-Type'] == 'Normal')
    ]
    print(f"Found {len(omt_matches)} matches in OMT")
    if len(omt_matches) > 0:
        print(omt_matches[['Wine Name', 'Item No. Int', 'Unit Price', 'Unit Price (EUR)', 'Minimum Quantity', 'Campaign Type']].to_string(index=False))
    else:
        print("\nTrying without Campaign filters:")
        omt_matches_nofilter = omt[
            (omt['Item No. Int'] == item_no) &
            (omt['Unit Price'] == 290.0) &
            (omt['Minimum Quantity'] == 36)
        ]
        print(f"Found {len(omt_matches_nofilter)} matches without Campaign filters")
        if len(omt_matches_nofilter) > 0:
            print(omt_matches_nofilter[['Wine Name', 'Item No. Int', 'Unit Price', 'Minimum Quantity', 'Campaign Type', 'Campaign Sub-Type']].to_string(index=False))
