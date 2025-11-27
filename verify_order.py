import pandas as pd

df = pd.read_excel(r'Outputs\Detailed match results\Main offer\Lines.xlsx')

print('First 20 wines in Lines.xlsx (in order from Multi.txt):')
print('='*100)

for i, row in df.head(20).iterrows():
    wine_name = row['Wine Name']
    vintage = row['Vintage Code']
    item_no = row['Unnamed: 11']
    min_qty = row['Minimum Quantity']
    chf = row['Unit Price']
    eur = row['Unit Price (€)']

    print(f"{i+1}. {wine_name} {vintage} | Item {item_no} | Min Qty {min_qty} | {chf} CHF -> {eur} EUR")

print('\n' + '='*100)
print(f'\nTotal wines in Lines.xlsx: {len(df)}')
print(f'Wines with Min Qty = 0: {(df["Minimum Quantity"] == 0).sum()}')
print(f'Wines with Min Qty = 36: {(df["Minimum Quantity"] == 36).sum()}')
