import pandas as pd
import sys
import io

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

df = pd.read_excel(r'Outputs\OMT lines\Matched_OMT Main Offer List_20251124_184830.xlsx')

print('Columns (first 5):', df.columns.tolist()[:5])
print('\nFirst 30 wines from "Extracted Wine Name (from Multi.txt)" column:')
print('='*80)

for i, name in enumerate(df['Extracted Wine Name (from Multi.txt)'].head(30), 1):
    print(f'{i}. {name}')

print('\n' + '='*80)
print('\nExpected order (from user):')
expected = [
    "Magnum Cristal Rosé 2014",
    "Bollinger RD 2008",
    "Krug Rosé 29ème Édition",
    "Champagne Brut Cristal - Louis Roederer 2012",
    "Champagne Brut – Delamotte",
    "Champagne Grande Année Bollinger 2015",
    "Château Haut-Marbuzet 2019",
    "Phélan Ségur 2018",
    "Crognolo 2022",
    "Château Grand-Puy-Lacoste 2021"
]

for i, name in enumerate(expected, 1):
    print(f'{i}. {name}')
