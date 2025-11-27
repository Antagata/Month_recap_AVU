import pandas as pd
import sys
import io

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

df = pd.read_excel(r'Outputs\OMT lines\Matched_OMT Main Offer List_20251124_185222.xlsx')

print('First 65 wines from Excel (Extracted from Multi.txt):')
print('='*100)

for i, name in enumerate(df['Extracted Wine Name (from Multi.txt)'].head(65), 1):
    print(f'{i}. {name}')

print('\n' + '='*100)
print('\nExpected wine names (from user):')
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
    "Château Grand-Puy-Lacoste 2021",
    "Château Phélan Ségur 2021",
    "Aromes de Pavie 2020",
    "Fieuzal Rouge 2023",
    "Tignanello - Antinori 2022",
    "Guado al Tasso 2022"
]

for i, name in enumerate(expected, 1):
    print(f'{i}. {name}')
