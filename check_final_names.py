import pandas as pd
import sys
import io

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

df = pd.read_excel(r'Outputs\OMT lines\Matched_OMT Main Offer List_20251124_185045.xlsx')

print('First 65 wines from "Extracted Wine Name (from Multi.txt)" column:')
print('='*100)

for i, name in enumerate(df['Extracted Wine Name (from Multi.txt)'].head(65), 1):
    print(f'{i}. {name}')
