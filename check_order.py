import pandas as pd

df = pd.read_excel(r'Outputs\OMT lines\Matched_OMT Main Offer List_20251124_181320.xlsx')

print('Total rows:', len(df))
print('\nExpected Item No. sequence from console output:')
expected = [35638, 49782, 58914, 62045, 59031, 62116, 62044, 51783, 55595, 49069,
            62101, 55632, 60043, 55595, 62116, 64736, 65281, 60436, 57206, 51783]

print('First 20 expected:', expected[:20])
print('\nActual Item No. in Excel (first 20):')
actual = df['Item No. Int'].head(20).tolist()
print(actual)

print('\nAre they in the same order?', expected[:20] == actual)
