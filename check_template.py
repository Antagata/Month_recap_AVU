import openpyxl

wb = openpyxl.load_workbook(r'Outputs\Detailed match results\Main offer\template\Lines Template.xlsx')
ws = wb.active

print('Template header row:')
for col in range(1, 13):
    cell_value = ws.cell(1, col).value
    col_letter = chr(64+col) if col <= 26 else f'A{chr(64+col-26)}'
    print(f'Column {col} ({col_letter}): {cell_value}')
