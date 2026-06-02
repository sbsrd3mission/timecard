import openpyxl

def fix_excel():
    file_path = 'タイムカード入力空バージョン.xlsx'
    wb = openpyxl.load_workbook(file_path, data_only=False)
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in range(2, 33):
            cell_j = sheet[f'J{row}']
            # If the cell has the old formula, update it to just F-C
            if cell_j.value == f'=F{row}-C{row}-(E{row}-D{row})':
                cell_j.value = f'=F{row}-C{row}'
            elif isinstance(cell_j.value, str) and str(cell_j.value).startswith('=') and f'-(E{row}-D{row})' in str(cell_j.value):
                # Replace the middle break part just in case
                cell_j.value = cell_j.value.replace(f'-(E{row}-D{row})', '')
                
    wb.save(file_path)
    print("Fixed タイムカード入力空バージョン.xlsx")

if __name__ == '__main__':
    fix_excel()
