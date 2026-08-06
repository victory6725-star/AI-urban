import openpyxl

def print_excel_cells():
    wb = openpyxl.load_workbook('test_real_land_use_table.xlsx')
    ws = wb.active
    print("Excel Table Content:")
    for r in range(1, 20):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 6)]
        if any(v is not None for v in row_vals):
            print(f"Row {r}: {row_vals}")

if __name__ == '__main__':
    print_excel_cells()
