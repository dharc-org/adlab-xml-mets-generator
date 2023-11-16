from openpyxl import load_workbook

wb = load_workbook(filename='../file.xlsx', read_only=True)
ws = wb['Sheet1']

for row in ws.rows:
    for cell in row:
        print(cell.value)

# Close the workbook after reading
wb.close()
