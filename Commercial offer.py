from openpyxl import load_workbook
import docx

all_product = {}
wb = load_workbook(filename='test.xlsx', data_only=True)
ws = wb.active

for data in ws.values:
    if isinstance(data[0], int):
        all_product[data[1]] = all_product.get(data[1], []) + [data[2:]]

print(all_product)

