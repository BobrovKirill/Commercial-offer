from openpyxl import load_workbook
import docx


        # считывание с excel


all_product = {}
wb = load_workbook(filename='test.xlsx', data_only=True)
ws = wb.active

for data in ws.values:
    if isinstance(data[0], int):
        all_product[data[1]] = data[2:]


        # поиск товара и добавление его в таблицу word


def request(key):
    if key in all_product:
        i = 1
        for key,val in all_product.items():
            data_cells = table.add_row().cells
            data_cells[0].text = str(i)
            data_cells[1].text = key
            data_cells[2].text = str(val[0])
            data_cells[3].text = str(val[1])
            data_cells[4].text = str(val[2])
            i+=1
    else:
        print('В списке нет данного товара')

file = 'task2.docx'
doc = docx.Document(file)
table = doc.tables[1]
table.style = 'Table Grid'
table.autofit = True
request('apple')
doc.save(file)
