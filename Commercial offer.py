from openpyxl import load_workbook
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

all_product = {}
wb = load_workbook(filename='test.xlsx', data_only=True)
ws = wb.active

for data in ws.values:
    if isinstance(data[0], int):
        all_product[data[1]] = all_product.get(data[1], []) + [data[2:]]

print(all_product)

document = Document()
first_data = document.add_paragraph('OOO "Ион плюс" \n Тел. (343) 383-27-39 \n E-mail: info@ion-plus.com \n www.ion-plus.com')
first_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
font.name = 'Times New Roman'
font.size = pt(12)

second_data = document.add_paragraph ('620144 г. Екатеринбург, ул. Фурманова, д. 60, оф. 4., ОГРН 1156671002861, ИНН 6671004558, КПП 667101001, р/с 40702810516540018409, кор./с. 30101810500000000674,Уральский банк ПАО «Сбербанк» гЕкатеринбург БИК 046577674')
second_data.aligment = WD_ALIGN_PARAGRAPH.CENTER

document.save('test.docx')