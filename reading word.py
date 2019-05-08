import docx
doc = docx.Document('task.docx')
print(doc.tables[1].rows[1].cells[5].text)

# table[№] номер таблицы rows[№] номер строки cells[№] номер столбца