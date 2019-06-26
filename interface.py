import tkinter as tk
from tkinter import ttk
import os
from tkinter import filedialog
from tkinter import *
from openpyxl import load_workbook
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH


def readyng_excel(): #Считывание с excel файла
    global all_product
    all_product = {}
    wb = load_workbook(filename = fileexcel, data_only=True)
    ws = wb.active

    for data in ws.values:
        if isinstance(data[0], int):
            all_product[data[1]] = data[2:]
    return (all_product)

def dir_excel(): #Выбор пути до файла
    global fileexcel
    fileexcel = filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx' )])
    path_excel.set (fileexcel)

def dir_word(): #Выбор пути до файла
    global fileword
    fileword = filedialog.askopenfilename(filetypes=[('Word files', '*.docx')])
    path_word.set(fileword)

def save_excel():
    f = fileexcel
    print(f)

def save_word():
    w = fileword
    print(fileword)

class Main(tk.Frame): #создание главного окна
    def __init__(self, root):
        super().__init__(root)
        self.init_main()

    def init_main(self):
        toolbar = tk.Frame(bg = '#d7d8e0', bd = 2)
        toolbar.place(width = 430, height = 150 )
        btn_open_excel = tk.Button(toolbar, text = 'Укажите путь до Excel файла', command = self.open_exel, bg = '#d7d8e0',bd = 2, compound = tk.TOP )
        btn_open_excel.place(x= 30, y = 10)
        btn_open_word = tk.Button(toolbar, text = 'Укажите путь до Word файла', command = self.open_word, bg = '#d7d8e0',bd = 2, compound = tk.TOP )
        btn_open_word.place(x= 220, y= 10)
        btn_start = tk.Button(toolbar, text = 'Создать файл', command = start)
        btn_start.place(x=175, y=100 )


    def open_exel(self):
        child_excel()
    def open_word(self):
        child_word()

class child_excel(tk.Toplevel): #создание окна для выбора файла
    def __init__(self):
        super().__init__(root)
        self.init_child()

    def init_child(self):
        self.title('Выбор Exel файла')
        self.geometry('550x150+700+350')
        self.resizable(False, False)
        self.grab_set()
        self.focus_set()

        labal = ttk.Label(self, text = 'Выберете путь к Excel файлу:')
        labal.place(x=10 , y=30)

        self.entry_excel = ttk.Entry(self, width = 45, textvariable = path_excel)
        self.entry_excel.place(x=180, y=30)

        btn_cancel = ttk.Button(self, text = 'Закрыть', command = self.destroy)
        btn_cancel.place(x=465, y=100)

        btn_ok = ttk.Button(self, text = 'Ok', command = readyng_excel)
        btn_ok.place(x=385, y=100)
        btn_ok.bind('<Button-1>')

        btn_browse = ttk.Button(self, text = 'Обзор', command = dir_excel)
        btn_browse.place(x=465, y=28)
        btn_browse.bind('<Button-1>')

class child_word(tk.Toplevel): #создание окна для выбора файла
    def __init__(self):
        super().__init__(root)
        self.init_child()

    def init_child(self):
        self.title('Выбор Word файла')
        self.geometry('550x150+700+350')
        self.resizable(False, False)
        self.grab_set()
        self.focus_set()

        labal = ttk.Label(self, text = 'Выберете путь к Word файлу:')
        labal.place(x=10 , y=30)

        self.entry_excel = ttk.Entry(self, width = 45, textvariable = path_word)
        self.entry_excel.place(x=180, y=30)

        btn_cancel = ttk.Button(self, text = 'Закрыть', command = self.destroy)
        btn_cancel.place(x=465, y=100)

        btn_ok = ttk.Button(self, text = 'Ok', command = readyng_word)
        btn_ok.place(x=385, y=100)
        btn_ok.bind('<Button-1>')

        btn_browse = ttk.Button(self, text = 'Обзор', command = dir_word)
        btn_browse.place(x=465, y=28)
        btn_browse.bind('<Button-1>')

        # поиск товара и добавление его в таблицу word

def readyng_word():
    def request(key):
        if key in all_product:
            i = 1
            global sum_price
            sum_price = 0
            for key,val in all_product.items():
                data_cells = table.add_row().cells
                data_cells[0].text = str(i)
                data_cells[1].text = key
                data_cells[2].text = str(val[0])
                data_cells[3].text = str(val[1])
                data_cells[4].text = str(val[2])
                data_cells[5].text = str(val[3])
                sum_price += val[3]
                i += 1
                for j in range(6):
                    paragraph = data_cells[j].paragraphs[0]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Открытие ворда и добавление стилей

    file = fileword # сделать так чтоб пользователь указывал название и путь шаблона
    global doc
    doc = docx.Document(file)
    table = doc.tables[1]
    request('apple')

            #формирование двух последних строк с объединением столбцов


    nums_cells = len(table.rows)
    data_cells = table.add_row().cells
    for i in range(2):
        merget = table.cell(nums_cells,i).merge(table.cell(nums_cells,i+1))
        merget = table.cell(nums_cells,i+3).merge(table.cell(nums_cells,i+4))
    data_cells[0].text = 'Итого, рублей'
    data_cells[3].text = str(sum_price)
    paragraph = data_cells[0].paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph = data_cells[3].paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    data_cells = table.add_row().cells
    for i in range(2):
        merget = table.cell(nums_cells+1,i).merge(table.cell(nums_cells+1,i+1))
        merget = table.cell(nums_cells+1,i+3).merge(table.cell(nums_cells+1,i+4))
    data_cells[0].text = 'В том числе НДС, руб:'
    data_cells[3].text = 'Без НДС*'
    paragraph = data_cells[0].paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph = data_cells[3].paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def start():
    doc.save('Новое коммерч.docx')

if __name__ == '__main__':
    root = tk.Tk()
    path_excel = StringVar()
    path_word = StringVar()
    app = Main(root)
    app.pack()
    root.title('Commercial Offer')
    root.geometry('430x150+700+350')
    root.resizable(False,False)
    root.mainloop()

