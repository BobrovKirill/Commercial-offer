import tkinter as tk
from tkinter import ttk
import os
from tkinter import filedialog

fileexcel,fileword = '', ''

def dir_excel():
    global fileexcel
    fileexcel = filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx' )])

def dir_word():
    global fileword
    fileword = filedialog.askopenfilename(filetypes=[('Word files', '*.docx')])

def save_excel():
    f = fileexcel
    print(f)

def save_word():
    w = fileword
    print(fileword)

class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()

    def init_main(self):
        toolbar = tk.Frame(bg = '#d7d8e0', bd = 2)
        toolbar.pack(side = tk.TOP, fill = tk.X)
        btn_open_excel = tk.Button(toolbar, text = 'Укажите путь до Excel файла', command = self.open_exel, bg = '#d7d8e0',bd = 2, compound = tk.TOP )
        btn_open_excel.pack(side = tk.LEFT)
        btn_open_word = tk.Button(toolbar, text = 'Укажите путь до Word файла', command = self.open_word, bg = '#d7d8e0',bd = 2, compound = tk.TOP )
        btn_open_word.pack(side = tk.LEFT)


    def open_exel(self):
        child_excel()
    def open_word(self):
        child_word()

class child_excel(tk.Toplevel):
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

        self.entry_excel = ttk.Entry(self, width = 45, textvariable = fileexcel)
        self.entry_excel.place(x=180, y=30)

        btn_cancel = ttk.Button(self, text = 'Закрыть', command = self.destroy)
        btn_cancel.place(x=465, y=100)

        btn_ok = ttk.Button(self, text = 'Ok', command = save_excel)
        btn_ok.place(x=385, y=100)
        btn_ok.bind('<Button-1>')

        btn_browse = ttk.Button(self, text = 'Обзор', command = dir_excel)
        btn_browse.place(x=465, y=28)
        btn_browse.bind('<Button-1>')

class child_word(tk.Toplevel):
    def __init__(self):
        super().__init__(root)
        self.init_child()

    def init_child(self):
        self.title('Выбор Exel файла')
        self.geometry('550x150+700+350')
        self.resizable(False, False)
        self.grab_set()
        self.focus_set()

        labal = ttk.Label(self, text = 'Выберете путь к Word файлу:')
        labal.place(x=10 , y=30)

        self.entry_excel = ttk.Entry(self, width = 45, textvariable = 'lal')
        self.entry_excel.place(x=180, y=30)

        btn_cancel = ttk.Button(self, text = 'Закрыть', command = self.destroy)
        btn_cancel.place(x=465, y=100)

        btn_ok = ttk.Button(self, text = 'Ok', command = save_word)
        btn_ok.place(x=385, y=100)
        btn_ok.bind('<Button-1>')

        btn_browse = ttk.Button(self, text = 'Обзор', command = dir_word)
        btn_browse.place(x=465, y=28)
        btn_browse.bind('<Button-1>')

if __name__ == '__main__':
    root = tk.Tk()
    app = Main(root)
    app.pack()
    root.title('Commercial Offer')
    root.geometry('500x300+700+350')
    root.resizable(False,False)
    root.mainloop()

