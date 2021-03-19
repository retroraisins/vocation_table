from tkinter import *
import tkinter.ttk as ttk
import xlrd
from otpusknoy_bilet import OtpusknoyBilet
ttk = ttk


class Window:

    def __init__(self):
        book = xlrd.open_workbook(filename='Base 5.xls', on_demand=True)
        book.sheet_by_index(0)
        self.root = Tk()
        self.root.resizable(True, True)
        self.root.geometry('1300x500')
        self.root.title("Строевая")
        self.root.wm_attributes('-alpha', 0.9)
        self.root['bg'] = 'grey'
        self.textId = Text(self.root, width=4, height=2, bd=1, wrap=WORD)
        self.textId.grid(row=0, column=0, pady=5)
        self.textDolznost = Text(self.root, width=45, height=2, bd=1, wrap=WORD)
        self.textDolznost.grid(row=0, column=1, pady=5)
        self.textZvanie = Text(self.root, width=15, height=2, bd=1, wrap=WORD)
        self.textZvanie.grid(row=0, column=2, pady=5)
        self.textSurname = Text(self.root, width=14, height=2, bd=1, wrap=WORD)
        self.textSurname.grid(row=0, column=3, pady=5)
        self.textName = Text(self.root, width=14, height=2, bd=1, wrap=WORD)
        self.textName.grid(row=0, column=4, pady=5)
        self.textSecondName = Text(self.root, width=14, height=2, bd=1, wrap=WORD)
        self.textSecondName.grid(row=0, column=5, pady=5)
        self.otpusknoi = Button(self.root, text="Выдать отпускной", command=self.otpusk)
        self.otpusknoi.grid(row=0, column=6, pady=5, padx=5)
        columns = ("#1", "#2", "#3", "#4", "#5", "#6")
        self.tree = ttk.Treeview(self.root, show="headings", columns=columns, height=20)
        self.tree.column('#1', width=35)
        self.tree.column('#2', width=363)
        self.tree.column('#3', width=115)
        self.tree.column('#4', width=115)
        self.tree.column('#5', width=115)
        self.tree.column('#6', width=115)
        self.tree.heading("#1", text="ID")
        self.tree.heading("#2", text="Должность")
        self.tree.heading("#3", text="Звание")
        self.tree.heading("#4", text="Фамилия")
        self.tree.heading("#5", text="Имя")
        self.tree.heading("#6", text="Отчество")
        self.tree.selection()
        self.ysb = ttk.Scrollbar(self.root, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=self.ysb.set)
        self.ysb.grid(row=1, column=6, sticky=W + N + S)
        # ПЕРЕМЕННЫЕ
        self.landString = StringVar()
        self.sheet = book.sheet_by_index(0)
        self.dolnosti = self.sheet.col_values(1, start_rowx=1)
        self.zvanie = self.sheet.col_values(4, start_rowx=1)
        self.surname = self.sheet.col_values(5, start_rowx=1)
        self.name = self.sheet.col_values(6, start_rowx=1)
        self.secondname = self.sheet.col_values(7, start_rowx=1)
        self.id = self.sheet.col_values(0, start_rowx=1)

        # КОД
        otborvoennyh = IntVar()
        otborvoennyh.set(1)

        row = 0
        for i in self.dolnosti:
            self.tree.insert('', END, values=[self.id[row], i, self.zvanie[row], self.surname[row], self.name[row], self.secondname[row]])
            self.tree.grid(row=1, column=0, columnspan=6, sticky=W, padx=2)
            row = 1 + row
        self.tree.bind("<<TreeviewSelect>>", self.print_selection)

    def print_selection(self, event):
        for selection in self.tree.selection():
            item = self.tree.item(selection)
            text = item["values"]
            self.textId.delete(0.0, END)
            self.textDolznost.delete(0.0, END)
            self.textZvanie.delete(0.0, END)
            self.textSurname.delete(0.0, END)
            self.textName.delete(0.0, END)
            self.textSecondName.delete(0.0, END)
            self.textId.insert(0.0, text[0])
            self.textDolznost.insert(0.0, text[1])
            self.textZvanie.insert(0.0, text[2])
            self.textSurname.insert(0.0, text[3])
            self.textName.insert(0.0, text[4])
            self.textSecondName.insert(0.0, text[5])


    def run(self):
        self.root.mainloop()

        # СПИСКИ
    def otpusk(self):
        OtpusknoyBilet(self.root)






    # ИНТЕРФЕЙС

    # icoadr = 'C:/Users/Acer/Downloads/1.ico'
    # root.iconbitmap(icoadr)



    def number_to_words(n):
        f = {0: 'ноль', 1: 'одни', 2: 'двое', 3: 'трое', 4: 'четверо', 5: 'пятеро',
             6: 'шесть', 7: 'семь', 8: 'восемь', 9: 'девять'}
        o = {10: 'десять', 20: 'двадцать', 30: 'тридцать', 40: 'сорок',
             50: 'пятьдесят', 60: 'шестьдесят', 70: 'семьдесят',
             80: 'восемьдесят', 90: 'девяносто'}
        s = {11: 'одиннадцать', 12: 'двенадцать', 13: 'тринадцать',
             14: 'четырнадцать', 15: 'пятнадцать', 16: 'шестнадцать',
             17: 'семнадцать', 18: 'восемнадцать', 19: 'девятнадцать'}
        if n < 10:
            return f.get(n)
        elif 10 < n < 20:
            return s.get(n)
        elif n >= 10 and n in o:
            return o.get(n)
        else:
            n1 = n % 10
            n2 = n - n1
            return o.get(n2) + ' ' + f.get(n1)


    # ФУНКЦИИ




    # СЛОВАРИ
    zvaniya_vinitelniy = {'гп мо':'гп мо','рядовой': 'рядового', 'ефрейтор': 'ефрейтора','младший сержант': 'младшего сержанта','сержант': 'сержанта','старший сержант': 'старшего сержанта', \
                          'старшина': 'старшину','прапорщик': 'прапорщика','старший прапорщик': 'старшего прапорщика', 'младший лейтенант': 'младшего лейтенанта', 'лейтенант': 'лейтенанта',\
                          'старший лейтенант': 'старшего лейтенанта','капитан': 'капитана','майор': 'майора','подполковник': 'подполковника','полковник': 'полковника',
                          'генерал-майор': 'генерал-майора', 'генерал-лейтенант': 'генерал-лейтенанта', 'генерал-полковник': 'генерал-полковника'}
    zvaniya_datelniy = {'гп мо':'гп мо','рядовой': 'рядовому', 'ефрейтор': 'ефрейтору','младший сержант': 'младшему сержанту','сержант': 'сержанту','старший сержант': 'старшему сержанту', \
                          'старшина': 'старшину','прапорщик': 'прапорщику','старший прапорщик': 'старшему прапорщику', 'младший лейтенант': 'младшему лейтенанту', 'лейтенант': 'лейтенанту',\
                          'старший лейтенант': 'старшего лейтенанта','капитан': 'капитана','майор': 'майора','подполковник': 'подполковника','полковник': 'полковника',
                          'генерал-майор': 'генерал-майору', 'генерал-лейтенант': 'генерал-лейтенанту', 'генерал-полковник': 'генерал-полковнику'}
    zvaniya_tvoritelniy = {'гп мо':'гп мо','рядовой': 'рядовым', 'ефрейтор': 'ефрейтором','младший сержант': 'младшим сержантом','сержант': 'сержантом','старший сержант': 'старшим сержантом', \
                          'старшина': 'старшиной','прапорщик': 'прапорщиком','старший прапорщик': 'старшим прапорщиком', 'младший лейтенант': 'младшим лейтенантом', 'лейтенант': 'лейтенантом',\
                          'старший лейтенант': 'старшим лейтенантом','капитан': 'капитаном','майор': 'майором','подполковник': 'подполковником','полковник': 'полковником',
                          'генерал-майор': 'генерал-майором', 'генерал-лейтенант': 'генерал-лейтенантом', 'генерал-полковник': 'генерал-полковником'}

if __name__ == "__main__":
    window = Window()
    window.run()

