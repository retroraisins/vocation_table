from tkinter import *
from docxtpl import DocxTemplate
import tkinter.messagebox as mb
from datetime import datetime, date
from tkcalendar import Calendar


class OtpusknoyBilet:
    def __init__(self,parent):
        self.windowotpusk = Toplevel(parent)
        self.windowotpusk.resizable(True, True)
        self.windowotpusk.wm_attributes('-alpha', 0.9)
        self.windowotpusk['bg'] = 'grey'
        self.windowotpusk.title("Отпускной билет")

        # context = (textId.get(0.0, END + "-1c") + "," + textDolznost.get(0.0, END + "-1c") + "," + textZvanie.get(0.0,
        #            END + "-1c") + "," + textSurname.get(0.0, END + "-1c") + "," + textName.get(0.0, END + "-1c") + "," + \
        #            textSecondName.get(0.0, END + "-1c")).split(",")
        # print(print_selection (context))
        # lblzvanie = Label(self.windowotpusk, text="Звание:", bd=3).grid(column=0, row=0, sticky=W)
        # entryzvanie = Entry(self.windowotpusk, width=60, bd=3)
        # entryzvanie.insert(0, context[2])
        # entryzvanie.grid(column=1, row=0, sticky=W)
        # lblsurname = Label(self.windowotpusk, text="Фамилия:", bd=3).grid(column=0, row=1, sticky=W)
        # entrysurname = Entry(self.windowotpusk, width=60, bd=3)
        # entrysurname.insert(0, context[3])
        # entrysurname.grid(column=1, row=1, sticky=W)
        # lblname = Label(self.windowotpusk, text="Имя:", bd=3).grid(column=0, row=2, sticky=W)
        # entryname = Entry(self.windowotpusk, width=60, bd=3)
        # entryname.insert(0, context[4])
        # entryname.grid(column=1, row=2, sticky=W)
        # lblsecondname = Label(self.windowotpusk, text="Отчество:", bd=3).grid(column=0, row=3, sticky=W)
        # entrysecondname = Entry(self.windowotpusk, width=60, bd=3)
        # entrysecondname.insert(0, context[5])
        # entrysecondname.grid(column=1, row=3, sticky=W)
        # lblvidotpuska = Label(self.windowotpusk, text="Вид отпуска:", bd=3).grid(column=0, row=4, sticky=W)
        # cmbvidotpuska = ttk.Combobox(self.windowotpusk,
        #                              values=["основной отпуск", "каникулярный отпуск", "дополнительный отпуск",
        #                                      "отпуск по беременности и родам", "отпуск по уходу за ребенком"], width=57)
        # cmbvidotpuska.current(0)
        # cmbvidotpuska.grid(column=1, row=4, sticky=W)
        # lblyear = Label(self.windowotpusk, text="За какой год:", bd=3).grid(column=0, row=5, sticky=W)
        # nowyear = int(date.today().strftime("%Y"))
        # years = []
        # for i in range(nowyear - 3, nowyear + 5):
        #     years.append(i)
        # entryyear = ttk.Combobox(self.windowotpusk, values=years, width=57)
        # entryyear.current(3)
        # entryyear.grid(column=1, row=5, sticky=W)
        # lbltown = Label(self.windowotpusk, text="Место проведения отпуска:", bd=3).grid(column=0, row=6, sticky=W)
        # entrytown = Entry(self.windowotpusk, width=60, bd=3)
        # entrytown.grid(column=1, row=6, sticky=W)
        # lbltimeon = Label(self.windowotpusk, text="Сроком на (дней):", bd=3).grid(column=0, row=7, sticky=W)
        # vacationDuration = StringVar()
        # vacationDuration.set(35)
        # entrytimeon = Entry(self.windowotpusk, width=60, bd=3, textvariable=vacationDuration)
        # entrytimeon.grid(column=1, row=7, sticky=W)
        # firstDayVacation = StringVar()
        # firstWorkDay = StringVar()
        # firstDayVacation.set('С какого числа\n убыть в отпуск?')
        # lblstart = Label(self.windowotpusk, textvariable=firstDayVacation, bd=3)
        # lblstart.grid(column=0, row=8, sticky=W)
        # now = datetime.now()
        # day1 = int(now.strftime("%d"))
        # month1 = int(now.strftime("%m"))
        # year1 = int(now.strftime("%Y"))
        # infoSave = StringVar()
        #
        # def print_sel(e):
        #     firstDayVacation.set('Убыть с ' + cal.get_date())
        #     dateOfVacation = cal.selection_get() + cal.timedelta(days=int(entrytimeon.get())) - cal.timedelta(days=1)
        #     dateOfVacation = dateOfVacation.strftime("%d.%m.%Y")
        #     entryendVacation.delete(0, END)
        #     entryendVacation.insert(0, dateOfVacation)
        #     fWorkDay = cal.selection_get() + cal.timedelta(days=int(entrytimeon.get()))
        #     fWorkDay = fWorkDay.strftime("%d.%m.%Y")
        #     firstWorkDay.set(fWorkDay)
        #     infoSave.set("")
        #
        # def listForOtpBilet():
        #
        #     firstDayVacation.set('Убыть с ' + cal.get_date())
        #     dateOfVacation = cal.selection_get() + cal.timedelta(days=int(entrytimeon.get())) - cal.timedelta(days=1)
        #     dateOfVacation = dateOfVacation.strftime("%d.%m.%Y")
        #     entryendVacation.delete(0, END)
        #     entryendVacation.insert(0, dateOfVacation)
        #     fWorkDay = cal.selection_get() + cal.timedelta(days=int(entrytimeon.get()))
        #     fWorkDay = fWorkDay.strftime("%d.%m.%Y")
        #     firstWorkDay.set(fWorkDay)
        #     if len(entrytown.get()) < 3:
        #         mb.showwarning("Внимание", "Заполните место проведения отпуска")
        #         return
        #     doc = DocxTemplate("Образцы документов/Отпускной билет.docx")
        #     dataForOtpysknoy = [entryzvanie.get(), entrysurname.get(), entryname.get(), entrysecondname.get(),
        #                         cmbvidotpuska.get(), entryyear.get(),
        #                         entrytown.get(), vacationDuration.get(), number_to_words(int(vacationDuration.get())),
        #                         cal.get_date(),
        #                         entryendVacation.get(), entryFirstWorkDay.get(), entryNumberOfVPD.get(),
        #                         entryTogether.get()]
        #     inicialname = list(dataForOtpysknoy[2])
        #     inicialname = inicialname[0]
        #     inicialsecondname = list(dataForOtpysknoy[3])
        #     inicialsecondname = inicialsecondname[0]
        #
        #     doctext = {'zvanie': dataForOtpysknoy[0], 'surname': dataForOtpysknoy[1], 'name': dataForOtpysknoy[2],
        #                'secondname': dataForOtpysknoy[3],
        #                'inicialname': inicialname, 'inicialsecondname': inicialsecondname,
        #                'vidotpuska': dataForOtpysknoy[4], 'year':
        #                    dataForOtpysknoy[5], 'town': dataForOtpysknoy[6], 'vacationDuration': dataForOtpysknoy[7],
        #                'number_to_words': dataForOtpysknoy[8],
        #                'firstDayVacation': dataForOtpysknoy[9], 'endVacation': dataForOtpysknoy[10],
        #                'FirstWorkDay': dataForOtpysknoy[11],
        #                'NumberOfVPD': dataForOtpysknoy[12], 'S_Kem': zvaniya_tvoritelniy.get(dataForOtpysknoy[0]),
        #                'Together': dataForOtpysknoy[13]}
        #     doc.render(doctext)
        #     doc.save("Отпускные билеты/" + dataForOtpysknoy[1] + " " + str(
        #         date.today().strftime("%d.%m.%Y")) + ' Отпускной билет.docx')
        #     infoSave.set('Отпускной билет сохранен в папке "Отпускные билеты"')
        #
        # cal = Calendar(self.windowotpusk, width=57, year=year1, day=day1, month=month1,
        #                background='darkblue', foreground='white', borderwidth=2, date_pattern='dd.mm.y', locale="ru")
        # cal.grid(column=1, row=8, padx=3, pady=3)
        # cal.bind("<<CalendarSelected>>", print_sel)
        # lblend = Label(self.windowotpusk, text="Отпуск по:", bd=3).grid(column=0, row=9, sticky=W)
        # entryendVacation = Entry(self.windowotpusk, width=60, bd=3)
        # entryendVacation.grid(column=1, row=9, sticky=W)
        # lblFirstWorkDay = Label(self.windowotpusk, text="Выйти на службу: ", bd=3).grid(column=0, row=10, sticky=W)
        # entryFirstWorkDay = Entry(self.windowotpusk, width=60, bd=3, textvariable=firstWorkDay)
        # entryFirstWorkDay.grid(column=1, row=10, sticky=W)
        # lblNumberOfVPD = Label(self.windowotpusk, text="Вписать номера ВПД: ", bd=3).grid(column=0, row=11, sticky=W)
        # entryNumberOfVPD = Entry(self.windowotpusk, width=60, bd=3)
        # entryNumberOfVPD.grid(column=1, row=11, sticky=W)
        # lblTogether = Label(self.windowotpusk, text="Кто следует вместе \n с военнослужащим: ", bd=3).grid(column=0, row=12,
        #                                                                                               sticky=W)
        # entryTogether = Entry(self.windowotpusk, width=60, bd=3)
        # entryTogether.grid(column=1, row=12, sticky=W)
        # spisok = Button(self.windowotpusk, text="Сформировать отпускной билет", command=listForOtpBilet)
        # spisok.grid(column=1, row=13, pady=5)
        #
        # lblInfo = Label(self.windowotpusk, textvariable=infoSave, bd=3)
        # lblInfo.grid(column=1, row=14, sticky=W)

