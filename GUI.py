import tkinter as tk
from tkinter import filedialog as fd
import customtkinter as ctk
# from customtkinter import *
from tkinter.ttk import Progressbar
import pandas as pd
import os
import sys
import subprocess
from time import sleep

from main import PI
from function import errors

times = 0

def load_file(button, entry):
    # global data_file
    # global file_name
    file_name = fd.askopenfilename(defaultextension=('.xlsx', '.xls'))
    print(file_name)
    # print(data_file)
    entry.configure(state='normal')
    entry.delete(0, tk.END)
    if file_name:
        entry.insert(0, file_name)
        entry.configure(state='disabled')
    else:
        entry.insert(0, 'Выберите файл')
    # return


def stop_script():
    btn_3.configure(state=tk.DISABLED)
    # win.destroy()
    print("Скрипт остановлен.")


def get_file(event):

    # win.after(5, get_file)
    data_init = entry_1.get()
    data_gtm = entry_2.get()
    max_distance = int(entry_3.get())

    if data_init == 'Выберите файл' or data_gtm == 'Выберите файл':
        Flag = 1
        errors(Flag)

    else:
        pb['value'] = 0
        win.update_idletasks()
        label_5['text'] = round(pb['value']), '%'
        win.update_idletasks()

        if r_var.get() == 1:
            Flag_smooth = 1 # сглаженная кривая
        else:
            Flag_smooth = 0 # не сглаженная кривая

        try:
            df_initial = pd.read_excel(os.path.join(os.path.dirname(__file__), data_init))  # Открытие экселя
            df_gtm = pd.read_excel(os.path.join(os.path.dirname(__file__), data_gtm), header=1)  # Открытие экселя с ГТМ

            times = PI(df_initial, df_gtm, data_init, max_distance, Flag_smooth, win, pb, label_5)
            print(times)

            label_3 = tk.Label(master=frame_2, text=f"Время расчёта: {int(times // 60)} мин. {int(times % 60)} сек.",
                               bg='#EDEDED',
                               fg='black',
                               font="Arial 9",
                               # padx=20, # отступ по x
                               # pady=30, # отступ по y
                               # width=20, # ширина
                               # height=10, # высота
                               # anchor='n', # расположение текста в лейбле
                               justify=tk.RIGHT)
            label_3.place(relx=0.5, rely=0.85, anchor=tk.CENTER)

            return times

        except:
            Flag = 2
            errors(Flag)

def open_report():
    if entry_1.get() == 'Выберите файл':
        Flag = 1
        errors(Flag)
    elif pb['value'] > 99.9:
        subprocess.Popen(str(entry_1.get()).replace(".xlsx", "") + "_out.xlsx", shell=True)
        # os.system(str(entry_1.get()).replace(".xlsx", "") + "_out.xlsx")
    else:
        Flag = 3
        errors(Flag)

win = tk.Tk()  # главное окно, еще называют root

win.title('Productivity index')
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(application_path, 'logo.png')
logo = tk.PhotoImage(file=config_path)
win.iconphoto(False, logo)
win.config(bg='#DADADA')  # фон, можно вместо названия цвета указать хеш; bg(background) – «фон». fg(foreground) ) - «передний план»
win.geometry("400x500+100+100")  # размер окна и расположение (его можно не указывать)
# win.minsize(400, 300) # минимальный возможный размер, если True в resizable
# win.minsize(800, 700) # максимальный возможный размер, если True в resizable
win.resizable(width=False, height=False)  # Нельзя изменять размер окна


# вставляем виджеты
frame_1 = ctk.CTkFrame(master=win,
                       width=250,
                       height=190,
                       corner_radius=20,
                       fg_color='#EDEDED'
                       )

frame_2 = ctk.CTkFrame(master=win,
                       width=250,
                       height=90,
                       corner_radius=20,
                       fg_color='#EDEDED'
                       )

label_1 = tk.Label(master=frame_1, text="Загрузка исходных данных",
                   bg='#EDEDED',
                   fg='black',
                   font="Arial 11 bold",
                   # padx=20, # отступ по x
                   # pady=30, # отступ по y
                   # width=50, # ширина
                   # height=10, # высота
                   # anchor='n', # расположение текста в лейбле
                   justify=tk.CENTER
                   )

label_2 = tk.Label(master=frame_2, text="Результаты",
                   bg='#EDEDED',
                   fg='black',
                   font="Arial 11 bold",
                   # padx=20, # отступ по x
                   # pady=30, # отступ по y
                   # width=20, # ширина
                   # height=10, # высота
                   # anchor='n', # расположение текста в лейбле
                   justify=tk.CENTER)

label_4 = tk.Label(win, text="Радиус поиска ППД:",
                   bg='#DADADA',
                   fg='black',
                   font="Arial 10",
                   # padx=20, # отступ по x
                   # pady=30, # отступ по y
                   # width=20, # ширина
                   # height=10, # высота
                   # anchor='n', # расположение текста в лейбле
                   justify=tk.CENTER)

label_5 = tk.Label(master=win, text="",
                   bg='#DADADA',
                   fg='black',
                   font="Arial 9",
                   justify=tk.CENTER)


btn_1 = tk.Button(master=frame_1, text='Эксплуатационные показатели', width=30, bg='#909090', fg='white')

btn_2 = tk.Button(master=frame_1, text='Выполненные ГТМ', width=30, bg='#909090', fg='white')

btn_3 = ctk.CTkButton(win, text='Расчёт', corner_radius=10, width=150, fg_color='#0070BA', font=('Arial', 12,  'bold'))

btn_4 = tk.Button(master=frame_2, text='Открыть отчёт',
                  command=open_report,
                  width=30,
                  bg='#909090',
                  fg='white'
                  )


entry_1 = tk.Entry(master=frame_1, width=31)
entry_1.insert(0, 'Выберите файл')

entry_2 = tk.Entry(master=frame_1, width=31)
entry_2.insert(0, 'Выберите файл')

entry_3 = tk.Entry(win, width=10, justify='center')
entry_3.insert(0, '1000')

pb = Progressbar(win, orient="horizontal", mode="determinate", length=150, maximum=100, value=0)

btn_1.bind('<ButtonRelease-1>', lambda event, button=btn_1, entry=entry_1: load_file(button, entry))
btn_2.bind('<ButtonRelease-1>', lambda event, button=btn_2, entry=entry_2: load_file(button, entry))
btn_3.bind('<ButtonRelease-1>', get_file)

r_var = tk.BooleanVar()
r_var.set(0)
r1 = tk.Radiobutton(master=win, text='Не сглаживать', variable=r_var, value=0, font="Arial 10", bg='#DADADA')
r2 = tk.Radiobutton(master=win, text='Сглаживать', variable=r_var, value=1, font="Arial 10", bg='#DADADA')

frame_1.place(relx=0.5, rely=0.23, anchor=tk.CENTER)
frame_2.place(relx=0.5, rely=0.88, anchor=tk.CENTER)
label_1.place(relx=0.5, rely=0.1, anchor=tk.CENTER)
label_2.place(relx=0.5, rely=0.2, anchor=tk.CENTER)
label_4.place(relx=0.2, rely=0.46)
label_5.place(relx=0.7, rely=0.71)
btn_1.place(relx=0.5, rely=0.26, anchor=tk.CENTER)
btn_2.place(relx=0.5, rely=0.66, anchor=tk.CENTER)
btn_3.place(relx=0.5, rely=0.66, anchor=tk.CENTER)
btn_4.place(relx=0.5, rely=0.54, anchor=tk.CENTER)
entry_1.place(relx=0.5, rely=0.44, anchor=tk.CENTER)
entry_2.place(relx=0.5, rely=0.84, anchor=tk.CENTER)
entry_3.place(relx=0.7, rely=0.48, anchor=tk.CENTER)
r1.place(relx=0.2, rely=0.51)
r2.place(relx=0.2, rely=0.56)
pb.place(relx=0.5, rely=0.73, anchor=tk.CENTER)


win.mainloop()  # запускает цикл обработки событий; пока мы не вызовем эту функцию, окно не откроется

