# pyinstaller --onedir --noconsole --hidden-import="sklearn.utils._typedefs" --hidden-import="sklearn.utils._heap" --hidden-import="sklearn.utils._sorting" --hidden-import="sklearn.utils._vector_sentinel" main.py

import tkinter as tk
from tkinter import filedialog as fd
import customtkinter as ctk
# from customtkinter import *

import pandas as pd
import os
import sys
import runpy
import time
# import threading
from threading import Thread
import multiprocessing as mp
import subprocess

from main import PI

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

    try:
        df_initial = pd.read_excel(os.path.join(os.path.dirname(__file__), data_init))  # Открытие экселя
        df_gtm = pd.read_excel(os.path.join(os.path.dirname(__file__), data_gtm), header=1)  # Открытие экселя с ГТМ
        # btn_3.configure(text="Стоп", state=tk.NORMAL)
        # # btn_3.set('Стоп')
        # win.update()
        # # Создание кнопки "Остановить"
        # btn_3.bind('<ButtonRelease-1>', stop_script)
        # win.after(5, lambda: btn_3.configure(state=tk.NORMAL))

        # win.after(1000, get_file)

        # update_timer()
        times = PI(df_initial, df_gtm, data_init)
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

    except:
        root = tk.Toplevel()
        root.title('Error')
        root.attributes('-toolwindow', True)
        if getattr(sys, 'frozen', False):
            # If the application is run as a bundle, the PyInstaller bootloader
            # extends the sys module by a flag frozen=True and sets the app
            # path into variable _MEIPASS'.
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(application_path, 'point.png')
        logo = tk.PhotoImage(file=config_path)
        root.iconphoto(False, logo)
        root.config(bg='light grey')  # фон, можно вместо названия цвета указать хеш; bg(background) – «фон». fg(foreground) ) - «передний план»
        root.geometry("200x80+200+200")  # размер окна и расположение (его можно не указывать)
        root.resizable(width=False, height=False)  # Нельзя изменять размер окна
        label_1 = tk.Label(root, text="Исходные данные \n "
                                      "не выбраны!",
                           bg='light grey',
                           fg='black',
                           font=("Arial", 9, "bold"),
                           # padx=20, # отступ по x
                           # pady=30, # отступ по y
                           # width=20, # ширина
                           # height=10, # высота
                           # anchor='n', # расположение текста в лейбле
                           justify=tk.CENTER)
        label_1.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    return times

    # runpy.run_path(path_name='GUI.py')
    # exit()

def open_report():
    if not entry_1.get() == 'Выберите файл':
        subprocess.Popen(str(entry_1.get()).replace(".xlsx", "") + "_out.xlsx", shell=True)
        # os.system(str(entry_1.get()).replace(".xlsx", "") + "_out.xlsx")
    else:
        root = tk.Toplevel()
        root.title('Error')
        root.attributes('-toolwindow', True)
        if getattr(sys, 'frozen', False):
            # If the application is run as a bundle, the PyInstaller bootloader
            # extends the sys module by a flag frozen=True and sets the app
            # path into variable _MEIPASS'.
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(application_path, 'point.png')
        logo = tk.PhotoImage(file=config_path)
        root.iconphoto(False, logo)
        root.config(bg='light grey')  # фон, можно вместо названия цвета указать хеш; bg(background) – «фон». fg(foreground) ) - «передний план»
        root.geometry("200x80+200+200")  # размер окна и расположение (его можно не указывать)
        root.resizable(width=False, height=False)  # Нельзя изменять размер окна
        label_1 = tk.Label(root, text="Расчет не выполнен!",
                           bg='light grey',
                           fg='black',
                           font=("Arial", 9, "bold"),
                           # padx=20, # отступ по x
                           # pady=30, # отступ по y
                           # width=20, # ширина
                           # height=10, # высота
                           # anchor='n', # расположение текста в лейбле
                           justify=tk.CENTER)
        label_1.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

def counter_label(label):
    a = 0
    label.config(text=str(a))
    def count():
        nonlocal a
        label.config(text=str(a))
        a += 1
        label.after(1000, count)
    return count

def start_stop(root, counter):
    first = True
    def call():
        nonlocal first
        if first:
            counter()
            first = False
        else:
            root.destroy()
    return call


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
win.geometry("400x450+100+100")  # размер окна и расположение (его можно не указывать)
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


# label_1.pack()
# btn_1.pack()  # виджет автоматически ставится по центру окна
# btn_2.grid(row=0, column=0) # расположение виджетов по колонкам
# entry_1.pack()

btn_1.bind('<ButtonRelease-1>', lambda event, button=btn_1, entry=entry_1: load_file(button, entry))
btn_2.bind('<ButtonRelease-1>', lambda event, button=btn_2, entry=entry_2: load_file(button, entry))
btn_3.bind('<ButtonRelease-1>', get_file)

# label = tk.Label(win, fg="green")
# label.pack()
# counter = counter_label(label)

def threaded_run(event):
    t = Thread(target=start_stop(win, counter))
    t.daemon = True
    t.start()


def run_task(event):
    p = mp.Process(target=start_stop(win, counter))
    p.start()


# button = tk.Button(root, textvariable=btn_text, width=25, command=start_stop(root, btn_text, counter))

frame_1.place(relx=0.5, rely=0.31, anchor=tk.CENTER)
frame_2.place(relx=0.5, rely=0.83, anchor=tk.CENTER)
label_1.place(relx=0.5, rely=0.1, anchor=tk.CENTER)
label_2.place(relx=0.5, rely=0.2, anchor=tk.CENTER)
btn_1.place(relx=0.5, rely=0.26, anchor=tk.CENTER)
btn_2.place(relx=0.5, rely=0.66, anchor=tk.CENTER)
btn_3.place(relx=0.5, rely=0.63, anchor=tk.CENTER)
btn_4.place(relx=0.5, rely=0.55, anchor=tk.CENTER)
entry_1.place(relx=0.5, rely=0.44, anchor=tk.CENTER)
entry_2.place(relx=0.5, rely=0.84, anchor=tk.CENTER)


win.mainloop()  # запускает цикл обработки событий; пока мы не вызовем эту функцию, окно не откроется

