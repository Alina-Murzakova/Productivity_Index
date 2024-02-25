import tkinter as tk
import os
import sys
def lenWell(x1, y1, x2, y2):
    lenght = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
    return lenght

def lenWell_gorizont(x1, y1, x2, y2, xx2, yy2):
    L1 = lenWell(x1, y1, x2, y2)
    L2 = lenWell(x1, y1, xx2, yy2)
    L = lenWell(x2, y2, xx2, yy2) # длина ГС
    if (L1 * L1 > L2 * L2 + L * L) or (L2 * L2 > L1 * L1 + L * L):
        P = min(L1, L2)
    else:
        if x2 == xx2:
            x3 = x2
            y3 = y1
        elif y2 == yy2:
            x3 = x1
            y3 = y2
        else:
            A = yy2 - y2
            B = x2 - xx2
            C = -1 * x2 * (yy2 - y2) + y2 * (xx2 - x2)
            x3 = (x2 * ((yy2 - y2)**2) + x1 * ((xx2 - x2)**2) + (xx2 - x2) * (yy2 - y2) * (y1 - y2)) / ((yy2 - y2)**2 + (xx2 - x2)**2)
            y3 = (xx2 - x2) * (x1 - x3) / (yy2 - y2) + y1

            x3 = (B * x1 / A - C / B - y1) * A * B / (A * A + B * B)
            y3 = B * x3 / A + y1 - B * x1 / A

        P = lenWell(x1, y1, x3, y3)

    return P


def errors(Flag):
    root = tk.Toplevel()
    root.title('Error')
    root.attributes('-toolwindow', True)
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    config_path = os.path.join(application_path, 'point.png')
    print(config_path)
    logo = tk.PhotoImage(file=config_path)
    root.iconphoto(False, logo)
    root.config(bg='light grey')  # фон, можно вместо названия цвета указать хеш; bg(background) – «фон». fg(foreground) ) - «передний план»
    root.geometry("200x80+200+200")  # размер окна и расположение (его можно не указывать)
    root.resizable(width=False, height=False)  # Нельзя изменять размер окна
    if Flag == 1:
        text = "Исходные данные \n ""не выбраны!"
    elif Flag == 2:
        text = "Ошибка в расчёте!"
    else:
        text = "Расчёт не выполнен!"
    label_1 = tk.Label(root, text=text,
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
