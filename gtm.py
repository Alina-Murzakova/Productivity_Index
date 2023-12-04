import numpy as np
import pandas as pd
import os
import sys

# Названия столбцов в таблице
date_gtm = 'Начало.1'
well_number = 'Скважина'
name_gtm = 'Тип'
# data_file_gtm = "Пограничное_ГТМ.xls" # Файл новая стратегия
# df_gtm = pd.read_excel(os.path.join(os.path.dirname(__file__), data_file_gtm), header=1) # Открытие экселя с ГТМ
# df_gtm = df_gtm.fillna(0)  # Заполнение пустых ячеек нулями

list_gtm = []

def gtm_date_convert(df_gtm, prod_well):

    df_gtm = df_gtm[[well_number, name_gtm, date_gtm]]
    if df_gtm[well_number].dtype == 'object':
    # print(type(prod_well))
        prod_well = str(prod_well)
    # print(type(df_gtm[well_number].iloc[0]))
    # print(df_gtm[well_number].dtype)


    df_gtm = df_gtm[df_gtm[well_number] == prod_well]
    wells_prod = df_gtm[well_number].unique()
    # print(type(df_gtm[well_number].iloc[0]))
    # print(df_gtm[well_number].iloc[0].type)
    list_gtms_for_legend = []
    list_colors_for_legend = []
    df_result = pd.DataFrame(columns=['Скважина', 'ГТМ', "Дата", "Сокращение", "Цвет", 'ВПП', 'Перфорация', 'ИДН', 'ОПЗ', 'Оптимизация', 'Промывка/нормализация', 'РИР',
                                      'ГРП', 'Смена ЭЦН', 'ЛАР', 'Прочие'])
    if not df_gtm.empty:
        # for prod_well in wells_prod:
        df_gtm_one_well = df_gtm.loc[(df_gtm[well_number] == prod_well)].copy()
        for i in range(df_gtm_one_well.shape[0]):
            problem = df_gtm_one_well.iloc[[i]]['Тип'].iloc[0]
            gtm = change_name_gtm(problem)
            start_date = df_gtm_one_well.iloc[[i]]['Начало.1'].iloc[0]
            start_date = start_date.replace(day=1, hour=00) # приведение даты к первому числу
            label, color = gtm_names(gtm)
            # создание списка и цветов ГТМ для легенды
            if label not in list_gtms_for_legend:
                list_gtms_for_legend.append(label)
                list_colors_for_legend.append(color)

            df_result = df_result._append({'Скважина': prod_well, 'ГТМ': gtm, "Дата": start_date, "Сокращение": label, "Цвет": color,
                                           'ВПП': np.nan, 'Перфорация': np.nan, 'ИДН': np.nan, 'ОПЗ': np.nan, 'Оптимизация': np.nan, 'Промывка/нормализация': np.nan, 'РИР': np.nan,
                                          'ГРП': np.nan, 'Смена ЭЦН': np.nan, 'ЛАР': np.nan, 'Прочие': np.nan}, ignore_index=True)

            # добавление столбцов
            df_result.loc[(df_result['ГТМ'] == gtm) & (df_result['Дата'] == start_date), [gtm]] = 0


    return df_result, list_gtms_for_legend, list_colors_for_legend

def gtm_names(gtm):

    label = 0
    t = 'Пр'
    color = "black"
    if gtm == 'ВПП':
        t = 'ВПП'
        color = 'royalblue'
    if gtm == 'Перфорация':
        t = 'Перф'
        color = 'brown'
    if gtm == 'ИДН':
        t = 'ИДН'
        color = 'lime'
    if gtm == 'ОПЗ':
        t = 'ОПЗ'
        color = 'orange'
    if gtm == 'Оптимизация':
        t = 'ОПТ'
        color = 'green'
    if gtm == 'Промывка/нормализация':
        t = 'П/Н'
        color = 'olive'
    if gtm == 'РИР':
        t = 'РИР'
        color = 'red'
    if gtm == 'ГРП':
        t = 'ГРП'
        color = 'blue'
    if gtm == 'Смена ЭЦН':
        t = 'ЭЦН'
        color = 'orangered'
    if gtm == 'ЛАР':
        t = 'ЛАР'
        color = 'tomato'
    if gtm == 'Прочие':
        t = 'Пр.'
        color = 'black'
    return t, color


def load_data_evt():
    try:
        if getattr(sys, 'frozen', False):
            # If the application is run as a bundle, the PyInstaller bootloader
            # extends the sys module by a flag frozen=True and sets the app
            # path into variable _MEIPASS'.
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(application_path, 'gtm.txt')
        f = open(config_path, 'r', encoding='utf-8')
        data = f.read()
        f.close()
        dict_gtm = eval(data)
        return dict_gtm
    except:
        pass

#     self.event = df_evt['event'].tolist()
def change_name_gtm(problem):
    # for i in range(len(self.event)):               # замена разных ГТМ типовыми (из файла)
    dict_gtm = load_data_evt()
    for k, v in dict_gtm.items():
    # for i in dict_gtm:
        if problem in v:
    #     if problem in dict_gtm[i]:
            return k
        else:
            k = 'Прочие'
    return k # problem


