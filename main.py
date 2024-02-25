import pandas as pd
import time, os
import numpy as np
import xlwings as xw
from function import lenWell, lenWell_gorizont
import matplotlib.pyplot as plt
import datetime as dt
from gtm import gtm_date_convert, load_data_evt
from Arps import get_max_Kprod, calc_arps
from scipy.optimize import curve_fit
from xlwings.utils import rgb_to_int
import win32api
import win32com.client
# print(win32api.FormatMessage(-2147352567))
from Charts_xlwings import charts
from Plots_matplotlib import plots
from time import sleep

data_file = "Крайнее.xlsx" # Файл для расчета
data_file_gtm = "Крайнее_ГТМ.xls" # Файл новая стратегия

max_distance = 1000 # Максимальное расстояние между скважинами, м
min_period = 12  # Минимальный период совместный работы

# Названия столбцов в Excel
date = 'Дата'
well_number = '№ скважины'
field = 'Месторождение'
work_marker = 'Характер работы'
well_status = 'Состояние'
prod = "НЕФ"
inj = "НАГ"
Qo_rate = 'Дебит нефти за последний месяц, т/сут'
Ql_rate = 'Дебит жидкости за посл.месяц, м3/сут'
water_cut = 'Обводненность за посл.месяц, % (вес)'
time_prod = 'Время работы в добыче, часы'
Winj_rate = 'Приемистость за последний месяц, м3/сут'
time_inj = "Время работы под закачкой, часы"
P_well = "Забойное давление (ТР), атм"
P_pressure = 'Пластовое давление (ТР), атм'
P_well_inj_weighted = 'Средневзвешенное Рзаб ППД'
Kprod = 'Коэффициент продуктивности (ТР), м3/сут/атм'
x1 = "Координата X"
y1 = "Координата Y"
x2 = "Координата забоя Х"
y2 = "Координата забоя Y"
Kprod_TP = 'Кпрод ТР расчет, м3/сут/атм'
Kprod_new = 'Кпрод (Рпл = Рзаб ППД), м3/сут/атм'
check_depletion_drive = 0
Flag_smooth = 0

def PI(df_initial, df_gtm, data_file, max_distance, Flag_smooth, win, pb, label_5):
    print(Flag_smooth)
    # df_initial = pd.read_excel(os.path.join(os.path.dirname(__file__), data_file))  # Открытие экселя
    # df_gtm = pd.read_excel(os.path.join(os.path.dirname(__file__), data_file_gtm), header=1)  # Открытие экселя с ГТМ

    start_time = time.time()
    # df_initial = pd.read_excel(os.path.join(os.path.dirname(__file__), data_file)) # Открытие экселя
    df_initial = df_initial.fillna(0)  # Заполнение пустых ячеек нулями

    # df_gtm = pd.read_excel(os.path.join(os.path.dirname(__file__), data_file_gtm), header=1) # Открытие экселя с ГТМ
    df_gtm = df_gtm.fillna(0)  # Заполнение пустых ячеек нулями
    type_gtm = df_gtm['Тип'].unique()

    last_date = df_initial[date].sort_values()
    last_date = last_date.unique()[-1]

    # Cписок скважин соответствующий характеру работы на последнюю дату
    wells_prod = df_initial[(df_initial[date] == last_date) & (df_initial[work_marker] == prod)][well_number].unique()
    wells_inj = df_initial[(df_initial[date] == last_date) & (df_initial[work_marker] == inj)][well_number].unique()

    df_P_well_inj = pd.DataFrame()
    df_result = pd.DataFrame() # итоговая таблица со всеми доб скважинами и историей
    df_wells = pd.DataFrame() # пары скважин в пределах радиуса
    df_result_well = pd.DataFrame() # итоговая таблица по каждой скважине
    df_r2 = pd.DataFrame()
    df_errors = pd.DataFrame()

    # подключение эксель
    app1 = xw.App(visible=False) # режим без видимого GUI - запуск новых экземпляров Excel
    new_wb = xw.Book() # открыть новый файл Excel

    for prod_well in wells_prod:
        df_injection_data = pd.DataFrame() # таблица с данными нагнетательных скважин ближайшего окружения
        df_wells_one = pd.DataFrame()
        list_inj_wells = [] # список нагнетательных скважин ближайшего окружения
        list_distance = [] # список расстояний до нагнетательных скважин ближайшего окружения
        print(prod_well)

        # для progressbara
        pb['value'] += 100 / len(wells_prod)
        label_5['text'] = round(pb['value']), '%'
        win.update_idletasks()
        sleep(0.05)

        # данные добывающей скважины
        data_prod = df_initial.loc[(df_initial[well_number] == prod_well) & (df_initial[work_marker] == prod)].copy()
        data_prod_not_null = data_prod[data_prod[Ql_rate] != 0]  # Оставляем ненулевые строки по добыче

        # Проверка, работала ли скважина
        if data_prod_not_null.empty:
            print(prod_well, "нулевая история у добывающей скважины")
            error = "нулевая история у добывающей скважины"
            dict_errors= {'Доб.скважина': prod_well, 'Нагн.скважина': 0, 'Причина': error}
            df_errors = df_errors._append(dict_errors, ignore_index=True)
            continue

        # a = data_prod_not_null.shape[0]
        if data_prod_not_null.shape[0] <= 6:
            print(prod_well, "короткая история работы доб.скважины")
            error = "короткая история работы доб.скважины"
            dict_errors= {'Доб.скважина': prod_well, 'Нагн.скважина': 0, 'Причина': error}
            df_errors = df_errors._append(dict_errors, ignore_index=True)
            continue

        last_work_date = data_prod_not_null[date].sort_values()
        last_work_date = last_work_date.iloc[-1]

        diff_date = (last_date - last_work_date).days / 365.25

        if diff_date >= 7:
            print(prod_well, 'скважина не работает более 7 лет')
            error = "скважина не работает более 7 лет"
            dict_errors= {'Доб.скважина': prod_well, 'Нагн.скважина': 0, 'Причина': error}
            df_errors = df_errors._append(dict_errors, ignore_index=True)
            continue

        # Проверка, есть ли у добывающей скважины данные Рзаб
        data_prod_not_null_Pwell = data_prod_not_null[data_prod_not_null[P_well] != 0]
        if data_prod_not_null_Pwell.empty:
            print(prod_well, "нет данных по Рзаб")
            error = "нет данных по Рзаб"
            dict_errors= {'Доб.скважина': prod_well, 'Нагн.скважина': 0, 'Причина': error}
            df_errors = df_errors._append(dict_errors, ignore_index=True)
            continue

        for inj_well in wells_inj:
            # Проверяем насколько далеко расположены скважины
            X_inj_T1 = int(df_initial[(df_initial[date] == last_date) & (df_initial[well_number] == inj_well)][x1])
            Y_inj_T1 = int(df_initial[(df_initial[date] == last_date) & (df_initial[well_number] == inj_well)][y1])
            X_prod_T1 = int(df_initial[(df_initial[date] == last_date) & (df_initial[well_number] == prod_well)][x1])
            Y_prod_T1 = int(df_initial[(df_initial[date] == last_date) & (df_initial[well_number] == prod_well)][y1])
            X_prod_T3 = int(df_initial[(df_initial[date] == last_date) & (df_initial[well_number] == prod_well)][x2])
            Y_prod_T3 = int(df_initial[(df_initial[date] == last_date) & (df_initial[well_number] == prod_well)][y2])

            # проверка вертикальная или горизонтальная скважина и расчет кратчайшего расстояния
            if X_prod_T3 == 0 and Y_prod_T3 == 0:
                distance = lenWell(X_inj_T1, Y_inj_T1, X_prod_T1, Y_prod_T1)
                distance_T3 = lenWell(X_inj_T1, Y_inj_T1, X_prod_T3, Y_prod_T3)
            else:
                distance = lenWell_gorizont(X_inj_T1, Y_inj_T1, X_prod_T1, Y_prod_T1, X_prod_T3, Y_prod_T3)

            # данные нагнетательной скважины
            data_inj = df_initial.loc[(df_initial[well_number] == inj_well) & (df_initial[work_marker] == inj)].copy()
            data_inj_not_null = data_inj[data_inj[Winj_rate] != 0]  # Оставляем ненулевые строки по закачке

            # Проверка работала ли нагнетательная скважина
            if data_inj_not_null.empty:
                print(inj_well, "нулевая история у нагнетательной скважины")
                error = "нулевая история у нагнетательной скважины"
                dict_errors = {'Доб.скважина': prod_well, 'Нагн.скважина': inj_well, 'Причина': error}
                df_errors = df_errors._append(dict_errors, ignore_index=True)
                continue

            # проверка входит ли нагнетательная скважина в заданный радиус
            if distance > max_distance:
                print(prod_well, inj_well, "доб и нагн скважины расположены далеко", distance)
                error = "доб и нагн скважины расположены далеко"
                dict_errors = {'Доб.скважина': prod_well, 'Нагн.скважина': inj_well, 'Причина': error}
                df_errors = df_errors._append(dict_errors, ignore_index=True)
                continue

            else:

                # Проверка совместной истории, если ли совместная история больше 12 месяцев
                data_all = data_prod_not_null.merge(data_inj_not_null, how='inner', on=date)
                pair_wells = data_all.shape[0]
                if pair_wells < min_period:
                    print(prod_well, inj_well, "совместная история работы меньше 12 месяцев")
                    error = "совместная история работы меньше 12 месяцев"
                    dict_errors = {'Доб.скважина': prod_well, 'Нагн.скважина': inj_well, 'Причина': error}
                    df_errors = df_errors._append(dict_errors, ignore_index=True)
                    continue

                # Проверка одновременного наличия данных Рзаб не меньше 6 месяцев
                data_inj_not_null = data_inj_not_null[data_inj_not_null[P_well] != 0] # удаление строк, где нет данных Рзаб
                data_all_with_P = data_prod_not_null_Pwell.merge(data_inj_not_null, how='inner', on=date)
                count_month_with_P = data_all_with_P.shape[0]
                if count_month_with_P < 6:
                    print(prod_well, inj_well, 'совместное наличие данных Рзаб меньше 6 месяцев')
                    error = 'совместное наличие данных Рзаб меньше 6 месяцев'
                    dict_errors = {'Доб.скважина': prod_well, 'Нагн.скважина': inj_well, 'Причина': error}
                    df_errors = df_errors._append(dict_errors, ignore_index=True)
                    continue

                # таблица одной нагнетательной скважины из окружения
                data_inj_not_null = data_inj_not_null[[date, well_number, Winj_rate, P_well]]

                list_inj_wells.append(inj_well) # список нагнетательных скважин окружения
                list_distance.append(distance) # список расстояний между доб скважиной и нагнетательными скважинами окружения

                # формирование таблицы с данными нагнетательных скважин ближайшего окружения
                df_injection_data = df_injection_data._append(data_inj_not_null)

                # Первая дата для графика на выбор
                # first_date = data_prod_not_null.sort_values(by=[date]).reset_index()
                # first_date = first_date.iloc[0][date] # первая дата с добычей

                # data_prod_not_null_Pwell = data_prod[data_prod[P_well] != 0]  # Оставляет ненулевые строки по Рзаб
                first_date = data_prod_not_null_Pwell.sort_values(by=[date]).reset_index()
                first_date = first_date.iloc[0][date]  # первая дата с Рзаб

        print(list_inj_wells) # список нагнетательных скважин ближайшего окружения

        # удаление ненужных столбцов
        data_prod = data_prod.drop([Winj_rate, time_inj, x1, y1, x2, y2], axis=1)

        array_time_work = np.array(data_prod[time_prod]) # массив время работы в часах
        array_time_work_days = np.array(data_prod[time_prod]) / 24 # массив время работы в днях
        data_prod['Накопленное время работы, дни'] = array_time_work_days.cumsum().round(0)
        array_cumulative_days = np.array(data_prod['Накопленное время работы, дни']) # массив накопленное время работы в днях

        production_liq = np.array(data_prod[Ql_rate] * data_prod[time_prod] / 24) # месячная добыча жидкости
        data_prod['Накопленная добыча жидкости, м3'] = production_liq.cumsum()
        array_cumulative_production_liq = np.array(data_prod['Накопленная добыча жидкости, м3']) # массив накопленная добыча жидкости

        # Проверка, есть ли рядом нагнетательные скважины
        # расчет через классический Кпрод
        if len(list_inj_wells) == 0:
            Flag = 0
            print('нет скважин ППД рядом')
            Kprod_result = Kprod_TP
            data_prod.reset_index(inplace=True, drop=True)
            data_prod['Кол-во действ скважин ППД в окружении'] = 0
            data_prod[Winj_rate] = 0
            data_prod['ППД в окружении'] = 0
            data_prod['Средневзвешенное Рзаб ППД'] = 0
            data_prod['Кпрод ТР расчет, м3/сут/атм'] = 0
            data_prod.loc[((data_prod[P_well] > 0) & (data_prod[Ql_rate] != 0) & (data_prod[P_pressure] > 0) |
                           (data_prod[Ql_rate] == 0) & (data_prod[P_pressure] > data_prod[P_well])), 'Кпрод ТР расчет, м3/сут/атм'] = data_prod[Ql_rate] / (data_prod[P_pressure] - data_prod[P_well])
            data_prod['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'] = 0

            data_prod_with_Kprod = data_prod[data_prod['Кпрод ТР расчет, м3/сут/атм'] > 0]  # только строки, где Кпрод больше 0
            Kprod_init = data_prod_with_Kprod['Кпрод ТР расчет, м3/сут/атм'].iloc[0]  # первое значение Кпрод

            array_Kprod_notnull = np.array(data_prod_with_Kprod['Кпрод ТР расчет, м3/сут/атм'])

            data_prod['Темп Кпрод ТР расчет'] = data_prod['Кпрод ТР расчет, м3/сут/атм'] / Kprod_init

            # continue
        else:
            # рядом есть скважины ППД
            # формирование таблицы вида "нагн скв - доб скв - расстояние" (только ближайшее окружение)
            Flag = 1
            Kprod_result = Kprod_new
            df_wells_one['Нагнетательные скважины'] = list_inj_wells
            df_wells_one['Добывающая скважина'] = prod_well
            df_wells_one['Расстояние'] = list_distance
            df_wells = df_wells._append(df_wells_one, ignore_index=True)


            # расчет средневзвешенного значения Рзаб ППД из списка
            df_injection_data["Pзаб*Приемистость"] = df_injection_data[P_well] * df_injection_data[Winj_rate]
            df_injection_data[well_number] = df_injection_data[well_number].astype(str) # измененение типа данных столбца с номером скважины
            df_injection_data['ППД в окружении'] = df_injection_data[well_number]

            inj_parameters_sum = df_injection_data.groupby(date).aggregate({well_number: 'count', Winj_rate: 'sum', 'Pзаб*Приемистость': 'sum', 'ППД в окружении': ', '. join})  # Суммирование всех произведений на определенную дату
            # inj_parameters_sum = df_injection_data.groupby(date).aggregate({well_number: 'count', Winj_rate: 'sum', 'Pзаб*Приемистость': 'sum'}) # Суммирование всех произведений на определенную дату
            inj_parameters_sum = inj_parameters_sum.rename(columns={well_number: 'Кол-во действ скважин ППД в окружении'}) # Переименование столбца
            # inj_parameters_sum['ППД в окружении'] = ", ".join(map(str, list_inj_wells))
            inj_parameters_sum['Средневзвешенное Рзаб ППД'] = inj_parameters_sum["Pзаб*Приемистость"] / inj_parameters_sum[Winj_rate]

            # удаление ненужных столбцов
            inj_parameters_sum = inj_parameters_sum.drop(['Pзаб*Приемистость'], axis=1)

            # добавление средневзв Рзаб ППД в таблицу добывающей скважины
            data_prod = data_prod.merge(inj_parameters_sum, how='left', on=date)
            data_prod['Кпрод ТР расчет, м3/сут/атм'] = 0
            data_prod.loc[((data_prod[P_well] > 0) & (data_prod[Ql_rate] != 0) & (data_prod[P_pressure] > 0) |
                           (data_prod[Ql_rate] == 0)) & (data_prod[P_pressure] > data_prod[P_well]), 'Кпрод ТР расчет, м3/сут/атм'] = data_prod[Ql_rate] / (data_prod[P_pressure] - data_prod[P_well])
            # if inj_parameters_sum['Средневзвешенное Рзаб ППД'].sum() != 0:
            data_prod['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'] = 0
            data_prod.loc[((data_prod[P_well] > 0) & (data_prod[Ql_rate] != 0) & (data_prod[P_pressure] > 0) |
                           (data_prod[Ql_rate] == 0)) & (data_prod[P_pressure] > data_prod[P_well]), 'Кпрод (Рпл = Рзаб ППД), м3/сут/атм'] = data_prod[Ql_rate] / (data_prod['Средневзвешенное Рзаб ППД'] - data_prod[P_well])
            # else:
            #     data_prod['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'] = 0

            data_prod_with_Kprod = data_prod[data_prod['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'] > 0]  # только строки, где Кпрод больше 0
            Kprod_init = data_prod_with_Kprod['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'].iloc[0]  # первое значение Кпрод

            array_Kprod_notnull = np.array(data_prod_with_Kprod['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'])

            data_prod['Темп Кпрод (Рпл = Рзаб ППД)'] = data_prod['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'] / Kprod_init
            # data_prod_with_Kprod_new['Темп Кпрод (Рпл = Рзаб ППД)'] = data_prod_with_Kprod_new['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'] / Kprod_init

        # Flag_smooth = 0
        # считываем Арпса
        data_prod_with_arps, df_r2_well = calc_arps(data_prod, Kprod_result, Flag_smooth)
        # data_prod_with_arps = data_prod_with_arps[[date, 'exponential', 'hyperbolic', 'harmonic', 'hyperbolic_power', 'double', 'minimize']]
        # data_prod_with_arps = data_prod_with_arps[[date, 'hyperbolic', 'hyperbolic_power', 'double', 'minimize']]
        data_prod_with_arps = data_prod_with_arps[[date, 'hyperbolic', 'hyperbolic_power', 'double']]
        count_arps = data_prod_with_arps.shape[1] - 1
        # data_prod = data_prod.merge(data_prod_with_arps, how='left', on=date)
        df_r2 = df_r2._append(df_r2_well, ignore_index=True)

        data_prod.replace([np.inf, -np.inf], 0, inplace=True) #  замена всех бесконечно и минус бесконечно больших чисел на 0
        data_prod.replace([np.nan, -np.nan], 0, inplace=True)  # замена всех бесконечно и минус бесконечно больших чисел на 0

        # считываем ГТМ
        gtm, list_gtms_for_legend, list_colors_for_legend = gtm_date_convert(df_gtm, prod_well)
        data_prod = data_prod.merge(gtm[['Дата', 'ВПП', 'Перфорация', 'ИДН', 'ОПЗ', 'Оптимизация', 'Промывка/нормализация', 'РИР',
                                          'ГРП', 'Смена ЭЦН', 'ЛАР', 'Прочие']], how='left', on=date)

        df_result = df_result._append(data_prod, ignore_index=True) # добавляем таблицу по скважине в общую таблицу

        # построение графиков Кпрод с ГТМ matplotlib
        # fig = plots(prod_well, data_prod, gtm, first_date, last_date, list_gtms_for_legend, list_colors_for_legend)


        # Запись в эксель - каждая добывающая скважина на свой лист
        if str(prod_well) in new_wb.sheets:
            xw.Sheet[str(prod_well)].delete()
        new_wb.sheets.add(str(prod_well))
        sht = new_wb.sheets(str(prod_well))
        if Flag==1:
            sht.api.Tab.ColorIndex = 17

        if not df_result.empty:
            sht.range('A1').options(index=False).value = data_prod
            # sht.pictures.add(fig, name='Plot', update=True, left=sht.range("AB3").left, top=sht.range("AB3").top) # добавление графика matplotlib в эксель

            # print(sht.name.type())
            # ws = new_wb[sht.name]
            # sht = sht.name

            data_prod_with_Kprod = data_prod[data_prod['Кпрод ТР расчет, м3/сут/атм'] > 0]
            index_first_not_null = data_prod_with_Kprod.index[0]
            index_last = data_prod.shape[0]

            # редактируемый график Кпрод
            charts(data_prod, xw, sht, index_first_not_null, index_last, count_arps, Kprod_result, Flag_smooth)

    #  Записываем историю добывающих скважин с расчетным Кпрод и закрываем книгу
    if not df_errors.empty:
        if "Ошибки" in new_wb.sheets:
            xw.Sheet["Ошибки"].delete()
        new_wb.sheets.add("Ошибки")
        sht = new_wb.sheets("Ошибки")
        sht.range('A1').options(index=False).value = df_errors

    if not df_r2.empty:
        if "R2" in new_wb.sheets:
            xw.Sheet["R2"].delete()
        new_wb.sheets.add("R2")
        sht = new_wb.sheets("R2")
        sht.range('A1').options(index=False).value = df_r2

    if not df_wells.empty:
        if "Скважины" in new_wb.sheets:
            xw.Sheet["Скважины"].delete()
        new_wb.sheets.add("Скважины")
        sht = new_wb.sheets("Скважины")
        sht.range('A1').options(index=False).value = df_wells

    if not df_result.empty:
        if "Доб.скважины с Кпрод" in new_wb.sheets:
            xw.Sheet["Доб.скважины с Кпрод"].delete()
        new_wb.sheets.add("Доб.скважины с Кпрод")
        sht = new_wb.sheets("Доб.скважины с Кпрод")
        sht.range('A1').options(index=False).value = df_result

    print(str(os.path.basename(data_file)))
    # new_wb.save(str(os.path.basename(data_file)).replace(".xlsx", "") + "_out.xlsx")
    new_wb.save(str(os.path.join(os.path.dirname(__file__), data_file)).replace(".xlsx", "") + "_out.xlsx")
    # app1.kill()
    app1.quit()
    del app1

    label_5['text'] = '' # при повторном нажатии на "расчет" почему-то процент расчета накладывается на 100%
    win.update_idletasks()

    end_time = time.time()
    elapsed_time = end_time - start_time
    print('Elapsed time: ', elapsed_time)
    return elapsed_time
