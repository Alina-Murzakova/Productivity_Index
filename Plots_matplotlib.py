import matplotlib.pyplot as plt
import numpy as np
import datetime as dt

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
Kprod = 'Коэффициент продуктивности (ТР), м3/сут/атм'
x1 = "Координата X"
y1 = "Координата Y"
x2 = "Координата забоя Х"
y2 = "Координата забоя Y"
Kprod_TP = 'Кпрод ТР расчет, м3/сут/атм'
Kprod_new = 'Кпрод (Рпл = Рзаб ППД), м3/сут/атм'


# построение графиков Кпрод с ГТМ matplotlib
def plots(prod_well, data_prod, gtm, first_date, last_date, list_gtms_for_legend, list_colors_for_legend):

    fig, ax = plt.subplots()  # создаем Рисунок fig и 1 график на нем ax
    fig.set_size_inches(6, 4)
    ax.plot(data_prod[date], data_prod["Кпрод (Рпл = Рзаб ППД), м3/сут/атм"], color='forestgreen', linewidth=2)
    ax.plot(data_prod[date], data_prod['Кпрод ТР расчет, м3/сут/атм'], color='orange', linewidth=2)
    ax.set_xlabel("Дата", fontsize=7)
    ax.set_ylabel("Кпрод, м3/сут/атм", fontsize=7)
    ax.grid(True)
    # ax.legend(['Кпрод (Рпл = Рзаб ППД)', 'Кпрод ТР расчет']) # loc='upper right'
    lims = [np.datetime64(first_date), np.datetime64(last_date)]
    ax.set_xlim(lims)
    legend1 = plt.legend(['Кпрод (Рпл = Рзаб ППД)', 'Кпрод ТР расчет'], loc=2, fontsize=7)
    ax.add_artist(legend1)

    # добавление ГТМ на график
    if gtm is not None:
        ax_second = ax.twinx()  # дополнительная ось
        ax_second.set_ylim(-0.046, 1)  # задание пределов дополнительной оси
        ax_second.grid(False)
        for i in range(0, gtm.shape[0]):
            ax_second.plot(gtm["Дата"][i], 0, color='white', linewidth=1, marker='D', markersize=4,
                           markerfacecolor=gtm["Цвет"][i], markeredgecolor='grey', markeredgewidth=0.5)
        gtm.apply(lambda x: ax_second.annotate(x['Сокращение'], xy=(x['Дата'], -0.03), fontsize=5), axis=1)

        # list_gtms = gtm["Сокращение"].tolist()
        # # list_colours = gtm["Цвет"].drop_duplicates().tolist()
        # legend2 = plt.legend(list_gtms, loc=1)
        # ax_second.add_artist(legend2)
        # ax_second.set_xticks(list_gtm_data)
        # ax_second.set_xticklabels(list_gtm, minor=False, rotation='vertical', fontsize=3)

        # добавление третьей оси с целью отображения легенды ГТМ
        ax_third = ax.twinx()
        ax_third.set_ylim(-0.046, 1)
        ax_third.grid(False)
        date_for_legend = dt.datetime(1980, 1, 1)  # дата, которую не должно быть видно на графике
        list_dates_for_legend = [date_for_legend for i in range(len(list_gtms_for_legend))]
        # list_colors_for_legend = ['orange', 'red', 'green', 'blue', 'grey']
        # list_gtms_for_legend = ['123', '456', '789', '147', '258']
        for i in range(len(list_gtms_for_legend)):
            ax_third.plot(list_dates_for_legend[i], 0, color='white', linewidth=1, marker='D', markersize=4,
                          markerfacecolor=list_colors_for_legend[i], markeredgecolor='grey', markeredgewidth=0.5)
        legend3 = plt.legend(list_gtms_for_legend, loc=1, fontsize=7)
        ax_third.add_artist(legend3)

    fig.savefig('Kprod_well_' + str(prod_well))
    return fig