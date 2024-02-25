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
import win32com.client

# редактируемый график Кпрод

def charts(data_prod, xw, sht, index_first_not_null, index_last, count_arps, Kprod, Flag_smooth):
    chart = xw.Chart()
    # chart = sht.charts.add(left=sht.range('A1').expand().width + 10, top=50, width=1000, height=450)
    chart.top = 30  # положение сверху
    chart.left = 180  # положение слева
    chart.name = 'Kprod'
    chart.set_source_data(sht.range('T' + str(index_first_not_null + 2) + ':U' + str(index_last + 1)))  # источник данных для графика (оси Y)
    sht.charts[0].api[1].SeriesCollection(1).XValues = sht.range('A' + str(index_first_not_null + 2) + ':A' + str(index_last + 1)).api  # диапазон для оси X

    chart.api[1].SeriesCollection(1).Name = "Кпрод ТР расчет, м3/сут/атм"
    sht.charts[0].api[1].SeriesCollection(1).AxisGroup = win32com.client.constants.xlPrimary
    sht.charts[0].api[1].SeriesCollection(1).ChartType = win32com.client.constants.xlLine

    chart.api[1].SeriesCollection(2).Name = "Кпрод (Рпл = Рзаб ППД), м3/сут/атм"
    sht.charts[0].api[1].SeriesCollection(2).AxisGroup = win32com.client.constants.xlPrimary
    sht.charts[0].api[1].SeriesCollection(2).ChartType = win32com.client.constants.xlLine

    last_column = sht.range('A1').end('right').column
    last_row = sht.range('A1').end('down').row
    series_number = chart.api[1].SeriesCollection().Count  # для столбцов с ГТМ

    column_first_gtm = sht.range((1, 1), (last_row, last_column)).api.Find('ВПП').Column
    columns_gtm = last_column - column_first_gtm + 1

    for i in range(columns_gtm):
        colors_marker = [41, 53, 51, 45, 10, 12, 30, 9, 43, 13, 16]
        style_marker = [8, 1, 3, 8, 3, 2, 2, 8, 3, 2, 2]
        if not data_prod[sht.cells(1, (column_first_gtm + i)).value].isna().all():
            series_number += 1
            chart.api[1].SeriesCollection().NewSeries()
            sht.charts[0].api[1].SeriesCollection(series_number).AxisGroup = win32com.client.constants.xlSecondary  # создание второй оси Y
            sht.charts[0].api[1].SeriesCollection(series_number).ChartType = win32com.client.constants.xlLineMarkers
            # sht.charts[0].api[1].SeriesCollection(3).Values = sht.range('AA'+str(index_first_not_null+2), 'AA'+str(index_last+1)).api # добавление еще одной кривой
            sht.charts[0].api[1].SeriesCollection(series_number).Values = sht.range(((index_first_not_null + 2), (column_first_gtm + i)), ((index_last + 1), (column_first_gtm + i))).api  # добавление еще одной кривой
            # sht.charts[0].api[1].SeriesCollection(3).Values = sht.range(sht.range((2, 1), (91, 1)).get_address(True, False, external=True)).api # добавление еще одной кривой
            sht.charts[0].api[1].SeriesCollection(series_number).XValues = sht.range('A' + str(index_first_not_null + 2) + ':A' + str(index_last + 1)).api  # диапазон для оси X
            chart.api[1].SeriesCollection(series_number).Name = sht.cells(1, (column_first_gtm + i)).api
            sht.charts[0].api[1].SeriesCollection(series_number).MarkerSize = 6
            sht.charts[0].api[1].SeriesCollection(series_number).MarkerStyle = style_marker[i]
            # sht.charts[0].api[1].SeriesCollection(series_number).MarkerForegroundColorIndex = 3 # border / -1 - automatic border
            sht.charts[0].api[1].SeriesCollection(series_number).MarkerBackgroundColorIndex = colors_marker[i]  # fill / -1 - automatic fill
            sht.charts[0].api[1].SeriesCollection(series_number).Format.Line.Visible = False
            # sht.charts[0].api[1].SeriesCollection(series_number).Format.Line.ForeColor.RGB = rgb_to_int((255, 255, 255))
            # sht.charts[0].api[1].SeriesCollection(series_number).Format.Line.Transparency = 1
        else:
            continue

    # chart.chart_type = 'line' # тип диаграммы
    chart.height = 250  # высота диаграммы в points
    chart.width = 500  # ширина диаграммы в points
    chart.api[1].SetElement(2)  # расположение заголовка диаграммы сверху
    chart.api[1].ChartTitle.Text = "Коэффициент продуктивности"  # название диаграммы
    chart.api[1].ChartTitle.Font.Size = 14  # размер шрифта диаграммы
    chart.api[1].Axes(1).TickLabelSpacing = 1  # значение 1, гарантирует отображение каждого элемента нашей оси
    # sht.shapes.api('Kprod').Line.Visible = 0 # удаление контура графика
    # chart.api[1].Axes(2).TickLabels.NumberFormat = "£0,,' M'" # формат меток вертикальной оси
    # chart.api[1].Axes(2).MajorGridlines.Delete()
    sht.charts[0].api[1].SeriesCollection(1).Format.Line.ForeColor.RGB = rgb_to_int((163, 203, 232))
    sht.charts[0].api[1].SeriesCollection(2).Format.Line.ForeColor.RGB = rgb_to_int((0, 64, 119))
    chart.api[1].Axes(2, 1).HasTitle = True  # This line creates the Y axis label.
    chart.api[1].Axes(2, 1).AxisTitle.Text = "Коэффициент продуктивности, м3/сут/атм"
    chart.api[1].Axes(2, 1).AxisTitle.Font.Size = 8
    chart.api[1].Legend.Font.Size = 7  # размер шрифта легенды
    # chart.api[1].Legend.Top = 50
    # chart.api[1].Legend.Left = 100
    # chart.api[1].HasLegend = 0 # удаление легенды
    chart.api[1].Axes(2, 1).MinimumScale = 0
    chart.api[1].Axes(2, 1).TickLabels.Font.Size = 8
    chart.api[1].Axes(1, 1).TickLabels.Font.Size = 8
    chart.api[1].Axes(1, 1).TickLabels.NumberFormat = "mm-yyyy"
    chart.api[1].PlotArea.Border.LineStyle = win32com.client.constants.xlContinuous  # граница области построения

    if series_number > 2:  # если добавилась вторая ось Y для ГТМ
        chart.api[1].Axes(2, 2).TickLabels.Font.Size = 8
        chart.api[1].Axes(2, 2).MajorUnit = 0.2
        chart.api[1].Axes(2, 2).MinimumScale = 0
        chart.api[1].Axes(2, 2).MaximumScale = 1


    # редактируемый график Арпс
    chart_arps = xw.Chart()
    chart_arps.top = 30
    chart_arps.left = 750
    chart_arps.name = 'Арпс'

    # chart_arps.set_source_data(sht.range('V' + str(index_first_not_null + 2) + ':AA' + str(index_last)))  # источник данных для графика (оси Y)
    # chart_arps.set_source_data(sht.range('V' + str(index_first_not_null + 2) + ':Y' + str(index_last)))  # источник данных для графика (оси Y)
    column_first_arps = column_first_gtm - count_arps
    column_last_arps = column_first_gtm - 1
    chart_arps.set_source_data(sht.range(((index_first_not_null + 2), column_first_arps), ((index_last + 1), column_last_arps)))  # источник данных для графика (оси Y)
    series_number_Arps = chart_arps.api[1].SeriesCollection().Count
    for i in range(series_number_Arps):  # 6 - кол-во столбцов с Арпсом
        # colors_line = [11711154, 9415388, 16764006, 28864, 15564081, 153] # десятичные значения цветов
        # colors_line = [(47, 180, 233), (174, 189, 21), (72, 160, 68), (243, 139, 0)]
        colors_line = [(47, 180, 233), (174, 189, 21), (72, 160, 68)]
        sht.charts[1].api[1].SeriesCollection(i + 1).XValues = sht.range('N' + str(index_first_not_null + 2) + ':N' + str(index_last + 1)).api  # диапазон для оси X
        chart_arps.api[1].SeriesCollection(i + 1).Name = sht.cells(1, (column_first_arps + i)).api
        sht.charts[1].api[1].SeriesCollection(i + 1).Format.Line.ForeColor.RGB = rgb_to_int(colors_line[i])

    # Кривая Кпрод
    chart_arps.api[1].SeriesCollection().NewSeries()
    sht.charts[1].api[1].SeriesCollection(series_number_Arps + 1).Values = sht.range('V' + str(index_first_not_null + 2) + ':V' + str(index_last + 1)).api
    sht.charts[1].api[1].SeriesCollection(series_number_Arps + 1).XValues = sht.range('N' + str(index_first_not_null + 2) + ':N' + str(index_last + 1)).api
    chart_arps.api[1].SeriesCollection(series_number_Arps + 1).Name = Kprod
    sht.charts[1].api[1].SeriesCollection(series_number_Arps + 1).Format.Line.ForeColor.RGB = rgb_to_int((0, 64, 119))

    if Flag_smooth == 1:
        chart_arps.api[1].SeriesCollection().NewSeries()
        sht.charts[1].api[1].SeriesCollection(series_number_Arps + 2).Values = sht.range('W' + str(index_first_not_null + 2) + ':W' + str(index_last + 1)).api
        sht.charts[1].api[1].SeriesCollection(series_number_Arps + 2).XValues = sht.range('N' + str(index_first_not_null + 2) + ':N' + str(index_last + 1)).api
        chart_arps.api[1].SeriesCollection(series_number_Arps + 2).Name = "decline_smoothed"
        sht.charts[1].api[1].SeriesCollection(series_number_Arps + 2).Format.Line.ForeColor.RGB = rgb_to_int((205, 34, 44))

    chart_arps.chart_type = 'xy_scatter_lines_no_markers'  # тип диаграммы
    chart_arps.height = 250  # высота диаграммы в points
    chart_arps.width = 500  # ширина диаграммы в points
    chart_arps.api[1].SetElement(2)  # расположение заголовка диаграммы сверху
    chart_arps.api[1].ChartTitle.Text = "Темп падения Кпрод (Арпс)"
    chart_arps.api[1].ChartTitle.Font.Size = 14  # размер шрифта заголовка
    # chart.api[1].HasLegend = 0 # удаление легенды
    chart_arps.api[1].Axes(1).TickLabelSpacing = 1  # значение 1, гарантирует отображение каждого элемента нашей оси
    chart_arps.api[1].Axes(2, 1).HasTitle = True  # This line creates the Y axis label.
    chart_arps.api[1].Axes(2, 1).AxisTitle.Text = "Темп падения Кпрод"
    chart_arps.api[1].Axes(2, 1).AxisTitle.Font.Size = 8
    chart_arps.api[1].Axes(2, 1).TickLabels.Font.Size = 8

    chart_arps.api[1].Axes(1, 1).HasTitle = True  # This line creates the X axis label.
    chart_arps.api[1].Axes(1, 1).AxisTitle.Text = "Накопленное время работы, дни"
    chart_arps.api[1].Axes(1, 1).AxisTitle.Font.Size = 8
    chart_arps.api[1].Axes(1, 1).TickLabels.Font.Size = 8
    chart_arps.api[1].Legend.Font.Size = 7  # размер шрифта легенды
    chart_arps.api[1].PlotArea.Border.LineStyle = win32com.client.constants.xlContinuous

    # chart_arps.api[1].Axes(1, 1).CategoryType = 2 # 2 - xlValue, 1 - xlCategory, 3 - xlSeriesAxis

    chart_arps.api[1].Axes(1, 1).MinimumScale = data_prod.iloc[index_first_not_null]['Накопленное время работы, дни'] // 10 * 10 if data_prod.iloc[index_first_not_null]['Накопленное время работы, дни'] > 10 else 0
    # chart_arps.api[1].Axes(1, 1).MaximumScale = 10
    # chart_arps.api[1].Axes(1).MinorUnit = 50
    # chart_arps.api[1].Axes(1).MajorUnit = 50

