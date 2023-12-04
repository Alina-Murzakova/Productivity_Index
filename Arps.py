import pandas as pd
import time, os
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt
import datetime as dt
# import seaborn as sns
from scipy.optimize import curve_fit
# from scipy.optimize import minimize, Bounds, NonlinearConstraint
from sklearn.metrics import r2_score
# from Arps_minimize import Arps_minimize



class FunctionFluidProduction:
    heroes = []
    """Класс кривой Кпрод"""

    def exponential(self, t, Kprod_i, Di):
        """
        Exponential decline curve equation
        Arguments:
            t: Float. Time since the well first came online, can be in various units
            (days, months, etc) so long as they are consistent.
            Kprod_i: Float. Initial Kprod.
            Di: Float. Nominal decline rate (constant)
        Output:
            Returns Kprod, or the expected Kprod at time t. Float.
        """
        # K, D, l, p = coeff
        return Kprod_i*np.exp(-Di*t)


    def hyperbolic(self, t, Kprod_i, b, Di):
        """
        Hyperbolic decline curve equation
        Arguments:
            t: Float. Time since the well first came online, can be in various units
            (days, months, etc) so long as they are consistent.
            Kprod_i: Float. Initial Kprod.
            b: Float. Hyperbolic decline constant
            Di: Float. Nominal decline rate at time t=0
        Output:
            Returns Kprod, or the expected Kprod at time t. Float.
        """
        return Kprod_i/((1.0+b*Di*t)**(1.0/b))


    def harmonic(self, t, Kprod_i, Di):
        """
        Harmonic decline curve equation
        Output:
            Returns Kprod, or the expected Kprod at time t. Float.
        """
        return Kprod_i/(1.0+Di*t)


    def hyperbolic_power(self, t, Kprod_i, a, b, Di, tau):
        return Kprod_i * np.exp(-a * np.log(1.0 + tau)) * np.exp(-1 / b * np.log(1.0 + b * Di * (t - tau)))


    def double(self, t, Kprod_i, b1, b2, Di, tau):
        return Kprod_i * np.exp(-1/b1*np.log(1.0+b1*Di*tau)) * np.exp(-1/b2*np.log(1.0+b2*Di*(t-tau)/(1.0+b1*Di*tau)))


def calc_arps(df, Kprod):
    df_r2 = pd.DataFrame(columns=['№ скважины', 'Тип Арпса', "R2"])
    list_index = df.index
    well_number = df['№ скважины'].iloc[0]
    Kprod_max = get_max_Kprod(df, Kprod, 'Дата')
    Kprod_init = df[Kprod].iloc[0]
    FP = FunctionFluidProduction()

    # types_arps = {FP.exponential: ([Kprod_max, 0], 0, [Kprod_max, 1]),
    #               FP.hyperbolic: ([Kprod_max, 1, 1], 0, [Kprod_max, 2, 1]),
    #               FP.harmonic: ([Kprod_max, 0], 0, [Kprod_max, 1]),
    #               FP.hyperbolic_power: ([Kprod_max, 1, 1, 1, 10], 0, [Kprod_max, 10, 10, 10, 500]),
    #               FP.double: ([Kprod_max, 1, 1, 1, 10], 0, [Kprod_max, 10, 10, 10, 500])}

    types_arps = {FP.hyperbolic: ([Kprod_max, 1, 1], 0, [Kprod_max, 2, 1]),
                  # FP.hyperbolic_power: ([Kprod_max, 1, 1, 1, 10], 0, [Kprod_max, 10, 10, 10, 500]),
                  FP.hyperbolic_power: ([Kprod_max, 0.0001, 0.0001, 0.0001, 10], 0, [Kprod_max, 10, 1, 10, 500]),
                  FP.double: ([Kprod_max, 0.0001, 0.0001, 0.0001, 10], 0, [Kprod_max, 10, 1, 10, 500])}

    for arps in types_arps:
        print(arps)
        # typ = [FP.exponential, FP.hyperbolic, FP.harmonic, FP.power, FP.double]

        popt_exp, pcov_exp = curve_fit(arps, df['Накопленное время работы, дни'], df[Kprod],
                                       p0=types_arps[arps][0], bounds=(types_arps[arps][1], types_arps[arps][2]))
        print(popt_exp)
        K_prod_arps = arps(df['Накопленное время работы, дни'], *popt_exp)
        decline_rate_Kprod = K_prod_arps / popt_exp[0]

        # df[(arps.__name__)] = K_prod_arps
        df[(arps.__name__)] = decline_rate_Kprod

        r_squared = r2_score(df[Kprod], K_prod_arps)
        dict_R2 = {'№ скважины': well_number, 'Тип Арпса': (arps.__name__), 'R2': r_squared}
        df_r2 = df_r2._append(dict_R2, ignore_index=True)
    # decline_rate_Kprod_minimize, dict_r2 = Arps_minimize(df, Kprod)
    # # df['minimize'] = Kprod_minimize
    # df['minimize'] = decline_rate_Kprod_minimize
    # df_r2 = df_r2._append(dict_r2, ignore_index=True)

    return df, df_r2


def get_max_Kprod(df, variable_column, date_column):
    """
    This function allows you to select the highest Kprod month as max Kprod
    Arguments:
        - df: Pandas dataframe.
        - variable_column: String. Column name for the column where we're attempting to get
        the max Kprod from (can be either 'Gas' or 'Oil' in this script)
        - date_column: String. Column name for the date that the data was taken at
    """
    #First, sort the data frame from earliest to most recent prod date
    df=df.sort_values(by=date_column)

    #Return the max value in the selected variable column
    return df[variable_column].max()





# data_file = "Крайнее_out.xlsx" # Файл для расчета
#
# # Названия столбцов в Excel
# # Названия столбцов в Excel
# date = 'Дата'
# well_number = '№ скважины'
# field = 'Месторождение'
# work_marker = 'Характер работы'
# well_status = 'Состояние'
# prod = "НЕФ"
# inj = "НАГ"
# Qo_rate = 'Дебит нефти за последний месяц, т/сут'
# Ql_rate = 'Дебит жидкости за посл.месяц, м3/сут'
# water_cut = 'Обводненность за посл.месяц, % (вес)'
# time_prod = 'Время работы в добыче, часы'
# Winj_rate = 'Приемистость за последний месяц, м3/сут'
# time_inj = "Время работы под закачкой, часы"
# P_well = "Забойное давление (ТР), атм"
# P_pressure = 'Пластовое давление (ТР), атм'
# K_prod = 'Коэффициентт продуктивности (ТР), м3/сут/атм'
# x1 = "Координата X"
# y1 = "Координата Y"
# x2 = "Координата забоя Х"
# y2 = "Координата забоя Y"
# Kprod_new = "Кпрод (Рпл = Рзаб ППД), м3/сут/атм"
#
#
# df_initial = pd.read_excel(r'C:\Users\Alina\Desktop\Python\Kprod\Крайнее_out.xlsx') # Открытие экселя
# df_initial = df_initial.fillna(0)  # Заполнение пустых ячеек нулями
#
# last_date = df_initial.sort_values(by=[date])
# last_date = df_initial[date].iloc[-1]
#
# wells_prod = df_initial[(df_initial[date] == last_date) & (df_initial[work_marker] == prod)][well_number].unique()
#
# df = df_initial[(df_initial[well_number] == 6197)]
#
# df = df[df[Kprod_new] != 0]
# array_Kprod = df['Кпрод (Рпл = Рзаб ППД), м3/сут/атм'].to_numpy()
#
# time_work = np.array(df[time_prod])
# time_work_days = np.array(df[time_prod])/24
# df['Время работы'] = time_work_days
# df['Накопленное время работы'] = time_work_days.cumsum()
# array_cumulative_days = df['Накопленное время работы'].to_numpy()
#
#
# class FunctionFluidProduction:
#     """Класс кривой добычи жидкости"""
#
#     def __init__(self, day_fluid_production, array_timeProduction):
#         self.day_fluid_production = day_fluid_production
#         self.array_timeProduction = array_timeProduction
#         self.first_m = -1
#         self.start_q = -1
#         self.ind_max = -1
#
#     def Adaptation(self, correlation_coeff):
#         """
#         :param correlation_coeff: коэффициенты корреляции функции
#         :return: сумма квадратов отклонений фактических точек от модели
#         """
#         k1, k2 = correlation_coeff
#         max_day_prod = np.amax(self.day_fluid_production)
#         index = list(np.where(self.day_fluid_production == np.amax(self.day_fluid_production)))[0][0]
#
#
#         indexes = np.arange(start=index, stop=self.day_fluid_production.size, step=1) - index
#         day_fluid_production_month = max_day_prod * (1 + k1 * k2 * indexes) ** (-1 / k2)
#         deviation = [(self.day_fluid_production[index:] - day_fluid_production_month) ** 2]
#         self.first_m = self.day_fluid_production.size - index + 1
#         self.start_q = max_day_prod
#         self.ind_max = index
#
#         return np.sum(deviation)
#
#     def Conditions_FP(self, correlation_coeff):
#         """Привязка (binding) к последней точке"""
#         k1, k2 = correlation_coeff
#         base_correction = self.day_fluid_production[-1]
#
#         max_day_prod = np.amax(self.day_fluid_production)
#         index = list(np.where(self.day_fluid_production == np.amax(self.day_fluid_production)))[0][0]
#
#         last_prod = max_day_prod * (1 + k1 * k2 * (self.day_fluid_production.size - 1 - index)) ** (-1 / k2)
#         binding = base_correction - last_prod
#         return binding
#
#
# def Calc_FFP(array_Kprod, array_timeProduction):
#     """
#     Функция для аппроксимации кривой добычи
#     :param array_production: массив добычи нефти
#     :param array_timeProduction: массив времени работы скважины
#     :return: output - массив с коэффициентами аппроксимации
#     [k1, k2, first_m, start_q, index, cumulative_oil]
#      0  1       2        3       4       5
#     """
# #     cumulative_oil = np.sum(array_production) / 1000
# #     array_rates = array_production / (array_timeProduction / 24)
# #     array_rates[array_rates == -np.inf] = 0
# #     array_rates[array_rates == np.inf] = 0
#
#     """ Условие, если в расчете только одна точка или последняя точка максимальная """
#     if (array_Kprod.size == 1) or (np.amax(array_Kprod) == array_Kprod[-1]):
#         index = list(np.where(array_Kprod == np.amax(array_Kprod)))[0][0]
#         first_m = array_Kprod.size - index + 1
#         start_q = array_Kprod[-1]
#         k1 = 0
#         k2 = 1
#         output = [k1, k2, first_m, start_q, index]
#     else:
#         # Ограничения:
#         k1_left = 0.0001
#         k2_left = 0.0001
#         k1_right = 1.1
#         k2_right = 50
#
#         k1 = 0.0001
#         k2 = 0.0001
#         c_cet = [k1, k2]
#         FP = FunctionFluidProduction(array_Kprod, array_timeProduction)
#         bnds = Bounds([k1_left, k2_left], [k1_right, k2_right])
#         non_linear_con = NonlinearConstraint(FP.Conditions_FP, [-0.00001], [0.00001])
#         try:
#             a = FP.Adaptation
#             res = minimize(FP.Adaptation, c_cet, method='trust-constr', bounds=bnds, constraints=non_linear_con,
#                            options={'disp': False, 'xtol': 1E-7, 'gtol': 1E-7, 'maxiter': 2000})
#             k1, k2 = res.x[0], res.x[1]
#             if k1 < 0:
#                 k1 = 0
#             if k2 < 0:
#                 k2 = 0
#             output = [k1, k2, FP.first_m, FP.start_q, FP.ind_max]
#         except:
#             index = list(np.where(array_Kprod == np.amax(array_Kprod)))[0][0]
#             first_m = array_Kprod.size - index + 1
#             start_q = array_Kprod[-1]
#             output = ["Невозможно", "Невозможно", first_m, start_q, index]
#     return output
#
# results_approximation = Calc_FFP(array_Kprod, array_cumulative_days)