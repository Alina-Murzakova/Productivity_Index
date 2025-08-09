import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
from sklearn.metrics import r2_score


class FunctionFluidProduction:
    heroes = []
    """Класс кривой Кпрод"""

    def exponential(self, t, Di):
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
        return np.exp(-Di*t)

    def hyperbolic(self, t, b, Di):
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
        return 1/((1.0+b*Di*t)**(1.0/b))

    def harmonic(self, t, Di):
        """
        Harmonic decline curve equation
        Output:
            Returns Kprod, or the expected Kprod at time t. Float.
        """
        return 1/(1.0+Di*t)

    def hyperbolic_power(self, t, a, b, Di, tau):
        return np.exp(-a * np.log(1.0 + tau)) * np.exp(-1 / b * np.log(1.0 + b * Di * (t - tau)))
        # return (1.0 / ((1.0 + tau)**a)) * (1 / ((1 + b * Di * (t - tau))**(1.0 / b)))

    def double(self, t, b1, b2, Di, tau):
        return np.exp(-1/b1*np.log(1.0+b1*Di*tau)) * np.exp(-1/b2*np.log(1.0+b2*Di*(t-tau)/(1.0+b1*Di*tau)))
        # return (1.0 / ((1.0 + b1 * Di * tau)**(1.0 / b1))) * (1.0 / (((1.0 + b2 * Di * (t - tau)) / (1.0 + b1 * Di * tau))**(1.0 / b2)))


def calc_arps(df, Kprod, Flag_smooth):
    df_with_Kprod = df[df[Kprod] > 0]  # только строки, где Кпрод больше 0
    df_r2 = pd.DataFrame(columns=['№ скважины', 'Тип Арпса', "R2"])
    list_index = df_with_Kprod.index
    well_number = df_with_Kprod['№ скважины'].iloc[0]
    Kprod_max = get_max_Kprod(df_with_Kprod, Kprod, 'Дата')
    Kprod_init = df_with_Kprod[Kprod].iloc[0]
    decline_rate_Kprod = df_with_Kprod[Kprod] / Kprod_init

    if Flag_smooth == 1:
        decline_rate_Kprod = decline_rate_Kprod.rolling(window=5, center=True).mean()
        df['smoothed'] = decline_rate_Kprod
        decline_rate_Kprod = decline_rate_Kprod[decline_rate_Kprod > 0]

        df_with_Kprod = df_with_Kprod.merge(decline_rate_Kprod.rename('Temp_smoothed'), left_index=True, right_index=True)
        # df = df.merge(df_with_Kprod[['Дата', 'Temp_smoothed']], how='left', on='Дата')
        # df = df.merge(decline_rate_Kprod.rename('smoothed'), how='left', left_index=True, right_index=True)

    sigma = np.ones(df_with_Kprod.shape[0])
    # sigma.fill(0.5)
    sigma[0] = 0.1 # = sigma[-3:]
    FP = FunctionFluidProduction()

    # types_arps = {FP.exponential: ([Kprod_max, 0], 0, [Kprod_max, 1]),
    #               FP.hyperbolic: ([Kprod_max, 1, 1], 0, [Kprod_max, 2, 1]),
    #               FP.harmonic: ([Kprod_max, 0], 0, [Kprod_max, 1]),
    #               FP.hyperbolic_power: ([Kprod_max, 1, 1, 1, 10], 0, [Kprod_max, 10, 10, 10, 500]),
    #               FP.double: ([Kprod_max, 1, 1, 1, 10], 0, [Kprod_max, 10, 10, 10, 500])}

    # types_arps = {FP.hyperbolic: ([Kprod_max, 1, 1], 0, [Kprod_max, 2, 1]),
    #               # FP.hyperbolic_power: ([Kprod_max, 1, 1, 1, 10], 0, [Kprod_max, 10, 10, 10, 500]),
    #               FP.hyperbolic_power: ([Kprod_max, 0.0001, 0.0001, 0.0001, 10], 0, [Kprod_max, 10, 1, 10, 500]),
    #               FP.double: ([Kprod_max, 0.0001, 0.0001, 0.0001, 10], 0, [Kprod_max, 10, 1, 10, 500])}

    types_arps = {FP.hyperbolic: ([1, 1], 0, [2, 1]),
                  FP.hyperbolic_power: ([0.0001, 0.0001, 0.0001, 10], 0, [10, 1, 10, 500]),
                  FP.double: ([0.0001, 0.0001, 0.0001, 10], 0, [10, 1, 10, 500])}

    for arps in types_arps:
        # print(arps)
        # typ = [FP.exponential, FP.hyperbolic, FP.harmonic, FP.power, FP.double]

        popt_exp, pcov_exp = curve_fit(arps, df_with_Kprod['Накопленное время работы, дни'], decline_rate_Kprod,
                                       p0=types_arps[arps][0],
                                       bounds=(types_arps[arps][1], types_arps[arps][2]),
                                       sigma = sigma)
        decline_rate_Kprod_Arps = arps(df_with_Kprod['Накопленное время работы, дни'], *popt_exp)
        decline_rate_Kprod_Arps_with_null = arps(df['Накопленное время работы, дни'], *popt_exp)
        # decline_rate_Kprod = K_prod_arps_with_null / popt_exp[0]

        df[(arps.__name__)] = decline_rate_Kprod_Arps_with_null

        r_squared = r2_score(decline_rate_Kprod, decline_rate_Kprod_Arps)
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
    df = df.sort_values(by=date_column)

    #Return the max value in the selected variable column
    return df[variable_column].max()