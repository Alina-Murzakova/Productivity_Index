import numpy as np
import pandas as pd
from scipy.optimize import minimize, Bounds, NonlinearConstraint
from sklearn.metrics import r2_score

class FunctionFluidProduction:
    """Класс кривой Кпрод"""

    def __init__(self, Kprod):
        self.Kprod = Kprod
        self.start_Kprod = -1
        self.ind_max = -1

    def Adaptation(self, correlation_coeff):
        """
        :param correlation_coeff: коэффициенты корреляции функции
        :return: сумма квадратов отклонений фактических точек от модели
        """
        k1, k2 = correlation_coeff
        index = 0
        first_Kprod = self.Kprod[0]
        # first_Kprod = np.amax(self.Kprod)
        # index = list(np.where(self.Kprod == np.amax(self.Kprod)))[0][0]

        indexes = np.arange(start=index, stop=self.Kprod.size, step=1) - index
        Kprod_approximation  = first_Kprod * (1 + k1 * k2 * indexes) ** (-1 / k2)
        deviation = [(self.Kprod[index:] - Kprod_approximation) ** 2]
        self.start_Kprod = first_Kprod
        self.ind_max = index
        return np.sum(deviation)

    def Conditions_FP(self, correlation_coeff):
        """Привязка (binding) к последней точке"""
        k1, k2 = correlation_coeff
        base_correction = self.Kprod[-1]

        first_Kprod = self.Kprod[0]
        index = 0
        # max_day_prod = np.amax(self.Kprod)
        # index = list(np.where(self.Kprod == np.amax(self.Kprod)))[0][0]

        last_prod = first_Kprod * (1 + k1 * k2 * (self.Kprod.size - 1 - index)) ** (-1 / k2)
        binding = base_correction - last_prod
        return binding


def Calc_FFP(array_Kprod):
    """
    Функция для аппроксимации кривой добычи
    :param array_production: массив добычи нефти
    :param array_timeProduction: массив времени работы скважины
    :return: output - массив с коэффициентами аппроксимации
    [k1, k2, start_Kprod, index]
     0   1        2         3
    """

    """ Условие, если в расчете только одна точка или последняя точка максимальная """

    if (array_Kprod.size == 1) or (np.amax(array_Kprod) == array_Kprod[-1]):
        index = list(np.where(array_Kprod == np.amax(array_Kprod)))[0][0]
        start_Kprod = array_Kprod[-1]
        k1 = 0
        k2 = 1
        output = [k1, k2, start_Kprod, index]
    else:
        # Ограничения:
        k1_left = 0.0001
        k2_left = 0.0001
        k1_right = 1.1
        k2_right = 50

        k1 = 0.0001
        k2 = 0.0001
        c_cet = [k1, k2]
        FP = FunctionFluidProduction(array_Kprod)
        bnds = Bounds([k1_left, k2_left], [k1_right, k2_right])
        non_linear_con = NonlinearConstraint(FP.Conditions_FP, [-0.00001], [0.00001])
        try:
            res = minimize(FP.Adaptation, c_cet, method='trust-constr', bounds=bnds, constraints=non_linear_con,
                           options={'disp': False, 'xtol': 1E-7, 'gtol': 1E-7, 'maxiter': 2000})
            k1, k2 = res.x[0], res.x[1]
            if k1 < 0:
                k1 = 0
            if k2 < 0:
                k2 = 0
            output = [k1, k2, FP.start_Kprod, FP.ind_max]
        except:
            index = list(np.where(array_Kprod == np.amax(array_Kprod)))[0][0]
            start_Kprod = array_Kprod[-1]
            output = ["Невозможно", "Невозможно", start_Kprod, index]
    return output


def Arps_minimize(df, Kprod):

    # max_Kprod = np.amax(self.Kprod)
    # index = list(np.where(self.Kprod == np.amax(self.Kprod)))[0][0]

    df_r2 = pd.DataFrame(columns=['№ скважины', 'Тип Арпса', "R2"])
    list_index = df.index
    well_number = df['№ скважины'].iloc[0]

    array_Kprod = np.array(df[Kprod])

    k1, k2, start_Kprod, num_m = Calc_FFP(array_Kprod)[:4]

    indexes = np.arange(start=0, stop=array_Kprod.size, step=1)
    Kprod_minimize = start_Kprod * (1 + k1 * k2 * indexes) ** (-1 / k2)
    decline_rate_Kprod = Kprod_minimize / start_Kprod

    r_squared = r2_score(df[Kprod], Kprod_minimize)
    dict_r2 = {'№ скважины': well_number, 'Тип Арпса': 'minimize', 'R2': r_squared}

    return decline_rate_Kprod, dict_r2
