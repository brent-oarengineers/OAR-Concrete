import math as m

import numpy
import scipy.integrate as integrate
import numpy as np
import scipy.optimize as optimize

class CircularPile:
    def __init__(self, D, nbars, cover, lig, fc, fy, Abar):
        self.D = D
        self.nbars = nbars
        self.cover = cover
        self.lig = lig
        self.fc = fc
        self.fy = fy
        self.R = D / 2
        self.bar_angle = 2 * m.pi / self.nbars
        self.PCD = 2 * (self.R - cover - lig)
        self.yieldStrain = fy / 200000.
        self.Abar = Abar

    def barDepthArray(self):
        X = np.zeros(self.nbars)
        for k in range(self.nbars):
            X[k] = self.PCD / 2 * m.cos(k * self.bar_angle)
        return X

class AS3600StressCurve:
    def __init__(self, fc, Ec, strain_max):
        self.fc = fc
        self.Ec = Ec
        self.strain_max = strain_max
        self.sigma_cp = 0.9 * fc
        self.strain_cp = (4.11 * self.sigma_cp ** 0.75) / self.Ec
        self.Ecp = self.sigma_cp / self.strain_cp
        self.n = self.Ec / (self.Ec - self.Ecp)

    def n_ratio(self, strain):
        return strain / self.strain_cp

    def k(self, strain):
        if strain <= self.strain_cp:
            return 1
        else:
            return max(0.67 + self.sigma_cp / 62, 1.0)

    def sigma_c(self, strain):
        return self.sigma_cp * (
                self.n * self.n_ratio(strain) / (self.n - 1 + self.n_ratio(strain) ** (self.n * self.k(strain))))

    def strain(self, x, dNA):
        return self.strain_max - self.strain_max * (x / dNA)

    def stress(self, x, dNA):
        return self.sigma_c(self.strain(x, dNA))


class Calculations:
    def __init__(self, pile: CircularPile, stressCurve: AS3600StressCurve):
        self.pile = pile
        self.stressCurve = stressCurve
        self.R = self.pile.R

    def chordLength(self, R, x):
        return 2 * m.sqrt(2 * R * x - (x ** 2))

    def Hc(self, dNA):
        return integrate.quad(lambda x: self.chordLength(self.R, x) * self.stressCurve.stress(x, dNA), 0, dNA)[0]

    def yc(self, dNA):
        return integrate.quad(lambda x: self.chordLength(self.R, x) * self.stressCurve.stress(x, dNA) * (self.R - x), 0, dNA)[0] / self.Hc(dNA)


    def barStrain(self, x, dNA):
        return (dNA - self.R + x) * self.stressCurve.strain_max / dNA

    def barStrainArray(self, dNA):
        return self.barStrain(self.pile.barDepthArray(), dNA)

    def barStress(self, strain):
        if strain <= -1 * self.pile.yieldStrain:
            return -1 * self.pile.yieldStrain * 200000
        elif strain >= self.pile.yieldStrain:
            return self.pile.yieldStrain * 200000
        else:
            return strain * 200000

    def barStressArray(self, dNA):
        barStress_vectorized = np.vectorize(self.barStress)
        return barStress_vectorized(self.barStrainArray(dNA))

    def grossBarForceArray(self, dNA):
        return self.barStressArray(dNA) * self.pile.Abar

    def barForceDeductions(self, dNA):
        X = np.zeros(self.pile.nbars)
        normalisedDepthArray = self.pile.R -self.pile.barDepthArray()
        for k in range(self.pile.nbars):
            if self.barStrainArray(dNA)[k] > 0:
                X[k] = self.stressCurve.stress(normalisedDepthArray[k], dNA)
            else:
                X[k] = 0
        return X * self.pile.Abar

    def netBarForceArray(self, dNA):
        return self.grossBarForceArray(dNA) - self.barForceDeductions(dNA)

    def Hsteel(self, dNA):
        return np.sum(self.netBarForceArray(dNA))

    def dNA(self):
        return optimize.newton(lambda x: self.Hc(x) + self.Hsteel(x), self.R / 2)

    def Ms(self):
        return np.sum(self.netBarForceArray(self.dNA()) * self.pile.barDepthArray())

    def Mc(self):
        return self.Hc(self.dNA()) * self.yc(self.dNA())

    def Mu(self):
        return self.Mc() + self.Ms()

    def kuo(self):
        bar_depth_array = self.pile.barDepthArray()
        do = numpy.max(bar_depth_array) + self.pile.R
        return self.dNA() / do

    def phi(self):
        value = 1.24 - 13 * self.kuo() / 12
        if value >= 0.85:
            return 0.85
        elif value <= 0.65:
            return 0.65
        else:
            return value



pile1 = CircularPile(450, 4, 100, 10, 40, 500, 310)
stresscurve1 = AS3600StressCurve(40, 30100, 0.003)
Calc = Calculations(pile1, stresscurve1)

print(Calc.Mu() * 0.75 / 1e6)
print(Calc.Mu() / 1e6)
print(Calc.phi())
print(Calc.kuo())
print(pile1.barDepthArray() + pile1.R)
print(Calc.phi() * Calc.Mu() / 1e6)
