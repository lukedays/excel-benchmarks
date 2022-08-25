import xlwings as xw
import numpy as np


@xw.func
@xw.arg("x", np.array, ndim=2)
def SumXlwings(x):
    return np.sum(x)
