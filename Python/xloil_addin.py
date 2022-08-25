import xloil as xlo
import numpy as np
from numba import njit


@xlo.func
def SumXloilNumpy(x: xlo.Array(float)) -> float:  # type: ignore
    return np.sum(x)


@xlo.func
@njit(parallel=True, nogil=True, cache=True, fastmath=True)
def SumXloilNumba(x: xlo.Array(float)) -> float:  # type: ignore
    sum = 0
    for i in x:
        for j in i:
            sum += j
    return sum


@xlo.func
def SumXloil(x) -> float:
    sum = 0
    for i in x:
        for j in i:
            sum += j
    return sum
