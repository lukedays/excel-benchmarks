import xloil as xlo
import numpy as np


@xlo.func
def SumXloilNumpy(x: xlo.Array(float)) -> float:  # type: ignore
    return np.sum(x)


@xlo.func
def SumXloil(x) -> float:
    sum = 0
    for i in x:
        for j in i:
            sum += j
    return sum
