#include "cpp_addin.h"
#include <ppl.h>

using namespace xll;

AddIn xai_sumcpp(
    Function(XLL_DOUBLE, "xll_sumcpp", "SumCpp")
        .Arguments({Arg(XLL_FPX, "input", "")}));

double WINAPI xll_sumcpp(_FPX *input)
{
#pragma XLLEXPORT

    int n = input->rows * input->columns;
    auto arr = input->array;

    return concurrency::parallel_reduce(arr, arr + n, 0);
}
