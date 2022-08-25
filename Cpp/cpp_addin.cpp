#include "cpp_addin.h"

using namespace xll;

AddIn xai_sumcpp(
    Function(XLL_DOUBLE, "xll_sumcpp", "SumCpp")
    .Arguments({
        Arg(XLL_FPX, "input", "")
        })
);
double WINAPI xll_sumcpp(_FPX* input)
{
#pragma XLLEXPORT

    int n = input->rows * input->columns;

	double sum{ 0 };
	for (int i = 0; i < n; ++i)
            sum += input->array[i];
	
	return sum;
}
