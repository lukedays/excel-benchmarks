namespace CSharpAddin;

public static class Addin
{
    public static double SumCSharp(double[,] input)
    {
        double sum = 0;
        for (int i = 0; i < input.GetLength(0);++i)
        {
            for (int j = 0; j < input.GetLength(1); ++j)
            {
                sum += input[i, j];
            }
        }
        return sum;
    }
}