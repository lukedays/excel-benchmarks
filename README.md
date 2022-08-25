# Excel Benchmarks

Gathering benchmarks for Excel add-ins. Inspired by [The Benchmarks Game](https://benchmarksgame-team.pages.debian.net/benchmarksgame/index.html) and some LinkedIn posts.

## Featuring

- Python via [xlOil](https://xloil.readthedocs.io/en/stable/Introduction.html) - thanks for the help [@cunnane](https://github.com/cunnane)!
- Python via [xlwings](https://www.xlwings.org/)
- C# via [Excel-DNA](https://excel-dna.net/) - thanks for the help [@govert](https://github.com/govert)!
- C++ via [xlladdins](https://github.com/xlladdins/xll) - credits to [@keithalewis](https://github.com/keithalewis)!
- VBA
- Native Excel Formulas

Unfortunately I couldn't make a [JavaScript addin](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) to work. Suggestions welcome!

## Benchmarks

### Sum a 10.000 x 100 matrix (1 million cells)

Time in miliseconds:
| Name   | Excel SUM(...) | C++<br>(xlladdins) | C#<br>(Excel-DNA) | Python + Numba<br>(xlOil) | Python + Numpy<br>(xlOil) | Python Only<br>(xlOil) | VBA   | Python + Numpy<br> (xlwings) |
| ------ | -------------- | ------------------ | ----------------- | ------------------------- | ------------------------- | ---------------------- | ----- | ---------------------------- |
| Median | 3,2            | 10,7               | 13,8              | 40,9                      | 42,3                      | 142,7                  | 242,5 | 928,3                        |

<img src="images/2022-08-25-15-46-44.png" alt="test1" width="600"/>

## How it works

I used the VBA code [provided by Microsoft](https://docs.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance) for performance measurement, in order to calculate the runtime of each cell with formulas.
I changed it to run a 100 times in order to compute the statistics.

## How to build (Windows only)

- Download Visual Studio Community 2022 (free for personal use)
- Install with .NET and C++ support
- Install Python 3.10 ([pyenv](https://github.com/pyenv-win/pyenv-win) is a nice way to control versions)
- Run the following commands in a terminal:
```
pip install xloil xlwings numba icc-rt cffi --upgrade
xlwings addin install
xloil install
```
- Clone the repo. On the instructions below, `ROOT` will be the repo folder path.
- Python - xlOil (has to be done before opening the workbook):
  - Open `%APPDATA%\xlOil\xlOil.ini`.
  - On the parameter `PYTHONPATH`, add `;ROOT\Python` to the end
  - Change parameter: `LoadModules=["xloil_addin"]`
- Open `ROOT\Benchmarks.xlsm`
- C#:
  - Build CSharpAddin project with Release configuration.
  - Drag this file to Excel: `ROOT\CSharp\bin\Release\net6.0-windows\CSharpAddin64.xll`
- C++:
  - Build CppAddin project with Release configuration.
  - Drag this file to Excel: `ROOT\Cpp\x64\Release\CppAddin.xll`
- Python - xlwings:
  - On the xlwings tab inside Excel, add `ROOT\Python` to `PYTHONPATH`
  - On `UDF Modules parameter`, write `xlwings_addin`
  - Click `Import Functions`
- Make sure the formulas for each framework work (yellow cells).
- Click the button Run to compute the times
  


