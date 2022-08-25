# PUT-YOUR-README-HERE

This project gets you up and running with the [xll library](https://github.com/xlladdins/xll.git)
for creating high performance Excel add-ins written in C++, or C, or any language
that can be called from C.

[Use this template](https://github.com/xlladdins/xll_template/generate) and open in Visual Studio 2019.  
Double click `xll_template.sln`.   
Press `F5` to build the add-in and start Excel with it loaded in the debugger.  

To build for 32-bit Excel set `Platform` to `Win32` in the Configuration Manager.

The solution and project are named `xll_template`. You may want to rename these,
and `xll_template.h`, `xll_template.cpp` while you're at it'.

The project creates a submodule for the xll library. To update to the latest xll version 
start a command prompt using `Tools\Command Line\Developer Command Prompt` then  
```
>cd xll
>git pull origin master
```