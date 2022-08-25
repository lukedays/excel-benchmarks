// defines.h - top level include
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once

// Parameterize by XLL_VERSION
// Define to be 12 for Excel 2007 and later or 4 otherwise
#ifndef XLL_VERSION
#define XLL_VERSION 12
#endif

#ifndef XLOPERX
#if XLL_VERSION == 12
#define XLOPERX XLOPER12
#else
#define XLOPERX XLOPER
#endif
#endif

#if XLL_VERSION == 12
	#undef _MBCS
	#ifndef _UNICODE
	#define _UNICODE
	#endif
	#ifndef UNICODE
	#define UNICODE
	#endif
#else
	#ifndef _MBCS
	#define _MBCS
	#endif
	#undef _UNICODE
	#undef UNICODE
#endif

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <tchar.h>
#include "XLCALL.H"

#if XLL_VERSION == 12
#define LPXLOPERX LPXLOPER12
#define TEXTX(s) L##s
#define TEXTR(s) RL"xyzyx(##s)xyzyx"
typedef struct _FP12 _FPX;
#else // 
#define LPXLOPERX LPXLOPER
#define TEXTX(s) u8##s
#define TEXTR(s) u8R"xyzyx(##s##)xyzyx"
typedef struct _FP _FPX;
#endif

#ifdef __cplusplus
//#include <type_traits>
#pragma warning(push)
#pragma warning(disable: 4724)
//??? where should this go
// mod with 0 <= x < y 
template<typename T>
//	requires std::is_integral_v<T>
inline T xmod(T x, T y)
{
	if (y == 0) 
		return 0;

	T z = x % y;

	return z >= 0 ? z : z + y;
}
#pragma warning(pop)
#endif

// xltypeX, XLOPERX::val.X, xX, XLL_X, desc
#define XLL_TYPE_SCALAR(X) \
    X(Num,     num,      num,  DOUBLE,  "IEEE 64-bit floating point")          \
    X(Bool,    xbool,    bool, BOOL,    "Boolean value")                       \
    X(Err,     err,      err,  WORD,    "Error type")                          \
    X(SRef,    sref.ref, ref,  LPOPER,  "Single refernce")                     \
    X(Int,     w,        int,  LONG,    "32-bit signed integer")               \
    X(Ref,     mref.lpmref,    lpmref, LPOPER,  "Multiple reference")          \

// types requiring allocation where xX is pointer to data
// xltypeX, XLOPERX::val.X, xX, XLL_X, desc
#define XLL_TYPE_ALLOC(X) \
    X(Str,     str,     str, PSTRING, "Pointer to a counted Pascal string")    \
    X(Multi,   array,   multi, LPOPER,  "Two dimensional array of OPER types") \
    X(BigData, bigdata.h.lpbData, bigdata, LPOPER,  "Blob of binary data")     \

// xllbitX, desc
#define XLL_BIT(X) \
	X(XLFree,  "Excel owns memory")    \
	X(DLLFree, "AutoFree owns memory") \

#define XLL_NULL_TYPE(X)                    \
	X(Missing, "missing function argument") \
	X(Nil,     "empty cell")                \

// xlerrX, Excel error string, error description
#define XLL_ERR(X)                                                          \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognised formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

// Defined in defines.c
// Arrays indexed by xlerrXXX
#ifdef __cplusplus 
extern "C" {
#endif
extern const LPCSTR xll_err_str[];  // Excel string
extern const LPCSTR xll_err_desc[]; // Human readable description
#ifdef __cplusplus
}
#endif

// Argument types for Excel Functions
// XLL_XXX, Excel4, Excel12, description
#define XLL_ARG_TYPE(X)                                                      \
X(BOOL,     "A", "A",  "short int used as logical")                          \
X(DOUBLE,   "B", "B",  "double")                                             \
X(CSTRING,  "C", "C%", "XCHAR* to C style NULL terminated unicode string")   \
X(PSTRING,  "D", "D%", "XCHAR* to Pascal style byte counted unicode string") \
X(DOUBLE_,  "E", "E",  "pointer to double")                                  \
X(CSTRING_, "F", "F%", "reference to a null terminated unicode string")      \
X(PSTRING_, "G", "G%", "reference to a byte counted unicode string")         \
X(USHORT,   "H", "H",  "unsigned 2 byte int")                                \
X(WORD,     "H", "H",  "unsigned 2 byte int")                                \
X(SHORT,    "I", "I",  "signed 2 byte int")                                  \
X(LONG,     "J", "J",  "signed 4 byte int")                                  \
X(FP,       "K", "K%", "pointer to struct FP")                               \
X(BOOL_,    "L", "L",  "reference to a boolean")                             \
X(SHORT_,   "M", "M",  "reference to signed 2 byte int")                     \
X(LONG_,    "N", "N",  "reference to signed 4 byte int")                     \
X(LPOPER,   "P", "Q",  "pointer to OPER struct (never a reference type)")    \
X(LPXLOPER, "R", "U",  "pointer to XLOPER struct")                           \
X(VOLATILE, "!", "!",  "called every time sheet is recalced")                \
X(UNCALCED, "#", "#",  "dereferencing uncalced cells returns old value")     \
X(VOID,     "",  ">",  "return type to use for asynchronous functions")      \
X(THREAD_SAFE,  "", "$", "declares function to be thread safe")              \
X(CLUSTER_SAFE, "", "&", "declares function to be cluster safe")             \
X(ASYNCHRONOUS, "", "X", "declares function to be asynchronous")             \

#ifdef __cplusplus 
extern "C" {
#endif
// Defined in defines.c
#define X(a,b,c,d) extern const LPCSTR XLL_##a;
	XLL_ARG_TYPE(X)
#undef X
#define X(a,b,c,d) extern const LPCSTR XLL_##a##4;
	XLL_ARG_TYPE(X)
#undef X
#define X(a,b,c,d) extern const LPCSTR XLL_##a##12;
	XLL_ARG_TYPE(X)
#undef X
#define X(a,b,c,d) extern const LPCSTR XLL_##a##X;
	XLL_ARG_TYPE(X)
#undef X
#ifdef __cplusplus
}
#endif

// 64-bit uses different symbol name decoration
#ifdef _M_X64 
#define XLL_DECORATE(s,n) s
#define XLL_X64(x) x
#define XLL_X32(x)
#else
#define XLL_DECORATE(s,n) "_" s "@" #n
#define XLL_X64(x)	
#define XLL_X32(x) x
#endif

// Function returning a constant value.
#define XLL_CONST(type, name, value, help, cat, topic) \
AddIn xai_ ## name (Function(XLL_##type, XLL_DECORATE("_xll_" #name, 0) , #name) \
.FunctionHelp(help).Category(cat).HelpTopic(topic)); \
extern "C" __declspec(dllexport) type WINAPI xll_ ## name () { return value; }

