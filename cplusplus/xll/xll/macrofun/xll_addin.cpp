// xll_addin.cpp - information related to addins
//#define XLL_VERSION 4
#include "../xll.h"

using namespace xll;

#define ARGS_HELP_URL "XLL.ARGS.html"

typedef const TCHAR* CSTRING;

XLL_CONST(CSTRING, XLL_ARGS_PROCEDURE, 
	_T("Procedure"),"Return \"Procedure\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_TYPE_TEXT, 
	_T("TypeText"), "Return \"TypeText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_FUNCTION_TEXT, 
	_T("FunctionText"), "Return \"FunctionText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_ARGUMENT_TEXT, 
	_T("ArgumentText"), "Return \"ArgumentText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_MACRO_TYPE, 
	_T("MacroType"), "Return \"MacroType\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_CATEGORY, 
	_T("Category"), "Return \"Category\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_SHORTCUT_TEXT, 
	_T("ShortcutText"), "Return \"ShortcutText\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_HELP_TOPIC, 
	_T("HelpTopic"), "Return \"HelpTopic\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGS_FUNCTION_HELP, 
	_T("FunctionHelp"), "Return \"FunctionHelp\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGUMENTS_COUNT,
	_T("ArgumentCount"), "Return \"ArgumentCount\"", "XLL", ARGS_HELP_URL);
// individual argument help
XLL_CONST(CSTRING, XLL_ARGUMENTS_NAME, 
	_T("ArgumentName"), "Return \"ArgumentName\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGUMENTS_HELP, 
	_T("ArgumentHelp"), "Return \"ArgumentHelp\"", "XLL", ARGS_HELP_URL);
XLL_CONST(CSTRING, XLL_ARGUMENTS_DEFAULT, 
	_T("ArgumentDefault"), "Return \"ArgumentDefault\"", "XLL", ARGS_HELP_URL);

AddIn xai_addins(
	Function(XLL_LPOPER, "xll_addins", "XLL.ADDINS")
	.FunctionHelp("Return array of all add-ins.")
	.Category("XLL")
	.Documentation(R"(
Return a one column array of all registered add-ins. These can
be used as the first argument to <code>XLL.ADDIN.ARGS</code> or <code>XLL.ADDIN.ARGUMENTS</code>.
)")
);
LPOPER WINAPI xll_addins()
{
#pragma XLLEXPORT
	static OPER result;

	result = ErrNA;
	try {
		result = Nil;
		for (const auto& [key, args] : AddIn::Map) {
			result.push_back(key);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("XLL.ADDINS: unknown exception");
	}

	return &result;
}

AddIn xai_addin_args(
	Function(XLL_LPOPER, "xll_addin_args", "XLL.ADDIN.ARGS")
	.Arguments({
		{XLL_LPOPER, "name", "is a function name or register id.", "\"XLL.ADDIN.ARGS\""},
		{XLL_LPOPER, "keys", "is an array of keys from XLL_ARGS_*.", ""},
	})
	.FunctionHelp("Return information about an add-in.")
	.Category("XLL")
	.Documentation(R"(
Return members of <code>xll::Args</code> for an add-in given its <code>name</code> or register id.
The <code>keys</code> are an array of strings from the <code>XLL_ARGS_*</code> constants.
If called with <code>name</code> missing return the default list of keys.
If called with <code>keys</code> missing return all known keys.
)")
);
LPOPER WINAPI xll_addin_args(LPOPER pname, LPOPER pkeys)
{
#pragma XLLEXPORT
	// known keys
	static OPER keys({
		OPER("ModuleText"),
		OPER("Procedure"),
		OPER("TypeText"),
		OPER("ArgumentText"),
		OPER("MacroType"),
		OPER("Category"),
		OPER("ShortcutText"),
		OPER("HelpTopic"),
		OPER("FunctionHelp"),
		OPER("ArgumentCount"),
	});
	static OPER result;

	result = ErrNA;
	try {
		if (pname->xltype == xltypeMissing) {
			return &keys;
		}

		Args* pargs = nullptr;
		if (pname->is_str()) {
			pargs = AddIn::Arguments(*pname);
		}
		else if (pname->is_num()) {
			pargs = AddIn::Arguments(pname->as_num());
		}
		else {
			XLL_ERROR(__FUNCTION__ ": name must be a string or number");
		}
		ensure(pargs || !__FUNCTION__ ": failed to find find function");

		if (pkeys->xltype == xltypeMissing or !*pkeys) {
			pkeys = &keys;
		}
		result = OPER(rows(*pkeys), columns(*pkeys));
		for (unsigned i = 0; i < result.size(); ++i) {
			const OPER& key = index(*pkeys, i);
			ensure(Excel(xlfMatch, key, keys, OPER(0))
				|| !__FUNCTION__ ": failed to find exact case insensitive match");
			if (key == "ArgumentCount") {
				result[i] = pargs->ArgumentCount();
			}
			else {
				result[i] = (*pargs)[key];
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("XLL.ADDIN.ARGS: unknown exception");
	}

	return &result;
}

AddIn xai_addin_arguments(
	Function(XLL_LPOPER, "xll_addin_arguments", "XLL.ADDIN.ARGUMENTS")
	.Arguments({
		{XLL_LPOPER, "name", "is a function name or register id.", "\"XLL.ADDIN.ARGS\""},
		{XLL_WORD, "index", "is the 1-based index of the individual function argument", "1"},
		{XLL_LPOPER, "keys", "is an array of keys from XLL_ARGUMENTS_*.", "\"ArgumentHelp\""},
		})
		.FunctionHelp("Return information about individual add-in arguments.")
	.Category("XLL")
	.Documentation(R"(
Return information about individual arguments for an add-in given 
its <code>name</code> or register id.
The <code>keys</code> are an array of strings from the <code>XLL_ARGUMENTS_*</code> constants.
If called with no arguments return the known list of keys. 
If the third argument is missing it will return all known keys of the first argument.
If <code>index</code> is <code>0</code> return the string for the default function.
)")
);
LPOPER WINAPI xll_addin_arguments(LPOPER pname, WORD i, LPOPER pkeys)
{
#pragma XLLEXPORT
	// known keys
	static OPER keys({
		OPER("ArgumentName"),
		OPER("ArgumentHelp"),
		OPER("ArgumentDefault"),
		});
	static OPER result;

	result = ErrNA;
	try {
		if (pname->xltype == xltypeMissing) {
			return &keys;
		}

		Args* pargs = nullptr;
		if (pname->is_str()) {
			pargs = AddIn::Arguments(*pname);
		}
		else if (pname->is_num()) {
			pargs = AddIn::Arguments(pname->as_num());
		}
		else {
			XLL_ERROR(__FUNCTION__ ": name must be a string or number");
		}
		ensure(pargs || !__FUNCTION__ ": failed to find find function");

		if (i == 0) {
			result = pargs->ArgumentDefault(0);

			return &result;
		}
		--i; // 1-based

		if (pkeys->xltype == xltypeMissing) {
			pkeys = &keys;
		}
		result = OPER(rows(*pkeys), columns(*pkeys));
		for (unsigned j = 0; j < result.size(); ++j) {
			const OPER& key = index(*pkeys, j);
			ensure(Excel(xlfMatch, key, keys, OPER(0))
				|| !__FUNCTION__ ": failed to find exact case insensitive match");
			result[j] = (*pargs)[key][i];
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return &result;
}

