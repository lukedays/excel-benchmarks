// handle.cpp - Embed C++ objects in Excel using xll::handle<T>
// https://xlladdins.com
//#define XLOPERX XLOPER
#include <concepts>
#include "../xll/xll.h"

using namespace xll;

// get and set a type
template<std::semiregular T>
class base {
	T x;
public:
	base() { }
	base(const T& x) 
		: x(x) 
	{ }
	base(const base&) = default;
	base& operator=(const base&) = default;
	virtual ~base()
	{ }

	T& get()
	{ 
		return x; 
	}
	void set(const T& _x)
	{
		x = _x;
	}
};

// embed base in Excel
AddIn xai_base(
	Function(XLL_HANDLEX, "xll_base", "\\XLL.BASE")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
		})
	.Uncalced() // required for functions creating handles
	.FunctionHelp("Return a handle to a base object.")
);
HANDLEX WINAPI xll_base(LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base<OPER>> h(new base<OPER>(*px));

	return h.get();
}

// call base member function
AddIn xai_base_get(
	Function(XLL_LPOPER, "xll_base_get", "XLL.BASE.GET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.BASE")
		})
	.FunctionHelp("Return the value stored in base.")
);
LPOPER WINAPI xll_base_get(HANDLEX _h)
{
#pragma XLLEXPORT
	static OPER o;
	xll::handle<base<OPER>> h(_h);

	o = h ? h->get() : ErrNA;

	return &o;
}

AddIn xai_base_set(
	Function(XLL_HANDLEX, "xll_base_set", "XLL.BASE.SET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.BASE."),
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
		})
	.FunctionHelp("Set the value of base to cell.")
);
HANDLEX WINAPI xll_base_set(HANDLEX _h, LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base<OPER>> h(_h);

	if (h) {
		h->set(*px);
	}

	return _h;
}

// single inheritance
template<std::semiregular T>
class derived : public base<T> {
	T x2;
public:
	derived()
	{ }
	// put x in base and x2 in derived
	derived(const T& x, const T& x2)
		: base<T>(x), x2(x2)
	{ }
	derived(const derived&) = default;
	derived& operator=(const derived&) = default;
	~derived() = default;

	T& get2()
	{
		return x2;
	}
};

AddIn xai_derived(
	Function(XLL_HANDLEX, "xll_derived", "\\XLL.DERIVED")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells."),
		Arg(XLL_LPOPER, "cell2", "is a cell or range of cells.")
	})
	.Uncalced() // required for functions creating handles
	.FunctionHelp("Return a handle to a derived object.")
);
HANDLEX WINAPI xll_derived(LPOPER px, LPOPER px2)
{
#pragma XLLEXPORT
	// derived isa base
	xll::handle<base<OPER>> h(new derived<OPER>(*px, *px2));

	return h.get();
}

// XLL.BASE.GET calls base::get for handles returned by XLL.DERIVED.

AddIn xai_derived_get(
	Function(XLL_LPOPER, "xll_derived_get", "XLL.DERIVED.GET2")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.DERIVED.")
	})
	.FunctionHelp("Return the second value stored in derived.")
);
LPOPER WINAPI xll_derived_get(HANDLEX _h)
{
#pragma XLLEXPORT
	static OPER o;
	// get handle to base class
	xll::handle<base<OPER>> h(_h);

	if (!h) {
		o = ErrNA;
	}
	else {
		// downcast to derived
		auto pd = h.as<derived<OPER>>();
		o = pd ? pd->get2() : ErrValue;
	}

	return &o;
}

// convert pointers to and from "\\base[0x<hexdigits>]"
static handle<base<OPER>>::codec<XLOPERX> base_codec; // ("\\base[0x", "]");

AddIn xai_ebase(
	Function(XLL_LPOPER, "xll_ebase", "\\XLL.EBASE")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
	})
	.Uncalced()
	.FunctionHelp("Return a handle to a base object.")
);
LPOPER WINAPI xll_ebase(LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base<OPER>> h(new base<OPER>(*px));

	auto& x = base_codec.encode(h.get());

	return const_cast<LPOPER>(&x);
}

AddIn xai_ebase_get(
	Function(XLL_LPOPER, "xll_ebase_get", "XLL.EBASE.GET")
	.Arguments({
		Arg(XLL_LPOPER, "handle", "is a handle returned by \\XLL.EBASE")
	})
	.FunctionHelp("Return the value stored in base.")
);
LPOPER WINAPI xll_ebase_get(LPOPER ph)
{
#pragma XLLEXPORT
	xll::handle<base<OPER>> h(base_codec.decode(*ph));

	return xll_base_get(h.get());
}

#ifdef _DEBUG

// Documentation for the Excel() function.
// https://xlladdins.github.io/Excel4Macros/

AddIn xai_handle_test(Macro("xll_handle_test", "HANDLE.TEST"));
int WINAPI xll_handle_test()
{
#pragma XLLEXPORT
	try {

		// enter data in cell RiCj
		// https://xlladdins.github.io/Excel4Macros/formula.html
		auto Formula = [](auto i, auto j, auto data) {
			return Excel(xlcFormula, OPER(data), OPER(REF(i,j)));
		};

		// get contents of cell RiCj;
		// https://xlladdins.github.io/Excel4Macros/get.cell.html
		auto Contents = [](auto i, auto j) {
			// 5 - contents
			return Excel(xlfGetCell, OPER(5), OPER(REF(i, j)));
		};

		// https://xlladdins.github.io/Excel4Macros/new.html
		// 1 - worksheet
		Excel(xlcNew, OPER(1), Missing, Missing); // worksheet

		// test base
		Formula(0, 0, 1.23);
		ensure(Contents(0, 0) == 1.23);
		Formula(1, 0, "=\\XLL.BASE(R[-1]C[0])");
		Formula(2, 0, "=XLL.BASE.GET(R[-1]C[0])");
		
		ensure(Contents(2, 0) == 1.23);
		
		Formula(0, 0, "foo");
		ensure(Contents(2, 0) == "foo");

		// test derived
		Formula(4, 0, 4.56);
		Formula(5, 0, "bar");
		Formula(6, 0, "=\\XLL.DERIVED(R[-2]C[0], R[-1]C[0])");
		Formula(7, 0, "=XLL.BASE.GET(R[-1]C[0])");
		Formula(8, 0, "=XLL.DERIVED.GET2(R[-2]C[0])");

		ensure(Contents(7, 0) == 4.56); // derived isa base
		ensure(Contents(8, 0) == "bar");

		// use pretty handles
		Formula(10, 0, "=\\XLL.EBASE(R1C1)");
		auto base = Contents(10, 0); // pretty name
		//ensure(Excel(xlfLeft, base, OPER(8)) == "\\base[0x"); // check prefix

		Formula(11, 0, "=XLL.EBASE.GET(R[-1]C[0])");
		auto R1C1 = Contents(0, 0);
		//ensure(Contents(11, 0) == R1C1);
		//ensure(Contents(2, 0) = R1C1); // XLL.BASE.GET also called
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_handle_test(xll_handle_test);

#endif // _DEBUG
