// xll_spreadsheet.cpp
#include "../splitpath.h"
#include "xll_paste.h"

#define ADDIN_URL "./" // "https://xlladdins.com/xlladdins/"

using namespace xll;

// system default
const OPER xBlack(1);
const OPER xWhite(2);

// custom colors
const EditColor grey(0xF0, 0xF0, 0xF0);
static OPER xGrey;
const EditColor green(0x00, 0xA1, 0x81);
static OPER xGreen;
const EditColor blue(0x31, 0x8B, 0xCE);
static OPER xBlue;

/*
inline Style H1()
{
	FormatFont format;
	format.color = xWhite;
	format.size_num = 24;
	Alignment align;
	align.vert_align = (int)Alignment::Vertical::Center;
	align.horiz_align = (int)Alignment::Horizontal::Left;
	Style s("H1");
	s.FormatFont(format).Alignment(align).Pattern(xGreen, xWhite);
	return s;
}
*/
inline void Header()
{
	FormatFont::Color(xWhite);
	FormatFont::SizeNum(24);
	Alignment::Align(Alignment::Vertical::Center);
	Alignment::Align(Alignment::Horizontal::Left);
	Patterns(xGreen, xWhite);
}
inline void Subheader()
{
	FormatFont::SizeNum(13);
}
inline void ArgumentName()
{
	Alignment::Align(Alignment::Horizontal::Right);
	FormatFont::Bold();
}
inline void Category()
{
	FormatFont::SizeNum(16);
	//FormatFont::Bold();
}
inline void TableHeader()
{
	FormatFont::SizeNum(13);
	FormatFont::Color(xWhite);
	Patterns(xBlue, xWhite);
	//Alignment::Align(Alignment::Vertical::Center);
	if (Excel(xlfColumn) == 2) {
		Alignment::Align(Alignment::Horizontal::Right);
		FormatFont::Bold();
	}
	else if (Excel(xlfColumn) == 3) {
		Alignment::Align(Alignment::Horizontal::Left);
	}
}

#pragma warning(disable: 100)
// create sample spreadsheet
bool Spreadsheet(const char* description = "", bool release = false)
{
	using cat_type_text_map = std::map<OPER, std::map<OPER, std::map<OPER, Args*>>>;

	// custom colors
	xGrey = grey.Color();
	xGreen = green.Color();
	//xGreen = Excel(xlfGetWorkbook, OPER(12));
	xBlue = blue.Color();


	// get dir and filename extension
	path sp(Excel4(xlGetName).to_string().c_str());
	OPER dir(release ? ADDIN_URL : sp.dirname().c_str());
	OPER Name(Excel(xlfUpper, OPER(sp.fname)));

	// sort by category, macro type, then function text
	cat_type_text_map cat_type_text;
	for (auto& [key, args] : AddIn::Map) {
		if (args.Documentation().size() > 0) {
			if (args.Category() != OPER("XLL")) {
				cat_type_text[args.Category()][args.Type()][args.FunctionText()] = &args;
			}
		}
	}

	std::list<std::pair<OPER, Args*>> tabs;
	for (auto& [cat, type_text] : cat_type_text) {
		for (auto& [type, text] : type_text) {
			for (auto& [key, args] : text) {
				tabs.push_front(std::pair(key, args)); // insert happens to the left
			}
		}
	}

	// global settings
	//OptionsView global;
	//global.gridlines = false;
	//global.View();

	Workbook::Insert();
	Workbook::Rename(Name);

	// insert sheets with sample call
	for (auto& [tab, args] : tabs) {
		if (!args->isFunction()) {
			continue;
		}
		//!!!Workbook::Select();
		Workbook::Insert();
		Workbook::Rename(tab.safe());

		Select sel("R1:R1");
		Header();

		sel(REF(0, 1)); // B1
		sel.Set(tab & OPER(" ") & args->Type());

		sel.Move(1,0);
		sel.Set(args->FunctionHelp());
		Subheader();
		sel.Move(1,0);

		if (args->isFunction()) {
			sel.Move(1,0);
			sel.Set(args->FunctionText());
			Alignment::Align(Alignment::Horizontal::Right);

			sel.Move(0,1);
			sel.Formula(OPER("=") & args->FunctionText() & OPER("(") & args->ArgumentText() & OPER(")"));
			xll::Name name0(args->FunctionText());
			name0.Define(sel, true);
			OPER xFunctionText = sel;

			sel.Move(0,-1); sel.Move(1,0);
			for (unsigned i = 1; i <= args->ArgumentCount(); ++i) {
				sel.Set(args->ArgumentName(i));
				ArgumentName();

				sel.Move(0,1);
				OPER ref = paste_default(args->ArgumentDefault(i));
				xll::Name namei(args->ArgumentName(i));
				namei.Define(ref, true); // local name

				sel.Move(0, ref.columns());
				sel.Set(args->ArgumentHelp(i));
				sel.Move(0, -static_cast<int>(ref.columns()));

				sel.Move(rows(ref),-1);
			}
			Excel(xlcSelect, xFunctionText);
		}
	}

	// move main sheet to first position
	Workbook::Select(Name);

	Select sel("R1:R1");
	Header();
	/*
	sel(REF(0, 1)); // B1
	// "=HYPERLINK(\"https://xlladdins.com/addins/xll_math.xll\", \"xll_math\")"));
	sel.Formula(OPER("=HYPERLINK(\"") & dir & OPER(sp.fname) & OPER(" & Bits() & ") & OPER(sp.ext) & OPER("\", \"") & Name & OPER("\")"));
	Excel(xlcNote, OPER("user: addinuser pass: @addm3"));
	//sel.Set(Name);
	*/

	sel.Move(1,0);
	sel.Set(OPER(description));

	sel.Move(2, 0);

	for (auto& [cat, type_text] : cat_type_text) {
		if (type_text.size() == 0) {
			continue;
		}
		sel.Set(OPER("Category ") & cat);
		Category();

		sel.Move(1,0);
		for (auto& [type, text] : type_text) {
			// Function/Macro
			sel.Set(Excel(xlfProper, type));
			TableHeader();

			sel.Move(0,1);
			sel.Set(OPER("Description"));
			TableHeader();

			sel.Move(1,-1);
			int i = 0;
			for (auto& [key, args] : text) {
				if (!args->FunctionHelp()) {
					continue;
				}
				// table row
				OPER ft = args->FunctionText();
				if (args->isFunction()) {
					OPER safeft = ft.safe();
					OPER hl = OPER("=HYPERLINK(CELL(\"address\", ") & safeft & OPER("!") & ft & OPER(")");
					Excel(xlcFormula, hl & OPER(", \"") & ft & OPER("\")"));
				}
				else {
					sel.Set(ft);
				}
				Alignment::Align(Alignment::Horizontal::Right);
				FormatFont::Bold();

				sel.Move(0,1);
				Excel(xlcFormula, args->FunctionHelp());
				if (i % 2 == 1) {
					Patterns(xGrey, xWhite);
				}
				sel.Move(1,-1);
				++i;
			}
		}
		sel.Move(1, 0);
	}
	// select column C and fit text width
	// COLUMN.WIDTH(width_num, reference, standard, type_num, standard_num)
	//??? Column col(3); col.FitWidth();
	Excel(xlcColumnWidth, OPER(), OPER("C3:C3"), OPER(), OPER(3));

	// select cell containing the hyperlink
	sel(REF(1, 1));
	Excel(xlcWorkbookMove, Document::Sheet(), Document::Book(), OPER(1));

	return true;
}

static AddIn xai_spreadsheet_doc(
	Macro("xll_spreadsheet_doc", "DOC")
	.Category("XLL")
);
int WINAPI xll_spreadsheet_doc(void)
{
#pragma XLLEXPORT
	int result = FALSE;

	//Excel(xlcEcho, OPER(false));
	//DWORD olevel = XLL_ALERT_LEVEL(0);
	OptionsCalculation oc;
	oc.type_num = OPER((int)OptionsCalculation::Type::Manual);
	oc.Calculation();
	try {
		result = Spreadsheet(R"(All functions and macros having documentation.)", true);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("unknown exception in xll::Spreadsheet");
	}
	oc.type_num = OPER((int)OptionsCalculation::Type::Automatic);
	oc.Calculation();
	//XLL_ALERT_LEVEL(olevel);
	//Excel(xlcEcho, OPER(true));

	return result;
}