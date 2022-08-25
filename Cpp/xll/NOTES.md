# NOTES

Documentation should use urls.

Are function names case insensitive? Yes!

Add date and handle types with "B1" and "B2"? Excel seems to ignore the 1 and 2.

#define XLLEXPORT __pragma(...), constexpr check of __FUNCTION__ ??? or generate add-in ???

https://blog.conan.io/2021/02/10/Dependencies-Visual-Studio-C++-property-files.html

https://docs.microsoft.com/en-us/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf
Adds formula auto complete.

x86 or x64
for /F "tokens=3" %A in ('reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v Platform') do (echo %A)

Generate JS from AddIn?

Use function tests instead of lambdas.
Get names from nm | grep '^test'
Allow wildcards.

Use xlfRegister("file.xll") to require add-in. Versions?

Create zip file for visualizers and project template to replace msi.

Add signing certificate
https://github.com/actions/starter-workflows/blob/main/ci/dotnet-desktop.yml

https://docs.microsoft.com/en-us/office/troubleshoot/office-suite-issues/click-hyperlink-to-sso-website
Add a custom HTTP header to the GET response for the Office file contents. "Content-Disposition: Attachment"

https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-mso_winother-mso_2013_release/excel-keeps-crashing-with-a-acces-violation-error/68c00399-7055-4359-94ea-0b8e4091f8c0

https://www.drdobbs.com/architecture-and-design/faking-dde-with-private-servers/184409151?pgno=3

Exception thrown at ... in EXCEL.EXE: 0xC0000005:
https://docs.microsoft.com/en-us/office/client-developer/excel/how-to-call-xll-functions-from-the-function-wizard-or-replace-dialog-boxes?redirectedfrom=MSDN

```
#define CLASS_NAME_BUFFSIZE  50
#define WINDOW_TEXT_BUFFSIZE  50
// Data structure used as input to xldlg_enum_proc(), called by
// called_from_paste_fn_dlg(), called_from_replace_dlg(), and
// called_from_Excel_dlg(). These functions tell the caller whether
// the current worksheet function was called from one or either of
// these dialog boxes.
typedef struct
{
  bool is_dlg;
  short low_hwnd;
  char *window_title_text; // set to NULL if don't care
}
  xldlg_enum_struct;
// The callback function called by Windows for every top-level window.
BOOL CALLBACK xldlg_enum_proc(HWND hwnd, xldlg_enum_struct *p_enum)
{
// Check if the parent window is Excel.
// Note: Because of the change from MDI (Excel 2010)
// to SDI (Excel 2013), comment out this step in Excel 2013.
  if(LOWORD((DWORD)GetParent(hwnd)) != p_enum->low_hwnd)
    return TRUE; // keep iterating
  char class_name[CLASS_NAME_BUFFSIZE + 1];
//  Ensure that class_name is always null terminated for safety.
  class_name[CLASS_NAME_BUFFSIZE] = 0;
  GetClassName(hwnd, class_name, CLASS_NAME_BUFFSIZE);
//  Do a case-insensitve comparison for the Excel dialog window
//  class name with the Excel version number truncated.
  size_t len; // The length of the window's title text
  if(_strnicmp(class_name, "bosa_sdm_xl", 11) == 0)
  {
// Check if a searching for a specific title string
    if(p_enum->window_title_text) 
    {
// Get the window's title and see if it matches the given text.
      char buffer[WINDOW_TEXT_BUFFSIZE + 1];
      buffer[WINDOW_TEXT_BUFFSIZE] = 0;
      len = GetWindowText(hwnd, buffer, WINDOW_TEXT_BUFFSIZE);
      if(len == 0) // No title
      {
        if(p_enum->window_title_text[0] != 0)
          return TRUE; // No match, so keep iterating
      }
// Window has a title so do a case-insensitive comparison of the
// title and the search text, if provided.
      else if(p_enum->window_title_text[0] != 0
      && _stricmp(buffer, p_enum->window_title_text) != 0)
        return TRUE; // Keep iterating
    }
    p_enum->is_dlg = true;
    return FALSE; // Tells Windows to stop iterating.
  }
  return TRUE; // Tells Windows to continue iterating.
}

bool called_from_paste_fn_dlg(void)
{
    XLOPER xHwnd;
// Calls Excel4, which only returns the low part of the Excel
// main window handle. This is OK for the search however.
    if(Excel4(xlGetHwnd, &xHwnd, 0))
        return false; // Couldn't get it, so assume not
// Search for bosa_sdm_xl* dialog box with no title string.
    xldlg_enum_struct es = {FALSE, xHwnd.val.w, ""};
    EnumWindows((WNDENUMPROC)xldlg_enum_proc, (LPARAM)&es);
    return es.is_dlg;
}
```

## AddinManager

states: uninstalled, known, installed, loaded

AIM: HKCU\Software\Microsoft\Office\<version>\Excel\Add-in Manager
    - value name: full path  
OPENn: HKCU\Software\Microsoft\Office\<version>\Excel\Options\Open<n>
    - value name: OPENn  data: "/R full_path"  
DIR: %AppData%\Microsoft\Templates\ 
    - trusted location;  

AddinManager uses AIM or OPENn but never both

New() puts full path in AIM
Add() moves AIM\full_path to OPENn
Remove() moves `OPENn ` to AIM\full_path
Delete() remove and then delete AIM registry entry
Install() copy to DIR

## Auto<OpenAfter>

Status of add-in manager.

If in OPENn check for newer.

## First Use
AIM and OPENn not set
    Prompt to install xlGetName in Template
    Prompt for New to add to AIM
    Prompt to Add (does not work?)
  
## Troubleshooting
  
From Govert:  
```
Some comments and first steps:
•	When you try to load the wrong bitness add-in, Excel opens the binary file as text, and you get exactly the content you show.
•	When you try to load the wrong bitness add-in, the code in the add-in never executes.
•	It sounds like the problem machines are somehow configured to not load .xll add-ins, hence are falling back to the open-as-text behaviour.
•	In this case Excel-DNA is probably not involved either, and any .xll add-in will fail. You can try these pure native add-ins from the Excel2013 XLL SDK: https://www.dropbox.com/sh/f7v40h7xgf2uj19/AADXCjIlsgQgKI_vpYjRxfFha?dl=0
(If you're not comfortable running the binaries, I can help you download the SDK and rebuild it yourself - there is a bit of fiddling needed.)
•	Once you've established that no .xll add-ins work on these machines, at least Excel-DNA and .NET are out of the way too.
My next suggestions (in decreasing confidence) would be to look for other configuration options that might be clocking add-ins.
•	Temporarily disable the Trend Micro to be quite sure it is not blocking the .xll files being loaded.
•	Check that VBA is installed (just press F11 and confirm, that the VBA IDE is displayed). VBA is an optional component at install time, and is required to be present for any add-ins (including .xll add-ins) to be supported by Excel.
•	Check whether .xla add-ins can be loaded. Maybe there is some setting blocking all add-ins.
•	Try to figure out whether Group Policy settings have been applied. This is not always easy, as Office virtualizes the registry so it can be super confusing. But for example in the past these was a Group Policy options that blocks the COM object model from working (something about Automation being disabled).
•	You can also do a Process Monitor trace to see what registry entries are read by the process as it tries to load the add-in. This is a bit tedious but has pointed me in the right direction once or twice.
Finally my suggestions deteriorate into the usual:
•	Uninstall and reinstall Office.
```
