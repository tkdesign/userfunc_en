What is Userfunc_en VBA?
=====================

Userfunc_ru VBA is add-in for Microsoft Office Excel. Contains subroutines and functions for automating routine tasks of editing and checking tables, invoices and other business documentation.

How to install?
---------------
1. Download file userfunc.xlam from dist folder to your PC.
2. Open path %UserProfile%\AppData\Roaming\Microsoft\AddIns in Explorer.
3. Copy userfunc.xlam file to this folder.
4. In Excel, create a new or open an existing workbook.
5. Go to File->Options.
6. The Excel Options window opens. Select Add-ins.
7. Select Excel Add-ins and click Go.
8. The Add-ins dialog appears. In Add-Ins dialog, click Browse....
9. Select file userfunc.xlam and click Open.
10. Click OK after back to Add-ins dialog.

This add-in contains descriptions sheet for Excel IntelliSense. For enable tooltips appearing when entering function name in cell, similar to built-in Excel functions, also you need to connect a third-party add-on from Excel-DNA IntelliSense (https://github.com/Excel-DNA/IntelliSense). For example, ExcelDna.IntelliSense64.xll file.

For support Excel IntelliSense, also copy file ExcelDna.IntelliSense64.xll from repository dist folder to folder %UserProfile%\AppData\Roaming\Microsoft\AddIns, or download latest version from Excel-DNA IntelliSense project repository. Install this third-party add-in same way as shown above.

How to use?
-----------
After installing this add-in, new tab will appear in Excel ribbon - Macros. Select one or more cells and click any button on Macros tab. For some subroutines, it is possible to undo the action (undo) or repeat it (repeat) as with built-in Excel subroutines.

System requirements
-------------------
* Windows 10-11 x64
* Microsoft Office Excel 2013-365 x64
* Support english language and US regional settings
* Installed english keyboard