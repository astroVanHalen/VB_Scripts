Excel VBA Scripts Collection
Welcome to my collection of VBA (Visual Basic for Applications) scripts for Microsoft Excel! These scripts, also called macros, are small programs that automate tasks in Excel, like formatting data, creating charts, or performing calculations. This guide explains how to add these scripts to your Excel so you can use them in any workbook, and how to run them easily using keyboard shortcuts (keybinds) or Excel's macro menu.
What You Need

Microsoft Excel (any recent version, like Excel 2016, 2019, 2021, or 365).
A computer with Excel installed (Windows or Mac).
No programming experience is required—just follow the steps below!

Step 1: Download the VBA Scripts

Visit the Repository: You’re already here on the GitHub page!
Download the Files:
Click the green Code button at the top of this page.
Select Download ZIP to get all the scripts.
Unzip the downloaded file to a folder on your computer (e.g., your Desktop or Documents).


Locate the Scripts: The folder will contain .bas or .txt files with the VBA code. These are the scripts you’ll add to Excel.

Step 2: Add Scripts to Your Personal Macro Workbook
The Personal Macro Workbook (PERSONAL.XLSB) is a special file in Excel that stores macros so they’re available in all your workbooks. Here’s how to add the scripts to it:

Open Excel:
Start Excel on your computer.


Enable the Developer Tab (if not already visible):
Go to File > Options > Customize Ribbon.
Check the box for Developer in the right-hand list, then click OK.
You should now see a Developer tab on the Excel ribbon.


Open the VBA Editor:
Click the Developer tab, then click Visual Basic (or press Alt + F11).
This opens the VBA Editor window.


Find or Create the Personal Macro Workbook:
In the VBA Editor, look at the Project Explorer (left side). If you don’t see it, go to View > Project Explorer.
Look for a project called PERSONAL.XLSB. If it’s not there, you need to create it:
Go back to Excel, click Developer > Record Macro.
In the dialog box, set "Store macro in" to Personal Macro Workbook.
Do any small action (e.g., type “test” in a cell), then click Stop Recording (in the Developer tab).
Return to the VBA Editor (Alt + F11), and you should now see PERSONAL.XLSB in the Project Explorer.




Import the Scripts:
In the VBA Editor, right-click on PERSONAL.XLSB in the Project Explorer.
Choose Insert > Module to create a new module (e.g., Module1).
To import a script:
Right-click the new module, select File > Import File, and browse to one of the .bas or .txt files you downloaded.
Repeat for each script you want to add.


Alternatively, you can open a script file in a text editor (like Notepad), copy the code, and paste it into a new module in the VBA Editor.


Save the Personal Macro Workbook:
In the VBA Editor, press Ctrl + S or go to File > Save PERSONAL.XLSB.
Close the VBA Editor.
In Excel, save and close any open workbooks. Excel will ask if you want to save changes to PERSONAL.XLSB—click Yes.



Now the scripts are stored in your Personal Macro Workbook and available in all Excel workbooks on your computer!
Step 3: Run the Macros
You can run the macros in two ways: using the macro dialog or by setting up keyboard shortcuts (keybinds).
Option 1: Use the Macro Dialog

Open any Excel workbook.
Go to the Developer tab and click Macros (or press Alt + F8).
In the dialog box, you’ll see a list of macros (e.g., PERSONAL.XLSB!MacroName).
Select the macro you want to run and click Run.
The macro will perform its task on the active workbook or sheet.

Option 2: Set Up Keyboard Shortcuts (Keybinds)
Keybinds let you run a macro by pressing a combination of keys (e.g., Ctrl + Shift + M). Here’s how to set them up:

Go to the Developer tab and click Macros (or press Alt + F8).
Select a macro from the list (e.g., PERSONAL.XLSB!MacroName).
Click Options.
In the Shortcut key field, type a letter (e.g., M).
You can choose Ctrl + [Letter] or Ctrl + Shift + [Letter].
Example: Typing M creates Ctrl + M; typing Shift + M creates Ctrl + Shift + M.
Avoid using common Excel shortcuts like Ctrl + C or Ctrl + V.


Click OK, then close the dialog.
Repeat for other macros, using different letters for each.
To run a macro, press its keybind (e.g., Ctrl + Shift + M) in any Excel workbook.

Tips

Enable Macros: When opening a workbook, Excel may ask you to enable macros. Click Enable Content to use your macros.
Check Macro Descriptions: Each script file in this repository includes a comment at the top explaining what it does. Open the .bas or .txt file in a text editor to read it.
Backup Your Work: Before running a macro on an important workbook, save a copy of

