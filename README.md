# Excel VBA Scripts Collection

Welcome to my collection of VBA (Visual Basic for Applications) scripts for Microsoft Excel! These scripts, also called macros, are small programs that automate tasks in Excel, like formatting data, creating charts, or performing calculations. This guide explains how to copy these scripts from GitHub and paste them into your Excel so you can use them in any workbook. You’ll also learn how to run them using a menu or keyboard shortcuts (keybinds). No programming experience is needed!

## What You Need
- Microsoft Excel (any recent version, like Excel 2016, 2019, 2021, or 365).
- A computer with Excel installed (Windows or Mac).
- A web browser to access GitHub.

## Step 1: Copy the VBA Scripts from GitHub
1. **Visit the Repository**:
   - Go to [https://github.com/astroVanHalen/VB_Scripts](https://github.com/astroVanHalen/VB_Scripts).
2. **Find a Script**:
   - Click on a script file (e.g., `CollectFromSheets.txt`, `GroupImageShape.txt`, `Resizer.txt`, or `SetZoom.txt`) to view its code.
3. **Copy the Code**:
   - Click the **Copy** button (clipboard icon next to 'Raw') above the code to copy the entire script to your clipboard.
   - Alternatively, click **Raw** to see the code in plain text, select all (**Ctrl + A** or **Cmd + A**), and copy (**Ctrl + C** or **Cmd + C**).
   - Repeat for each script you want to use.

## Step 2: Paste Scripts into Your Personal Macro Workbook
The Personal Macro Workbook (`PERSONAL.XLSB`) is a special file in Excel that stores macros so they’re available in all your workbooks. Here’s how to paste the scripts into it:

1. **Open Excel**:
   - Start Excel on your computer.
2. **Enable the Developer Tab** (if not visible):
   - Go to **File** > **Options** > **Customize Ribbon**.
   - Check the box for **Developer** in the right-hand list, then click **OK**.
   - You should see a **Developer** tab on the Excel ribbon.
3. **Open the VBA Editor**:
   - Click the **Developer** tab, then click **Visual Basic** (or press **Alt + F11**).
   - This opens the VBA Editor window.
4. **Find or Create the Personal Macro Workbook**:
   - In the VBA Editor, look at the **Project Explorer** (left side). If you don’t see it, go to **View** > **Project Explorer**.
   - Look for `PERSONAL.XLSB`. If it’s not there, create it:
     - Go back to Excel, click **Developer** > **Record Macro**.
     - Set "Store macro in" to **Personal Macro Workbook**.
     - Do a small action (e.g., type “test” in a cell), then click **Stop Recording** (in the Developer tab).
     - Return to the VBA Editor (**Alt + F11**), and you should see `PERSONAL.XLSB`.
5. **Paste the Script**:
   - In the VBA Editor, right-click `PERSONAL.XLSB` in the Project Explorer.
   - Choose **Insert** > **Module** to create a new module (e.g., `Module1`).
   - Double-click the new module to open its code window (right side).
   - Paste the copied code (**Ctrl + V** or **Cmd + V**) into the code window.
   - Repeat for each script, creating a new module for each if you prefer (or paste multiple scripts into one module).
6. **Save the Personal Macro Workbook**:
   - In the VBA Editor, press **Ctrl + S** or go to **File** > **Save PERSONAL.XLSB**.
   - Close the VBA Editor.
   - In Excel, save and close any open workbooks. If prompted to save `PERSONAL.XLSB`, click **Yes**.

Your scripts are now stored in `PERSONAL.XLSB` and available in all Excel workbooks on your computer!

## Step 3: Run the Macros
You can run the macros in two ways: using the macro menu or by setting keyboard shortcuts.

### Option 1: Use the Macro Menu
1. Open any Excel workbook.
2. Go to the **Developer** tab and click **Macros** (or press **Alt + F8**).
3. In the dialog box, you’ll see a list of macros (e.g., `PERSONAL.XLSB!MacroName`).
4. Select a macro and click **Run**.
5. The macro will perform its task on the active workbook or sheet.

### Option 2: Set Up Keyboard Shortcuts (Keybinds)
Keybinds let you run a macro by pressing keys (e.g., Ctrl + Shift + M). Here’s how:
1. Go to the **Developer** tab and click **Macros** (or press **Alt + F8**).
2. Select a macro (e.g., `PERSONAL.XLSB!MacroName`).
3. Click **Options**.
4. In the **Shortcut key** field, type a letter (e.g., `M`).
   - Choose **Ctrl + [Letter]** or **Ctrl + Shift + [Letter]** (e.g., `Shift + M` for **Ctrl + Shift + M**).
   - Avoid common Excel shortcuts like **Ctrl + C** or **Ctrl + V**.
5. Click **OK**, then close the dialog.
6. Repeat for other macros, using different letters.
7. Run a macro by pressing its keybind (e.g., **Ctrl + Shift + M**).

## Tips
- **Enable Macros**: When opening a workbook, Excel may ask to enable macros. Click **Enable Content** to use your macros.
- **Check Macro Descriptions**: Each script file on GitHub has a comment at the top explaining what it does. Read it in the GitHub browser before copying.
- **Backup Your Work**: Before running a macro on an important workbook, save a copy of your file.
- **Troubleshooting**:
   - If a macro doesn’t appear in the menu, ensure it’s in `PERSONAL.XLSB` and you saved it.
   - If a macro doesn’t work, check its description on GitHub for specific instructions (e.g., it may need a certain sheet or data format).
- **Sharing**: To share macros, share the GitHub repository link (`https://github.com/astroVanHalen/VB_Scripts`) so others can copy the code.

## Example Macro
- **Macro Name**: `ActivateFirstSheet`
- **Description**: Sets the first worksheet in your workbook as the active sheet.
- **How to Use**: Run it via the macro menu or assign a keybind like **Ctrl + Shift + F**.

## Contributing
Have ideas for new macros or improvements? Feel free to:
- Open an **Issue** on this GitHub repository to suggest changes.
- Submit a **Pull Request** with new scripts or updates.

## Contact
For help or questions, open an issue on this GitHub repository, and I’ll assist!

Happy automating with Excel!
