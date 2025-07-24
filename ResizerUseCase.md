# Excel Screenshot Macro Setup

This guide shows how to set a hotkey for the `Resizer` macro and create a new macro to paste a screenshot, resize it with `Resizer`, and align it with a selected cell in Excel using a hotkey.

## What You Need
- Excel (2016, 2019, 2021, or 365) with Developer tab enabled.
- `PERSONAL.XLSB` with the `Resizer` macro (from [https://github.com/astroVanHalen/VB_Scripts](https://github.com/astroVanHalen/VB_Scripts)).
- A screenshot tool (e.g., Windows **Win + Shift + S**, Mac **Cmd + Shift + 4**).

## Step 1: Set Hotkey for Resizer Macro
1. In Excel, go to **Developer** > **Macros** (or **Alt + F8**).
2. Select `PERSONAL.XLSB!Resizer`, click **Options**.
3. Enter a letter (e.g., `y` for **Ctrl + y**). Avoid **Ctrl + C/V**.
4. Click **OK**, then close the dialog.

## Step 2: Record Screenshot Macro
1. Take a screenshot (e.g., **Win + Shift + S** or **Cmd + Shift + 4**) to copy it to your clipboard.
2. Select a cell in Excel (e.g., A1).
3. Go to **Developer** > **Record Macro**.
   - Name: `PasteAndResizeScreenshot`or whatever you would like.
   - Store in: **Personal Macro Workbook**.
   - Click **OK** to start recording.
4. Paste the screenshot (**Ctrl + V**).
5. Press the `Resizer` hotkey (e.g., **Ctrl + y**).
6. Click the screenshot, cut it (**Ctrl + X**), ensure the same cell is selected, and paste (**Ctrl + V**).
7. Go to **Developer** > **Stop Recording**.
8. Verify the macro in VBA Editor (**Alt + F11**, under `PERSONAL.XLSB` > **Modules**).
9. Save `PERSONAL.XLSB` (**Ctrl + S** in VBA Editor).

## Step 3: Assign Hotkey to Screenshot Macro
1. Go to **Developer** > **Macros** (or **Alt + F8**).
2. Select `PERSONAL.XLSB!PasteAndResizeScreenshot`, click **Options**.
3. Enter a letter (e.g., `g` for **Ctrl + g**). Avoid conflicts.
4. Click **OK**, close the dialog, and save `PERSONAL.XLSB`.

## Step 4: Test the Macro
1. Take a screenshot.
2. Select a cell, press the macro’s hotkey (e.g., **Ctrl + Shift + S**).
3. Check that the screenshot is pasted, resized, and aligned with the cell’s top-left corner.

## Tips
- **Enable Macros**: Click **Enable Content** if Excel prompts when opening a workbook.
- **Troubleshooting**: If the macro fails, ensure `Resizer`’s hotkey is correct and both macros are in `PERSONAL.XLSB`.
- **Repository**: Get scripts at [https://github.com/astroVanHalen/VB_Scripts](https://github.com/astroVanHalen/VB_Scripts).

## Example
- **Macro**: `PasteAndResizeScreenshot`
- **Purpose**: Pastes, resizes, and aligns a screenshot with a cell.
- **Hotkey**: E.g., **Ctrl + g**.

## Help
Open an issue at [https://github.com/astroVanHalen/VB_Scripts](https://github.com/astroVanHalen/VB_Scripts) for support.