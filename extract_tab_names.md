## Legacy Macro
### Steps
- Create a Named Range
- Go to Formulas ➜ Name Manager ➜ New.
- Name: SheetNames
- Refers to:
```=GET.WORKBOOK(1)```
- Click OK.

- Spill the sheet names into cells
- In a worksheet (e.g., A1), enter the following formula:
```=TEXTAFTER(SheetNames,"]")```
This will spill all tab names down the column.
Compatible with older versions:
```=IFERROR(MID(INDEX(SheetNames,ROW(A1)),FIND("]",INDEX(SheetNames,ROW(A1)))+1,255),"")```

- Fill down until blanks appear (or use dynamic arrays if available).


## Power Query can list all sheet names in a clean table.
### Steps

- Save the workbook.
- Go to Data ➜ Get Data ➜ From File ➜ From Workbook.
- Browse to this same file (the one you're working in) and open it.
- In the Navigator, you’ll see Sheets with their names. Select them or click Transform Data.
- In Power Query, you’ll see a table including a Name column (these are the worksheet names).
- Click Close & Load to load the list to a new sheet.
