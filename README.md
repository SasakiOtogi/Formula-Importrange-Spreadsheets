# Formula-Importrange-Spreadsheets

A Google Apps Script for Google Sheets with helpers for workbooks that pull data from another spreadsheet using `IMPORTRANGE`. It does three things:

- Shows the name of the previous sheet (tab) inside a cell.
- Forces every `IMPORTRANGE` formula in the workbook to refresh.
- Duplicates the current sheet and moves its `HARIAN!` references down one row, so you can make a new daily sheet in one click.

All code is in [`Formula Importrange Spreadsheets.js`](Formula%20Importrange%20Spreadsheets.js).

---

## Installation

1. Open your Google Sheets file.
2. Go to **Extensions → Apps Script**.
3. Delete the default code and paste in the contents of `Formula Importrange Spreadsheets.js`.
4. Save the project (💾 or `Ctrl + S`).
5. Reload the spreadsheet. A new menu called **Duplicate Sheet** appears in the toolbar.
6. The first time you run a function, Google asks you to grant permissions. Accept them so the script can read and edit the spreadsheet.

---

## Functions

### 1. `PREVIOUS_SHEET_NAME(trigger)`: custom formula

Returns the name of the sheet directly to the **left** of the sheet the formula is in. If the formula is in the first sheet, it returns `"N/A"`.

**Usage in a cell:**

```
=PREVIOUS_SHEET_NAME()
```

**Why the `trigger` parameter?**
Google Sheets caches the results of custom functions and doesn't recalculate them when you rename or move sheets. The value you pass as `trigger` is ignored by the code. Its only job is to make the formula recalculate whenever that value changes, for example:

```
=PREVIOUS_SHEET_NAME(RANDBETWEEN(1, 100))
=PREVIOUS_SHEET_NAME(NOW())
```

**Common use:** build references to the previous day's sheet, for example combined with `INDIRECT`:

```
=INDIRECT("'" & PREVIOUS_SHEET_NAME(NOW()) & "'!B2")
```

---

### 2. `refreshAllImportRange()`: refresh all IMPORTRANGE formulas

Sometimes `IMPORTRANGE` keeps showing old data. This function goes through **every sheet** in the spreadsheet and, for each cell whose formula contains `IMPORTRANGE` (any capitalization), it:

1. Saves the formula.
2. Clears the cell (`setFormula('')`).
3. Applies the change immediately (`SpreadsheetApp.flush()`).
4. Puts the original formula back.

Re-entering the formula makes Google Sheets fetch the data again.

**Running it automatically (recommended):**

1. In the Apps Script editor, open **Triggers** (the ⏰ clock icon on the left).
2. Click **+ Add Trigger**.
3. Set:
   - Function: `refreshAllImportRange`
   - Event source: **Time-driven**
   - Interval: for example, **every 5 minutes** or **every hour**
4. Save.

> ⚠️ The script processes each matching cell one at a time and calls `flush()` for each one. Workbooks with many `IMPORTRANGE` cells can take a while to run and can hit Apps Script's execution time limit (6 minutes per run).

---

### 3. `duplicateSheetAndIncrementFormula()`: duplicate sheet and move to the next row

Built for a daily-report workflow where each new sheet reads the **next row** of a source sheet called `HARIAN` ("daily").

Steps:

1. Copies the active sheet into the same spreadsheet.
2. Reads every formula in the new sheet.
3. Finds references in the form `HARIAN!<column><row>` where the column is **N, P, R, J, or L**.
4. Adds **1** to the row number of each matching reference.
5. Switches to the new sheet.

**Example:**

| Original sheet          | Duplicated sheet        |
|-------------------------|-------------------------|
| `=HARIAN!N460`          | `=HARIAN!N461`          |
| `=HARIAN!P12 + HARIAN!R12` | `=HARIAN!P13 + HARIAN!R13` |
| `=HARIAN!A5` (column A)  | `=HARIAN!A5` (unchanged) |

**How to run it:** use the **Duplicate Sheet → Duplicate** menu in the spreadsheet.

**Notes:**
- Only columns `N, P, R, J, L` are changed. To change other columns, edit the regex in the code:
  ```js
  const targetPattern = /HARIAN!([NPRJL])(\d+)/gi;
  ```
- References with `$` (for example `HARIAN!$N$460`) or with a quoted sheet name (for example `'HARIAN'!N460`) do **not** match the pattern and are left unchanged.
- The new sheet gets Google's default name (for example "Copy of …"). Rename it as needed.

---

### 4. `onOpen()`: custom menu

Runs automatically when the spreadsheet opens and adds this menu:

```
Duplicate Sheet
 └── Duplicate   → runs duplicateSheetAndIncrementFormula()
```

---

## Example Workflow

1. The `HARIAN` sheet holds daily data, one row per day.
2. A report sheet (for example `Day 1`) reads values with `IMPORTRANGE` and/or `HARIAN!N460`, `HARIAN!P460`, and so on.
3. Each new day, click **Duplicate Sheet → Duplicate**. A new sheet is created that reads row `461`.
4. In the new sheet, `PREVIOUS_SHEET_NAME()` shows the previous day's sheet name, which you can use to compare against yesterday's data.
5. A time-driven trigger runs `refreshAllImportRange()` so imported data stays up to date.
