# Data Logger – Local CSV-to-Excel Automation (separate add-on)

This package does not modify the live Vercel Data Logger. It watches the browser's
CSV download folder on your Windows laptop and imports new Data Logger reports
into `Daily_Tracker_Recreated.xlsx`.

## 1. Files
- `datalog_csv_sync.py` – local folder watcher and Excel importer
- `Daily_Tracker_Recreated.xlsx` – recreated workbook based on your screenshots
- `START_SYNC.bat` – example launcher (edit the two paths first)

## 2. Install once
Install Python 3.10+ for Windows, then open Command Prompt and run:

    py -m pip install openpyxl

## 3. Configure
1. Copy `Daily_Tracker_Recreated.xlsx` to a permanent folder, e.g. `C:\DataLoggerAutomation`.
2. Open `START_SYNC.bat` in Notepad.
3. Change `YOUR_WINDOWS_USERNAME` to your Windows account folder name.
4. Save the batch file.

The default watcher folder is your Windows Downloads folder. If your browser saves
Data Logger CSV files elsewhere, change the `--watch` path accordingly.

## 4. Start
Double-click `START_SYNC.bat`. Keep its black Command Prompt window open while
you are collecting data. When a new CSV appears, it will be imported automatically.
Successfully processed CSVs are moved into `Processed_DataLogger_CSV`.

## 5. Import rules
- Uses `Drive ID` as the duplicate key. If a Drive ID already exists, it skips the
  entire new CSV and does not overwrite the existing row or manual corrections.
- Imports Driver, Annotator, Date, Vehicle, Drive ID, RSU Number, RSU Storage,
  Overall Session Time, Comments, and category durations.
- Converts durations to Excel time values with `[h]:mm:ss` formatting.
- Sorts imported data by Date. The current Data Logger CSV does not include a
  drive-cycle start time, so it cannot reliably sort same-day drives by start time.
- Leaves Route, State/City, Start/End Point, Drive Cycle Start/End Time,
  Calendar Week, Z Frame Checker, and Sanity Check fields blank for manual entry.
- Keep the workbook closed in Excel during imports to avoid file-lock errors.

## Important current-source limitation
Your supplied Data Logger CSV export code uses `recordingSessions[category].find(...)`,
which exports only the first recorded session for a condition. If a condition is
started/stopped multiple times, the CSV may omit later sessions. Before relying on
the totals, update the export loop to use the cumulative `conditionData` total
instead. The exact code replacement is included in `DataLogger_Export_Fix.js.txt`.
