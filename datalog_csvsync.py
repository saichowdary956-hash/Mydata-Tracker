#!/usr/bin/env python3
"""
Watches a local Data Logger CSV export folder and appends each new Drive ID
to the DailyTracker sheet in the recreated Excel workbook.

Requirements:
    py -m pip install openpyxl

Run:
    py datalog_csv_sync.py --watch "C:\\Users\\YOURNAME\\Downloads" ^
      --workbook "C:\\DataLoggerAutomation\\Daily_Tracker_Recreated.xlsx"
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import time
from datetime import datetime, date
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

CATEGORY_COLUMNS = {
    "Weather": {
        "Sunny": "Sunny", "Low Sun": "Low Sun", "Cloudy": "Cloudy",
        "Rain": "Rain", "Fog": "Fog", "Snow": "Snow",
    },
    "Roadtype": {
        "City": "City (intense traffic)", "Country": "Country",
        "Highway": "Highway", "Construction site": "Construction Site",
        "Tunnel": "Tunnel",
    },
    "Lighting": {
        "Day": "Day", "Dawn": "Dawn", "Dusk": "Dusk",
        "Lit Night": "Lit Night", "Dark Night": "Dark Night",
    },
    "Traffic": {"Flow": "Flow", "Jam": "Jam"},
    "Speed": {
        "3-18mph": "03_18 mph", "19-37mph": "18_37 mph",
        "38-55mph": "37_55 mph", "56-80mph": "55_80 mph",
        "81-155mph": "81_155 mph",
    },
}
CATEGORY_ORDER = ["Weather", "Roadtype", "Lighting", "Traffic", "Speed"]


def duration_to_excel(value: str) -> float:
    """Convert HH:MM:SS or H:MM:SS to an Excel day fraction."""
    value = (value or "").strip()
    match = re.fullmatch(r"(\d+):([0-5]?\d):([0-5]?\d)", value)
    if not match:
        return 0.0
    hours, minutes, seconds = map(int, match.groups())
    return (hours * 3600 + minutes * 60 + seconds) / 86400


def read_csv_report(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    metadata: dict[str, str] = {}
    for row in rows:
        if len(row) >= 2:
            key = row[0].strip().rstrip(":").lower()
            if key in {
                "driver", "annotator", "date", "vehicle", "rsu no",
                "rsu storage", "drive id", "overall session time", "comments",
            }:
                metadata[key] = row[1].strip()

    drive_id = metadata.get("drive id", "").strip()
    if not drive_id:
        raise ValueError(f"No Drive ID found in {path.name}; file left for review.")

    header_idx = next(
        (i for i, row in enumerate(rows) if row and row[0].strip().lower() == "category"),
        None,
    )
    if header_idx is None:
        raise ValueError(f"Category header not found in {path.name}.")

    category_positions: dict[str, int] = {}
    header = rows[header_idx]
    for idx, cell in enumerate(header):
        label = cell.strip()
        if label in CATEGORY_ORDER:
            category_positions[label] = idx

    durations: dict[str, float] = {}
    for row in rows[header_idx + 1:]:
        for category, start_idx in category_positions.items():
            if start_idx + 2 >= len(row):
                continue
            condition = row[start_idx].strip()
            duration = row[start_idx + 1].strip()
            target = CATEGORY_COLUMNS.get(category, {}).get(condition)
            if target:
                # Sum duplicate condition rows if the source CSV contains them.
                durations[target] = durations.get(target, 0.0) + duration_to_excel(duration)

    raw_date = metadata.get("date", "")
    parsed_date = None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            parsed_date = datetime.strptime(raw_date, fmt).date()
            break
        except ValueError:
            pass

    return {
        "driver": metadata.get("driver", ""),
        "annotator": metadata.get("annotator", ""),
        "date": parsed_date,
        "vehicle": metadata.get("vehicle", ""),
        "rsu_no": metadata.get("rsu no", ""),
        "rsu_storage": metadata.get("rsu storage", "").replace("%", "").strip(),
        "drive_id": drive_id,
        "session_time": duration_to_excel(metadata.get("overall session time", "")),
        "comments": metadata.get("comments", ""),
        "durations": durations,
        "source_file": path.name,
    }


def sync_one(csv_path: Path, workbook_path: Path, sheet_name: str = "DailyTracker") -> str:
    report = read_csv_report(csv_path)
    wb = load_workbook(workbook_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Worksheet '{sheet_name}' not found.")
    ws = wb[sheet_name]

    headers = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    required = ["Driver", "Annotator", "Date", "Vehicle", "Drive ID"]
    missing = [h for h in required if h not in headers]
    if missing:
        raise ValueError(f"Workbook is missing required columns: {', '.join(missing)}")

    drive_col = headers["Drive ID"]
    for r in range(2, ws.max_row + 1):
        existing = ws.cell(r, drive_col).value
        if existing is not None and str(existing).strip() == report["drive_id"]:
            return f"SKIPPED duplicate Drive ID {report['drive_id']} (existing row {r}); manual values untouched."

    # Insert before the TOTAL row if present; otherwise append after current records.
    total_row = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().upper() == "TOTAL":
            total_row = r
            break
    target_row = total_row if total_row else ws.max_row + 1
    if total_row:
        ws.insert_rows(total_row, 1)

    def put(header_name: str, value: Any):
        col = headers.get(header_name)
        if col:
            ws.cell(target_row, col).value = value

    put("Driver", report["driver"])
    put("Annotator", report["annotator"])
    put("Date", report["date"])
    put("Vehicle", report["vehicle"])
    put("Drive ID", report["drive_id"])
    put("RSU Number", report["rsu_no"])
    put("RSU Storage %", float(report["rsu_storage"]) if report["rsu_storage"] else None)
    put("Drive cycle minutes", report["session_time"])
    put("Comments", report["comments"])

    for col_name, seconds_as_days in report["durations"].items():
        put(col_name, seconds_as_days)

    # Date is available from the Data Logger. Route/day/locations and cycle
    # start/end times are intentionally left blank for manual completion.
    for name in ("Date", "RSU Start Date"):
        col = headers.get(name)
        if col and ws.cell(target_row, col).value:
            ws.cell(target_row, col).number_format = "m/d/yyyy"
    for name in list(report["durations"].keys()) + ["Drive cycle minutes"]:
        col = headers.get(name)
        if col:
            ws.cell(target_row, col).number_format = "[h]:mm:ss"

    # Keep records chronological by Date; ties remain stable. Rows without dates
    # stay after dated records. Preserve all row values when sorting.
    if total_row:
        data_end = target_row
    else:
        data_end = ws.max_row
    if data_end >= 2:
        data_rows = []
        for r in range(2, data_end + 1):
            values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if str(values[0] or "").strip().upper() == "TOTAL":
                continue
            d = values[headers["Date"] - 1] if headers.get("Date") else None
            sort_key = (d is None, d.isoformat() if hasattr(d, "isoformat") else str(d or ""), str(values[headers["Drive ID"] - 1] or ""))
            data_rows.append((sort_key, values))
        data_rows.sort(key=lambda item: item[0])
        for idx, (_, values) in enumerate(data_rows, start=2):
            for c, value in enumerate(values, start=1):
                ws.cell(idx, c).value = value

    # Rebuild TOTAL formulas in the last row.
    total_row_now = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().upper() == "TOTAL":
            total_row_now = r
            break
    if total_row_now:
        for c in range(17, min(41, ws.max_column + 1)):
            letter = ws.cell(1, c).column_letter
            ws.cell(total_row_now, c).value = f"=SUM({letter}2:{letter}{total_row_now-1})"
            ws.cell(total_row_now, c).number_format = "[h]:mm:ss"

    wb.save(workbook_path)
    return f"IMPORTED {report['drive_id']} from {csv_path.name} into row {target_row}."


def wait_until_stable(path: Path, checks: int = 3, delay: float = 1.0) -> bool:
    previous = -1
    stable = 0
    for _ in range(checks * 4):
        try:
            current = path.stat().st_size
        except FileNotFoundError:
            return False
        if current > 0 and current == previous:
            stable += 1
            if stable >= checks:
                return True
        else:
            stable = 0
        previous = current
        time.sleep(delay)
    return False


def main():
    parser = argparse.ArgumentParser(description="Automatically import Data Logger CSV exports into Excel.")
    parser.add_argument("--watch", required=True, help="Folder where Data Logger CSV files are downloaded.")
    parser.add_argument("--workbook", required=True, help="Path to Daily_Tracker_Recreated.xlsx.")
    parser.add_argument("--processed", default="Processed_DataLogger_CSV", help="Archive folder for successfully imported CSVs.")
    parser.add_argument("--once", action="store_true", help="Process existing CSVs once, then exit.")
    args = parser.parse_args()

    watch_dir = Path(args.watch).expanduser().resolve()
    workbook = Path(args.workbook).expanduser().resolve()
    archive = watch_dir / args.processed
    watch_dir.mkdir(parents=True, exist_ok=True)
    archive.mkdir(parents=True, exist_ok=True)

    if not workbook.exists():
        raise SystemExit(f"Workbook not found: {workbook}")

    print(f"Watching: {watch_dir}")
    print(f"Workbook: {workbook}")
    print("Close the workbook in Excel while importing, or save a copy first.")
    seen: set[str] = set()

    while True:
        for csv_path in sorted(watch_dir.glob("*.csv")):
            if csv_path.parent == archive or csv_path.name in seen:
                continue
            if not wait_until_stable(csv_path):
                continue
            try:
                result = sync_one(csv_path, workbook)
                print(result)
                if result.startswith("IMPORTED") or result.startswith("SKIPPED"):
                    destination = archive / csv_path.name
                    if destination.exists():
                        destination = archive / f"{csv_path.stem}_{int(time.time())}{csv_path.suffix}"
                    shutil.move(str(csv_path), str(destination))
                seen.add(csv_path.name)
            except PermissionError:
                print(f"Workbook/file is locked. Close Excel and retry: {csv_path.name}")
            except Exception as exc:
                print(f"ERROR processing {csv_path.name}: {exc}")
        if args.once:
            break
        time.sleep(3)


if __name__ == "__main__":
    main()
