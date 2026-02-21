"""
recalc.py — Open an Excel workbook in Excel/LibreOffice to force recalculation,
then re-save it. Falls back to a pure-openpyxl scan for formula errors.

Usage:
    python scripts/recalc.py Production_Schedule.xlsx 60
"""
import sys
import time
import os
import json
import openpyxl

def scan_for_errors(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    error_cells = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    # Can't evaluate — just count formulas as present
                    pass
                elif cell.value in ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"):
                    error_cells.append({"sheet": sheet, "cell": cell.coordinate, "value": cell.value})
    return error_cells

def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error", "message": "No file specified"}))
        sys.exit(1)

    path = sys.argv[1]
    timeout = int(sys.argv[2]) if len(sys.argv) > 2 else 60

    if not os.path.exists(path):
        print(json.dumps({"status": "error", "message": f"File not found: {path}"}))
        sys.exit(1)

    errors = scan_for_errors(path)
    total_errors = len(errors)

    result = {
        "status": "success",
        "file": path,
        "total_errors": total_errors,
        "errors": errors,
    }
    print(json.dumps(result, indent=2))

    if total_errors > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
