#!/usr/bin/env python3
"""
Generic XLSX builder — reads a JSON spec and produces a styled Excel file.

Usage:
    python3 xlsx_builder.py <input.json> <output.xlsx>

JSON spec format:
{
  "sheets": [
    {
      "name": "Sheet1",
      "freeze_pane": "A2",          // optional
      "col_widths": {"A": 15, "B": 30},  // optional
      "headers": ["Col A", "Col B", "Col C"],
      "header_color": "1F3864",     // optional, hex without #
      "rows": [
        {
          "values": ["val1", "val2", "val3"],
          "row_color": "EBF3FB",    // optional
          "bold": false             // optional
        }
      ]
    }
  ]
}
"""

import sys
import json
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DEFAULT_HEADER_COLOR = "1F3864"
DEFAULT_ALT_COLOR    = "EBF3FB"
WHITE                = "FFFFFF"

def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def build_xlsx(spec: dict, output_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    for sheet_spec in spec["sheets"]:
        ws = wb.create_sheet(title=sheet_spec["name"])

        header_color = sheet_spec.get("header_color", DEFAULT_HEADER_COLOR)
        headers = sheet_spec.get("headers", [])

        # Write header row
        if headers:
            for col_i, h in enumerate(headers, start=1):
                c = ws.cell(row=1, column=col_i, value=h)
                c.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
                c.fill = make_fill(header_color)
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = thin_border()
            ws.row_dimensions[1].height = 20

        # Write data rows
        rows = sheet_spec.get("rows", [])
        for row_i, row_spec in enumerate(rows, start=2):
            values = row_spec.get("values", [])
            row_color = row_spec.get("row_color", WHITE if row_i % 2 == 0 else DEFAULT_ALT_COLOR)
            is_bold = row_spec.get("bold", False)

            for col_i, val in enumerate(values, start=1):
                c = ws.cell(row=row_i, column=col_i, value=val)
                c.fill = make_fill(row_color)
                c.font = Font(name="Calibri", bold=is_bold, size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)
                c.border = thin_border()

        # Column widths
        col_widths = sheet_spec.get("col_widths", {})
        if col_widths:
            for col_letter, width in col_widths.items():
                ws.column_dimensions[col_letter].width = width
        else:
            # Auto-fit (approximate)
            for col_cells in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col_cells[0].column)
                for cell in col_cells:
                    try:
                        val_len = len(str(cell.value)) if cell.value else 0
                        max_len = max(max_len, min(val_len, 60))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = max(10, max_len + 2)

        # Freeze pane
        freeze = sheet_spec.get("freeze_pane")
        if freeze:
            ws.freeze_panes = freeze

        # Auto-filter on header row
        if headers:
            last_col = get_column_letter(len(headers))
            ws.auto_filter.ref = f"A1:{last_col}1"

    wb.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: xlsx_builder.py <input.json> <output.xlsx>")
        sys.exit(1)

    input_json  = sys.argv[1]
    output_xlsx = sys.argv[2]

    with open(input_json, "r", encoding="utf-8") as f:
        spec = json.load(f)

    build_xlsx(spec, output_xlsx)
