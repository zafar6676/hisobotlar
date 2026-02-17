#!/usr/bin/env python3
"""Extract all fields from government reporting forms."""
import os
import sys
import json
import traceback
from pathlib import Path

# --- DOCX extraction ---
def extract_docx(filepath):
    from docx import Document
    doc = Document(filepath)
    result = {"paragraphs": [], "tables": []}

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            result["paragraphs"].append({
                "index": i,
                "style": para.style.name if para.style else "",
                "text": text
            })

    for ti, table in enumerate(doc.tables):
        tbl = {"table_index": ti, "rows": []}
        for ri, row in enumerate(table.rows):
            cells = []
            for ci, cell in enumerate(row.cells):
                cells.append(cell.text.strip())
            tbl["rows"].append(cells)
        result["tables"].append(tbl)

    return result

# --- XLSX / XLTX / XLSM extraction ---
def extract_xlsx(filepath):
    import openpyxl
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    result = {"sheets": []}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_data = {"name": sheet_name, "rows": [], "merged_cells": []}

        # Get merged cells info
        try:
            for mc in ws.merged_cells.ranges:
                sheet_data["merged_cells"].append(str(mc))
        except:
            pass

        row_count = 0
        for row in ws.iter_rows(values_only=False):
            row_data = []
            for cell in row:
                try:
                    val = cell.value
                    col = getattr(cell, 'column', None)
                    r = getattr(cell, 'row', None)
                    nf = getattr(cell, 'number_format', '') if val is not None else ''
                    row_data.append({
                        "col": col,
                        "row": r,
                        "value": str(val) if val is not None else "",
                        "number_format": nf
                    })
                except Exception:
                    row_data.append({"col": None, "row": None, "value": "", "number_format": ""})
            sheet_data["rows"].append(row_data)
            row_count += 1
            if row_count > 200:  # safety limit per sheet
                sheet_data["rows"].append([{"value": f"... TRUNCATED at 200 rows, sheet has more ..."}])
                break

        result["sheets"].append(sheet_data)

    wb.close()
    return result

# --- DOC extraction (legacy binary format) ---
def extract_doc(filepath):
    """Try to extract text from .doc using antiword or python-docx fallback."""
    import subprocess
    result = {"text": "", "method": ""}

    # Try antiword
    try:
        out = subprocess.run(["antiword", filepath], capture_output=True, text=True, timeout=30)
        if out.returncode == 0 and out.stdout.strip():
            result["text"] = out.stdout
            result["method"] = "antiword"
            return result
    except:
        pass

    # Try catdoc
    try:
        out = subprocess.run(["catdoc", filepath], capture_output=True, text=True, timeout=30)
        if out.returncode == 0 and out.stdout.strip():
            result["text"] = out.stdout
            result["method"] = "catdoc"
            return result
    except:
        pass

    # Try strings extraction
    try:
        out = subprocess.run(["strings", filepath], capture_output=True, text=True, timeout=30)
        if out.returncode == 0:
            result["text"] = out.stdout[:50000]
            result["method"] = "strings"
            return result
    except:
        pass

    result["text"] = "COULD NOT EXTRACT"
    result["method"] = "failed"
    return result


def main():
    base = Path("/home/user/hisobotlar")
    all_files = []

    for ext in ["*.doc", "*.docx", "*.xlsx", "*.xltx", "*.xlsm", "*.xls"]:
        all_files.extend(base.rglob(ext))

    # Skip 'old' directory and .git
    all_files = [f for f in all_files if '/old/' not in str(f) and '/.git/' not in str(f)]
    all_files = sorted(all_files, key=lambda x: str(x))

    output = {}

    for filepath in all_files:
        rel = str(filepath.relative_to(base))
        print(f"Processing: {rel}", file=sys.stderr)

        try:
            ext = filepath.suffix.lower()
            if ext == '.docx':
                data = extract_docx(str(filepath))
            elif ext in ['.xlsx', '.xltx', '.xlsm']:
                data = extract_xlsx(str(filepath))
            elif ext == '.doc':
                data = extract_doc(str(filepath))
            else:
                data = {"error": f"Unsupported format: {ext}"}

            output[rel] = data
        except Exception as e:
            output[rel] = {"error": str(e), "traceback": traceback.format_exc()}

    # Write output
    with open(base / "extracted_data.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\nProcessed {len(output)} files. Output: extracted_data.json", file=sys.stderr)

if __name__ == "__main__":
    main()
