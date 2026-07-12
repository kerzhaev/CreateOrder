#!/usr/bin/env python
"""
Create or refresh the workbook Localization sheet from ModuleLocalization.bas.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path

import pythoncom
import win32com.client.gencache
from win32com.client.dynamic import Dispatch


LOCALIZATION_SHEET_NAME = "Localization"
HEADER_ROW = ("key", "ru")
TRANSLATION_PREFIX = re.compile(
    r'^AddTranslation\s+"(?P<lang>[^"]+)",\s+"(?P<key>[^"]+)",\s+(?P<expr>.+)$',
    re.IGNORECASE,
)
STRING_LITERAL_PATTERN = re.compile(r'"([^"]*)"')


def read_module_text(module_path: Path) -> str:
    for encoding in ("utf-8", "cp1251"):
        try:
            return module_path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("decode", b"", 0, 1, f"Unsupported encoding for {module_path}")


def reset_excel_gen_cache() -> None:
    gen_path = Path(win32com.client.gencache.GetGeneratePath())
    for child in gen_path.glob("00020813-0000-0000-C000-000000000046*"):
        if child.is_dir():
            shutil.rmtree(child, ignore_errors=True)
        elif child.exists():
            child.unlink(missing_ok=True)


def get_excel_open_path(workbook_path: Path) -> str:
    return workbook_path.resolve().as_uri() if sys.platform == "win32" else str(workbook_path.resolve())


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Refresh workbook Localization sheet.")
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument(
        "--module-path",
        type=Path,
        default=Path("CreateOrder.xlsm.modules/ModuleLocalization.bas"),
        help="Path to ModuleLocalization.bas",
    )
    return parser


def parse_translations(module_path: Path) -> dict[str, dict[str, str]]:
    content = read_module_text(module_path)
    translations: dict[str, dict[str, str]] = defaultdict(dict)

    for raw_line in content.splitlines():
        line = raw_line.strip()
        if not line.lower().startswith("addtranslation "):
            continue

        match = TRANSLATION_PREFIX.match(line)
        if not match:
            continue

        lang = match.group("lang").strip().lower()
        key = match.group("key").strip().lower()
        expr = match.group("expr")
        literal_parts = STRING_LITERAL_PATTERN.findall(expr)
        value = "\n".join(literal_parts)
        translations[key][lang] = value

    return dict(translations)


def get_or_create_sheet(workbook):
    for index in range(1, workbook.Worksheets.Count + 1):
        ws = workbook.Worksheets(index)
        if ws.Name == LOCALIZATION_SHEET_NAME:
            return ws, False

    ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    ws.Name = LOCALIZATION_SHEET_NAME
    return ws, True


def write_localization_sheet(ws, translations: dict[str, dict[str, str]]) -> None:
    ws.Cells.Clear()

    for col_index, header in enumerate(HEADER_ROW, start=1):
        ws.Cells(1, col_index).Value = header

    for row_index, key in enumerate(sorted(translations.keys()), start=2):
        ws.Cells(row_index, 1).Value = key
        ws.Cells(row_index, 2).Value = translations[key].get("ru", "")

    ws.Columns("A:B").AutoFit()


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    module_path = args.module_path.resolve()

    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1
    if not module_path.exists():
        print(f"Localization module not found: {module_path}", file=sys.stderr)
        return 1

    try:
        translations = parse_translations(module_path)
    except Exception as exc:  # noqa: BLE001
        print(f"LOCALIZATION PARSE ERROR: {exc}", file=sys.stderr)
        return 1

    pythoncom.CoInitialize()
    reset_excel_gen_cache()
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None

    try:
        workbook = excel.Workbooks.Open(get_excel_open_path(workbook_path))
        ws, created = get_or_create_sheet(workbook)
        write_localization_sheet(ws, translations)
        workbook.Save()
        print(f"sheet={'created' if created else 'updated'} rows={len(translations)}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"LOCALIZATION SHEET ERROR: {exc}", file=sys.stderr)
        return 1
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
