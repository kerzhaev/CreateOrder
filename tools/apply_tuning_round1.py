#!/usr/bin/env python3
"""Round 1 tuning: only field widths and the premium row offset.
Column positions (Left) stay unchanged, so layout keeps the same
horizontal spacing — just the elements shrink to fit their content.
"""
from pathlib import Path

src = Path("CreateOrder.xlsm.modules/frmEnrollmentWizard.frm")
data = src.read_text(encoding="cp1251", newline="")

replacements = [
    # Percent fields: 52 -> 28 (only numeric values fit in 28 px).
    ("Private Const CTRL_PERCENT_WIDTH As Single = 52\r\n",
     "Private Const CTRL_PERCENT_WIDTH As Single = 28\r\n"),
    # Combo for secrecy/class/fizo: 136 -> 96.
    ("Private Const CTRL_PARAM_WIDTH As Single = 136\r\n",
     "Private Const CTRL_PARAM_WIDTH As Single = 96\r\n"),
    # Combo for medal: 174 -> 120.
    ("Private Const CTRL_PARAM_ACHIEVEMENT_WIDTH As Single = 174\r\n",
     "Private Const CTRL_PARAM_ACHIEVEMENT_WIDTH As Single = 120\r\n"),
    # Medal "amount" field: 70 -> 36.
    ("Private Const FRA430_COL_RIGHT_AMOUNT_WIDTH As Single = 70\r\n",
     "Private Const FRA430_COL_RIGHT_AMOUNT_WIDTH As Single = 36\r\n"),
    # Medal "Номер приказа": 192 -> 160 (header text is short).
    ("Private Const CTRL_DOC_WIDTH As Single = 192\r\n",
     "Private Const CTRL_DOC_WIDTH As Single = 160\r\n"),
    # Premium row top: 114 -> 126 (more vertical breathing room from secrecy row).
    ("Private Const FRA727_ROW3_CHK_TOP As Single = 114",
     "Private Const FRA727_ROW3_CHK_TOP As Single = 126"),
]

for old, new in replacements:
    assert old in data, f"Anchor not found:\n{old!r}"
    assert data.count(old) == 1, f"Anchor appears multiple times:\n{old!r}"
    data = data.replace(old, new, 1)

src.write_text(data, encoding="cp1251", newline="")
print("Round-1 tuning applied (widths only).")