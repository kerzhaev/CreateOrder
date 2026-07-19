#!/usr/bin/env python3
"""Geometry verifier — only checks numeric arguments (coordinates / sizes)."""
import re
import subprocess
from pathlib import Path

old_bytes = subprocess.check_output(
    ["git", "show", "HEAD:CreateOrder.xlsm.modules/frmEnrollmentWizard.frm"],
    cwd=".",
)
old = old_bytes.decode("cp1251", errors="replace")
new = Path("CreateOrder.xlsm.modules/frmEnrollmentWizard.frm").read_text(encoding="cp1251")

# Parse constants block.
const_block_match = re.search(
    r"' ----- Layout constants for the monthly payments page.*?(?=Private mpWizard As Object)",
    new, re.DOTALL,
)
if not const_block_match:
    print("Constants block not found!")
    raise SystemExit(1)
consts_text = const_block_match.group(0)
constants = {}
for line in consts_text.splitlines():
    if "'" in line:
        line = line[:line.index("'")]
    line = line.strip()
    if not line.startswith("Private Const "):
        continue
    m = re.match(r"Private Const (\w+) As Single = (.+)", line)
    if m:
        try:
            constants[m.group(1)] = float(eval(m.group(2), {"__builtins__": {}}, {}))
        except Exception:
            pass

print(f"Parsed {len(constants)} constants.")

# Find CreateMonthlyPage in both versions.
def find_cpm(text):
    m = re.search(r"Private Sub CreateMonthlyPage\(\).*?End Sub", text, re.DOTALL)
    return m.group(0) + "\n" if m else None

old_cpm = find_cpm(old)
new_cpm = find_cpm(new)
assert old_cpm and new_cpm, "CreateMonthlyPage missing"

# Signatures to know which arg indices are numeric:
# AddPageCheckBoxT(pageHost, locKey, fallbackText, left, top)
#   numeric: [3]=left, [4]=top
# AddPageTextBoxT(pageHost, locKey, fallbackText, left, top, [width, [height, [multiline, [readOnly]]]])
#   numeric: [3]=left, [4]=top, [5]=width, [6]=height, [7]=multiline(bool), [8]=readOnly(bool)
# AddPageComboBoxT(pageHost, locKey, fallbackText, left, top, [width])
#   numeric: [3]=left, [4]=top, [5]=width
# AddPageFrame(pageHost, name, caption, left, top, width, height)
#   numeric: [3]=left, [4]=top, [5]=width, [6]=height
# AddLabelToPage(pageHost, labelText, left, top, width)
#   numeric: [2]=left, [3]=top, [4]=width

numeric_indices = {
    "AddPageCheckBoxT": {3, 4},
    "AddPageTextBoxT": {3, 4, 5, 6, 7, 8},
    "AddPageComboBoxT": {3, 4, 5},
    "AddPageFrame": {3, 4, 5, 6},
    "AddLabelToPage": {2, 3, 4},
}


def split_args(s):
    """Split by top-level commas, respecting quotes."""
    parts, buf, in_quote = [], "", False
    for ch in s:
        if ch == '"':
            in_quote = not in_quote
            buf += ch
        elif ch == "," and not in_quote:
            parts.append(buf.strip())
            buf = ""
        else:
            buf += ch
    if buf.strip():
        parts.append(buf.strip())
    return parts


def parse_calls(text):
    out = {}
    for line in text.splitlines():
        s = line.strip()
        m = re.match(r"Set (\w+) = (\w+)\((.+)\)$", s)
        if not m:
            continue
        var, func, args_str = m.group(1), m.group(2), m.group(3)
        out[var] = (func, split_args(args_str))
    return out

old_calls = parse_calls(old_cpm)
new_calls = parse_calls(new_cpm)


def to_float(tok, env):
    """Try to interpret a token as float; returns float or raises ValueError."""
    t = tok.strip()
    if t.startswith('"') and t.endswith('"'):
        raise ValueError("string literal")
    if t in env:
        return env[t]
    return float(t)


errors = []
for var, (old_func, old_args) in sorted(old_calls.items()):
    if var not in new_calls:
        continue
    new_func, new_args = new_calls[var]
    if old_func != new_func:
        errors.append(f"{var}: func {old_func} != {new_func}")
        continue
    idx_set = numeric_indices.get(old_func, set())
    for i in idx_set:
        if i >= len(old_args) or i >= len(new_args):
            continue
        oa, na = old_args[i], new_args[i]
        if oa == na:
            continue
        try:
            old_num = to_float(oa, constants)
        except ValueError:
            continue
        try:
            new_num = to_float(na, constants)
        except ValueError:
            errors.append(f"{var} arg[{i}]: old={oa}, new={na} is not numeric")
            continue
        if abs(old_num - new_num) > 0.001:
            errors.append(
                f"{var} arg[{i}]: old={old_num}, new='{na}' -> {new_num}  MISMATCH"
            )

if errors:
    print(f"\nFound {len(errors)} mismatch(es):")
    for e in errors:
        print("  -", e)
    raise SystemExit(1)

matched = sum(1 for k in old_calls if k in new_calls)
print(f"All {matched} control assignments verified — no numeric mismatches.")

# Also: ensure all NEW constant refs in CreateMonthlyPage are defined.
cpm_const_refs = set()
for line in new_cpm.splitlines():
    for m in re.finditer(r"\b(FRA\w+|CTRL\w+)\b", line):
        cpm_const_refs.add(m.group(1))
missing = cpm_const_refs - set(constants.keys())
if missing:
    print(f"Undefined constants in CreateMonthlyPage: {sorted(missing)}")
    raise SystemExit(1)
print(f"All {len(cpm_const_refs)} constant references in CreateMonthlyPage are defined.")