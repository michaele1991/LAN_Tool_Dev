# GBE Image Builder — AI Workflow Instructions

## Trigger
When the user says "create GBE image", "build image", "make NVM image", or similar — execute the workflow below automatically. Ask ONE question at a time.

---

## Step 1 — Project
Run this and show as a numbered list, then ask "Which project?":
```powershell
Get-ChildItem "..\GBE_Image_Creator\GBE_Image" -Directory | Select-Object -ExpandProperty Name | Sort-Object
```

## Step 2 — Offset
Ask: "NVM word offset? (hex, e.g. 0x59)"

Then look up and display all bits at that offset from the project XLSM:
```python
from openpyxl import load_workbook
from pathlib import Path
project = "<chosen>"
xlsm = next(Path(f"../GBE_Image_Creator/GBE_Image/{project}").glob("*.xlsm"))
wb = load_workbook(str(xlsm), read_only=True, data_only=True, keep_vba=False)
ws = wb["full nvm map"]
target = "<offset>".lower().lstrip("0x").rstrip("h") + "h"
for row in ws.iter_rows(min_row=7, values_only=True):
    if row[0] and str(row[0]).lower().strip() == target:
        print(f"Bit {row[1]:>4}  {str(row[2]):<45}  V={row[6]}  LM={row[7]}")
wb.close()
```
Show table: Bit | RTL name | Description | Current V | Current LM

## Step 3 — Bit
Ask: "Which bit?"

## Step 4 — Mode
Ask: "V, LM, or Both?"

## Step 5 — New value
Ask: "New value? (0 or 1)"
Warn if new value equals current value.

## Done check
Ask: "Any more changes? (y/n)"
- y → back to Step 2
- n → show full change summary, ask "Build now? (y/n)"

---

## Build
```python
import struct, shutil
from pathlib import Path

proj_dir = Path(f"../GBE_Image_Creator/GBE_Image/<project>")
out_dirs = [d for d in proj_dir.iterdir() if d.is_dir()]
cons_dir = next((d for d in out_dirs if any(x in d.name.lower() for x in ("cons","_v_","lan"))), out_dirs[0])
corp_dir = next((d for d in out_dirs if any(x in d.name.lower() for x in ("corp","lm","non"))), out_dirs[-1])

def patch(bin_path, offset_int, bit_int, new_bit):
    data = bytearray(bin_path.read_bytes())
    pos = offset_int * 2
    old = struct.unpack_from("<H", data, pos)[0]
    new = (old | (1 << bit_int)) if new_bit else (old & ~(1 << bit_int))
    struct.pack_into("<H", data, pos, new)
    cpos = 0x3F * 2
    struct.pack_into("<H", data, cpos, 0)
    s = sum(struct.unpack_from("<H", data, i)[0] for i in range(0, cpos, 2))
    csum = (0xBABA - s) & 0xFFFF
    struct.pack_into("<H", data, cpos, csum)
    bin_path.write_bytes(data)
    return old, new, csum
```
- Backup `.bin` → `.bin.bak` before patching
- mode V  → patch cons_dir/*.bin
- mode LM → patch corp_dir/*.bin
- mode Both → patch both
- Also patch *.txt (same word positions, space-separated hex)
- Save `change_request_<timestamp>.json` + `change_report_<timestamp>.txt` in project folder

---

## XLSM columns (sheet "full nvm map", data from row 7)
A=Offset  B=Bits  C=RTL name  D=C-Spec  F=Description  G=V value  H=LM value

## Checksum
`sum(word[0x00..0x3E]) + word[0x3F] + 0xBABA = 0  (mod 0x10000)`
