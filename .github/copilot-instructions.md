# GBE Image Builder — AI Workflow Instructions

## Trigger
When the user says anything like "create GBE image", "build image", "make NVM image",
or "create image for [project]" — execute the workflow below automatically.
Do NOT ask the user to explain anything. Just start with Q1.

---

## Step-by-step workflow

### Q1 — Project
Run this and show as a numbered list:
```powershell
Get-ChildItem "GBE_Image_Creator\GBE_Image" -Directory | Select-Object -ExpandProperty Name | Sort-Object
```
Ask: "Which project? (enter number or name prefix)"

### Q2 — Offset
Ask: "NVM word offset? (hex, e.g. 0x59)"

Then immediately look up and display all bits at that offset:
```python
from openpyxl import load_workbook
from pathlib import Path

project = "<chosen project>"
xlsm = next(Path(f"GBE_Image_Creator/GBE_Image/{project}").glob("*.xlsm"))
wb = load_workbook(str(xlsm), read_only=True, data_only=True, keep_vba=False)
ws = wb["full nvm map"]
offset_target = "<user offset>"  # normalize: strip 0x, add h -> e.g. "59h"
for row in ws.iter_rows(min_row=7, values_only=True):
    if row[0] and row[0].lower().strip() == offset_target:
        print(f"Bit {row[1]:>4}  {row[2]:<45}  V={row[6]}  LM={row[7]}")
wb.close()
```
Show as table: Bit | RTL name | Description | Current V | Current LM

### Q3 — Bit
Ask: "Which bit?"

### Q4 — Mode
Ask: "V, LM, or Both?"

### Q5 — New value
Ask: "New value? (0 or 1)"
If new value == current value → warn: "Already set to that value."

### Done check
Ask: "Any more changes? (y/n)"
- y → go back to Q2 (same project)
- n → show full change summary table, then ask "Build now? (y/n)"

---

## Build — patch the binaries

When user confirms build, execute this Python logic for each change:

```python
import struct, shutil
from pathlib import Path

project = "<chosen project>"
proj_dir = Path(f"GBE_Image_Creator/GBE_Image/{project}")

# Find V (Consumer) and LM (Corporate) output dirs
out_dirs = [d for d in proj_dir.iterdir() if d.is_dir()]
cons_dir = next((d for d in out_dirs if any(x in d.name.lower() for x in ("cons","_v_","lan"))), out_dirs[0])
corp_dir = next((d for d in out_dirs if any(x in d.name.lower() for x in ("corp","lm","non"))), out_dirs[-1])

def patch(bin_path, offset_int, bit_int, new_bit):
    data = bytearray(bin_path.read_bytes())
    pos = offset_int * 2
    old = struct.unpack_from("<H", data, pos)[0]
    new = (old | (1 << bit_int)) if new_bit else (old & ~(1 << bit_int))
    struct.pack_into("<H", data, pos, new)
    # recalculate checksum at 0x3F
    cpos = 0x3F * 2
    struct.pack_into("<H", data, cpos, 0)
    s = sum(struct.unpack_from("<H", data, i)[0] for i in range(0, cpos, 2))
    csum = (0xBABA - s) & 0xFFFF
    struct.pack_into("<H", data, cpos, csum)
    bin_path.write_bytes(data)
    return old, new, csum

# For each change:
#   if mode in (V, Both)  -> patch cons_dir/*.bin
#   if mode in (LM, Both) -> patch corp_dir/*.bin
# Also patch the matching *.txt file (same word positions, space-separated hex)
```

Always backup `.bin` → `.bin.bak` before patching.

---

## Checksum rule
`sum(word[0x00] .. word[0x3E]) + word[0x3F] + 0xBABA = 0  (mod 0x10000)`

---

## XLSM column layout (sheet "full nvm map", data starts row 7)
| Col | Content |
|-----|---------|
| A   | Word offset (e.g. "59h") |
| B   | Bit(s) |
| C   | RTL name |
| D   | C-Spec name |
| F   | Description |
| G   | V value (Consumer / LAN SW) |
| H   | LM value (Corporate / Non-LAN SW) |

---

## Output files
Save in the project folder after build:
- `change_request_<YYYYMMDD_HHMMSS>.json` — machine-readable change list
- `change_report_<YYYYMMDD_HHMMSS>.txt`  — human-readable diff report

---

## Rules
- Ask ONE question at a time. Wait for the answer before the next.
- Do NOT ask anything beyond the 5 questions above.
- Always show current V and LM values before asking for new value.
- Warn if no-op (new value == current value).
- Projects root folder: `GBE_Image_Creator/GBE_Image/` (relative to repo root).
