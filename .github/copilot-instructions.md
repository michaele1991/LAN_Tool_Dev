# LAN Tool Dev — Copilot Instructions

## GBE Image Build Workflow  ← PRIMARY WORKFLOW

When the user says anything like "build GBE image", "create NVM image", "make image for [project]",
or "build [platform]", follow this exact 5-question workflow.
Ask ONE question at a time. Wait for the answer before asking the next.

### Questions to ask (in order):

**Q1 — Platform**
Run this command and show the numbered list to the user:
```
.venv\Scripts\python.exe src\build_nvm.py --platform list 2>&1
```
If the venv does not exist, tell the user to run `setup.bat` first (or `bash setup.sh` on Linux/macOS).
List all available platform folder names and ask: "Which platform? (enter number)"

**Q2 — Silicon Step**
Ask: "Silicon stepping? (e.g. A0, B0, C0) [default: A0]"

**Q3 — Image Version**
Ask: "Image version? (e.g. 1.4, 2.0) [default: 1.4]"

**Q4 — Variant**
Ask:
```
Which variant(s)?
  [0] Both  — V (Consumer/LAN SW) + LM (Corporate/Non-LAN SW)  ← default
  [1] V     — Consumer / LAN SW only
  [2] LM    — Corporate / Non-LAN SW only
Enter number:
```

**Q5 — NVM Modifications**
Ask: "Any NVM register changes? Specify by register name, bit, and value. Type 'none' to skip."

The user specifies changes as natural language, e.g.:
- `"FEXTNVM12 bit 4 = 0 for Both"`
- `"Device ID = 0x57BA for V"`
- `"0x58 bit 4 = 0"`
- `"none"` — build base map without changes

For each change:
1. Search the xlsm C-Spec/RTL name columns for a case-insensitive partial match to find the offset.
2. Show the offset (e.g. 0x58), the register name, and current V/LM values for that bit.
3. Warn if the new value equals the current value (no-op).
4. After all changes collected, show the full change list.

**Confirm**
Show the complete build summary including NVM changes, then ask: "Build now? [Y/n]"

### Execute the build
When the user confirms, run:
```
.venv\Scripts\python.exe src\wizard.py
```
(wizard handles the full interactive flow including nvm_changes)

Or for headless/scripted builds, run:
```
.venv\Scripts\python.exe src\build_nvm.py \
  --platform <platform_folder_name> \
  --step <step> \
  --version <version> \
  --variant <V|LM|Both>
```
Show the output. If it succeeds, tell the user where the .bin files are.

### Rules
- Do NOT ask any questions beyond the 5 defined above
- Do NOT require Excel or any paid tools — build_nvm.py is pure Python
- If venv is missing, instruct: `setup.bat` (Windows) or `bash setup.sh` (Linux/macOS) — run from inside `NVM_AI_Assistant/`
- Output goes to `NVM_AI_Assistant/output/<platform>/`

---

## NVM Bit-Patch Workflow  ← for patching individual bits in existing images

When the user says "patch", "change bit", "set offset", or asks about a specific NVM bit:

**Q1 - Project:** List folders under `GBE_Image_Creator/GBE_Image/` and ask which one.
**Q2 - Offset:** Ask: "NVM word offset? (hex, e.g. 0x59)" — then show the bits table from the XLSM.
**Q3 - Bit:** Ask: "Which bit to change?"
**Q4 - Mode:** Ask: "V, LM, or Both?"
**Q5 - Value:** Ask: "New value? (0 or 1)" — warn if it matches the current value.

Then ask: "Any more changes? (y/n)" — loop back to Q2 or show summary and run `gbe_build.py`.

---

## Repository Structure

```
LAN_Tool_Dev/
├── .gitignore
├── README.md
├── NVM_AI_Assistant/           ← all NVM tools live here
│   ├── build.bat / build.sh    ← double-click to run the 5-question wizard
│   ├── setup.bat / setup.sh    ← run once after cloning to create venv
│   ├── RUN_APP.bat             ← run the Claude-powered NVM advisor GUI
│   ├── requirements.txt        ← anthropic + openpyxl
│   ├── output/                 ← generated .bin/.txt files (git-ignored)
│   └── src/
│       ├── app.py              ← Claude-powered NVM advisor GUI
│       ├── build_nvm.py        ← pure Python NVM assembler (PRIMARY BUILD TOOL)
│       └── wizard.py           ← interactive 5-question CLI wrapper
└── GBE_Image_Creator/
    └── GBE_Image/              ← platform folders, each with *.xlsm NVM map
        ├── Nahum13_ptl_pcd_p_h/
        ├── Nahum11_mtl_m_p/
        └── ... (29 platforms)
```

## XLSM Column Layout (sheet: "full nvm map", data from row 7)
- Col A: LAN Word Offset (e.g. "0ah", "1Fh")
- Col B: Bits (e.g. "15:0", "7", "11:8")
- Col C: RTL name
- Col D: C-Spec name
- Col F: Description
- Col G: V value  — Consumer / LAN SW
- Col H: LM value — Corporate / Non-LAN SW

## Checksum rule
Word 0x3F must satisfy: `sum(word[0x00..0x3E]) + word[0x3F] + 0xBABA ≡ 0 (mod 0x10000)`
`build_nvm.py` calculates and applies this automatically.

## General rules
- Always backup .bin files before patching (.bin.bak)
- Keep answers short and focused
- Platforms folder: `GBE_Image_Creator/GBE_Image/` (or `Eng_GBE_Image/` on target PCs)
