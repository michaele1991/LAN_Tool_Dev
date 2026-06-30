# LAN Tool Dev — Copilot Instructions

## GBE Image Build Workflow  ← PRIMARY WORKFLOW

When the user says anything like "build GBE image", "create NVM image", "make image for [project]",
or "build [platform]", ask these 5 questions **all at once in a single message**:

```
1. Platform?  (run: .venv\Scripts\python.exe src\build_nvm.py 2>&1 to list options if needed)
2. Silicon step?  (e.g. A0, B0)
3. Image version?  (e.g. 1.4)
4. Variant?  Both / V / LM  [default: Both]
5. NVM changes?  e.g. "0x58[4]=0"  or  "none"
```

Once you have all 5 answers, run a **single command** — no interactive prompts:

```powershell
.venv\Scripts\python.exe src\build_nvm.py `
  --platform <platform> `
  --step <step> `
  --version <version> `
  --variant <Both|V|LM> `
  --change <offset>[<bit>]=<value>
```

**NVM change format:** `--change 0x58[4]=0`  (repeat for multiple changes)
Variant suffix optional: `--change 0x58[4]=0:LM`  (default = Both)

### Example — complete one-liner:
```powershell
.venv\Scripts\python.exe src\build_nvm.py --platform Nahum11_lnl_m --step A0 --version 0.8 --variant Both --change 0x58[4]=0
```

### Rules
- Ask all 5 questions together in one message, then build with one command
- Do NOT use wizard.py (it is for manual interactive use only)
- Do NOT require Excel — build_nvm.py is pure Python
- If venv is missing: run `setup.bat` (Windows) or `bash setup.sh` (Linux/macOS) from inside `NVM_AI_Assistant/`
- Output goes to `output/<platform>/`  — path is printed at end of build

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
