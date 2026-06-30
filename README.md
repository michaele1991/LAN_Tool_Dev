# LAN Tool Dev — GBE NVM Builder

Build Intel® GBE NVM images for 29 platforms (Nahum 7–13, MTL, ARL, LNL, PTL, WCL, NVL) from any PC — **no Excel required**.

---

## ▶ START HERE — Steps in Order

Follow these steps in order every time you use this tool:

| Step | What to do | Script / Command |
|------|-----------|-----------------|
| **1** | **Clone the repo** (once) | `git clone https://github.com/michaele1991/LAN_Tool_Dev` |
| **2** | **Set up the Python env** (once per machine) | Windows: `NVM_AI_Assistant\setup.bat` · Linux/macOS: `bash NVM_AI_Assistant/setup.sh` |
| **3** | **Verify the NVM register values** against the GBE RTL spec before building *(see note below)* | Open `GBE_Image_Creator\GBE_Image\<platform>\*.xlsm` |
| **4** | **Build the image** | Windows: `NVM_AI_Assistant\build.bat` · or use the CLI one-liner below |
| **5** | **Collect your output files** | `NVM_AI_Assistant\output\<platform>\` |

> **⚠ Important — Always verify register values against the GBE RTL / C-Spec before building.**
> The `.xlsm` file is the source of truth for default bit values. Before modifying any bit,
> cross-check the register name and bit offset with the official GBE RTL spec or C-Spec document
> to confirm the field name, reset value, and expected behavior. Mistakes here will produce
> a silently incorrect NVM image that may cause link failures or misidentified devices.

---

## Scripts — What Each One Does

```
NVM_AI_Assistant/
├── setup.bat / setup.sh      ← STEP 2 — Run ONCE after cloning. Creates .venv, installs openpyxl.
├── build.bat / build.sh      ← STEP 4 (interactive) — Launches the 5-question wizard in terminal.
└── src/
    ├── build_nvm.py          ← STEP 4 (scripted) — Core engine. Use this for one-liner / AI builds.
    └── wizard.py             ← Called by build.bat — interactive prompts (manual use only).
```

**Use `build_nvm.py` directly** when building from a script, CI, or AI agent — it accepts all
parameters on the command line and produces output with no prompts:

```powershell
# Windows
.venv\Scripts\python.exe src\build_nvm.py --platform Nahum11_lnl_m --step A0 --version 0.8 --variant Both

# With NVM bit override
.venv\Scripts\python.exe src\build_nvm.py --platform Nahum11_lnl_m --step A0 --version 0.8 --variant Both --change 0x58[4]=0

# Multiple changes
.venv\Scripts\python.exe src\build_nvm.py --platform Nahum11_lnl_m --step A0 --version 0.8 --variant Both --change 0x58[4]=0 --change 0x1C[0]=1

# Linux / macOS
.venv/bin/python src/build_nvm.py --platform Nahum11_lnl_m --step A0 --version 0.8 --variant Both
```

**Use `build.bat` / `wizard.py`** when running manually and you want interactive prompts
(platform browser, register search by name, etc.).

---

## The 5 Questions (wizard / AI agent)

| # | Question | Example |
|---|----------|---------|
| Q1 | Which platform? | `Nahum11_lnl_m` |
| Q2 | Silicon step? | `A0` |
| Q3 | Image version? | `0.8` |
| Q4 | Variant — V (Consumer), LM (Corporate), or Both? | `Both` |
| Q5 | NVM bit changes? (offset, bit, value) | `0x58[4]=0` or `none` |

---

## NVM Change Format

```
--change <hex_offset>[<bit>]=<0|1>           # applies to Both variants
--change <hex_offset>[<bit>]=<0|1>:<variant> # V, LM, or Both
```

Examples:
```
--change 0x58[4]=0          # FEXTNVM12 bit 4 → 0 for both V and LM
--change 0x58[4]=0:LM       # Corporate only
--change 0x1C[0]=1:V        # Consumer only
```

---

## Output

```
output/<platform>/
  <ImageName>_<Step>_<Version>_Release_Cons_Prod_NA.bin   ← V variant
  <ImageName>_<Step>_<Version>_Release_Corp_Prod_NA.bin   ← LM variant
  <ImageName>_<Step>_<Version>_Release_Cons_Prod_NA.txt   ← hex dump
  <ImageName>_<Step>_<Version>_Release_Corp_Prod_NA.txt
```

Full paths are printed at end of every build.

---

## Repository Layout

```
LAN_Tool_Dev/
├── .github/
│   └── copilot-instructions.md   ← teaches GitHub Copilot the build workflow
├── NVM_AI_Assistant/
│   ├── build.bat / build.sh      ← interactive build wizard
│   ├── setup.bat / setup.sh      ← one-time env setup
│   ├── requirements.txt          ← openpyxl (+ anthropic optional)
│   ├── output/                   ← generated .bin/.txt files (git-ignored)
│   └── src/
│       ├── build_nvm.py          ← pure Python NVM assembler (core engine)
│       └── wizard.py             ← interactive 5-question CLI
└── GBE_Image_Creator/
    └── GBE_Image/                ← 29 platform folders, each with *.xlsm
        ├── Nahum13_ptl_pcd_p_h/
        ├── Nahum11_lnl_m/
        └── ...
```

---

## How It Works (pure Python, no Excel)

1. Reads the `full nvm map` sheet from the platform's `.xlsm` file via `openpyxl`
2. Assembles 16-bit NVM words from bit-field rows (col A = offset, col B = bits, col G = V value, col H = LM value)
3. Applies any `--change` overrides on top of the base map
4. Writes checksum at word `0x3F`: `(sum(word[0x00..0x3E]) + checksum) & 0xFFFF == 0xBABA`
5. Outputs `.bin` (little-endian 16-bit words) and `.txt` (hex dump, 8 words/line)

---

## Platforms Supported

| Family | Platforms |
|--------|-----------|
| Nahum 7  | CMV-H, KBL-H, KBL-LP, LBG, SPT-H, SPT-LP |
| Nahum 8  | CML-H, CML-LP, CNL-H, CNL-LP, ICL-LP |
| Nahum 9  | EBG-LP, TGL-H, TGL-LP |
| Nahum 10 | ADP-LP, ADP-S, RPL-LP, RPL-S |
| Nahum 11 | ARL-SOC-M, GNR, LNL-M, MTL-M/P, MTL-S |
| Nahum 13 | NVL-PCD-H, NVL-PCH-S, PTL-PCD-P/H, WCL-PCD-N |
| Other    | TTLR, MTL-M/P (legacy) |

---

## AI-Assisted Build (GitHub Copilot)

Open the repo in VS Code and ask Copilot: *"build GBE image"*

Copilot will ask all 5 questions at once, then run a single `build_nvm.py` command — no interactive prompts, minimal tokens.

Requires Python venv already set up (`setup.bat` done).

