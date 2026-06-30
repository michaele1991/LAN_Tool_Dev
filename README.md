# LAN Tool Dev — GBE NVM Builder

Build Intel® GBE NVM images for 29 platforms (Nahum 7–13, MTL, ARL, LNL, PTL, WCL, NVL) from any PC — **no Excel required**.

## Quick Start

### 1. Clone
```bash
git clone https://github.com/michaele1991/LAN_Tool_Dev
cd LAN_Tool_Dev
```

### 2. Setup (once)
| OS | Command |
|----|---------|
| Windows | `setup.bat` |
| Linux / macOS | `bash setup.sh` |

Requires **Python 3.10+**. Only dependency: `openpyxl`.

### 3. Build
| OS | Command |
|----|---------|
| Windows | `build.bat` |
| Linux / macOS | `bash build.sh` |

Or ask **GitHub Copilot**: *"build GBE image"* — it will ask the 5 questions and run the build automatically.

---

## The 5 Questions

When you (or the AI) trigger a build, you are asked:

| # | Question | Example |
|---|----------|---------|
| Q1 | Which platform? | `Nahum13_ptl_pcd_p_h` |
| Q2 | Silicon step? | `A0` |
| Q3 | Image version? | `1.4` |
| Q4 | Variant — V (Consumer), LM (Corporate), or Both? | `Both` |
| Q5 | Confirm build? | `Y` |

Output `.bin` and `.txt` files are written to `output/<platform>/`.

---

## Output File Naming

```
<ImageName>_<Step>_<Version>_Release_<Suffix>.bin
```

| Variant | Suffix |
|---------|--------|
| V  (Consumer)  | `Cons_Prod_NA` |
| LM (Corporate) | `Corp_Prod_NA` |

Example:
```
output/Nahum13_ptl_pcd_p_h/
  GBE_PTL-PCD-P-H_ALL_A0_1.4_Release_Cons_Prod_NA.bin
  GBE_PTL-PCD-P-H_ALL_A0_1.4_Release_Corp_Prod_NA.bin
```

---

## Repository Layout

```
LAN_Tool_Dev/
├── build.bat / build.sh          ← run to start the build wizard
├── setup.bat / setup.sh          ← run once after cloning
├── requirements.txt              ← openpyxl only
│
├── GBE_Builder/
│   ├── build_nvm.py              ← pure Python NVM assembler (core engine)
│   ├── wizard.py                 ← interactive 5-question CLI
│   └── gbe_build.py              ← legacy bit-patch tool
│
├── GBE_Image_Creator/
│   └── GBE_Image/                ← 29 platform folders, each with *.xlsm
│       ├── Nahum13_ptl_pcd_p_h/
│       ├── Nahum11_mtl_m_p/
│       └── ...
│
└── output/                       ← generated files (git-ignored)
```

---

## How It Works (pure Python, no Excel)

1. Reads the `full nvm map` sheet from the platform's `.xlsm` file using `openpyxl`
2. Assembles 16-bit NVM words by combining bit-field rows (col A = offset, col B = bits, col G = V value, col H = LM value)
3. Writes checksum at word `0x3F`: `(sum(word[0x00..0x3E]) + checksum) & 0xFFFF == 0xBABA`
4. Outputs `.bin` (little-endian 16-bit words) and `.txt` (hex dump, 8 words/line)

---

## Platforms Supported

| Family | Platforms |
|--------|-----------|
| Nahum 7 | CMV-H, KBL-H, KBL-LP, LBG, SPT-H, SPT-LP |
| Nahum 8 | CML-H, CML-LP, CNL-H, CNL-LP, ICL-LP |
| Nahum 9 | EBG-LP, TGL-H, TGL-LP |
| Nahum 10 | ADP-LP, ADP-S, RPL-LP, RPL-S |
| Nahum 11 | ARL-SOC-M, GNR, LNL-M, MTL-M/P, MTL-S |
| Nahum 13 | NVL-PCD-H, NVL-PCH-S, PTL-PCD-P/H, WCL-PCD-N |
| Other | TTLR |

---

## Advanced: CLI Usage

```bash
# Build a specific platform directly (no wizard prompts)
.venv/bin/python GBE_Builder/build_nvm.py \
  --platform Nahum13_ptl_pcd_p_h \
  --step A0 \
  --version 1.4 \
  --variant Both

# List all platforms
.venv/bin/python GBE_Builder/build_nvm.py --platform list
```

---

## AI-Assisted Build (GitHub Copilot)

This repo includes `.github/copilot-instructions.md` which teaches Copilot the full build workflow.
When you open the repo in VS Code and ask Copilot:

> *"Build a GBE image"*

Copilot will:
1. Ask the 5 questions
2. Run `build_nvm.py` with your answers
3. Report the output file paths
