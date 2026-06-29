# LAN Tool Dev — Copilot Instructions

## GBE Image Creation Workflow

When the user says anything like "create GBE image", "build NVM image", or "make image for [project]",
follow this exact step-by-step workflow — ask ONE question at a time, wait for the answer, then ask the next.

### Questions to ask (in order):

**Q1 - Project:**
List all folders under GBE_Image_Creator/GBE_Image/ as a numbered list and ask the user to pick one.

**Q2 - Offset:**
Ask: "What is the NVM word offset? (hex, e.g. 0x59)"
Then look up and display all bits at that offset from the XLSM in the project folder.
Show the bits table: Bit | RTL name | Description | Current V | Current LM

**Q3 - Bit:**
Ask: "Which bit do you want to change?"

**Q4 - Mode:**
Ask: "V, LM, or Both?"

**Q5 - New value:**
Ask: "New value for this bit? (0 or 1)"
If the new value equals the current value, warn the user.

**Done check:**
Ask: "Any more changes? (y/n)"
- If y, go back to Q2
- If n, show the full change summary and ask: "Build now? (y/n)"

### Build:
When confirmed, run gbe_build.py or patch binaries directly:
- Read the XLSM full nvm map sheet to find the current word value
- Set the specified bit in the binary at offset*2
- Recalculate checksum at word 0x3F: sum(word[0x00..0x3E]) + csum + 0xBABA = 0 mod 0x10000
- Save change_request_<timestamp>.json and change_report_<timestamp>.txt in the project folder

---

## Folder Structure

```
LAN_Tool_Dev/
├── GBE_Builder/
│   ├── gbe_build.py        <- standalone CLI build script (no API needed)
│   └── RUN.bat             <- double-click to run
├── GBE_Image_Creator/
│   ├── GBE_Image/          <- all NVM projects (XLSMs + binaries)
│   │   ├── Nahum13_ptl_pcd_p_h/
│   │   ├── Nahum11_mtl_m_p/
│   │   └── ... (29 projects total)
│   └── src/app.py          <- GUI tool
└── NVM_AI_Assistant/
    └── src/app.py          <- Claude-powered NVM advisor GUI
```

## XLSM Column Layout (sheet: "full nvm map", data from row 7)
- Col A: LAN Word Offset
- Col B: Bits
- Col C: RTL name
- Col D: C-Spec name
- Col F: Description
- Col G: V value (LAN SW / Consumer)
- Col H: LM value (Non-LAN SW / Corporate)

## Checksum rule
Word 0x3F must satisfy: sum(word[0x00]..word[0x3E]) + word[0x3F] + 0xBABA = 0 (mod 0x10000)
Always recalculate and update word 0x3F after any binary patch.

## General rules
- Always backup .bin files before patching (.bin.bak)
- Warn if new value equals current value (no-op change)
- Projects folder name on target PCs: Eng_GBE_Image (fallback: GBE_Image)
- Keep answers short and focused
- Do NOT ask any questions beyond the 5 defined above
