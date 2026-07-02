# GBE Image Creator

Intel® GBE NVM Configuration Studio — build, inspect and patch GBE NVM images for 28+ platforms (Nahum 7–13, MTL).

## Quick Start (any Windows PC)

1. **Install Python 3.10+** from https://www.python.org/downloads/  
   ✅ Check **"Add Python to PATH"** during install.

2. **Clone or download** this repository:
   ```
   git clone https://github.com/michaele1991/GBE_Image_Creator
   ```

3. **Double-click `run.bat`**  
   It will automatically create the virtual environment, install dependencies, and launch the app.

4. **Click "Clone NVM Images"** in the toolbar to download all 28 NVM project folders from the official repository.

## Manual launch (after first run)

```
.venv\Scripts\pythonw.exe src\app.py
```

## Requirements

- Windows 10/11
- Python 3.10+
- Microsoft Excel (for VBA Build flow only)
- Git (for "Clone NVM Images" button)

## Notes

- The app loads `.xlsm` files and lets you edit 16-bit NVM words or set/clear individual bits.
- Checksum update follows the existing workflow (word index 0x3F, value 0xBABA).
- GBE_Image project folders are **not** included in the repo — use the "Clone NVM Images" button to fetch them.
