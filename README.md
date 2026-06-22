# GBE_Image

GUI app to edit GBE NVM .bin images (word/bit edits) and update checksum.

## Run

- From the repo root, run: python src/app.py

## Notes

- The app loads .bin files and lets you edit 16-bit words or set/clear bits.
- Checksum update follows the existing workflow (word index 0x3F, value 0xBABA).
- Optional TXT checksum update uses line 7 in the .txt file, matching the existing script.
