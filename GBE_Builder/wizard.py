"""
GBE NVM Build Wizard — interactive 5-question CLI
==================================================
Run this directly: python wizard.py
The AI (GitHub Copilot) will invoke this automatically when asked
to "build a GBE image" or "create an NVM image".

Questions asked:
  Q1  Platform  — which silicon/NVM map
  Q2  Step      — silicon stepping (e.g. A0)
  Q3  Version   — image version string (e.g. 1.4)
  Q4  Variant   — V / LM / Both
  Q5  Confirm   — show summary and build

Requirements: openpyxl  (pip install openpyxl)
"""

import sys
from pathlib import Path

# Allow running from any directory
sys.path.insert(0, str(Path(__file__).parent))
from build_nvm import _find_platforms_root, list_platforms, find_xlsm, read_general_variable, build


# ── helpers ────────────────────────────────────────────────────────────────

def ask(prompt: str, default: str = "") -> str:
    full_prompt = f"{prompt} [{default}]: " if default else f"{prompt}: "
    try:
        answer = input(full_prompt).strip()
    except (EOFError, KeyboardInterrupt):
        print()
        sys.exit(0)
    return answer if answer else default


def choose(prompt: str, options: list[str], default_idx: int = 0) -> int:
    print(f"\n{prompt}")
    for i, opt in enumerate(options):
        marker = " *" if i == default_idx else ""
        print(f"  [{i}] {opt}{marker}")
    while True:
        raw = ask(f"  Enter number", str(default_idx))
        try:
            idx = int(raw)
            if 0 <= idx < len(options):
                return idx
        except ValueError:
            pass
        print(f"  Please enter a number between 0 and {len(options)-1}.")


# ── wizard ─────────────────────────────────────────────────────────────────

def run_wizard():
    print("=" * 60)
    print("  GBE NVM Image Build Wizard")
    print("=" * 60)

    # Locate platforms root
    root = _find_platforms_root()
    if root is None:
        print("\nERROR: Cannot locate GBE_Image platforms folder.")
        print("Make sure you are running from inside the LAN_Tool_Dev repo.")
        sys.exit(1)

    platforms = list_platforms(root)
    if not platforms:
        print(f"\nERROR: No platform folders (with .xlsm) found under:\n  {root}")
        sys.exit(1)

    # ── Q1: Platform ───────────────────────────────────────────────────────
    print("\n── Q1 / 5 ── Platform")
    platform_names = [p.name for p in platforms]
    # Pre-select first platform alphabetically as default
    default_plat = next(
        (i for i, n in enumerate(platform_names) if "ptl_pcd_p_h" in n.lower()), 0
    )
    plat_idx = choose("Which platform / silicon project?",
                      platform_names, default_idx=default_plat)
    platform_folder = platforms[plat_idx]

    # Read defaults from xlsm general variable sheet
    try:
        xlsm = find_xlsm(platform_folder)
        info = read_general_variable(xlsm)
    except Exception:
        info = {}

    default_step = info.get("step", "A0")
    default_ver  = "1.4"

    # ── Q2: Step ───────────────────────────────────────────────────────────
    print("\n── Q2 / 5 ── Silicon Stepping")
    step = ask("Silicon step (e.g. A0, B0, C0)", default_step).upper().strip()
    if not step:
        step = default_step

    # ── Q3: Version ────────────────────────────────────────────────────────
    print("\n── Q3 / 5 ── Image Version")
    version = ask("Image version (e.g. 1.4, 2.0)", default_ver)
    # Validate format
    parts = version.split(".")
    if len(parts) != 2:
        print(f"  Using default version {default_ver!r}.")
        version = default_ver

    # ── Q4: Variant ────────────────────────────────────────────────────────
    print("\n── Q4 / 5 ── Build Variant")
    var_options = [
        "Both  — V (Consumer/LAN SW) + LM (Corporate/Non-LAN SW)",
        "V     — Consumer / LAN SW only",
        "LM    — Corporate / Non-LAN SW only",
    ]
    var_idx = choose("Which variant(s) to build?", var_options, default_idx=0)
    variants = ["V", "LM"] if var_idx == 0 else (["V"] if var_idx == 1 else ["LM"])

    # ── Q5: Confirm ────────────────────────────────────────────────────────
    output_dir = Path(__file__).resolve().parent.parent / "output" / platform_folder.name
    base_name  = (info.get("image file name") or info.get("Project name")
                  or platform_folder.name)
    major, minor = (version.split(".", 1) + ["0"])[:2]
    ver_str = f"{major}.{minor}"

    print("\n── Q5 / 5 ── Confirm Build")
    print()
    print(f"  Platform  : {platform_folder.name}")
    print(f"  XLSM      : {xlsm.name}")
    print(f"  Step      : {step}")
    print(f"  Version   : {ver_str}")
    print(f"  Variant(s): {' + '.join(variants)}")
    print(f"  Output    : {output_dir}")
    print()
    print("  Files that will be created:")
    from build_nvm import VAR_SUFFIX
    for v in variants:
        fname = f"{base_name}_{step}_{ver_str}_Release_{VAR_SUFFIX[v]}"
        print(f"    {fname}.bin")
        print(f"    {fname}.txt")
    print()
    confirm = ask("Build now? [Y/n]", "Y").upper()
    if confirm not in ("Y", "YES", ""):
        print("\nBuild cancelled.")
        sys.exit(0)

    # ── Run build ──────────────────────────────────────────────────────────
    print()
    try:
        outputs = build(platform_folder, step, version, variants, output_dir)
    except Exception as exc:
        print(f"\nERROR during build: {exc}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print(f"\n{'='*60}")
    print("  BUILD COMPLETE")
    print(f"  Output folder: {output_dir}")
    for p in outputs:
        sz = p.stat().st_size
        print(f"    {p.name}  ({sz} bytes)")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    run_wizard()
