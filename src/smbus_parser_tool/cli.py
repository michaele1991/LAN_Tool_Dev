import argparse

from .app import launch_gui
from .parser import parse_csv, summarize


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Portable SMBus parser tool")
    parser.add_argument("--input", help="Input SMBus CSV file")
    parser.add_argument("--summary", action="store_true", help="Print a short parse summary")
    parser.add_argument("--gui", action="store_true", help="Open the GUI")
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    if args.gui or not args.input:
        launch_gui()
        return 0
    records = parse_csv(args.input)
    if args.summary:
        print(summarize(records))
    else:
        print(f"Parsed {len(records)} SMBus record(s)")
    return 0
