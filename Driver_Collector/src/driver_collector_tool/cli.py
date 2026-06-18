import argparse

from .app import launch_gui
from .config import DRIVERS, FAMILIES, FLOWS, build_plan
from .exporter import export_etl_to_csv
from .runner import run_config, start_collection, stop_collection


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Driver collector and ETL-to-CSV exporter")
    subparsers = parser.add_subparsers(dest="command")

    plan_parser = add_selection_args(subparsers.add_parser("plan", help="Show script mapping for a selection"))
    plan_parser.set_defaults(func=cmd_plan)

    config_parser = add_selection_args(subparsers.add_parser("config", help="Run verbose/logging config script when mapped"))
    config_parser.set_defaults(func=cmd_config)

    start_parser = add_selection_args(subparsers.add_parser("start", help="Start collection"))
    start_parser.add_argument("--tag", default="", help="Optional suffix/session tag")
    start_parser.set_defaults(func=cmd_start)

    stop_parser = add_selection_args(subparsers.add_parser("stop", help="Stop collection"))
    stop_parser.set_defaults(func=cmd_stop)

    export_parser = subparsers.add_parser("export-csv", help="Export ETL to CSV using tracerpt.exe")
    export_parser.add_argument("--etl", required=True)
    export_parser.add_argument("--output", required=True)
    export_parser.add_argument("--symbols", help="Optional PDB folder/file or symbol path")
    export_parser.set_defaults(func=cmd_export_csv)

    gui_parser = subparsers.add_parser("gui", help="Open GUI")
    gui_parser.set_defaults(func=cmd_gui)
    return parser


def add_selection_args(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    parser.add_argument("--family", choices=FAMILIES, required=True)
    parser.add_argument("--driver", choices=DRIVERS, required=True)
    parser.add_argument("--flow", choices=FLOWS, required=True)
    return parser


def main(argv=None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    if args.command is None:
        launch_gui()
        return 0
    return args.func(args)


def selected_plan(args):
    return build_plan(args.family, args.driver, args.flow)


def cmd_plan(args) -> int:
    print(selected_plan(args).describe())
    return 0


def cmd_config(args) -> int:
    print(run_config(selected_plan(args)))
    return 0


def cmd_start(args) -> int:
    print(start_collection(selected_plan(args), tag=args.tag))
    return 0


def cmd_stop(args) -> int:
    print(stop_collection(selected_plan(args)))
    return 0


def cmd_export_csv(args) -> int:
    output = export_etl_to_csv(args.etl, args.output, args.symbols)
    print(f"Wrote CSV: {output}")
    return 0


def cmd_gui(args) -> int:
    launch_gui()
    return 0
