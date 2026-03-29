from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .config import DEFAULT_OUTPUT, DEFAULT_SOURCE, SOURCE_SHEET
from .source import SourceContext
from .tabs import TABS
from .verify import assert_bridge


def run(source: Path, output: Path, verify: bool = True) -> None:
    wb_in = load_workbook(source, data_only=True)
    if SOURCE_SHEET not in wb_in.sheetnames:
        raise KeyError(f"Missing sheet {SOURCE_SHEET!r}")
    ws_in = wb_in[SOURCE_SHEET]

    ctx = SourceContext.from_worksheet(ws_in)

    if verify:
        assert_bridge(ctx.years, ctx.snaps)

    wb_out = Workbook()
    for tab in TABS:
        tab.build(wb_out, ctx)

    wb_out.save(output)
    wb_in.close()


def main() -> None:
    p = argparse.ArgumentParser(description="Build Annual Waterfall Excel from Source.xlsx")
    p.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    p.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    p.add_argument("--no-verify", action="store_true")
    args = p.parse_args()
    run(args.source, args.output, verify=not args.no_verify)
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
