from __future__ import annotations

from pathlib import Path

SOURCE_SHEET = "Source Data"
OUTPUT_SHEET = "Annual Waterfall"
ORIGINAL_SHEET = "Original Data"
DEFAULT_SOURCE = Path(__file__).resolve().parent.parent / "Source.xlsx"
DEFAULT_OUTPUT = Path(__file__).resolve().parent.parent / "Annual_Waterfall_Output.xlsx"

MRR_TO_ARR = 12
