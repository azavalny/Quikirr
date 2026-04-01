from __future__ import annotations

from pathlib import Path

SOURCE_SHEET = "Source Data"
OUTPUT_SHEET = "Annual Waterfall"
ORIGINAL_SHEET = "Original Data"
QUARTERLY_SHEET = "Quarterly Waterfall"
TOP_CUSTOMERS_SHEET = "Top Customers"
DEFAULT_SOURCE = Path(__file__).resolve().parent.parent / "Source.xlsx"
DEFAULT_OUTPUT = Path(__file__).resolve().parent.parent / "Output.xlsx"

MRR_TO_ARR = 12
