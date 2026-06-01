from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from quikirr.__main__ import run_to_bytes  # noqa: E402


def run_waterfall_analysis(file_bytes: bytes) -> bytes:
    return run_to_bytes(file_bytes)
