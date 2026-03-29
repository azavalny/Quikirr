from __future__ import annotations

from .annual_waterfall import AnnualWaterfallTab
from .base import SheetTab
from .original_data import OriginalDataTab

TABS: list[SheetTab] = [
    AnnualWaterfallTab(),
    OriginalDataTab(),
]
