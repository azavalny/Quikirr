from __future__ import annotations

from .annual_waterfall import AnnualWaterfallTab
from .base import SheetTab
from .original_data import OriginalDataTab
from .top_customers import TopCustomersTab

TABS: list[SheetTab] = [
    AnnualWaterfallTab(),
    TopCustomersTab(),
    OriginalDataTab(),
]
