from __future__ import annotations

from .annual_waterfall import AnnualWaterfallTab
from .base import SheetTab
from .intermediate_calcs import IntermediateCalculationsTab
from .original_data import OriginalDataTab
from .quarterly_waterfall import QuarterlyWaterfallTab
from .top_customers import TopCustomersTab

TABS: list[SheetTab] = [
    IntermediateCalculationsTab(),
    AnnualWaterfallTab(),
    QuarterlyWaterfallTab(),
    OriginalDataTab(),
    TopCustomersTab(),
]
