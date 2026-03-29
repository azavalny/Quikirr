from __future__ import annotations

from openpyxl import Workbook

from ..config import ORIGINAL_SHEET
from ..source import SourceContext
from ..writer import copy_original_sheet


class OriginalDataTab:
    title = ORIGINAL_SHEET

    def build(self, wb: Workbook, ctx: SourceContext) -> None:
        ws = wb.create_sheet(self.title[:31])
        copy_original_sheet(ctx.ws_in, ws)
