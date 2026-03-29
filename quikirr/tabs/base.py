from __future__ import annotations

from typing import Protocol

from openpyxl import Workbook

from ..source import SourceContext


class SheetTab(Protocol):
    title: str

    def build(self, wb: Workbook, ctx: SourceContext) -> None: ...
