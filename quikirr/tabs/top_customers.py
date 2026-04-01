from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border
from openpyxl.utils import get_column_letter

from ..config import MRR_TO_ARR, ORIGINAL_SHEET, TOP_CUSTOMERS_SHEET
from ..source import (
    SourceContext,
    _is_customer_label_row,
    _parse_header_date,
    _to_float,
    dec_label,
    find_table_bounds,
    year_end_column_index,
)
from ..styles import (
    ACCOUNTING_FMT,
    FILL_BLUE,
    FILL_GREEN_COND,
    FILL_RED_COND,
    FONT_BOLD,
    FONT_WHITE_BOLD,
    MULT_FMT,
    PCT_FMT,
    PCT_PARENS_NEG_FMT,
    THIN,
)

TOP_N = 10
DATA_START = 4
SN = f"'{ORIGINAL_SHEET}'"


@dataclass
class _Customer:
    name: str
    src_row: int
    last_arr: float = 0.0
    since: str = ""


def _read_customers(ctx: SourceContext) -> list[_Customer]:
    ws = ctx.ws_in
    header_row = ctx.header_row
    cust_col = ctx.cust_col
    y2c = year_end_column_index(header_row, ws)
    last_year = max(ctx.years)
    last_col = y2c[last_year]

    _, _, first_dc = find_table_bounds(ws)
    date_cols: list[tuple[int, date]] = []
    for c in range(first_dc, ws.max_column + 1):
        d = _parse_header_date(ws.cell(header_row, c).value)
        if d is not None:
            date_cols.append((c, d))
    date_cols.sort(key=lambda t: t[1])

    customers: list[_Customer] = []
    for r in range(ctx.data_start_row, ctx.data_end_row + 1):
        label = ws.cell(r, cust_col).value
        if not _is_customer_label_row(label):
            continue
        cust = _Customer(
            name=str(label).strip(),
            src_row=r,
            last_arr=_to_float(ws.cell(r, last_col).value) * MRR_TO_ARR,
        )
        for col_idx, d in date_cols:
            if _to_float(ws.cell(r, col_idx).value) > 0:
                cust.since = f"{d:%b}-{d:%y}"
                break
        customers.append(cust)
    return customers


def _col_layout(n: int) -> dict[str, Any]:
    """Return column index info (1-based) for N years."""
    arr_start = 3
    cagr = arr_start + n
    pct_arr = cagr + 1
    pct_cum = pct_arr + 1
    dg_start = pct_cum + 1
    pct_of_g = dg_start + (n - 1)
    mult = pct_of_g + 1
    pg_start = mult + 1
    since = pg_start + (n - 1)
    total = since
    return dict(
        arr_start=arr_start, cagr=cagr, pct_arr=pct_arr, pct_cum=pct_cum,
        dg_start=dg_start, pct_of_g=pct_of_g, mult=mult,
        pg_start=pg_start, since=since, total=total,
    )


def _src_cell(src_cl: str, src_row: int) -> str:
    return f"{SN}!${src_cl}${src_row}"


def _write_headers(ws, years: list[int], layout: dict, total_cols: int) -> None:
    n = len(years)
    last_col_letter = get_column_letter(total_cols)
    first_lbl = dec_label(years[0])
    last_lbl = dec_label(years[-1])

    ws.merge_cells(f"A1:{last_col_letter}1")
    title = ws["A1"]
    title.value = f"Top Customers as of {last_lbl}"
    title.fill = FILL_BLUE
    title.font = FONT_WHITE_BOLD
    title.alignment = Alignment(horizontal="center")

    for c in range(1, total_cols + 1):
        for r in (2, 3):
            cell = ws.cell(r, c)
            cell.fill = FILL_BLUE
            cell.font = FONT_WHITE_BOLD

    ws.cell(2, 1).value = "($ Actuals)"
    ws.cell(2, 1).alignment = Alignment(horizontal="left")

    _merge_super(ws, 2, layout["arr_start"], layout["pct_cum"], "ARR")
    _merge_super(ws, 2, layout["dg_start"], layout["mult"], "$ Growth")
    _merge_super(ws, 2, layout["pg_start"], layout["since"] - 1, "% Growth")
    ws.cell(2, layout["since"]).value = ""

    col = 1
    ws.cell(3, col).value = "Rank"; col += 1
    ws.cell(3, col).value = "Customer"; col += 1

    for y in years:
        ws.cell(3, col).value = dec_label(y)
        ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1

    ws.cell(3, col).value = f"{first_lbl} - {last_lbl} CAGR"
    ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1
    ws.cell(3, col).value = f"% of {last_lbl} ARR"
    ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1
    ws.cell(3, col).value = "% Total Cum."
    ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1

    for i in range(n - 1):
        ws.cell(3, col).value = f"{dec_label(years[i])} - {dec_label(years[i+1])}"
        ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1
    ws.cell(3, col).value = f"% of {dec_label(years[-2])} - {last_lbl} Growth"
    ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1
    ws.cell(3, col).value = "Mult. of Initial"
    ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1

    for i in range(n - 1):
        ws.cell(3, col).value = f"{dec_label(years[i])} - {dec_label(years[i+1])}"
        ws.cell(3, col).alignment = Alignment(horizontal="right"); col += 1

    ws.cell(3, col).value = "Customer Since"
    ws.cell(3, col).alignment = Alignment(horizontal="right")


def _merge_super(ws, row: int, start: int, end: int, value: str) -> None:
    if start == end:
        ws.cell(row, start).value = value
        ws.cell(row, start).alignment = Alignment(horizontal="center")
        return
    ws.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
    ws.cell(row, start).value = value
    ws.cell(row, start).alignment = Alignment(horizontal="center")


def _col_formats(layout: dict, n: int) -> list[tuple[int, str]]:
    """Return (1-based col, format_string) pairs."""
    fmts: list[tuple[int, str]] = []
    for i in range(n):
        fmts.append((layout["arr_start"] + i, ACCOUNTING_FMT))
    fmts.append((layout["cagr"], PCT_FMT))
    fmts.append((layout["pct_arr"], PCT_FMT))
    fmts.append((layout["pct_cum"], PCT_FMT))
    for i in range(n - 1):
        fmts.append((layout["dg_start"] + i, ACCOUNTING_FMT))
    fmts.append((layout["pct_of_g"], PCT_FMT))
    fmts.append((layout["mult"], MULT_FMT))
    for i in range(n - 1):
        fmts.append((layout["pg_start"] + i, PCT_PARENS_NEG_FMT))
    return fmts


def _apply_fmt(ws, row: int, fmts: list[tuple[int, str]]) -> None:
    for col, fmt in fmts:
        ws.cell(row, col).number_format = fmt


def _set_row_style(ws, row: int, total_cols: int, bold: bool = False, border: bool = False) -> None:
    for c in range(1, total_cols + 1):
        cell = ws.cell(row, c)
        cell.alignment = Alignment(horizontal="right" if c > 2 else "left")
        if bold:
            cell.font = FONT_BOLD
        if border:
            cell.border = Border(top=THIN, bottom=THIN)


def _write_customer_formulas(
    ws, row: int, rank: int, cust: _Customer,
    years: list[int], src_cols: dict[int, str],
    layout: dict, total_row: int,
    cust_cl: str, rank_cl: str, ds: int, de: int,
) -> None:
    n = len(years)
    span = n - 1
    ws.cell(row, 1).value = rank
    ws.cell(row, 2).value = (
        f'=IFERROR(INDEX({SN}!${cust_cl}${ds}:${cust_cl}${de},'
        f'MATCH($A{row},{SN}!${rank_cl}${ds}:${rank_cl}${de},0),1),"")'
    )

    acl: list[str] = []
    for i, y in enumerate(years):
        col = layout["arr_start"] + i
        cl = get_column_letter(col)
        acl.append(cl)
        ws.cell(row, col).value = (
            f'=IFERROR(INDEX({SN}!${src_cols[y]}${ds}:${src_cols[y]}${de},'
            f'MATCH($A{row},{SN}!${rank_cl}${ds}:${rank_cl}${de},0),1)*{MRR_TO_ARR},0)'
        )

    first, last = acl[0], acl[-1]
    ws.cell(row, layout["cagr"]).value = (
        f"=IF({first}{row}=0,0,({last}{row}/{first}{row})^(1/{span})-1)"
    )

    pct_cl = get_column_letter(layout["pct_arr"])
    ws.cell(row, layout["pct_arr"]).value = (
        f"=IF({last}{total_row}=0,0,{last}{row}/{last}{total_row})"
    )
    ws.cell(row, layout["pct_cum"]).value = (
        f"=SUM({pct_cl}${DATA_START}:{pct_cl}{row})"
    )

    dgcl: list[str] = []
    for i in range(n - 1):
        col = layout["dg_start"] + i
        cl = get_column_letter(col)
        dgcl.append(cl)
        ws.cell(row, col).value = f"={acl[i+1]}{row}-{acl[i]}{row}"

    last_dg = dgcl[-1]
    ws.cell(row, layout["pct_of_g"]).value = (
        f"=IF({last_dg}{total_row}=0,0,{last_dg}{row}/{last_dg}{total_row})"
    )
    ws.cell(row, layout["mult"]).value = (
        f"=IF({first}{row}=0,0,{last}{row}/{first}{row})"
    )

    for i in range(n - 1):
        col = layout["pg_start"] + i
        ws.cell(row, col).value = f"=IF({acl[i]}{row}=0,0,{acl[i+1]}{row}/{acl[i]}{row}-1)"

    ws.cell(row, layout["since"]).value = cust.since
    ws.cell(row, layout["since"]).alignment = Alignment(horizontal="right")


def _write_sum_row(
    ws, row: int, label: str, sum_start: int, sum_end: int,
    years: list[int], layout: dict, total_row: int,
    bold: bool = False, border: bool = False, total_cols: int = 0,
) -> None:
    n = len(years)
    span = n - 1
    ws.cell(row, 2).value = label

    acl: list[str] = []
    for i in range(n):
        col = layout["arr_start"] + i
        cl = get_column_letter(col)
        acl.append(cl)
        ws.cell(row, col).value = f"=SUM({cl}{sum_start}:{cl}{sum_end})"

    first, last = acl[0], acl[-1]
    ws.cell(row, layout["cagr"]).value = (
        f"=IF({first}{row}=0,0,({last}{row}/{first}{row})^(1/{span})-1)"
    )
    pct_cl = get_column_letter(layout["pct_arr"])
    ws.cell(row, layout["pct_arr"]).value = (
        f"=IF({last}{total_row}=0,0,{last}{row}/{last}{total_row})"
    )
    ws.cell(row, layout["pct_cum"]).value = (
        f"=SUM({pct_cl}{sum_start}:{pct_cl}{sum_end})"
    )

    dgcl: list[str] = []
    for i in range(n - 1):
        col = layout["dg_start"] + i
        cl = get_column_letter(col)
        dgcl.append(cl)
        ws.cell(row, col).value = f"=SUM({cl}{sum_start}:{cl}{sum_end})"

    last_dg = dgcl[-1]
    ws.cell(row, layout["pct_of_g"]).value = (
        f"=IF({last_dg}{total_row}=0,0,{last_dg}{row}/{last_dg}{total_row})"
    )
    ws.cell(row, layout["mult"]).value = (
        f"=IF({first}{row}=0,0,{last}{row}/{first}{row})"
    )

    for i in range(n - 1):
        col = layout["pg_start"] + i
        ws.cell(row, col).value = f"=IF({acl[i]}{row}=0,0,{acl[i+1]}{row}/{acl[i]}{row}-1)"

    _set_row_style(ws, row, total_cols, bold=bold, border=border)


def _write_total_customers_row(
    ws, row: int, years: list[int], src_cols: dict[int, str],
    layout: dict, data_start_row: int, data_end_row: int,
    total_cols: int,
) -> None:
    n = len(years)
    span = n - 1
    ws.cell(row, 2).value = "Total Customers"

    acl: list[str] = []
    for i, y in enumerate(years):
        col = layout["arr_start"] + i
        cl = get_column_letter(col)
        acl.append(cl)
        rng = f"{SN}!${src_cols[y]}${data_start_row}:${src_cols[y]}${data_end_row}"
        ws.cell(row, col).value = f"=SUM({rng})*{MRR_TO_ARR}"

    first, last = acl[0], acl[-1]
    ws.cell(row, layout["cagr"]).value = (
        f"=IF({first}{row}=0,0,({last}{row}/{first}{row})^(1/{span})-1)"
    )
    ws.cell(row, layout["pct_arr"]).value = (
        f"=IF({last}{row}=0,0,{last}{row}/{last}{row})"
    )
    ws.cell(row, layout["pct_cum"]).value = (
        f"=IF({last}{row}=0,0,{last}{row}/{last}{row})"
    )

    dgcl: list[str] = []
    for i in range(n - 1):
        col = layout["dg_start"] + i
        cl = get_column_letter(col)
        dgcl.append(cl)
        ws.cell(row, col).value = f"={acl[i+1]}{row}-{acl[i]}{row}"

    last_dg = dgcl[-1]
    ws.cell(row, layout["pct_of_g"]).value = (
        f"=IF({last_dg}{row}=0,0,{last_dg}{row}/{last_dg}{row})"
    )
    ws.cell(row, layout["mult"]).value = (
        f"=IF({first}{row}=0,0,{last}{row}/{first}{row})"
    )

    for i in range(n - 1):
        col = layout["pg_start"] + i
        ws.cell(row, col).value = f"=IF({acl[i]}{row}=0,0,{acl[i+1]}{row}/{acl[i]}{row}-1)"

    _set_row_style(ws, row, total_cols, bold=True, border=True)


def _write_other_row(
    ws, row: int, label: str, top10_row: int, total_row: int,
    years: list[int], layout: dict, total_cols: int,
) -> None:
    n = len(years)
    span = n - 1
    ws.cell(row, 2).value = label

    acl: list[str] = []
    for i in range(n):
        col = layout["arr_start"] + i
        cl = get_column_letter(col)
        acl.append(cl)
        ws.cell(row, col).value = f"={cl}{total_row}-{cl}{top10_row}"

    first, last = acl[0], acl[-1]
    ws.cell(row, layout["cagr"]).value = (
        f"=IF({first}{row}=0,0,({last}{row}/{first}{row})^(1/{span})-1)"
    )
    ws.cell(row, layout["pct_arr"]).value = (
        f"=IF({last}{total_row}=0,0,{last}{row}/{last}{total_row})"
    )
    ws.cell(row, layout["pct_cum"]).value = ""

    dgcl: list[str] = []
    for i in range(n - 1):
        col = layout["dg_start"] + i
        cl = get_column_letter(col)
        dgcl.append(cl)
        ws.cell(row, col).value = f"={cl}{total_row}-{cl}{top10_row}"

    last_dg = dgcl[-1]
    ws.cell(row, layout["pct_of_g"]).value = (
        f"=IF({last_dg}{total_row}=0,0,{last_dg}{row}/{last_dg}{total_row})"
    )
    ws.cell(row, layout["mult"]).value = (
        f"=IF({first}{row}=0,0,{last}{row}/{first}{row})"
    )

    for i in range(n - 1):
        col = layout["pg_start"] + i
        ws.cell(row, col).value = f"=IF({acl[i]}{row}=0,0,{acl[i+1]}{row}/{acl[i]}{row}-1)"

    _set_row_style(ws, row, total_cols)


def _add_conditional_formatting(ws, layout: dict, n: int, first_row: int, last_row: int) -> None:
    for i in range(n - 1):
        cl = get_column_letter(layout["dg_start"] + i)
        rng = f"{cl}{first_row}:{cl}{last_row}"
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_GREEN_COND)
        )
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED_COND)
        )


class TopCustomersTab:
    title = TOP_CUSTOMERS_SHEET

    def build(self, wb: Workbook, ctx: SourceContext) -> None:
        years = ctx.years
        n = len(years)
        layout = _col_layout(n)
        total_cols = layout["since"]

        y2c = year_end_column_index(ctx.header_row, ctx.ws_in)
        src_cols = {y: get_column_letter(c) for y, c in y2c.items()}

        customers = _read_customers(ctx)
        ranked = sorted(customers, key=lambda c: c.last_arr, reverse=True)
        ranked = [c for c in ranked if c.last_arr > 10]
        top10 = ranked[:TOP_N]
        num_top = len(top10)
        last_src_col = src_cols[max(years)]
        other_label = (
            f'="Other ("&COUNTIFS({SN}!${last_src_col}${ctx.data_start_row}:'
            f'${last_src_col}${ctx.data_end_row},">"&10/{MRR_TO_ARR})'
            f'-{num_top}&" Customers)"'
        )

        ws_orig = wb[ORIGINAL_SHEET]
        rank_col_idx = ctx.ws_in.max_column + 1
        rank_cl = get_column_letter(rank_col_idx)
        cust_cl = get_column_letter(ctx.cust_col)
        ds = ctx.data_start_row
        de = ctx.data_end_row
        for r in range(ds, de + 1):
            ws_orig.cell(r, rank_col_idx).value = (
                f"=IF(${last_src_col}{r}*{MRR_TO_ARR}>10,"
                f"RANK(${last_src_col}{r},${last_src_col}${ds}:${last_src_col}${de},0)"
                f"+COUNTIF(${last_src_col}${ds}:${last_src_col}{r},${last_src_col}{r})-1,"
                f'"N/A")'
            )

        top10_end = DATA_START + min(TOP_N, len(top10)) - 1
        top10_total_row = top10_end + 1
        other_row = top10_total_row + 2
        total_row = other_row + 1
        top5_row = total_row + 2
        top10_row2 = top5_row + 1

        ws = wb.create_sheet()
        ws.title = self.title[:31]

        _write_headers(ws, years, layout, total_cols)
        fmts = _col_formats(layout, n)

        for rank_i, cust in enumerate(top10, start=1):
            row = DATA_START + rank_i - 1
            _write_customer_formulas(
                ws, row, rank_i, cust, years, src_cols, layout, total_row,
                cust_cl, rank_cl, ds, de,
            )
            _apply_fmt(ws, row, fmts)
            _set_row_style(ws, row, total_cols)

        _write_sum_row(
            ws, top10_total_row, "Total Top 10",
            DATA_START, top10_end, years, layout, total_row,
            bold=True, border=True, total_cols=total_cols,
        )
        _apply_fmt(ws, top10_total_row, fmts)

        _write_other_row(
            ws, other_row, other_label,
            top10_total_row, total_row, years, layout, total_cols,
        )
        _apply_fmt(ws, other_row, fmts)

        _write_total_customers_row(
            ws, total_row, years, src_cols, layout,
            ctx.data_start_row, ctx.data_end_row, total_cols,
        )
        _apply_fmt(ws, total_row, fmts)

        top5_end = DATA_START + min(5, len(top10)) - 1
        _write_sum_row(
            ws, top5_row, "Total Top 5",
            DATA_START, top5_end, years, layout, total_row,
            total_cols=total_cols,
        )
        _apply_fmt(ws, top5_row, fmts)

        _write_sum_row(
            ws, top10_row2, "Total Top 10",
            DATA_START, top10_end, years, layout, total_row,
            total_cols=total_cols,
        )
        _apply_fmt(ws, top10_row2, fmts)

        _add_conditional_formatting(ws, layout, n, DATA_START, top10_row2)

        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 18
        for j in range(3, total_cols + 1):
            ws.column_dimensions[get_column_letter(j)].width = 14
