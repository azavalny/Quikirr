from __future__ import annotations

from dataclasses import dataclass


@dataclass
class YearMetrics:
    bop_arr: float | None
    new_arr: float | None
    upsell: float | None
    downsell: float | None
    churn_arr: float | None
    eop_arr: float
    bop_cust: int | None
    new_cust: int | None
    churn_cust: int | None
    eop_cust: int
    needs_prior: bool


def compute_metrics_for_year(
    prev: list[float] | None, curr: list[float]
) -> YearMetrics:
    if prev is None:
        curr_pad = list(curr)
        eop_arr = float(sum(curr_pad))
        eop_cust = sum(1 for x in curr_pad if x > 0)
        return YearMetrics(
            None, None, None, None, None,
            eop_arr, None, None, None, eop_cust, False,
        )

    n = max(len(prev), len(curr))
    prev = list(prev) + [0.0] * (n - len(prev))
    curr = list(curr) + [0.0] * (n - len(curr))

    eop_arr = float(sum(curr))
    eop_cust = sum(1 for x in curr if x > 0)

    new_arr = 0.0
    upsell = 0.0
    downsell = 0.0
    churn_arr = 0.0
    new_cust = 0
    churn_cust = 0

    for p, c in zip(prev, curr):
        if p <= 0 and c > 0:
            new_arr += c
            new_cust += 1
        elif p > 0 and c == 0:
            churn_arr -= p
            churn_cust += 1
        elif p > 0 and c > 0:
            d = c - p
            if d > 0:
                upsell += d
            elif d < 0:
                downsell += d

    bop_arr = sum(prev)
    bop_cust = sum(1 for x in prev if x > 0)

    return YearMetrics(
        bop_arr, new_arr, upsell, downsell, churn_arr,
        eop_arr, bop_cust, new_cust, churn_cust, eop_cust, True,
    )


def assert_bridge(years: list[int], snaps: dict[int, list[float]], tol: float = 0.01) -> None:
    for i in range(1, len(years)):
        prev = snaps[years[i - 1]]
        curr = snaps[years[i]]
        m = compute_metrics_for_year(prev, curr)
        lhs = (
            (m.bop_arr or 0)
            + (m.new_arr or 0)
            + (m.upsell or 0)
            + (m.downsell or 0)
            + (m.churn_arr or 0)
        )
        rhs = m.eop_arr
        if abs(lhs - rhs) > tol:
            raise AssertionError(f"Bridge mismatch {years[i]}: {lhs} vs {rhs}")
