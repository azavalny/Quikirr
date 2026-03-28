---
name: Formula-based waterfall output
overview: Replace hardcoded values in the Annual Waterfall summary tab with live Excel formulas that reference the Original Data tab, so the output workbook is self-updating when source data changes.
todos:
  - id: formula-specs
    content: Add build_formula_specs() that generates formula strings for each summary row, referencing 'Original Data' columns dynamically
    status: completed
  - id: write-formulas
    content: Update write_summary_sheet to handle formula strings (openpyxl writes them natively when value starts with '=')
    status: completed
  - id: cli-flag
    content: Add --formulas / --no-formulas CLI flag, default to formulas mode
    status: completed
  - id: verify-formulas
    content: Regenerate output, open in Excel, and confirm formulas evaluate to same values as before
    status: completed
isProject: false
---

# Formula-based Annual Waterfall

## Approach

Instead of writing computed Python values into summary cells, write **Excel formulas** referencing the `'Original Data'` tab. The Python script still detects the table layout (header row, customer rows, year-end columns) and generates the correct cell references dynamically, but the output cells contain formulas, not numbers.

## Key references (from Source.xlsx layout)

- Header row: **3**, Customer col: **B**, Data rows: **4:103**
- Year-end columns: C (Dec-22), O (Dec-23), AA (Dec-24), AM (Dec-25)
- These are discovered at runtime, not hardcoded

## Data sheet name

Formulas will reference `'Original Data'` (the 2nd tab we already create).

## Formula mapping for each summary row

For each year-end column letter `$COL` and data range `$4:$103`:

**EoP ARR**

```
=SUMIF('Original Data'!$COL$4:$COL$103,"<>"&0)*12
```

**BoP ARR** (year N) = EoP ARR formula of year N-1:

```
=<EoP_ARR_cell_of_prior_year_column>
```

(Direct cell reference to the EoP ARR row in the previous year column on the same sheet.)

**New** -- sum current-year values where prior-year was 0:

```
=SUMPRODUCT(('Original Data'!$PREV$4:$PREV$103=0)*('Original Data'!$CURR$4:$CURR$103<>0)*'Original Data'!$CURR$4:$CURR$103)*12
```

**Upsell** -- sum of (curr-prev) where both > 0 and curr > prev:

```
=SUMPRODUCT(('Original Data'!$PREV$4:$PREV$103>0)*('Original Data'!$CURR$4:$CURR$103>0)*('Original Data'!$CURR$4:$CURR$103>'Original Data'!$PREV$4:$PREV$103)*('Original Data'!$CURR$4:$CURR$103-'Original Data'!$PREV$4:$PREV$103))*12
```

**(Downsell)** -- sum of (curr-prev) where both > 0 and curr < prev (result is negative):

```
=SUMPRODUCT(('Original Data'!$PREV$4:$PREV$103>0)*('Original Data'!$CURR$4:$CURR$103>0)*('Original Data'!$CURR$4:$CURR$103<'Original Data'!$PREV$4:$PREV$103)*('Original Data'!$CURR$4:$CURR$103-'Original Data'!$PREV$4:$PREV$103))*12
```

**(Churn)** -- sum of -prev where prev > 0 and curr = 0:

```
=-SUMPRODUCT(('Original Data'!$PREV$4:$PREV$103>0)*('Original Data'!$CURR$4:$CURR$103=0)*'Original Data'!$PREV$4:$PREV$103)*12
```

**Derived rows** -- formulas referencing other cells on the summary sheet itself:

- **% Growth**: `=EoP/BoP-1`
- **Net New ARR**: `=EoP-BoP` (same-sheet cell refs)
- **New + Upsell**: `=New+Upsell` (same-sheet refs)
- **Downsell + Churn**: `=Downsell+Churn` (same-sheet refs)
- **Gross Retention %**: `=(BoP+Downsell+Churn)/BoP`
- **Loss-Only Retention %**: `=(BoP+Churn)/BoP`
- **Net Retention %**: `=(BoP+Upsell+Downsell+Churn)/BoP`

**Customer counts** -- COUNTIF-based:

- **BoP Customers**: `=COUNTIF('Original Data'!$PREV$4:$PREV$103,">"&0)`
- **New**: `=SUMPRODUCT(($PREV=0)*($CURR<>0))` (count version)
- **(Churn)**: `=SUMPRODUCT(($PREV>0)*($CURR=0))`
- **EoP Customers**: `=COUNTIF('Original Data'!$CURR$4:$CURR$103,">"&0)`

**Logo-derived rows** -- same-sheet cell references:

- **Net Logos Added**: `=New-Churn` (customer rows)
- **Logo % Growth**: `=EoP_Cust/BoP_Cust-1`
- **Gross Logo Retention %**: `=(BoP_Cust-Churn_Cust)/BoP_Cust`

**Average logo sizes** -- same-sheet cell references:

- **Avg. Logo Size**: `=EoP_ARR/EoP_Cust`
- **Avg Logo Size YoY Growth**: `=CurrAvg/PrevAvg-1`
- **Avg. New Logo Size**: `=New_ARR/New_Cust`
- **Avg. Churned Logo Size**: `=ABS(Churn_ARR)/Churn_Cust`
- YoY growth rows: `=CurrAvg/PrevAvg-1`

## Changes to [build_annual_waterfall.py](d:/Quikirr/build_annual_waterfall.py)

- Add a new function `build_formula_specs(years, year_col_letters, data_first_row, data_last_row)` that returns the same row-spec structure but with formula strings instead of values
- Modify `write_summary_sheet` to detect when a cell value is a string starting with `=` and write it as a formula (openpyxl handles this natively -- any string starting with `=` written via `cell.value = "=..."` is stored as a formula)
- Keep the existing `build_summary_values` + `_display_currency` path as a `--no-formulas` fallback
- Add `--formulas` flag (default on) to the CLI

## First-year column (Dec-22)

Same rule: only EoP ARR and EoP Customers get formulas; everything else is hatched.

## Formatting

All existing formatting (fills, fonts, borders, number formats) stays identical -- only the cell values change from numbers to formula strings.