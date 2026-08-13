# Additions Cleaning & Validation Rules

Date: 2026-07-13
Updated: 2026-07-31
File: `NormalizeAdditions.py`
Status: Implemented

## Purpose

`NormalizeAdditions.py` is a PyQt6 desktop tool that opens a propane-tank
"additions" CSV, normalizes each column, flags problems by color, and saves the
cleaned CSV. This doc records the per-column rules and the color model.

## File format

Input is an **Excel workbook** (`.xlsx/.xlsm/.xls`, first sheet) or a **CSV**;
output is always **CSV**. `read_table` picks the reader by extension and hands
back an all-text DataFrame either way.

Excel is the preferred input. A CSV export only carries the text a cell
*displays*, so a date shown as `Dec-95` loses the real `12/1/1995` behind it and
has to be re-guessed. Reading the workbook gives a true `datetime`, which
`_cell_to_text` renders as ISO `YYYY-MM-DD` before the rules run — no guessing.
CSV is read with `dtype=str, keep_default_na=False`, Excel with `dtype=object,
keep_default_na=False`, so nothing is float-coerced and a literal `N/A` survives
as text instead of silently becoming an empty cell.
## Per-column rules

Rules apply only when the column is present; all other columns pass through.

| Column | Normalization | Flag |
|---|---|---|
| CSC | strip commas/spaces | RED if empty or non-numeric |
| Serv loc | strip commas/spaces | ORANGE if empty or non-numeric |
| Equipment Type | none | RED if not exactly `1` |
| Serial | forced to text (kills `E` sci-notation corruption) | — |
| Type | first A -> `ASME`, first D -> `DOT` (case-insensitive, junk chars ignored) | RED if empty/unrecognizable |
| SizeGallons | strip commas/spaces, drop the decimal part | RED if not a whole number in [57, 5000] |
| SizePounds | strip commas/spaces, drop the decimal part | ORANGE if empty AND (Type=DOT or Manufacturer=WORTHINGTON) |
| Manufacturer | UPPER+trim; `?`/`N/A` -> `UNKNOWN`; >12 chars looked up in shortcut dict | ORANGE if still >12 chars (no shortcut) |
| ManufacturerDate | readable: ASME->`YYYY`, DOT->`MM YY` (month zero-padded); empty or unreadable: ASME->`1900`, DOT->`12 60` | ORANGE if the cell was filled but unreadable |
| Customer Owned | `y*`->`Y`, `n*`->`N` | — |
| Activity | overwrite every row with `1` | — |
| Source | overwrite every row with `Miscellaneous` | — |
| Activity date | overwrite every row with today's date, `MM/DD/YYYY` | — |

Manufacturer shortcut dict lives in `MANUFACTURERS` in the script.

## Color model

Validation flags are computed once (`normalize_dataframe` returns
`{(row, column): 'RED'|'ORANGE'}`).

Cleaned side, highest precedence first:
1. RED - hard error (CSC, Equipment Type, SizeGallons, Type)
2. ORANGE - review warning (Serv loc, SizePounds, Manufacturer, ManufacturerDate)
3. GREEN - value was changed by cleaning and passed validation
4. none

Original side: same RED/ORANGE flag if the cell is flagged (so the problem is
visible in the source), otherwise BLUE where cleaning changed the value.

## Decisions

- CSC and Serv loc flag on both empty and non-numeric ("must have a number"),
  but commas are stripped first so `1,234` is accepted as `1234`.
- Activity / Source / Activity date overwrite every row (batch stamp for new additions).
- Both size columns share one rule and parse as whole numbers. A decimal part is
  dropped, not deleted, so `100.5` -> `100` and never `1005`.
- Empty Manufacturer becomes `UNKNOWN`.
- Type resolves on the first A/D letter; anything else is left as-is and flagged.
- `UNKNOWN` / `N/A` / any unreadable date is treated the same as an empty cell
  and takes the type default, but flags ORANGE so the guess stays visible. Two-
  digit years pivot at 50: `<50` -> `20xx`, `>=50` -> `19xx`.
- Dates are parsed by pulling month/year out of the tokens rather than matching
  one fixed layout, because Excel writes the cell's *displayed* text to CSV: the
  same tank arrives as `Dec-95`, `12/1/1995`, `1995-12-01` or `12 95` depending
  on cell formatting. Any day component is discarded. A leading `'` (Excel's
  text marker) is stripped. Three-part numeric dates read as M/D/Y unless the
  first part is 4 digits (ISO) or greater than 12 (D/M/Y).
- A date landing in **1905** is a typed 4-digit year that Excel read as a day
  serial (`2020` -> `1905-07-12`). Every year 1900-2100 lands in 1905 and
  nothing else does, so `_recover_serial_year` converts the date back to its
  serial and uses that as the year — the code equivalent of Excel's Format
  Cells -> Number. The repair flags ORANGE so a human confirms it. The matching
  2-digit case (which lands in 1900) is deliberately **not** handled: 1900 is
  also the ASME placeholder, so repairing there risks rewriting real data.

## Verification

Logic is a Qt-free module function (`normalize_dataframe`) tested headless
against `sample_additions.csv`, which exercises every rule and flag path.
