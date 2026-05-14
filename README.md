# TIM GL Data — Interactive Pivot Table

Processes the TIM SAP GL export CSV files (semicolon-delimited, UTF-8) and generates a fully self-contained interactive HTML pivot table that works in any browser — no server, no Excel required.

## What it produces

`gl_pivot.html` — a single HTML file (~13 MB for all 6 months) with:

- **45 000+ aggregated rows** from ~88 million raw GL lines (July–December 2025)
- **Drag-and-drop pivot** — move any field between Rows, Columns, and the unused area
- **OWC enrichment** — every GL account is joined to `OWC accounts.xlsx` for `OWC Categoria`, `OWC Codice DB`, and balance-sheet position text
- **SEGMENTS enrichment** (optional) — segment codes joined to `SEGMENTS.xlsx` for `SEGMENTI` (CO / ENT / STAFF / TECH / TELECOM) and `FUNZIONE` drill-down
- **Quick presets**: Importo by Company & Month, Net by GL Account (heatmap), OWC Categoria breakdown, and more

## Requirements

```
Python 3.9+
pandas >= 1.5
openpyxl
```

Install:

```bash
pip install pandas openpyxl
```

## Folder layout expected

```
GL Data/
├── build_gl_pivot_html.py   ← this script
├── OWC accounts.xlsx
├── SEGMENTS.xlsx
└── TIM GL Data/
    ├── 2025_July_*.csv
    ├── 2025_August_*.csv
    ├── 2025_September_*.csv
    ├── 2025_October_*.csv
    ├── 2025_November_*.csv
    └── 2025_December_*.csv
```

> The CSV files are **not** included in this repo (each is 2–3 GB).

## Usage

```bash
# All months, OWC enriched (default) — ~2.5 min
python3 build_gl_pivot_html.py

# Single month only
python3 build_gl_pivot_html.py --months December

# Add SEGMENTI / FUNZIONE dimensions (single month recommended, ~44 MB output)
python3 build_gl_pivot_html.py --months December --with-segment

# Custom output path
python3 build_gl_pivot_html.py --output /path/to/report.html

# Larger chunks if you have RAM to spare (faster)
python3 build_gl_pivot_html.py --chunk 200000
```

## Pivot dimensions available

| Dimension | Source | Example values |
|---|---|---|
| Società | GL CSV | C120, C130 … |
| Anno | GL CSV | 2025 |
| Mese | GL CSV | Jan, Feb … Dec |
| Conto Co.Ge. | GL CSV | GL account code |
| Tipo doc. | GL CSV | DA, RE … |
| Div. Soc. | GL CSV | EUR, USD … |
| OWC Testo | OWC accounts.xlsx | "Crediti verso clienti" |
| OWC Categoria | OWC accounts.xlsx | Trade Receivable, Inventories … |
| OWC Codice DB | OWC accounts.xlsx | `crecoml`, `tfr` … |
| SEGMENTI* | SEGMENTS.xlsx | CO, ENT, STAFF, TECH, TELECOM |
| FUNZIONE* | SEGMENTS.xlsx | CO-CO, STAFF-HRO, TECH-IT … |
| Seg Descrizione* | SEGMENTS.xlsx | Full segment description |

\* Requires `--with-segment`

## Measures

| Measure | Meaning |
|---|---|
| Importo | Dare (debit) amount |
| Imp. Avere | Avere (credit) amount |
| Netto | Importo − Imp. Avere |
| Transazioni | Row count |

## Notes

- The HTML requires an internet connection to load PivotTable.js and jQuery UI from CDN. For fully offline sharing, ask the maintainer for the inline-bundled version.
- Raw CSV files should never be committed — they contain sensitive financial data and are too large for Git.
