# zhu_course06

Course work of week 6: MedicStore sales reporting.

MedicStore sells medical products across Europe. The team used to consolidate the sales data by hand in Excel. This project automates it in Python: merge the files, answer the questions from the Data Team and the operational team, and export the results and graphs to Excel for the board presentation.

## Installation

```bash
git clone https://github.com/emlyon-hids-python-2026/zhu_course06.git
cd zhu_course06
uv sync
```

## Data

The data is not tracked by git. Download it into `data/`:

```bash
cd scripts && uv run download_data.py && cd ..
```

| File | Content |
|---|---|
| `Sales_MedicStore_part1.csv`, `Sales_MedicStore_part2.csv` | Sales lines (`;` separator) |
| `Mapping.xlsx` | `Sub-Category` → `Category` |


## Usage

| Version | Run | Output |
|---|---|---|
| Notebook | Open `notebooks/analysis.ipynb`, *Run All* | `data/output_v1/` |
| Script with functions | `uv run scripts/analysis_functions.py` | `data/output_v2/` |

Each output folder contains the consolidated dataset, an analysis workbook (one sheet per table plus the
graphs) and the graphs as PNG.

## Pipeline

1. Load both sales files and stack them (`pd.concat`).
2. Left merge with the mapping on `Sub-Category`, with checks that no line is lost.
3. Compute `total_sales = Quantity * Unit Price`.
4. Answer the technical and business questions.
5. Plot orders per country and sales per category.
6. Export everything to Excel.

## Key results

| Question | Answer |
|---|---|
| Number of orders | 3,744 |
| Total sales | 2,290,823 |
| Top customers (order count) | Frank Olsen (13), then a tie at 12: Paul Stevenson, Yoseph Carroll, Natalie Fritzler |
| Sales for France | 650,730 (28.4 % of total) |

## Project structure

- `data/`: datasets and outputs (not tracked by git)
- `notebooks/`: Jupyter notebook version
- `scripts/`: download script and script with functions
- `src/`: Python package

## Author

ZHU Lingxuan <lingxuan.zhu@edu.em-lyon.com>
