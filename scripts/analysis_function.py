"""
MedicStore - Sales reporting

Same analysis as notebooks/analysis.ipynb just organised in functions:
each step of the pipeline is a function, and `main()` calls them in order.

Run it from the project root:

    uv run scripts/analysis_functions.py
"""

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.drawing.image import Image

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
OUTPUT_DIR = DATA_DIR / "output_v2"


# --- Part 1: consolidation ----------------------------------------------


def load_data(data_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Load the two parts of the sales and the mapping table and return them as three DataFrames"""
    sales_part1 = pd.read_csv(data_dir / "Sales_MedicStore_part1.csv", sep=";")
    sales_part2 = pd.read_csv(data_dir / "Sales_MedicStore_part2.csv", sep=";")
    mapping = pd.read_excel(data_dir / "Mapping.xlsx")
    return sales_part1, sales_part2, mapping


def consolidate(
    sales_part1: pd.DataFrame, sales_part2: pd.DataFrame, mapping: pd.DataFrame
) -> pd.DataFrame:
    """Concatenate the two files, merge them with the mapping table and compute total_sales"""
    sales = pd.concat([sales_part1, sales_part2], ignore_index=True)
    df_final = sales.merge(mapping, on="Sub-Category", how="left")

    assert len(df_final) == len(sales), "The merge changed the number of lines"
    assert df_final["Category"].notna().all(), (
        "Some sub-categories are not in the mapping table"
    )

    df_final["total_sales"] = df_final["Quantity"] * df_final["Unit Price"]
    return df_final


# --- Part 2: questions ----------------------------------------------------


def compute_shapes(datasets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Technical Q1: number of lines and columns of each dataset"""
    return pd.DataFrame(
        {
            "dataset": list(datasets),
            "nb_lines": [df.shape[0] for df in datasets.values()],
            "nb_columns": [df.shape[1] for df in datasets.values()],
        }
    )


def top_customers(df: pd.DataFrame, n: int = 3) -> pd.Series:
    """Business Q3: the `n` customers with the most orders (unique Order IDs)."""
    return (
        df.groupby("Customer Name")["Order ID"]
        .nunique()
        .sort_values(ascending=False)
        .head(n)
    )


def answer_business_questions(df: pd.DataFrame) -> dict[str, object]:
    """Business Q1 to Q5: return a dictionary {question: answer}."""
    total_sales_amount = df["total_sales"].sum()
    total_sales_france = df.loc[df["Country"] == "France", "total_sales"].sum()

    return {
        "How many orders do we have": df["Order ID"].nunique(),
        "What is the total sales amount": total_sales_amount,
        "Who are the top 3 customers (order count)": ", ".join(top_customers(df).index),
        "What is the total Sales for France": total_sales_france,
        "How much does the total Sales for France represent compare to the total global Sales": total_sales_france
        / total_sales_amount,
    }


# --- Part 3: graphs -------------------------------------------------------


def save_bar_chart(series: pd.Series, title: str, path: Path, rot: int = 90) -> Path:
    """Save a bar chart of `series` as a PNG file."""
    fig, ax = plt.subplots()
    series.plot.bar(ax=ax, title=title, rot=rot)
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)
    return path


def make_graphs(df: pd.DataFrame, output_dir: Path) -> list[Path]:
    """Create the graphs of Objective 3 and return their paths."""
    orders_by_country = (
        df.groupby("Country")["Order ID"].nunique().sort_values(ascending=False)
    )
    sales_by_category = (
        df.groupby("Category")["total_sales"].sum().sort_values(ascending=False)
    )
    return [
        save_bar_chart(
            orders_by_country,
            "Number of orders per country",
            output_dir / "orders_per_country.png",
        ),
        save_bar_chart(
            sales_by_category,
            "Total sales per category",
            output_dir / "sales_per_category.png",
            rot=0,
        ),
    ]


# --- Part 4: export -------------------------------------------------------


def export_analysis(
    path: Path,
    df_shapes: pd.DataFrame,
    df_describe: pd.DataFrame,
    df_analysis: pd.DataFrame,
    top3: pd.Series,
    graphs: list[Path],
) -> None:
    """Export the tables and the graphs into one Excel file, one sheet each"""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_shapes.to_excel(writer, sheet_name="technical_q1_shapes", index=False)
        df_describe.to_excel(writer, sheet_name="technical_q5_describe")
        df_analysis.to_excel(writer, sheet_name="business_questions", index=False)
        top3.to_excel(writer, sheet_name="business_q3_top3")

        sheet = writer.book.create_sheet("graphs")
        for i, graph in enumerate(graphs):
            sheet.add_image(Image(graph), f"A{1 + 25 * i}")


# --- Pipeline ------------------------------------------------------------------


def main() -> None:
    """Run the whole pipeline: load, consolidate, analyse, export"""
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Part 1
    sales_part1, sales_part2, mapping = load_data(DATA_DIR)
    df_final = consolidate(sales_part1, sales_part2, mapping)
    df_final.to_excel(OUTPUT_DIR / "medicstore_final.xlsx", index=False)

    # Part 2
    df_shapes = compute_shapes(
        {
            "Sales_MedicStore_part1.csv": sales_part1,
            "Sales_MedicStore_part2.csv": sales_part2,
            "Mapping.xlsx": mapping,
            "final dataset": df_final,
        }
    )
    df_describe = df_final.describe(include="all")
    dict_analysis = {
        "How many lines are in Sales_MedicStore_part1.csv": len(sales_part1),
        "How many lines are in Sales_MedicStore_part2.csv": len(sales_part2),
        "How many lines are in Mapping.xlsx": len(mapping),
        **answer_business_questions(df_final),
    }
    df_analysis = pd.DataFrame(
        list(dict_analysis.items()), columns=["Question", "Answer"]
    )

    # Parts 3 and 4
    graphs = make_graphs(df_final, OUTPUT_DIR)
    export_analysis(
        OUTPUT_DIR / "medicstore_analysis.xlsx",
        df_shapes,
        df_describe,
        df_analysis,
        top_customers(df_final),
        graphs,
    )

    print(df_analysis.to_string(index=False))
    print(f"\nAnalysis exported to {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
