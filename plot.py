# /// script
# requires-python = ">=3.10"
# dependencies = ["matplotlib", "openpyxl"]
# ///

"""
Read the local Smithsonian GVP Excel workbook, filter eruptions for 1960-2025,
and save a first-pass pulse-style plot without modifying the original file.

    uv run plot.py
"""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from openpyxl import load_workbook

FILE = "GVP_Eruption_List_Holocene_20260424.xlsx"
PICTURE = "volcanic_pulse_v2.png"

HERE = Path(__file__).parent
DATA = HERE / "data" / FILE
OUT = HERE / "out"


def load_eruption_rows(path):
    """Read the original workbook in read-only mode and return eruptions in 1960-2025."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["Eruption List"]

    header_row = next(sheet.iter_rows(min_row=2, max_row=2, values_only=True))
    header_map = {name: idx for idx, name in enumerate(header_row)}

    required = [
        "Volcano Name",
        "VEI",
        "Start Year",
        "Start Month",
        "Start Day",
    ]
    missing = [name for name in required if name not in header_map]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    rows = []
    parse_errors = 0
    for row in sheet.iter_rows(min_row=3, values_only=True):
        try:
            volcano_name = row[header_map["Volcano Name"]]
            vei_value = row[header_map["VEI"]]
            start_year = row[header_map["Start Year"]]
            start_month = row[header_map["Start Month"]]
            start_day = row[header_map["Start Day"]]

            if start_year is None:
                continue
            year = int(start_year)
            if not (1960 <= year <= 2025):
                continue

            rows.append(
                {
                    "year": year,
                    "volcano": volcano_name,
                    "vei": vei_value,
                    "month": start_month,
                    "day": start_day,
                    "order": len(rows),
                }
            )
        except Exception:
            parse_errors += 1

    return rows, parse_errors


def main():
    eruptions, parse_errors = load_eruption_rows(DATA)
    print(f"excel file: {DATA.name}")
    print(f"filtered eruptions: {len(eruptions)}")
    vei_present = sum(1 for item in eruptions if item["vei"] is not None and str(item["vei"]).strip() not in {"", "NULL", "null", "None"})
    unknown_vei = sum(1 for item in eruptions if item["vei"] is None or str(item["vei"]).strip() in {"", "NULL", "null", "None"})
    print(f"VEI present count: {vei_present}")
    print(f"unknown VEI count: {unknown_vei}")
    print(f"parse_errors: {parse_errors}")

    year_counts = {}
    for item in eruptions:
        year_counts.setdefault(item["year"], 0)
        year_counts[item["year"]] += 1
    max_yearly_count = max(year_counts.values())
    max_years = sorted(year for year, count in year_counts.items() if count == max_yearly_count)
    print(f"maximum yearly eruption count: {max_yearly_count}")
    print(f"year(s) with maximum yearly eruption count: {max_years}")

    fig, ax = plt.subplots(figsize=(12, 7), facecolor="#f8f8f8")
    ax.set_facecolor("#f8f8f8")

    x_values = []
    y_values = []
    sizes = []
    edge_colors = []
    face_colors = []

    for year in sorted(year_counts):
        year_items = [item for item in eruptions if item["year"] == year]
        for idx, item in enumerate(year_items, start=1):
            x = item["year"]
            y = idx
            x_values.append(x)
            y_values.append(y)

            vei = item["vei"]
            if vei is None or str(vei).strip() in {"", "NULL", "null", "None"}:
                sizes.append(80)
                edge_colors.append("#2f2f2f")
                face_colors.append("none")
            else:
                try:
                    numeric_vei = float(vei)
                except (TypeError, ValueError):
                    numeric_vei = 0.0
                sizes.append(max(80, 100 + numeric_vei * 140))
                edge_colors.append("#1d4ed8")
                face_colors.append("#93c5fd")

    ax.scatter(
        x_values,
        y_values,
        s=sizes,
        facecolors=face_colors,
        edgecolors=edge_colors,
        linewidths=1.3,
        alpha=0.9,
    )

    ax.set_xlim(1959, 2026)
    ax.set_xlabel("Start Year")
    ax.set_ylabel("Eruption order within year")
    ax.set_title("Volcanic Pulse — Global Volcanic Eruptions, 1960–2025")
    ax.grid(True, axis="x", linestyle="--", alpha=0.25)

    legend_handles = []
    legend_labels = []
    for vei in [0, 1, 2, 3, 4, 5, 6]:
        size = max(80, 100 + float(vei) * 140)
        handle = plt.Line2D(
            [0], [0],
            marker='o',
            linestyle='',
            markersize=math_sqrt(size / 10),
            markerfacecolor="#93c5fd",
            markeredgecolor="#1d4ed8",
            markeredgewidth=1.3,
        )
        legend_handles.append(handle)
        legend_labels.append(f"VEI {vei}")

    unknown_handle = plt.Line2D(
        [0], [0],
        marker='o',
        linestyle='',
        markersize=5,
        markerfacecolor='none',
        markeredgecolor="#2f2f2f",
        markeredgewidth=1.3,
    )
    legend_handles.append(unknown_handle)
    legend_labels.append("Unknown VEI")

    ax.legend(legend_handles, legend_labels, title="VEI", loc="upper right", frameon=True)
    fig.tight_layout()

    OUT.mkdir(exist_ok=True)
    fig.savefig(OUT / PICTURE, dpi=150)
    print(f"saved {OUT / PICTURE}")
    print(f"expected_count=2229, plotted_count={len(eruptions)}")
    print(f"unknown_vei_count_expected=28, actual_unknown={unknown_vei}")
    plt.close(fig)


if __name__ == "__main__":
    import math

    def math_sqrt(x):
        return math.sqrt(x)

    main()
