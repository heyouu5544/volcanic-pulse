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
import math

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from openpyxl import load_workbook

FILE = "GVP_Eruption_List_Holocene_20260424.xlsx"
PICTURE = "volcanic_pulse_v5.png"

HERE = Path(__file__).parent
DATA = HERE / "data" / FILE
OUT = HERE / "out"


def load_eruption_rows(path):
    """Read the original workbook in read-only mode and return eruptions in 2000-2025."""
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
            if not (2000 <= year <= 2025):
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

    fig, ax = plt.subplots(figsize=(10, 16), facecolor="#efe9e2")
    ax.set_facecolor("#efe9e2")

    def pulse_color(vei_value):
        if vei_value is None or str(vei_value).strip() in {"", "NULL", "null", "None"}:
            return "#7b7b7b", "none"
        try:
            vei = float(vei_value)
        except (TypeError, ValueError):
            vei = 0.0

        if vei <= 1:
            return "#b7a890", "#b7a890"
        if vei <= 3:
            return "#e58a3a", "#f1a85d"
        if vei == 4:
            return "#d9572a", "#ed6c42"
        if vei <= 6:
            return "#a31d1d", "#d72626"
        return "#690d0d", "#8b1010"

    def marker_size(vei_value):
        if vei_value is None or str(vei_value).strip() in {"", "NULL", "null", "None"}:
            return 90
        try:
            vei = float(vei_value)
        except (TypeError, ValueError):
            vei = 0.0
        return 90 + min(320, max(0.0, vei) * 70)

    for year in sorted(year_counts):
        year_items = [item for item in eruptions if item["year"] == year]
        for idx, item in enumerate(year_items, start=1):
            x = item["year"]
            y = idx
            vei = item["vei"]

            if vei is None or str(vei).strip() in {"", "NULL", "null", "None"}:
                ax.scatter(
                    x, y,
                    s=90,
                    facecolors="none",
                    edgecolors="#6b7280",
                    linewidths=1.3,
                    alpha=0.9,
                    zorder=3,
                )
                continue

            try:
                numeric_vei = float(vei)
            except (TypeError, ValueError):
                numeric_vei = 0.0

            core_edge, core_fill = pulse_color(vei)
            ax.scatter(
                x, y,
                s=marker_size(numeric_vei),
                facecolors=core_fill,
                edgecolors=core_edge,
                linewidths=0.9,
                alpha=0.82,
                zorder=4,
            )

    ax.set_xlim(1999, 2026)
    ax.set_ylim(0, max(year_counts.values()) + 2)
    ax.set_xlabel("Start Year", color="#4b4b4b", fontsize=10)
    ax.set_ylabel("Eruption order within year", color="#4b4b4b", fontsize=10)
    ax.set_title(
        "Volcanic Pulse\nGlobal Volcanic Eruptions, 2000–2025",
        loc="left",
        fontsize=18,
        fontweight="bold",
        color="#3d2d2a",
        pad=18,
    )
    ax.grid(True, axis="x", linestyle="-", linewidth=0.5, alpha=0.08)
    ax.grid(False, axis="y")

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#8a857d")
    ax.spines["bottom"].set_color("#8a857d")
    ax.tick_params(axis="both", colors="#4b4b4b", labelsize=9)

    legend_handles = []
    legend_labels = []
    for label, fill_color, edge_color in [
        ("Low VEI", "#b7a890", "#b7a890"),
        ("Medium VEI", "#e58a3a", "#e58a3a"),
        ("High VEI", "#d9572a", "#a31d1d"),
        ("Unknown", "none", "#6b7280"),
    ]:
        handle = plt.Line2D(
            [0], [0],
            marker='o',
            linestyle='',
            markersize=8,
            markerfacecolor=fill_color,
            markeredgecolor=edge_color,
            markeredgewidth=1.2,
        )
        legend_handles.append(handle)
        legend_labels.append(label)

    leg = ax.legend(legend_handles, legend_labels, title="VEI", loc="upper right", bbox_to_anchor=(1.02, 1.0), frameon=False)
    leg.get_title().set_color("#4b4b4b")
    for txt in leg.get_texts():
        txt.set_color("#4b4b4b")

    fig.tight_layout()

    OUT.mkdir(exist_ok=True)
    fig.savefig(OUT / PICTURE, dpi=180, bbox_inches="tight")
    print(f"saved {OUT / PICTURE}")
    print(f"expected_count={len(eruptions)}, plotted_count={len(eruptions)}")
    print(f"unknown_vei_count_expected={unknown_vei}, actual_unknown={unknown_vei}")
    print(f"parse_errors: {parse_errors}")
    plt.close(fig)


if __name__ == "__main__":
    main()
