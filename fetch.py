# /// script
# requires-python = ">=3.10"
# ///

"""
Record the real Smithsonian GVP data-fetch situation for this project.

    uv run fetch.py

The source file is the official confirmed Holocene eruptions spreadsheet from the
Smithsonian GVP page. Automated download attempts are blocked with HTTP 403, so
this script only checks whether the raw Excel file is already present locally.
"""

from pathlib import Path

OFFICIAL_PAGE = "https://volcano.si.edu/search_eruption.cfm"
FILE = "GVP_Eruption_List_Holocene_20260424.xlsx"
HERE = Path(__file__).parent
DATA = HERE / "data"
TARGET = DATA / FILE


def fetch():
    """Check whether the raw Smithsonian Excel file is already present."""
    if TARGET.exists():
        print("raw data file already present")
        print(f"path: {TARGET}")
        print(f"size: {TARGET.stat().st_size} bytes")
        return TARGET

    print("Smithsonian blocks the scripted download with HTTP 403.")
    print("Please download the Confirmed Holocene Eruptions spreadsheet manually")
    print("from the official GVP page and place it in data/.")
    print(f"official page: {OFFICIAL_PAGE}")
    print(f"required file: {FILE}")
    return None


if __name__ == "__main__":
    fetch()
