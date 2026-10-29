#!/usr/bin/env python3
"""Remove the unused columns from a Divvy transaction export.

Prompts for the .xlsx to clean and writes the result next to it as
<name>_cleaned.xlsx, keeping the columns named in KEPT_COLUMNS and dropping
everything else. Rows come out oldest first, and the time column comes out
headed Date, holding the date alone.
"""

import sys
import warnings
from pathlib import Path

import pandas as pd

# The columns worth keeping, in the order they should appear.
KEPT_COLUMNS = [
    "Cleared Time in Statement (MT)",
    "First Name",
    "Last Name",
    "Clean Merchant Name",
    "Amount",
    "Budget",
    "Receipt?",
    "Notes (Memo)",
    "Account",
    "Class",
]

# The export lists the newest transaction first. Rows come out sorted on this
# column instead, oldest first.
SORT_COLUMN = "Cleared Time in Statement (MT)"

SCRIPT_DIR = Path(__file__).resolve().parent


def prompt_for_export():
    """Ask for the export to clean and return its path and contents.

    Keeps asking until the answer names a spreadsheet that opens, so a typo or
    the wrong file costs nothing. A bare filename or a relative path is looked
    up in the folder this script lives in, which makes the usual answer just
    the name of the export. Absolute paths and paths starting with ~ work as
    well.
    """
    while True:
        try:
            # Terminal drag-and-drop wraps a path containing spaces in quotes.
            answer = input(f"Path to .xlsx (relative to {SCRIPT_DIR}): ")
        except (EOFError, KeyboardInterrupt):
            sys.exit("\nCancelled.")
        answer = answer.strip().strip("'\"")
        if not answer:
            print("No path given.", file=sys.stderr)
            continue
        path = Path(answer).expanduser()
        if not path.is_absolute():
            path = SCRIPT_DIR / path
        if not path.is_file():
            print(f"Not found: {path}", file=sys.stderr)
            continue
        try:
            # These exports carry no default style, which openpyxl warns about.
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                return path, pd.read_excel(path)
        except Exception as error:
            # A file that is not a spreadsheet, or one that is damaged.
            print(f"Could not read {path.name}: {error}", file=sys.stderr)


def main():
    source, export = prompt_for_export()

    # Columns are picked by name, so a differently shaped export is fine as
    # long as it still has these. Say which one is missing when it does not.
    missing = [name for name in KEPT_COLUMNS if name not in export.columns]
    if missing:
        sys.exit(f"{source.name} has no column named {', '.join(missing)}")

    # Sort on the parsed time rather than on the text, so the order holds even
    # when an export stores some times as text and others as dates. A row whose
    # time is blank or unreadable sorts to the top, where it is easy to spot;
    # these exports cover cleared transactions, so every row should carry one.
    # The whole export is sorted and the columns picked afterwards, which means
    # the sort column does not have to be one of the kept ones.
    cleaned = export.sort_values(
        SORT_COLUMN,
        key=lambda column: pd.to_datetime(column, errors="coerce"),
        na_position="first",
    )[KEPT_COLUMNS].copy()

    # The export stamps a time of day that is not wanted, so the sort column is
    # cut back to the date it falls on. Rows were ordered on the full timestamp
    # above, so transactions on the same day keep the order they happened in.
    cleaned[SORT_COLUMN] = pd.to_datetime(
        cleaned[SORT_COLUMN], errors="coerce"
    ).dt.strftime("%Y-%m-%d")
    cleaned = cleaned.rename(columns={SORT_COLUMN: "Date"})

    destination = source.with_name(f"{source.stem}_cleaned.xlsx")
    try:
        cleaned.to_excel(destination, index=False)
    except OSError as error:
        sys.exit(f"Could not write {destination.name}: {error}")
    print(f"Wrote {destination}")


if __name__ == "__main__":
    main()
