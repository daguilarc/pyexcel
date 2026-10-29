# divvy-xlsx-cleanup

Removes the unused columns from a Divvy/Bill transaction export so the
spreadsheet is easier to read for me (Diego).

## Requirements

Python 3, pandas, and openpyxl, which pandas uses to read and write .xlsx
files.

```
python3 -m pip install pandas openpyxl
```

## Usage

```
python3 divvy_cleanup.py
```

The script asks for the path to the `.xlsx`. A bare filename is looked up in
the same folder as the script; absolute paths, `~`, and paths dragged in from
Finder all work.

The cleaned copy is written next to the input as `<name>_cleaned.xlsx`. The
original file is left as it was.

The export lists the newest transaction first. The cleaned copy is ordered the
other way, oldest first by `Cleared Time in Statement (MT)`. That column comes
out headed `Date`, holding the date alone; the time of day is used for the sort
and then dropped.

## Columns it keeps

Ten of the export's 38 columns, in the order `KEPT_COLUMNS` lists them.

Columns are picked by name, so an export with extra or reordered columns is
still fine. If one of the ten is missing, the script stops and names it.

The cleaned copy is a plain sheet of values. Number formats, styles, and
column widths from the export are not carried over.

## Changing which columns are kept

`KEPT_COLUMNS` in `divvy_cleanup.py` is the list of headers to keep. Add or
remove a name to change the output.
