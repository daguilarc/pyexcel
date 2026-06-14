# pyexcel

Python tools that connect Excel to scikit-learn and forecasting models via [xlwings](https://www.xlwings.org/).

| Tool | Script | Excel integration |
|------|--------|-------------------|
| **Isolation Forest** | `isolation_forest.py` | UDFs: `=ISOLATION_FOREST(...)`, `=ANOMALY_BINARY(...)` |
| **TabPFN forecast** | `tabpfn_forecast.py` | Macro: `RunForecastWithSettings` (uses `ForecastSettings.xlsx`) |

| Mode | What it does |
|------|----------------|
| **CLI** | Terminal: Python opens the workbook via xlwings, writes results, saves |
| **UDF / macro** | In Excel: formulas in cells or **Alt+F8** macro (after xlwings **Import Functions**) |

Platform differences (setup, paths, what to use when): **[Excel on Mac vs Windows](#excel-on-mac-vs-windows)**.

---

## Excel on Mac vs Windows

Both platforms need **Microsoft Excel installed** and the **xlwings** add-in pointed at the same Python env. Neither CLI nor UDF runs without Excel on that machine.

### At a glance

| Topic | Excel on Mac | Excel on Windows |
|-------|----------------|------------------|
| **Supported?** | Yes | Yes |
| **Setup script** | `bash setup.sh` | `setup.bat` |
| **Portable Miniconda folder** | `miniconda/` | `miniconda\` |
| **Typical Python command** | `python3` or `python` (after `conda activate excelpy`) | `python` |
| **Default terminal** | Terminal.app / iTerm (bash/zsh) | Command Prompt or PowerShell |
| **Path separators in CLI** | `/Users/you/...` forward slashes | `C:\Users\you\...` backslashes |
| **Quote paths with spaces** | `"..."` in bash/zsh | `"..."` in CMD and PowerShell |
| **CLI (terminal)** | Works well; common choice for authors on Mac | Works the same way |
| **UDF formulas** | Works after Import Functions; setup can be fussier (interpreter path, permissions) | Works well; common for sharing live workbooks |
| **Forecast macro (Alt+F8)** | Yes | Yes |
| **Legacy VBA** (`ISOLATION_FOREST_UDF.bas`) | Not the main path | Optional Windows-oriented alternative to xlwings Import Functions |

### Recommended workflow

| You are… | Suggested approach |
|----------|-------------------|
| On **Mac**, running your own analysis | **CLI** — see [CLI (terminal commands)](#cli-terminal-commands); fewer moving parts than in-cell UDFs |
| On **Mac**, need formulas in the sheet | **UDF** — Import Functions once; keep `isolation_forest.py` in a fixed path |
| On **Windows**, sharing workbooks with colleagues | **UDF** — `=ISOLATION_FOREST(...)` recalculates in the file they already use |
| On **Windows**, batch / reproducible runs | **CLI** — same commands as Mac, different path quoting |
| **Forecast** on either platform | Configure `ForecastSettings.xlsx`, then macro or CLI |

CLI is **not** Mac-only and UDF is **not** Windows-only — both work on both; the table above is about convenience, not capability.

### Installation differences

| Step | Mac | Windows |
|------|-----|---------|
| Install Miniconda | `.sh` installer or extract to `miniconda/` | `.exe` installer or portable `miniconda\` |
| One-click setup | `bash setup.sh` from Terminal | Double-click `setup.bat` |
| xlwings add-in | `python -m xlwings addin install` in `excelpy` env | Same |
| Import Functions | xlwings ribbon → pick `.py` file | Same |

After setup, point xlwings at the **`excelpy` conda env** interpreter (xlwings settings / conf if formulas fail with “module not found”).

### CLI on Mac vs Windows

Same command order on both platforms:

```text
python <script.py> <workbook.xlsx> [options...]
```

| | Mac (bash/zsh) | Windows (CMD) |
|--|----------------|-----------------|
| Activate env | `conda activate excelpy` | `conda activate excelpy` |
| Change directory | `cd ~/Desktop/excelpy` | `cd C:\Users\you\Desktop\excelpy` |
| Path with spaces | `python isolation_forest.py "/Users/you/My Reports/data.xlsx"` | `python isolation_forest.py "C:\Users\you\My Reports\data.xlsx"` |
| Sheet name option | `--sheet "Sales Data"` | `--sheet "Sales Data"` |

Full examples and flag rules: **[CLI (terminal commands)](#cli-terminal-commands)**.

### UDF / macro on Mac vs Windows

| | Mac | Windows |
|--|-----|---------|
| **Isolation Forest** | `=ISOLATION_FOREST(...)`, `=ANOMALY_BINARY(...)` | Same |
| **One-time setup** | xlwings → Import Functions → `isolation_forest.py` | Same |
| **Forecast** | Alt+F8 → `RunForecastWithSettings` (import `tabpfn_forecast.py` too if needed) | Same |
| **Recalculation** | Excel recalculates when the sheet changes; first call may be slow | Same |
| **If `#ERROR` or import fails** | Confirm conda env, re-run `xlwings addin install`, re-import after moving `.py` | Same; also check Trust Center / macro settings if using `.bas` VBA |

### Mac-specific notes

- Use **`python3`** if `python` is not on PATH after activating conda.
- **Gatekeeper / privacy:** allow Terminal (and sometimes Excel) to control Excel when macOS prompts you.
- **Apple Silicon:** use the **arm64** Miniconda installer so sklearn/xlwings wheels match your chip.
- Saving: if CLI says the file is locked, **close the workbook in Excel** and run again.

### Windows-specific notes

- Prefer **quoted paths** in CMD when spaces appear in user folder names (`Desktop`, `My Documents`).
- **PowerShell** works; use the same quoted paths (single quotes are fine).
- Optional **`ISOLATION_FOREST_UDF.bas`**: VBA wrapper + RunPython for shops that already use VBA modules (not required if you use xlwings Import Functions).

### Sharing files across platforms

Workbooks (`.xlsx` / `.xlsm`) open on both platforms. Colleagues on Windows can use your UDF formulas if they complete the same xlwings + Import Functions setup on their PC. CLI-produced files (extra sheets, scores, forecasts) are normal Excel files and open anywhere without Python.

---

## Code layout (no shared module yet)

Both scripts keep helpers **inline at the top of each file** (arg parsing, array intake, transforms, output builders, etc.). A shared module (e.g. `_core.py`) would be reasonable, but is **not used on purpose for now** so each tool can be developed and tested in isolation. That may change later.

---

## Dependencies

| File | Use when |
|------|----------|
| [`requirements-base.txt`](requirements-base.txt) | Shared xlwings / pandas / sklearn stack (included by the files below) |
| [`requirements-isolation-forest.txt`](requirements-isolation-forest.txt) | Anomaly detection only |
| [`requirements-tabpfn-forecast.txt`](requirements-tabpfn-forecast.txt) | Forecast macro only |
| [`requirements.txt`](requirements.txt) | Both tools (includes the two per-tool files above) |

Conda: [`environment.yml`](environment.yml) installs `requirements-base.txt` via pip (add forecast deps with `pip install -r requirements-tabpfn-forecast.txt` if needed).

```bash
pip install -r requirements-isolation-forest.txt
pip install -r requirements-tabpfn-forecast.txt
pip install -r requirements.txt
```

---

## Installation

### Quick setup

**Windows:** extract Miniconda to `miniconda\`, run `setup.bat`  
**Mac/Linux:** extract to `miniconda/`, run `bash setup.sh`

### Step-by-step

1. **Miniconda** — [download](https://docs.conda.io/en/latest/miniconda.html), or portable extract into `miniconda/`:
   - Mac: `bash Miniconda3-latest-MacOSX-arm64.sh -b -p miniconda`
   - Linux: `bash Miniconda3-latest-Linux-x86_64.sh -b -p miniconda`
2. **Setup script** — creates `excelpy` env and installs xlwings add-in.
3. **Import Functions** (one-time per script) — xlwings ribbon → select `isolation_forest.py` and/or `tabpfn_forecast.py`.
4. **Verify** — `=ISOLATION_FOREST(` should autocomplete.

### Manual install

```bash
conda env create -f environment.yml
conda activate excelpy
python -m xlwings addin install
```

---

## CLI (terminal commands)

Both CLI tools use the same shell pattern: **Python opens your workbook through xlwings** (Excel must be installed). They are not standalone “read xlsx without Excel” tools.

### Command shape

```text
python <script.py> <path-to-workbook.xlsx> [options...]
```

| Part | Rule |
|------|------|
| `python` | Use the interpreter from the `excelpy` conda env (`conda activate excelpy`). On some Macs the command is `python3` instead of `python`. |
| `<script.py>` | Path to `isolation_forest.py` or `tabpfn_forecast.py`. |
| `<path-to-workbook.xlsx>` | **First and only required argument** — the `.xlsx` / `.xlsm` file to process. |
| `[options...]` | Optional flags **after** the workbook path (`--flag value` or `--flag` for switches). |

### Before you run

1. Activate the environment: `conda activate excelpy`
2. `cd` to the folder that contains the script (or use full paths for both script and workbook).
3. Prefer the workbook **saved and closed** in Excel (avoids lock/save conflicts). The script opens it, writes sheets, and saves.

### Paths and quoting

Paths are read by your shell. **Quote any path that contains spaces or special characters.**

**Mac / Linux (bash or zsh)**

```bash
cd /Users/you/Desktop/excelpy
conda activate excelpy

# Relative workbook path (relative to current directory)
python isolation_forest.py ../data/my_report.xlsx

# Absolute path with spaces — quotes required
python isolation_forest.py "/Users/you/Desktop/My Reports/Q1 data.xlsx"

# Sheet name with a space — quotes on the option value
python isolation_forest.py data.xlsx --sheet "Sales Data" --contamination 0.05
```

**Windows (Command Prompt)**

```batch
cd C:\Users\you\Desktop\excelpy
conda activate excelpy

python isolation_forest.py "C:\Users\you\Desktop\My Reports\Q1 data.xlsx"
python isolation_forest.py data.xlsx --sheet "Sales Data" --contamination 0.05
```

**Windows (PowerShell)** — same quoting; use `python` from the activated env:

```powershell
python isolation_forest.py 'C:\Users\you\Desktop\My Reports\Q1 data.xlsx'
```

| Situation | Format |
|-----------|--------|
| Spaces in file or folder name | Wrap the whole path in `"..."` (or `'...'` in PowerShell) |
| Relative path | `data.xlsx` or `./data.xlsx` from your current `cd` |
| Absolute path | Full path to the file; quotes if needed |
| Option with spaces | `--sheet "Sheet 2"` not `--sheet Sheet 2` |

### Flags (isolation forest only)

Boolean switches take **no value** — presence means on:

```bash
python isolation_forest.py data.xlsx --ascending
python isolation_forest.py data.xlsx --ascending --epsilon
```

Numeric / text options use a value **after** the flag:

```bash
python isolation_forest.py data.xlsx --contamination 0.05 --monte-carlo-samples 1000 --output-sheet "Outlier Scores"
```

Flag names use **hyphens** as shown (`--monte-carlo-samples`, not `--monte_carlo_samples`).

### TabPFN CLI (single argument)

Forecast CLI only accepts the **data workbook** path. Configuration is **not** passed on the command line — it is read from `ForecastSettings.xlsx` in the **same directory as `tabpfn_forecast.py`** (create/configure it by running the macro once, or copy it there by hand).

```bash
cd /path/to/excelpy
conda activate excelpy
python tabpfn_forecast.py "/path/to/your/analysis.xlsx"
```

The workbook you pass is the file that contains your data sheets; settings live beside the script, not necessarily beside that file.

### Exit status

- **0** — finished; check the new/updated sheet in the workbook (`AnomalyScores` by default, or `Forecast` for TabPFN).
- **1** — error message printed to the terminal (missing data, bad settings, missing deps, etc.).

---

# Isolation Forest (`isolation_forest.py`)

Unsupervised anomaly detection on numeric features. No Settings sheet; no target column.

- **Default scores:** lower = more anomalous (standard Isolation Forest).
- **Monte Carlo:** stationary bootstrap when *n* ≥ 10, else i.i.d. bootstrap; throttled ~100 iterations / 0.5s.

## Quick reference

```
=ISOLATION_FOREST(data, [contamination], [monte_carlo], [ascending], [epsilon], [binary], [bayes_freq], [array_index])
=ANOMALY_BINARY(data, [prior_contamination], [monte_carlo], [bayes_freq], [array_index])
```

Optional parameters after data ranges can appear **in any order** (matched by type/value). Booleans are assigned in order seen: `epsilon`, `ascending`, `binary`, `bayes_freq` for `ISOLATION_FOREST`; `bayes_freq` only for `ANOMALY_BINARY`.

## `=ISOLATION_FOREST(...)`

Returns continuous anomaly scores (and optional 0/1 binary column).

| Parameter | Default | Notes |
|-----------|---------|--------|
| Data range(s) | — | Rows = observations, columns = features. Multiple ranges combined for fitting; same column count required. |
| `contamination` | `0.1` | Expected anomaly fraction, `0 < x ≤ 0.5` (Liu et al., 2008). |
| `monte_carlo_samples` | `0` | `0` or `100`–`10000`. |
| `ascending` | `FALSE` | `TRUE`: higher score = more anomalous. |
| `epsilon` | `FALSE` | Shift scores positive; **requires** `ascending=TRUE` (gamma GLM use). |
| `binary` | `FALSE` | `TRUE`: adds 0/1 column beside scores. |
| `bayes_freq` | `FALSE` | With `binary=TRUE` and 2+ arrays: sequential Beta updating for binary only. |
| `array_index` | `0` | `1`…`n`: return only that array’s column(s). |

**Examples**

```excel
=ISOLATION_FOREST(A2:E100)
=ISOLATION_FOREST(A2:E100, 0.05, 1000)
=ISOLATION_FOREST(A2:E100, FALSE, TRUE)          // ascending=TRUE (FALSE does not stick on epsilon)
=ISOLATION_FOREST(A2:E50, G2:J50, TRUE, TRUE)  // binary + bayes_freq on multi-array
=ISOLATION_FOREST(A2:E50, G2:J50, 2)           // second array only
```

**Returns:** single value; one column; or one column per array (padded to max length). With `binary=TRUE`, columns interleave score / binary per array.

## `=ANOMALY_BINARY(...)`

Returns `1` = anomaly, `0` = normal.

| Parameter | Default | Notes |
|-----------|---------|--------|
| Data range(s) | — | Same layout as above. |
| `prior_contamination` | `0.1` | Fixed rate if `bayes_freq=FALSE`; prior for first array if `bayes_freq=TRUE`. |
| `monte_carlo_samples` | `0` | Same as above. |
| `bayes_freq` | `FALSE` | `TRUE` + 2+ arrays: fit cumulatively, update Beta prior between arrays. |
| `array_index` | `0` | Same as above. |

**Examples**

```excel
=ANOMALY_BINARY(A2:E100)
=ANOMALY_BINARY(A2:E50, G2:J50, 0.1, TRUE)
=ANOMALY_BINARY(A2:E50, G2:J50, 1000, TRUE, 2)
```

Frequentist mode (`bayes_freq=FALSE`): all arrays combined, one fit, sklearn `predict` (-1 → 1).  
Bayesian mode: scores current array only; prior updated from binary counts at posterior-mean percentile (see limitation below).

### Bayesian updating (`bayes_freq=TRUE`)

Uses Beta(α, β) with prior strength 100: α = `prior × 100`, β = `(1 - prior) × 100`. After each array (except the last), α += anomalies, β += normals. Contamination for the next array = `beta.mean(α, β)` capped at 0.5.

**Limitation:** counts use a threshold tied to the parameter being updated (posterior-mean percentile). The posterior tracks **cutoff stability**, not independent evidence about a population anomaly rate.

### Multi-array rules

- Same number of **columns** across arrays; row counts may differ.
- Cross-sheet ranges OK: `=ISOLATION_FOREST(Sheet1!A2:E10, Sheet2!A2:E10)`.
- Multiple columns are built per array then transposed for Excel row layout.

## CLI

See **[CLI (terminal commands)](#cli-terminal-commands)** for env activation, paths, quoting, and flag syntax.

```bash
python isolation_forest.py <workbook.xlsx> [--sheet NAME] [--contamination 0.1] [--monte-carlo-samples 0] [--ascending] [--epsilon] [--output-sheet AnomalyScores]
```

| Option | Default |
|--------|---------|
| `--sheet` | first sheet (quote if the name has spaces) |
| `--contamination` | `0.1` |
| `--monte-carlo-samples` | `0` |
| `--ascending` | off (flag only, no value) |
| `--epsilon` | off (requires `--ascending`) |
| `--output-sheet` | `AnomalyScores` |

**CLI data layout:** row 1 = headers, data from row 2; all numeric columns are features; rows with any NaN dropped.

**Output sheet:** Index, original features, Anomaly Score, Is Anomaly.

```bash
python isolation_forest.py data.xlsx --contamination 0.05 --monte-carlo-samples 1000
python isolation_forest.py "My Reports/data.xlsx" --sheet "Raw Input" --ascending --epsilon
```

## Isolation Forest troubleshooting

| Issue | Fix |
|-------|-----|
| Import / sklearn errors | `pip install -r requirements-isolation-forest.txt` |
| `#ERROR` in cell | ≥2 valid rows; matching feature counts; valid parameter ranges |
| Function unknown | xlwings → Import Functions → `isolation_forest.py` |
| CLI “no data” | Headers row 1; numeric columns present |
| Feature mismatch | All ranges must have same column count |

## Performance (typical)

| Size | Time |
|------|------|
| 1k × 10 features | &lt; 1 s |
| 10k × 50 | ~5–10 s |
| 100k × 100 | ~1–2 min |

Monte Carlo multiplies runtime (throttled). Bayesian multi-array mode fits once per array.

---

# TabPFN forecast (`tabpfn_forecast.py`)

Supervised forecasting from a **Settings** workbook (`ForecastSettings.xlsx`, created on first macro run).

1. Fill **Settings**: data sheets, header row, target column, features, optional time indices, model type (`TabPFN`, `GradientBoosting`, `OLS`, `Gamma`), MC samples (linear models), etc.
2. **Alt+F8** → `RunForecastWithSettings` (or CLI below).

**Models**

- **TabPFN / GradientBoosting:** single forecast + approximate 95% band from RMSE.
- **OLS / Gamma:** frequentist pairs bootstrap or optional Bayesian (Bambi/PyMC); writes `Forecast` + `Diagnostics` sheets.

**CLI**

See **[CLI (terminal commands)](#cli-terminal-commands)** for path quoting and env setup.

```bash
python tabpfn_forecast.py <workbook.xlsx>
```

No CLI flags — settings come from `ForecastSettings.xlsx` next to `tabpfn_forecast.py` (not next to the data workbook).

**Deps:** `pip install -r requirements-tabpfn-forecast.txt` (TabPFN may need Hugging Face auth; falls back to GradientBoosting).

---

## Isolation Forest vs forecast

| | Isolation Forest | TabPFN / OLS / Gamma |
|--|------------------|----------------------|
| Type | Unsupervised | Supervised |
| Needs target | No | Yes |
| Output | Anomaly scores / flags | Forecast + intervals |
| Typical use | Outliers, QC | Time series / regression |

---

## Project layout

```
excelpy/
├── README.md
├── requirements-base.txt
├── requirements.txt
├── requirements-isolation-forest.txt
├── requirements-tabpfn-forecast.txt
├── environment.yml
├── isolation_forest.py
├── tabpfn_forecast.py
├── ISOLATION_FOREST_UDF.bas    # legacy VBA + RunPython (Windows-oriented)
├── setup.bat / setup.sh
└── miniconda/                  # you add this
```

---

## System requirements

- Microsoft Excel (Windows or Mac)
- ~500 MB for Miniconda + packages
- Portable Miniconda works without admin rights
