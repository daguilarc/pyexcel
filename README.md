# pyexcel

Two Python tools connect Excel to Python via [xlwings](https://www.xlwings.org/): `isolation_forest.py` scores rows for anomalies with scikit-learn's Isolation Forest, and `tabpfn_forecast.py` forecasts a target column with TabPFN, GradientBoosting, OLS, or a Gamma GLM. Each tool runs as an Excel macro or as a standalone CLI command.

## Install

- [Miniconda](https://docs.conda.io/en/latest/miniconda.html) is required; a portable install goes in a `miniconda/` folder beside the scripts.
- Automated: `bash setup.sh` (Mac/Linux) or `setup.bat` (Windows) — creates the `excelpy` conda env, installs dependencies, and registers the xlwings Excel add-in.
- Manual:
  ```bash
  conda env create -f environment.yml
  conda activate excelpy
  pip install -r requirements.txt
  ```
- Microsoft Excel must be installed on the machine — both the macro and the CLI drive a live Excel instance through xlwings.

## Excel macro

- Alt+F11 -> File -> Import File -> select `pyexcel_macros.bas`.
- Set the `SCRIPT_DIR` constant in the module to the folder containing `isolation_forest.py` and `tabpfn_forecast.py`.
- Save the workbook as `.xlsm`.
- Alt+F8 -> run `RunIsolationForest` or `RunForecast`.
- The first run creates `PyexcelSettings.xlsx` beside the scripts; configure it (see Settings and flags below), then run the macro again.

## CLI

```
python <tool>.py <workbook.xlsx> [flags]
```
`<tool>` is `isolation_forest` or `tabpfn_forecast`.

- Precedence: explicit CLI flag > `PyexcelSettings.xlsx` > built-in default.
- `--settings <path>` uses a different settings workbook (default: `PyexcelSettings.xlsx` beside the script).
- Quote any path or option value that contains spaces.
- Exit code `0` on success, `1` on error (message printed to stderr).

## Settings and flags

Labels are read from the `IsolationForest` / `Forecast` sheets in `PyexcelSettings.xlsx`; the matching CLI flag sets the same value.

**Isolation Forest**

| Setting | Flag | Default |
|---|---|---|
| Data Sheets (comma-separated) | `--sheets` / `--sheet` | first sheet |
| Contamination | `--contamination` | `0.1` |
| Monte Carlo Samples | `--monte-carlo-samples` | `0` |
| Ascending | `--ascending` | `False` |
| Epsilon | `--epsilon` | `False` |
| Cumulative Sheets | `--cumulative` | `False` |
| Output Sheet | `--output-sheet` | `AnomalyScores` |

**Forecast**

| Setting | Flag | Default |
|---|---|---|
| Data Sheets (comma-separated) | `--data-sheets` | — (required) |
| Header Row (1-based) | `--header-row` | `1` |
| Data Start Row (1-based) | `--data-start-row` | `2` |
| Target Column | `--target-column` | — (required) |
| Feature Columns (or "all") | `--feature-columns` | `all` |
| Categorical Columns | `--categorical-columns` | — |
| Primary Time Index | `--primary-time-index` | — |
| Secondary Time Index | `--secondary-time-index` | — |
| Model Type | `--model-type` | `TabPFN` |
| MC Samples (linear only) | `--mc-samples` | `5000` |
| Ensemble Size (TabPFN) | `--ensemble-size` | `50` |
| GB Estimators (GradientBoosting) | `--gb-estimators` | `100` |
| Uncertainty Mode | `--uncertainty-mode` | `Frequentist` |
| Prior Confidence (1-10, Bayesian only) | `--prior-confidence` | `5` |
| Future Rows | `--future-rows` | `10` |
| Output Sheet | `--output-sheet` | `Forecast` |

## Isolation Forest

- Unsupervised anomaly scoring over one or more sheets, using scikit-learn's IsolationForest.
- Output sheet columns: `Source Sheet`, `Index`, the numeric feature columns, `Anomaly Score`, `Is Anomaly`.
- `Cumulative Sheets` (when scoring more than one sheet): sheet *k* is fit on sheets 1..*k* only, with no lookahead into later sheets; contamination is held constant across all sheets.
- `Monte Carlo Samples` > 0 replaces the single fit with stationary-bootstrap resampling, averaging scores across iterations.

## TabPFN forecast

- Supervised forecasting of `Target Column`, using `Model Type`: TabPFN, GradientBoosting, OLS, or Gamma (GLM).
- TabPFN falls back to GradientBoosting automatically when the TabPFN package or its model weights are unavailable.
- OLS and Gamma get Monte Carlo prediction intervals: frequentist pairs bootstrap by default, or Bayesian MCMC (via Bambi/PyMC) when `Uncertainty Mode` is Bayesian.
- Both surfaces write the forecast to `Output Sheet` and a `Diagnostics` sheet (residual metrics and, for OLS/Gamma, Monte Carlo sensitivity statistics).

## Migration

- UDF formulas were removed; use the Excel macro or the CLI instead.
- `ForecastSettings.xlsx` is replaced by `PyexcelSettings.xlsx` — copy your values across.
- The legacy VBA module was deleted.

## Troubleshooting

| Issue | Fix |
|---|---|
| `#ERROR` / exception in Excel | Check `PyexcelSettings.xlsx` values; re-run after fixing |
| `ModuleNotFoundError` | `pip install -r requirements.txt` in the `excelpy` env |
| Macro not listed under Alt+F8 | Re-import `pyexcel_macros.bas`; confirm workbook is saved as `.xlsm` |
| CLI can't find `PyexcelSettings.xlsx` | Pass `--settings <path>` or place it beside the script |
| Workbook save fails | Close the workbook in Excel, then re-run |
