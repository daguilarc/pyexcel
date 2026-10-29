# pyexcel

Excel forecasting with TabPFN or gradient boosting via [xlwings](https://www.xlwings.org/), plus Divvy/Bill transaction export cleanup in `divvy_xlsx/`. Forecasting runs as an Excel macro or CLI command.

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
- Set the `SCRIPT_DIR` constant in the module to the folder containing `tabpfn_forecast.py`.
- Save the workbook as `.xlsm`.
- Alt+F8 -> run `RunForecast`.
- The first run creates `PyexcelSettings.xlsx` beside the scripts; configure it (see Settings and flags below), then run the macro again.

## CLI

```
python tabpfn_forecast.py <workbook.xlsx> [flags]
```

- Precedence: explicit CLI flag > `PyexcelSettings.xlsx` > built-in default.
- `--settings <path>` uses a different settings workbook (default: `PyexcelSettings.xlsx` beside the script).
- Quote any path or option value that contains spaces.
- Exit code `0` on success, `1` on error (message printed to stderr).

## Settings and flags

Labels are read from the `Forecast` sheet in `PyexcelSettings.xlsx`; the matching CLI flag sets the same value.

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
| Ensemble Size (TabPFN) | `--ensemble-size` | `50` |
| GB Estimators (GradientBoosting) | `--gb-estimators` | `100` |
| Future Rows | `--future-rows` | `10` |
| Output Sheet | `--output-sheet` | `Forecast` |

## TabPFN forecast

- Choose `TabPFN` or `GradientBoosting` as Model Type.
- TabPFN falls back to gradient boosting if the package is missing or a gated-model/authentication RuntimeError occurs; other errors are raised.
- Both entry points write predictions and a `Diagnostics` sheet, then save the workbook. Existing output and diagnostics sheets are cleared when reused.
- Features use linear trends with a numeric time index, otherwise historical means plus noise.
- Categorical columns use group demeaning. Predictions remain on that transformed scale with sampled category labels; they are not group-specific predictions restored to the original scale.
- Bounds are prediction ± 1.96 × training RMSE, not calibrated 95% prediction intervals. MAE/RMSE are training metrics, not holdout accuracy.

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
| macOS keeps asking permission for the same folder | Excel is sandboxed, and each new Excel process needs a fresh grant to open files a script names. Leave Excel running between runs, or open the workbook yourself first and the CLI will attach to it. Granting Excel access to the folder under System Settings -> Privacy & Security -> Files and Folders also helps. |

## TabPFN access and licensing

Verified October 26, 2026 against the upstream sources linked below. **Recent TabPFN model weights (2.5 and later) distributed on Hugging Face have noncommercial, nonproduction licenses** permitting research and evaluation within their specific terms. Commercial decision-making, client deliverables, paid services, and production deployment require appropriate commercial rights. Accepting a Hugging Face gate or using an access token does not grant those rights.

This script uses the local `tabpfn` Python package, which can download model weights from Hugging Face. It does not call Prior Labs' hosted inference API. The checkpoint selected by the installed package determines the applicable model license; this repository does not pin a checkpoint. TabPFN-2 weights have a different license (Apache 2.0 with an additional attribution requirement), so the noncommercial restriction must not be stated as applying to every TabPFN version or to Hugging Face API access generally.

Prior Labs' hosted API is a separate offering: its pricing page explicitly permits commercial API use under its terms. Commercial licensing is also available for local deployments. Check the exact checkpoint license and current service terms before use:

- [Upstream TabPFN licensing and model versions](https://github.com/PriorLabs/TabPFN#license)
- [TabPFN-2.5 model license](https://huggingface.co/Prior-Labs/tabpfn_2_5/blob/main/LICENSE)
- [Prior Labs pricing and commercial API FAQ](https://priorlabs.ai/pricing)

## Divvy/Bill cleanup

```sh
python divvy_xlsx/divvy_cleanup.py
```

Enter the export path when prompted. Bare filenames are resolved relative to `divvy_xlsx/`; absolute paths also work. The script keeps ten named columns, sorts transactions oldest first, converts the cleared timestamp to a `Date` column, and writes `<name>_cleaned.xlsx` beside the source. The source stays unchanged. See the [Divvy README](divvy_xlsx/README.md) for column selection and formatting details.

## Checks

```sh
python -m unittest discover -s tests -v
```

Tests use an in-memory workbook adapter and a stand-in for TabPFN; they run real gradient boosting and Divvy XLSX round trips. Desktop Excel and actual TabPFN inference require separate integration checks.

## License

Repository code is MIT licensed. Imported Divvy code retains its MIT notice. TabPFN model weights have separate licenses; this repository's MIT license does not override them.
