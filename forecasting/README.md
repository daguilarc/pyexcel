# Forecasting

Choose historical sheets, a target and forecast length. `tabpfn_forecast.py` automatically backtests and forecasts subsequent observations in the target's original units. Desktop Excel and Python are required for command-line and macro use.

TabPFN is a pretrained tabular model. Gradient boosting builds successive trees that correct earlier errors. Both here learn target changes from past observations, then add predicted changes to the latest target recursively.

## Install

From the repository folder, use a separate environment (Python 3.11 recommended):

```sh
python3.11 -m venv .venv-forecast
source .venv-forecast/bin/activate
python -m pip install -r forecasting/requirements.txt
python -m xlwings addin install
```

On Windows activate with `.venv-forecast\Scripts\activate`. Configure xlwings to use that interpreter. For gradient boosting alone, `tabpfn` is optional; the other dependencies are required. Divvy has separate dependencies.

Conda is optional. With conda on PATH, run `bash forecasting/setup.sh` or `forecasting\setup.bat`. Each wrapper runs from its own directory and updates the isolated Python 3.11 `pyexcel-forecast` environment using the neighboring environment.yml and requirements.txt, then installs the add-in through it. They do not delete environments or require bundled Miniconda.

## Run

A target-only `History` sheet can have a `Sales` header followed by historical values in chronological row order:

```sh
python forecasting/tabpfn_forecast.py /path/to/data.xlsx \
  --data-sheets History --target-column Sales --future-rows 10
```

This produces ten subsequent observations per series. With no time column, configured sheet order then source row order defines chronology. Entirely blank rows are ignored. Headers must be unique nonblank text, and sheets must share headers. Invalid required cells stop the run; rows are never silently dropped or imputed.

Optional refinements:

- `--primary-time-index Year` sorts each series. Use finite numeric values/numeric strings, naive date/datetime cells, or unambiguous ISO dates/datetimes. Use one kind throughout history. `--secondary-time-index Quarter` adds a finite numeric ordering key and requires primary time. Repeated complete keys within a series fail. Time and target/identity roles must differ.
- `--categorical-columns Store,Product` identifies independent series, each with its own model and origin. This setting partitions history; it does not make categorical regression adjustments. Identities must be nonblank strings, booleans or finite numbers. Numeric 1 and 1.0 share identity; text "1" remains distinct. Series need not have aligned starts or ends.
- `--feature-columns all` (default) selects all historical columns except target, time and identity columns. `none` selects no covariates. Otherwise supply an explicit comma-separated list of numeric covariates. Numeric strings are accepted; boolean or nonfinite numeric values are rejected. Explicit lists must be unique and cannot include target/time/identity roles. Unselected columns are ignored. Exclude text columns or designate them as identities when they identify series.
- `--model-type GradientBoosting` selects gradient boosting instead of default TabPFN.

Every model input uses the previous target, the mean of up to three preceding targets, observation index and previous selected covariates. During each forecast, covariates stay fixed at their last observed prefix values. This cannot anticipate future price or promotion changes. Diagnostics list feature names and frozen values.

Regular numeric time can supply future display labels. Recognizable date frequencies, including calendar months, can also be extended. Irregular, compound or insufficient time histories get blank future-time labels and complete Step numbers. These labels never feed the model. Each origin retains original time fields and source sheet/row. Row-order mode depends on rows actually being chronological.

## Settings and macro

CLI flags override nonblank settings, which override defaults. `--settings /path/to/settings.xlsx` selects a settings workbook; otherwise use `forecasting/PyexcelSettings.xlsx`. Data and settings must be separate workbooks.

Import `forecasting/pyexcel_macros.bas` into an `.xlsm` workbook. Set `SCRIPT_DIR` to the forecasting folder or configure xlwings PYTHONPATH. Run `RunForecast`; the first run creates settings. Fill its `Forecast` sheet (labels in A, values in B) and run again.

| Setting | Flag | Default |
|---|---|---|
| Data Sheets (comma-separated) | `--data-sheets` | required |
| Header Row (1-based) | `--header-row` | 1 |
| Data Start Row (1-based) | `--data-start-row` | 2 |
| Target Column | `--target-column` | required |
| Feature Columns (or "all") | `--feature-columns` | all (or none/list) |
| Categorical Columns | `--categorical-columns` | blank |
| Primary Time Index | `--primary-time-index` | blank |
| Secondary Time Index | `--secondary-time-index` | blank |
| Model Type | `--model-type` | TabPFN |
| Ensemble Size (TabPFN) | `--ensemble-size` | 50 |
| GB Estimators (GradientBoosting) | `--gb-estimators` | 100 |
| Future Rows | `--future-rows` | 10 |
| Output Sheet | `--output-sheet` | Forecast |

Counts must be positive integers; data starts below headers. Historical sheets, output and Diagnostics must be distinct, including casing aliases. All results and payloads are validated before clearing either destination, then written and saved once. Native Excel write/save failures can still leave partial changes.

## Automatic backtesting and limits

For n records and requested H, the block length is B=min(H,max(1,floor(n/4))). Prefix lengths n−2B and n−B are evaluated when they contain at least three records. Each predicts B subsequent observations using the same recursive function as final forecasting. Evaluation blocks do not overlap, and held-out targets/covariates never enter prediction. This is bounded validation, without tuning or model selection; it needs up to three fits per series.

Diagnostics include each origin, training count, lead, actual, prediction and signed error in original units, plus aggregate/per-lead MAE, RMSE and counts. Leads without observations are Unmeasured. Global scores weight available ledger rows equally. Scores describe earlier smaller models under fixed-covariate assumptions, not measured accuracy of the final all-history fit or guaranteed future error. There are no uncertainty bands or calibrated coverage claims.

One or two records produce a disclosed last-value baseline, without fitting a model or using covariates. Three records can fit but have no backtest; four have one origin. Forecasting proceeds when no backtest is available. Long horizons can have unmeasured leads. Recursive changes can drift or become negative; no clipping, seasonal advantage or statistical adequacy is promised. Extremely large finite values may exceed the estimator’s numeric range and stop execution.

Missing TabPFN or a RuntimeError explicitly naming authentication/gated access restarts every evaluation and final fit with gradient boosting. Diagnostics disclose the effective family/reason; arbitrary failures remain visible. Many series can be slow, and model access or Excel configuration can block a run.

## TabPFN licensing

The local `tabpfn` package may download weights. Licensing depends on the selected checkpoint, which this script does not pin. Check its license before use; download access does not grant commercial rights. The hosted API has separate terms and is not used here.

- [TabPFN model licenses](https://github.com/PriorLabs/TabPFN#license)
- [TabPFN-2.5 license](https://huggingface.co/Prior-Labs/tabpfn_2_5/blob/main/LICENSE)
- [Prior Labs API terms](https://priorlabs.ai/pricing)
