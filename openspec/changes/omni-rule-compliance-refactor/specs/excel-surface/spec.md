# excel-surface

## ADDED Requirements

### Requirement: RunPython macros are the only in-Excel surface
Each tool SHALL expose exactly one in-Excel entry point: a plain Python function (no `@xw.func`, no `@xw.sub` — both removed) invoked via `RunPython` from a Sub in the shipped `pyexcel_macros.bas`, using `xw.Book.caller()` for the active workbook. This mechanism SHALL be identical on Windows and macOS.

#### Scenario: Forecast macro runs cross-platform
- **WHEN** a user imports `pyexcel_macros.bas` and runs `RunForecast` via Alt+F8 on either platform
- **THEN** the forecast executes against the open workbook and writes Forecast + Diagnostics sheets

#### Scenario: Isolation forest macro exists
- **WHEN** a user runs `RunIsolationForest` via Alt+F8 with a populated IsolationForest settings sheet
- **THEN** anomaly scores are written to the configured output sheet of the open workbook

#### Scenario: UDF surface removed
- **WHEN** grepping the repo for `@xw.func`, `@xw.sub`, `ISOLATION_FOREST_UDF.bas`, or the UDF function names post-change
- **THEN** there are zero hits outside the README migration note

### Requirement: Single settings workbook
Macros SHALL read configuration from `PyexcelSettings.xlsx` (sheets `Forecast` and `IsolationForest`) located beside the scripts, auto-creating it with commented defaults on first run. `ForecastSettings.xlsx` SHALL no longer be read; migration is documented, not automatic.

#### Scenario: First macro run scaffolds settings
- **WHEN** a macro runs and no settings workbook exists
- **THEN** it creates `PyexcelSettings.xlsx` with both tool sheets populated with defaults and tells the user to configure and re-run

#### Scenario: Isolation forest multi-sheet mode is a single pooled fit
- **WHEN** the IsolationForest settings list multiple data sheets
- **THEN** all sheets' valid rows are pooled into one fit and scored together, and the output sheet's `Source Sheet` column identifies each row's origin

#### Scenario: The former Bayesian/sequential mode is gone
- **WHEN** searching isolation_forest.py for Beta-prior contamination updating or cumulative per-sheet scoring
- **THEN** both are absent. The Beta updating was measured during apply to be a mathematical no-op (the threshold is the contamination-quantile, so the flagged fraction equals the prior by construction; drift across 6 heterogeneous sheets was 0.002-0.004, all quantile rounding). The cumulative no-lookahead fitting underneath it was then removed as well, by decision, leaving one scoring path. The genuine Bayesian MCMC mode in tabpfn_forecast.py (Bambi/PyMC) is unaffected.

### Requirement: Setup scripts work on current conda and cover both tools
`setup.sh`/`setup.bat` SHALL NOT use `conda env create --force` (absent from current conda; use `--yes`/recreate logic) and SHALL print next steps for both tools including the VBA module import.

#### Scenario: Setup on current conda
- **WHEN** setup runs against a current conda release
- **THEN** env creation succeeds and the printed steps name `pyexcel_macros.bas`, both scripts, and the settings workbook
