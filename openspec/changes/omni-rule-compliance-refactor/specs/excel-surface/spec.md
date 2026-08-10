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

#### Scenario: Isolation forest multi-sheet cumulative mode
- **WHEN** the IsolationForest settings list multiple data sheets with `Cumulative Sheets` enabled
- **THEN** sheet k is scored against a model fit on sheets 1..k only (no lookahead), with contamination held constant

#### Scenario: No-op Beta updating removed
- **WHEN** searching either script for Beta-prior contamination updating
- **THEN** it is absent — measurement during apply showed the update is a mathematical no-op (the threshold is the contamination-quantile, so the flagged fraction equals the prior by construction; measured drift across 6 heterogeneous sheets was 0.002-0.004, all quantile rounding), so the claim was removed rather than documented

### Requirement: Setup scripts work on current conda and cover both tools
`setup.sh`/`setup.bat` SHALL NOT use `conda env create --force` (absent from current conda; use `--yes`/recreate logic) and SHALL print next steps for both tools including the VBA module import.

#### Scenario: Setup on current conda
- **WHEN** setup runs against a current conda release
- **THEN** env creation succeeds and the printed steps name `pyexcel_macros.bas`, both scripts, and the settings workbook
