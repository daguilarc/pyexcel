# cli-contract

## ADDED Requirements

### Requirement: Uniform command shape
Every tool SHALL be invocable as `python <tool>.py <workbook.xlsx> [--flags]`: workbook path as the only positional argument, all options as `--kebab-case` flags.

#### Scenario: Isolation forest flags preserved
- **WHEN** a user runs `python isolation_forest.py data.xlsx --contamination 0.05 --monte-carlo-samples 100`
- **THEN** the run completes with those parameters and writes the output sheet

#### Scenario: Forecast accepts flags
- **WHEN** a user runs `python tabpfn_forecast.py data.xlsx --target-column Sales --model-type OLS`
- **THEN** the forecast runs with those values without requiring a settings workbook

### Requirement: Uniform configuration precedence
Both CLIs SHALL resolve every configuration key as: explicit flag > settings workbook (`PyexcelSettings.xlsx`, tool's sheet) > built-in default. Both SHALL accept `--settings <path>`; the default settings location is beside the scripts.

#### Scenario: Flag overrides settings workbook
- **WHEN** the settings workbook sets Model Type to TabPFN and the user passes `--model-type OLS`
- **THEN** the run uses OLS and prints which source each overridden key came from

#### Scenario: Settings workbook alone suffices
- **WHEN** a user passes only a workbook path and the tool's settings sheet is populated
- **THEN** the run completes using those settings

#### Scenario: Missing required config
- **WHEN** the forecast runs with no target column from flag or settings
- **THEN** it exits 1 naming the missing key and both ways to supply it

### Requirement: Exit codes and help text are truthful
Both CLIs SHALL exit 0 on success, 1 on error with a stderr message, and `--help` SHALL correctly state where configuration is read from (fixes tabpfn_forecast.py:941-943, which claims a "Settings sheet" in the data workbook).

#### Scenario: Corrected forecast help text
- **WHEN** a user runs `python tabpfn_forecast.py --help`
- **THEN** the text states configuration comes from flags or the settings workbook, and names the default settings path
