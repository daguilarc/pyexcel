# documentation

## ADDED Requirements

### Requirement: README is the single usage reference
README.md SHALL be the sole normative usage documentation, target ≤ ~120 lines: what the tools are; install; the macro surface (once — import `pyexcel_macros.bas`, configure `PyexcelSettings.xlsx`, Alt+F8); the CLI surface (once — command shape, precedence, flags per tool); migration notes (UDF removal, settings rename); one compact platform-notes table. Each fact SHALL appear exactly once.

#### Scenario: No duplicated prose
- **WHEN** reviewing the rewritten README
- **THEN** quoting rules, env activation, settings location, and macro import steps each appear in exactly one place

#### Scenario: Stale naming removed
- **WHEN** searching the README for `excelpy` as a path/folder (currently README:73, 420) or for `ForecastSettings.xlsx` outside the migration note
- **THEN** there are no hits (the conda env name may be stated once)

### Requirement: Platform claims trace to xlwings documentation
The README SHALL state that both surfaces (RunPython macro, CLI) work identically on Windows and macOS, and SHALL NOT document any Windows-only mechanism (UDFs, `@xw.script`/`@xw.sub`). This corrects README:35, 38, 44, 83-87, whose Mac UDF/macro claims contradict the xlwings docs.

#### Scenario: No untraceable platform claims
- **WHEN** reviewing platform statements in the rewritten README
- **THEN** every mechanism named is one the repo actually ships, with no per-platform capability forks

### Requirement: No editorializing constructions
The README SHALL state facts directly — no contrastive framing ("not X, but Y", "X is **not** Y-only"), no rhetorical reassurance, no self-referential process notes (current instances: README:49, 110, 154, 245).

#### Scenario: Snark purge verified
- **WHEN** reviewing the rewritten README for these constructions
- **THEN** none remain

### Requirement: Docstrings defer to README
Function docstrings SHALL be at most a one-line summary, compact parameter list, and README pointer; the parameter essays and hardware-sizing speculation die with the UDFs, and no surviving docstring exceeds ~25 lines or documents a parameter normatively in two places.

#### Scenario: Docstring size bounded
- **WHEN** the change is applied
- **THEN** no docstring in either script exceeds ~25 lines
