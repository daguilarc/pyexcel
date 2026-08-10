# Proposal: omni-rule-compliance-refactor

## Why

Two problems, one root cause. First, the invocation story is fractured: four configuration mechanisms across two tools (type-sniffed UDF varargs, argparse flags, a settings workbook hardcoded beside the script, a legacy VBA wrapper), documented by a 437-line README. Second — found by tracing the README's claims against the xlwings docs instead of repeating them — most of that story is fictional on macOS: UDFs are Windows-only ("UDFs are currently only available on Windows", docs.xlwings.org/en/stable/udfs.html) and the Alt+F8 macro path (`@xw.sub`, tabpfn_forecast.py:733) is the Windows-only Import Functions mechanism; the repo ships no cross-platform `RunPython` VBA stub (git ls-files: only `ISOLATION_FOREST_UDF.bas`, a UDF-shaped Function). On a Mac, only the CLI has ever worked. On top of that, two confirmed defects mean the headline paths don't run anywhere: TabPFN/GradientBoosting always raises `Unknown model_type` (tabpfn_forecast.py:405-406, misattached `else`, AST-verified), and Monte Carlo mode crashes on ~1023/1024 iterations (isolation_forest.py:387, seeds up to 2^42 vs numpy's 2^32-1 cap, verified `ValueError`).

Decision (user, 2026-08-09): drop the UDF surface entirely. Each tool gets exactly two surfaces — an Excel macro via `RunPython` (the only cross-platform macro mechanism) and a CLI — with one shared configuration story.

## What Changes

- **Fix broken model dispatch** — re-attach the `Unknown model_type` guard to the model-dispatch chain (tabpfn_forecast.py:364-406) so TabPFN/GradientBoosting runs succeed.
- **Fix Monte Carlo seeding** — draw per-iteration seeds from `[0, 2**32)` (isolation_forest.py:387).
- **BREAKING: remove the UDF surface** — delete `ISOLATION_FOREST`/`ANOMALY_BINARY` `@xw.func` functions (isolation_forest.py:401-747) and the UDF-only plumbing that dies with them (`_parse_udf_args`, `_arrays_to_numpy`, `_build_excel_output`, `_build_array_column`, and the ~200 lines of docstring essays). Existing workbook formulas stop working; migration is the macro or CLI.
- **BREAKING: delete `ISOLATION_FOREST_UDF.bas`**; ship `pyexcel_macros.bas` instead — two Subs (`RunIsolationForest`, `RunForecast`) that call `RunPython`, imported once per workbook. Works on Windows and Mac; the Windows-only `@xw.sub` decorator is removed.
- **One configuration story**: a single settings workbook (`PyexcelSettings.xlsx`, one sheet per tool, auto-created on first macro run, replacing `ForecastSettings.xlsx`). Macros read it; CLIs resolve flag > settings workbook > default, with `--settings <path>` on both tools.
- **Isolation forest gains a macro entry + settings sheet** (it has none today — grep: only `@xw.func` at isolation_forest.py:401, 584). Its settings mirror its CLI flags and accept multiple data sheets, carrying the multi-array sequential mode (isolation_forest.py:152-232) over as multi-sheet mode.
- **Remove the no-op "Bayesian updating"** (found during apply, by measurement): the Beta update cannot move, because the binary threshold is the contamination-quantile, so the flagged fraction equals the prior by construction — posterior mean = (c·100 + c·n)/(100+n) = c. Measured drift across 6 heterogeneous sheets (n=31..500): 0.0018-0.0042, entirely quantile rounding. The real behavior underneath — cumulative no-lookahead fitting — is kept and renamed `Cumulative Sheets` / `--cumulative`.
- **Fix the Diagnostics sheet, which never contained anything** (found during apply, by end-to-end test): its section markers were written as `=== MODEL PERFORMANCE ===`. Excel parses a cell string starting with `=` as a formula, that one is invalid, and the invalid formula makes the ENTIRE bulk assignment silently discard — no exception, exit code 0, `used_range` stuck at `$A$1`. Isolated by a 5-variant control set: `===` writes nothing; `---`, plain text, and apostrophe-escaping all persist; a single `=x` cell crashed Excel outright. Markers are now `---`. This predates the refactor — every Diagnostics sheet this tool ever produced was empty.
- **Give the CLI a Diagnostics sheet at all** (found during apply): `run_forecast_configured` saved before Diagnostics was built and nothing saved afterward, which a macro survives (workbook stays open) but a CLI does not (process exits). There is now a single save point after every sheet is written.
- **Fix silent feature loss on every sheet read** (found during apply): `.options(pd.DataFrame, header=1)` (isolation_forest.py:787 originally) relies on xlwings' converter default `index=1` (verified in the installed package source), so the first column silently became the DataFrame index and was dropped from the features — contradicting README:349's "all numeric columns are features". Now reads with `index=False`.
- **Delete dead code** — `detect_header_row`, `clean_dataframe`, `_is_numeric_string` (tabpfn_forecast.py:77-136; grep-verified zero external callers).
- **§8 repetition elimination (within-file, per stated no-shared-module constraint, README:108-110)** — merge the near-identical Bayesian cumulative loops (isolation_forest.py:152-190 vs 193-232); single-source the macro/CLI settings plumbing (tabpfn_forecast.py:820-838 vs 974-992), get-or-create-sheet (tabpfn_forecast.py:666-670, 847-851), gamma offset (×3: 390/547/577), `add_constant` guard (×5: 163/176/374/388/502), MC aggregation (×2: 565-568/608-611), mean+noise (×2: 485-488/492-495).
- **Fix setup scripts** — `conda env create --force` no longer exists on current conda (flag absent from docs.conda.io env/create flag list; `--yes` is current) — setup.sh:24 and setup.bat:24 fail today. Also cover both tools, not just isolation forest.
- **README rewrite** (~120 lines): one macro section, one CLI section, accurate platform statements (macro + CLI on both platforms, nothing else), stale `excelpy/` paths (README:73, 420) fixed, editorializing constructions removed (README:49, 110, 154, 245), CLI help text corrected (tabpfn_forecast.py:941-943 claims settings come from the data workbook; they come from the settings file, tabpfn_forecast.py:951).

## Capabilities

### New Capabilities
- `cli-contract`: uniform command shape, flags, precedence, and exit codes for both tools.
- `excel-surface`: the single in-Excel path — `RunPython` macros + settings workbook — on both platforms.
- `runtime-correctness`: dispatch and MC seeding work; no dead or unreachable code; §8 duplication eliminated.
- `documentation`: single-source README with platform claims traced to xlwings docs.

### Modified Capabilities
(none — first spec-driven change in this repo)

## Impact

- **Code**: isolation_forest.py (UDFs removed, macro + settings core added, CLI kept), tabpfn_forecast.py (dispatch fix, flags, `@xw.sub` removed, macro entry kept as plain function), ISOLATION_FOREST_UDF.bas deleted, pyexcel_macros.bas added, setup.sh/setup.bat fixed.
- **Users**: UDF formula users (Windows-only by construction) must switch to the macro or CLI; `ForecastSettings.xlsx` users rename/migrate to `PyexcelSettings.xlsx` (documented; macro auto-creates the new file).
- **Dependencies**: none added or removed.
