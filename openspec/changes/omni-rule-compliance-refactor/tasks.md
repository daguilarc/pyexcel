# Tasks: omni-rule-compliance-refactor

Execution note (omni-rule §4/§15): dispatch implementation subagents on the lightest adequate model (Sonnet or Haiku, never the session default); each subagent receives the task text as immutable, and reports found-vs-changed counts where §8 applies. Structural instructions below are traced claims — if honoring one forces a worse structure, stop and challenge per §1, don't build the workaround.

## 1. Bug fixes (baseline must run before refactoring)

- [ ] 1.1 Fix misattached `else` in tabpfn_forecast.py: move the `Unknown model_type` raise onto the model-dispatch chain (lines 364-397); remove the `else` at 405-406 from the line-400 coefficient-extraction `if`. Verify by AST (line-400 `If` has empty `orelse`) and a completing GradientBoosting run.
- [ ] 1.2 Fix MC seed range at isolation_forest.py:387 to `[0, 2**32)`. Verify with a real `--monte-carlo-samples 100` run; report iteration count actually executed (§9.1 positive control).
- [ ] 1.3 Delete `detect_header_row`, `clean_dataframe`, `_is_numeric_string` (tabpfn_forecast.py:77-136); grep zero references.

## 2. Settings core (one per tool, shared by macro + CLI within each file)

- [ ] 2.1 Implement `PyexcelSettings.xlsx` reader/scaffolder in tabpfn_forecast.py: sheet `Forecast` carrying today's keys (tabpfn_forecast.py:190-206); auto-create with defaults on first run; drop all `ForecastSettings.xlsx` reads.
- [ ] 2.2 Implement sheet `IsolationForest` in isolation_forest.py: data sheets (comma-separated), contamination, monte-carlo samples, ascending, epsilon, bayes updating, output sheet — mirroring CLI flags 1:1.
- [ ] 2.3 Route each file's macro and CLI entries through ONE settings→run function (replaces duplicated plumbing tabpfn_forecast.py:820-838 / 974-992); verify single-sourcing at the call graph.

## 3. CLI contract

- [ ] 3.1 Both tools: `--settings <path>` flag; precedence flag > settings sheet > default, printing the source of each overridden key.
- [ ] 3.2 tabpfn_forecast.py: add flags mirroring every Forecast settings key; correct help/docstring (lines 941-943) to the actual settings resolution.
- [ ] 3.3 Uniform exit codes (0 success / 1 error to stderr) — verify both tools.

## 4. Macro surface (RunPython) and UDF removal

- [ ] 4.1 Add macro entries `run_isolation_forest()` / `run_forecast()` as plain functions using `xw.Book.caller()`; remove the `@xw.sub` decorator (tabpfn_forecast.py:733) — Windows-only mechanism, traced in design.md.
- [ ] 4.2 Create `pyexcel_macros.bas`: Subs `RunIsolationForest` and `RunForecast` calling `RunPython "import <tool>; <tool>.<entry>()"`. Smoke-test via Alt+F8 on a live workbook.
- [ ] 4.3 Delete `@xw.func` UDFs (isolation_forest.py:401-747) and `ISOLATION_FOREST_UDF.bas`. Then run a call-graph orphan sweep over `_parse_udf_args`, `_is_contamination_value`, `_is_mc_samples_value`, `_is_array_index_value`, `_arrays_to_numpy`, `_build_excel_output`, `_build_array_column`, `BOOL_SLOTS_*`: delete each unless the multi-sheet path verifiably calls it (cite the call site to keep).
- [ ] 4.4 Preserve multi-array Bayesian semantics as multi-sheet mode: sequential per-sheet scoring with Beta updates matching isolation_forest.py:152-232; verify against a two-sheet workbook.

## 5. Setup scripts

- [ ] 5.1 Replace `conda env create ... --force` (setup.sh:24, setup.bat:24) with current-conda syntax (`--yes`, or remove-then-create); verify against installed conda at apply time.
- [ ] 5.2 Update next-steps output: both tools, `pyexcel_macros.bas` import, settings workbook — no Import Functions references.

## 6. §8 repetition elimination (within-file, §6 gate per helper)

- [ ] 6.1 Merge the Bayesian cumulative loops (isolation_forest.py:152-190 / 193-232) into one parameterized loop; the line-221 threshold variant is a classified deliberate difference → parameter, not forced merge.
- [ ] 6.2 tabpfn_forecast.py micro-dedup: gamma offset (390/547/577), `add_constant` guard (163/176/374/388/502), MC aggregation (565-568/608-611), mean+noise (485-488/492-495), get-or-create-sheet (666-670/847-851). Single-use occurrences stay inline per §6.
- [ ] 6.3 Report found-vs-changed per concept (operand-based enumeration); behavior check: identical CLI outputs on same inputs before (post-Group-1) vs after.

## 7. Documentation

- [ ] 7.1 Rewrite README.md (≤ ~120 lines) per documentation spec: macro section once, CLI section once, migration notes (UDF removal, settings rename), one platform table with only traced claims (macro + CLI on both platforms); purge editorializing (README:49, 110, 154, 245); fix `excelpy/` paths (README:73, 420).
- [ ] 7.2 Trim all surviving docstrings to ≤ ~25 lines, deferring to README.
- [ ] 7.3 Cross-check README against implemented flags, settings keys, and macro names — every documented item exists, every existing item documented.

## 8. Postflight (omni-rule §14)

- [ ] 8.1 Strict implementation-vs-tasks divergence report; no reinterpretation.
- [ ] 8.2 Fresh §8 enumeration of every helper/constant introduced by Groups 2-6, across both files.
- [ ] 8.3 Evidence bundle: AST output, MC run output with iteration count, GradientBoosting + OLS runs, two-sheet Bayesian run, macro smoke-test note, README line count, greps for removed names and `excelpy`.
