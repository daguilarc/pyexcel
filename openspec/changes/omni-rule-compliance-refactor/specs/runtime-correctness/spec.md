# runtime-correctness

## ADDED Requirements

### Requirement: Model dispatch runs every documented model type
`run_forecast_configured` SHALL execute for TabPFN, GradientBoosting, OLS, and Gamma, raising `Unknown model_type` only for values outside that set (fixes the misattached `else`, tabpfn_forecast.py:400/405-406, AST-verified).

#### Scenario: Default TabPFN run completes
- **WHEN** a forecast runs with Model Type `TabPFN` (falling back to GradientBoosting when tabpfn is absent)
- **THEN** it writes the Forecast sheet and succeeds instead of raising `Unknown model_type`

#### Scenario: Invalid model type still rejected
- **WHEN** a forecast runs with Model Type `RandomForest`
- **THEN** it fails naming the four valid options

### Requirement: Monte Carlo seeding stays within numpy's seed range
All per-iteration `random_state` seeds SHALL be drawn from `[0, 2**32)` (fixes isolation_forest.py:387).

#### Scenario: MC run completes with positive control (omni-rule §9.1)
- **WHEN** `python isolation_forest.py data.xlsx --monte-carlo-samples 100` runs on valid data
- **THEN** all 100 iterations demonstrably execute (iteration count reported) and the output sheet is written

### Requirement: No dead or unreachable code
The following SHALL be removed: the never-called trio `detect_header_row`/`clean_dataframe`/`_is_numeric_string` (tabpfn_forecast.py:77-136, grep-verified), and all UDF-only plumbing orphaned by the surface removal — `_parse_udf_args`, `_is_contamination_value`, `_is_mc_samples_value`, `_is_array_index_value`, `_arrays_to_numpy`, `_build_excel_output`, `_build_array_column`, `BOOL_SLOTS_*` (isolation_forest.py:41-118, 235-321) — except any piece the multi-sheet mode verifiably still calls (trace at removal time, cite the call site or delete).

#### Scenario: Orphan check is call-graph based
- **WHEN** the removal task completes
- **THEN** every retained private helper has a cited caller, and grep finds zero references to the removed names

### Requirement: Repetition eliminated within each file (omni-rule §8)
Near-identical logic occurring 2+ times within a file SHALL be single-sourced at the call graph, each new helper passing §6's 2-of-4 gate: macro/CLI settings plumbing per file, per-setting label/flag/type/default metadata (one `SETTINGS_SCHEMA` table per file, consumed by the reader, the scaffolder, the macro entry and the CLI parser), get-or-create-sheet (tabpfn_forecast.py:666-670/847-851), gamma offset (×3), `add_constant` guard (×5), MC aggregation (×2), mean+noise (×2).

The two Bayesian cumulative loops (isolation_forest.py:152-190/193-232) were originally slated for merging on the premise that the line-221 threshold was a deliberate variant. That premise was FALSE — it is byte-equivalent to `_scores_to_binary(scores, c, ascending=False)`, so the two functions were the same function. They were merged, and the merged result was subsequently deleted outright when the whole cumulative mode was removed.

#### Scenario: Found-vs-changed reporting
- **WHEN** the consolidation task completes
- **THEN** the report lists count found vs count changed per concept, enumerated by operands not syntax, with deliberate non-merges justified

#### Scenario: Behavior preserved where surfaces survive
- **WHEN** identical inputs run through the CLI paths before (post-bug-fix) and after consolidation
- **THEN** outputs are identical
