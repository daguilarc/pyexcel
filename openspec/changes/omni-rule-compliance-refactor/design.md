# Design: omni-rule-compliance-refactor

## Context

Trace ledger (every mechanism claim read or labeled; re-audited 2026-08-09 after the first pass repeated two unread README claims):

- **UDFs are Windows-only** — docs.xlwings.org/en/stable/udfs.html: "UDFs are currently only available on Windows." VERIFIED (docs quote). README:35/44/83-87 claim otherwise.
- **`@xw.sub` macros are Windows-only** — same page: "On Windows, as an alternative to calling macros via RunPython, you can also use the `@xw.script` decorator" (`@xw.sub` is its deprecated alias). VERIFIED (docs quote). So `RunForecastWithSettings` (tabpfn_forecast.py:733) is unreachable on Mac; the repo ships no `RunPython` VBA stub (git ls-files) — README:38/85 claim otherwise.
- **`RunPython` is the cross-platform macro mechanism** — quickstart docs set it up "from either a command prompt on Windows or a terminal on Mac"; requires the xlwings add-in (or standalone mode) plus a VBA Sub. VERIFIED (docs). RunPython executes code with results written back through `xw.Book.caller()`; it does not hand expression values back to VBA — CORROBORATED by the docs' canonical example shape, exact contract UNVERIFIED-by-quote (moot: the Function-shaped ISOLATION_FOREST_UDF.bas:74 is deleted regardless).
- **Defects**: misattached `else` (tabpfn_forecast.py:400/405-406, AST: line-400 `If` owns the line-406 `orelse`); MC seed range (isolation_forest.py:387, numpy `ValueError` reproduced); dead trio (tabpfn_forecast.py:77-136, grep); `conda env create --force` absent from current conda's documented flags — setup.sh:24/setup.bat:24 fail on current conda. All VERIFIED.
- **Isolation forest has no macro entry** — grep: only `@xw.func` (isolation_forest.py:401, 584). VERIFIED.
- **Constraint (author-stated)**: no shared `_core.py`; each script self-contained (README:108-110).

## Goals / Non-Goals

**Goals:**
- Exactly two surfaces per tool, identical on Mac and Windows: `RunPython` macro and `python <tool>.py <workbook> [--flags]`.
- One configuration story: settings workbook for macros; flag > settings > default for CLIs.
- Both headline model paths run; §8/§12 compliance; README ≤ ~120 lines with only traced platform claims.

**Non-Goals:**
- No shared module across scripts (constraint honored; the settings-workbook *format* is shared convention, not shared code).
- No packaging, no test-suite bootstrap beyond the verification runs specced here.
- No new statistical methods invented. (Amended during apply: the existing "Bayesian updating" was MEASURED to be a no-op and was removed rather than reimplemented — removing a false claim is not designing a new one. Building a genuine cross-sheet drift estimator would be a separate change.)
- No compatibility shim for deleted UDF formulas (user decision: surface removed).

## Decisions

1. **`RunPython` + shipped VBA module, not `@xw.sub`/`@xw.script`.** The decorators are Windows-only (traced above); keeping them means platform-forked docs — the exact disease being cured. `pyexcel_macros.bas` carries two Subs (`RunIsolationForest`, `RunForecast`) each doing `RunPython "import <tool>; <tool>.<entry>()"`; users import the module once per workbook (Alt+F11 → Import, then Alt+F8). *Alternative rejected*: keep `@xw.sub` on Windows as a convenience — two documented paths for one action.

2. **Single settings workbook `PyexcelSettings.xlsx`, one sheet per tool (`Forecast`, `IsolationForest`), auto-created on first macro run.** Macros cannot take arguments, so a settings source is mandatory; one file with per-tool sheets beats two files (fewer artifacts, one README section). Default location stays beside the scripts (macro compatibility, matching today's tabpfn_forecast.py:72-74 behavior); CLIs override with `--settings`. *Alternatives rejected*: per-tool settings files (artifact sprawl); settings inside the data workbook (pollutes user data, breaks the multi-workbook case).

3. **Isolation forest settings mirror its CLI flags 1:1** (data sheets, contamination, monte-carlo samples, ascending, epsilon, output sheet) **and accept multiple data sheets**, mapping the UDF multi-array sequential mode (isolation_forest.py:152-232) onto sheets as `Cumulative Sheets`. Apply-time measurement split that mode in two: cumulative no-lookahead fitting is real (max |cumulative - pooled| score delta = 0.165 on a two-sheet control) and is kept; the Beta contamination updating layered on top is a no-op and is deleted. The type-sniffing parser and Excel-array output shaping (`_parse_udf_args`, `_arrays_to_numpy`, `_build_excel_output`, `_build_array_column`) die with the UDFs — output now lands on a sheet with row alignment, like the CLI already does (isolation_forest.py:853-865).

4. **Bug fixes land first** (dispatch `else`, MC seed) so refactor steps are checked against a baseline that runs.

5. **§8 within-file consolidation, §6 2-of-4 gate per helper.** The two Bayesian loops merge into one parameterized sequential-scoring path (the line-221 threshold variant becomes a parameter — classified deliberate, not forced-merged). Macro and CLI entry points in each file route through one settings→run function, verified at the call graph. tabpfn micro-dedups per proposal list; get-or-create-sheet stays inline where it occurs once.

6. **README owns usage; docstrings shrink to ≤ ~25 lines.** Two normative sources already contradict (README:237-245 vs tabpfn_forecast.py:941-943).

## Risks / Trade-offs

- [Deleting UDFs breaks existing Windows workbooks silently — formulas show #NAME?] → README migration section + release note; git history retains the code.
- [`PyexcelSettings.xlsx` rename strands ForecastSettings.xlsx users] → macro auto-creates the new file; README one-line migration (copy values over). No silent auto-import (magic fallback paths are how the current mess started).
- [RunPython startup latency per macro call vs persistent UDF server] → acceptable; forecast runs are seconds-to-minutes anyway; noted in README.
- [MC results differ from any pre-fix run that ever succeeded] → pre-fix MC crashed ~always; note non-reproducibility once.
- [VBA module import is a manual step users may miss] → macro section of README is step 1-2-3; setup scripts print it.

## Migration Plan

Single branch. Order: bug fixes → settings core + CLI flags → macro module + UDF removal → dedup → docs. Rollback = revert. Verification (positive controls per §9.1): GradientBoosting and OLS forecast runs complete; isolation CLI with `--monte-carlo-samples 100` completes with iterations confirmed executed; macro Subs invoke both entries on a live workbook (Windows or Mac, whichever is at hand — the mechanism is identical by construction); `--help` on both tools matches README.

## Open Questions

- None blocking. (Resolved this revision: conda `--force` is gone on current conda → use `--yes`; UDF-on-Mac and macro-on-Mac claims were false → surfaces redesigned around RunPython.)
