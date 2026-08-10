#!/usr/bin/env python3
"""
Isolation Forest Anomaly Detection

Two entry points, both routed through run_isolation_forest_configured():
- run_isolation_forest() - macro entry point, invoked from VBA via RunPython
- main()                 - CLI entry point (python isolation_forest.py <workbook.xlsx> [flags])

See README.md for configuration details (PyexcelSettings.xlsx, 'IsolationForest' sheet).
"""

import xlwings as xw  # type: ignore
import pandas as pd
import numpy as np
import sys
import os
import argparse
import time
from typing import List
from sklearn.ensemble import IsolationForest


# =============================================================================
# GLOBAL HELPER FUNCTIONS
# =============================================================================

# Constants
RANDOM_SEED = 42  # Base seed for reproducibility

SETTINGS_FILENAME = 'PyexcelSettings.xlsx'
SETTINGS_SHEET_NAME = 'IsolationForest'

# Single source of truth for every settings key: (key, cli flags, settings-sheet
# label, value kind, built-in default). Drives read_settings, create_default_settings,
# the macro entry's cfg dict, and the CLI's argparse setup + precedence resolution.
SETTINGS_SCHEMA = [
    ('sheets', ('--sheets', '--sheet'), 'Data Sheets (comma-separated)', 'list', []),
    ('contamination', ('--contamination',), 'Contamination', 'float', 0.1),
    ('monte_carlo_samples', ('--monte-carlo-samples',), 'Monte Carlo Samples', 'int', 0),
    ('ascending', ('--ascending',), 'Ascending', 'bool', False),
    ('epsilon', ('--epsilon',), 'Epsilon', 'bool', False),
    ('output_sheet', ('--output-sheet',), 'Output Sheet', 'str', 'AnomalyScores'),
]

# CLI help text per key (falls back to a generated string from the label if absent).
_SETTINGS_HELP = {
    'sheets': 'Comma-separated sheet names to process (default: first sheet)',
    'contamination': 'Expected proportion of anomalies (default: 0.1)',
    'monte_carlo_samples': 'Monte Carlo resampling iterations (0=off, 100-10000)',
    'ascending': 'Use ascending scores (high score = highly anomalous)',
    'epsilon': 'Shift scores to ensure all values > 0 (requires --ascending)',
    'output_sheet': 'Output sheet name (default: AnomalyScores)',
}


def _apply_score_transforms(scores: np.ndarray, ascending: bool, epsilon: bool) -> np.ndarray:
    if ascending:
        scores = -scores
    if epsilon:
        percentile_10 = np.percentile(scores, 0.1)
        epsilon_val = max(1e-6, percentile_10 - scores.min() + 1e-6)
        scores = scores - percentile_10 + epsilon_val
    return scores


def _scores_to_binary(scores: np.ndarray, contamination: float, ascending: bool) -> np.ndarray:
    pct = (100 - contamination * 100) if ascending else (contamination * 100)
    threshold_val = np.percentile(scores, pct)
    if ascending:
        return (scores >= threshold_val).astype(int)
    return (scores <= threshold_val).astype(int)


def monte_carlo_score_samples(X_clean, contamination, monte_carlo_samples, stationary_bootstrap_p=0.1):
    """
    Monte Carlo resampling for Isolation Forest scoring.
    Uses stationary bootstrap (Politis & Romano, 1994) with fallback to i.i.d. bootstrap for small samples.

    Stationary bootstrap: starts at a random position; with probability p continues to the
    next observation (wraps around), with probability (1-p) jumps to a random position.
    Expected block length = 1/p. Falls back to i.i.d. bootstrap if sample size < 1/p.

    Args:
        X_clean: Clean data array (no NaN rows)
        contamination: Contamination rate (0-0.5)
        monte_carlo_samples: Number of Monte Carlo iterations (100-10000)
        stationary_bootstrap_p: Probability parameter for stationary bootstrap (default 0.1)

    Returns:
        Aggregated scores (mean across all Monte Carlo iterations)
    """
    throttle_interval = 0.005  # 100 samples per 0.5 seconds (0.5/100)
    start_time = time.time()
    n_samples = len(X_clean)
    scores_buf = np.empty((monte_carlo_samples, n_samples), dtype=np.float64)

    # Calculate minimum sample threshold from expected block length
    expected_block_length = 1.0 / stationary_bootstrap_p
    min_samples_threshold = int(np.ceil(expected_block_length))

    # Use i.i.d. bootstrap if sample size is too small for stationary bootstrap
    use_stationary = n_samples >= min_samples_threshold

    # Create seed sequence for proper parallelization support
    child_ss = np.random.SeedSequence(RANDOM_SEED).spawn(monte_carlo_samples)

    # Pre-allocate indices array for stationary bootstrap (O(1) optimization)
    indices = np.empty(n_samples, dtype=np.int64) if use_stationary else None

    for sample in range(monte_carlo_samples):
        rng = np.random.default_rng(child_ss[sample])

        # Generate bootstrap indices
        if use_stationary:
            # Stationary bootstrap algorithm (Politis & Romano, 1994)
            current_pos = rng.integers(0, n_samples)
            for i in range(n_samples):
                indices[i] = current_pos
                # With probability p, continue to next (wrap around); with probability (1-p), jump to random
                if rng.random() < stationary_bootstrap_p:
                    current_pos = (current_pos + 1) % n_samples
                else:
                    current_pos = rng.integers(0, n_samples)
            X_boot = X_clean[indices]
        else:
            # i.i.d. bootstrap fallback for small samples
            X_boot = X_clean[rng.integers(0, n_samples, size=n_samples)]

        scores_buf[sample] = IsolationForest(
            contamination=contamination,
            random_state=rng.integers(0, 2**32)
        ).fit(X_boot).score_samples(X_clean)

        sleep_time = (sample + 1) * throttle_interval - (time.time() - start_time)
        if sleep_time > 0:
            time.sleep(sleep_time)

    return np.mean(scores_buf, axis=0)


# =============================================================================
# SETTINGS
# =============================================================================

def get_settings_path() -> str:
    """Return the path to PyexcelSettings.xlsx beside this script."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), SETTINGS_FILENAME)


def settings_sheet_exists(path: str) -> bool:
    """Return True if `path` exists and contains the SETTINGS_SHEET_NAME sheet."""
    if not os.path.exists(path):
        return False
    settings_wb = xw.Book(path)
    try:
        return SETTINGS_SHEET_NAME in {s.name for s in settings_wb.sheets}
    finally:
        settings_wb.close()


def _to_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().upper() == 'TRUE'


def read_settings(path: str) -> dict:
    """
    Read the 'IsolationForest' settings sheet (label/value pairs in columns A/B).
    Returns only keys that are present and non-blank; absent or blank keys are
    omitted so callers fall through to built-in defaults.
    """
    settings_wb = xw.Book(path)
    try:
        sheet = settings_wb.sheets[SETTINGS_SHEET_NAME]
        used = sheet.used_range
        rows = used.value if used is not None else []
        if rows and not isinstance(rows[0], (list, tuple)):
            rows = [rows]
        labels = {}
        for row in rows or []:
            if not row or row[0] is None:
                continue
            label = str(row[0]).strip()
            labels[label] = row[1] if len(row) > 1 else None
    finally:
        settings_wb.close()

    def present(label):
        if label not in labels:
            return False
        value = labels[label]
        return not (value is None or (isinstance(value, str) and not value.strip()))

    settings: dict = {}
    for key, _flags, label, kind, _default in SETTINGS_SCHEMA:
        if not present(label):
            continue
        raw_value = labels[label]
        if kind == 'list':
            settings[key] = [s.strip() for s in str(raw_value).split(',') if s.strip()]
        elif kind == 'bool':
            settings[key] = _to_bool(raw_value)
        elif kind == 'int':
            settings[key] = int(raw_value)
        elif kind == 'float':
            settings[key] = float(raw_value)
        else:  # 'str'
            settings[key] = str(raw_value).strip()
    return settings


def create_default_settings(path: str) -> None:
    """
    Create PyexcelSettings.xlsx with a default 'IsolationForest' settings sheet.
    If the workbook already exists (e.g. tabpfn_forecast.py created it with its
    own sheet), open it and add/replace the 'IsolationForest' sheet rather than
    overwriting the file.
    """
    default_rows = [
        [label, '' if kind == 'list' else default]
        for _key, _flags, label, kind, default in SETTINGS_SCHEMA
    ]

    file_exists = os.path.exists(path)
    settings_wb = xw.Book(path) if file_exists else xw.Book()
    try:
        if SETTINGS_SHEET_NAME in {s.name for s in settings_wb.sheets}:
            sheet = settings_wb.sheets[SETTINGS_SHEET_NAME]
            sheet.clear()
        elif file_exists:
            sheet = settings_wb.sheets.add(SETTINGS_SHEET_NAME, after=settings_wb.sheets[-1])
        else:
            sheet = settings_wb.sheets[0]
            sheet.name = SETTINGS_SHEET_NAME
        sheet.range('A1').value = default_rows
        sheet.range('A:A').column_width = 34
        sheet.range('B:B').column_width = 26
        settings_wb.save(path)
    finally:
        settings_wb.close()


# =============================================================================
# SHARED RUN FUNCTION (both entry points call this)
# =============================================================================

def run_isolation_forest_configured(
    wb,
    sheets: List[str],
    contamination: float,
    monte_carlo_samples: int,
    ascending: bool,
    epsilon: bool,
    output_sheet: str,
) -> dict:
    """
    Run Isolation Forest anomaly scoring over one or more sheets and write the
    combined result to `output_sheet`. Shared by the macro entry point and the CLI.

    sheets: list of sheet names to read (empty = [wb.sheets[0].name]).
    """
    sheet_names = list(sheets) if sheets else [wb.sheets[0].name]

    numeric_cols = None
    per_sheet_df: List[pd.DataFrame] = []
    for name in sheet_names:
        sheet = wb.sheets[name]
        data_range = sheet.used_range
        raw = data_range.options(pd.DataFrame, index=False, header=1).value if data_range is not None else None
        df = pd.DataFrame(raw) if raw is not None else pd.DataFrame()
        df = df.dropna(how='all').reset_index(drop=True)
        cols = list(df.select_dtypes(include=[np.number]).columns)
        if numeric_cols is None:
            numeric_cols = cols
        elif cols != numeric_cols:
            raise ValueError(f"Sheet '{name}' has different numeric columns than the first sheet")
        per_sheet_df.append(df)

    if not numeric_cols:
        raise ValueError("No numeric columns found")

    if contamination <= 0 or contamination > 0.5:
        raise ValueError("contamination must be > 0 and <= 0.5")
    if monte_carlo_samples != 0 and (monte_carlo_samples < 100 or monte_carlo_samples > 10000):
        raise ValueError("monte_carlo_samples must be 0 (off) or between 100-10000")
    if epsilon and not ascending:
        raise ValueError("epsilon requires ascending. Epsilon only works with ascending scores.")

    sheet_arrays = [df[numeric_cols].values.astype(np.float64) for df in per_sheet_df]

    # Fit once on every clean row across all sheets, then score.
    valid_masks = [~np.isnan(arr).any(axis=1) for arr in sheet_arrays]
    X_clean = np.vstack([arr[m] for arr, m in zip(sheet_arrays, valid_masks)])

    if len(X_clean) < 2:
        raise ValueError("Need at least 2 valid data points")

    model = IsolationForest(contamination=contamination, random_state=RANDOM_SEED)
    model.fit(X_clean)

    if monte_carlo_samples > 0:
        scores_clean = monte_carlo_score_samples(X_clean, contamination, monte_carlo_samples)
    else:
        scores_clean = model.score_samples(X_clean)

    scores_clean = _apply_score_transforms(scores_clean, ascending, epsilon)
    binary_clean = _scores_to_binary(scores_clean, contamination, ascending)

    scores_per_sheet = []
    binary_per_sheet = []
    pos = 0
    for arr, m in zip(sheet_arrays, valid_masks):
        n_valid = int(np.sum(m))
        s = np.full(len(arr), np.nan)
        b = np.full(len(arr), np.nan)
        s[m] = scores_clean[pos:pos + n_valid]
        b[m] = binary_clean[pos:pos + n_valid]
        scores_per_sheet.append(s)
        binary_per_sheet.append(b)
        pos += n_valid

    output_rows = []
    n_anomalies = 0
    n_samples = 0
    for name, arr, scores, binary in zip(sheet_names, sheet_arrays, scores_per_sheet, binary_per_sheet):
        valid_mask = ~np.isnan(arr).any(axis=1)
        for row_idx in range(len(arr)):
            if not valid_mask[row_idx] or np.isnan(scores[row_idx]):
                continue
            is_anom = binary[row_idx]
            label = 'Anomaly' if is_anom == 1 else 'Normal'
            output_rows.append([name, row_idx + 1] + arr[row_idx].tolist() + [float(scores[row_idx]), label])
            n_samples += 1
            if is_anom == 1:
                n_anomalies += 1

    if output_sheet not in {s.name for s in wb.sheets}:
        out_sht = wb.sheets.add(output_sheet, after=wb.sheets[-1])
    else:
        out_sht = wb.sheets[output_sheet]
        out_sht.clear()

    header = ['Source Sheet', 'Index'] + list(numeric_cols) + ['Anomaly Score', 'Is Anomaly']
    out_sht.range('A1').value = [header]
    if output_rows:
        out_sht.range('A2').value = output_rows

    wb.save()

    return {
        'n_sheets': len(sheet_names),
        'n_samples': n_samples,
        'n_features': len(numeric_cols),
        'n_anomalies': n_anomalies,
        'output_sheet': output_sheet,
    }


# =============================================================================
# MACRO ENTRY POINT (call via Alt+F8 / RunPython)
# =============================================================================

def run_isolation_forest():
    """Macro entry point. Invoked from VBA via RunPython."""
    try:
        wb = xw.Book.caller()
        settings_path = get_settings_path()

        if not settings_sheet_exists(settings_path):
            create_default_settings(settings_path)
            wb.app.alert(
                f"Created '{SETTINGS_FILENAME}' with default configuration.\n\n"
                f"Location: {settings_path}\n\n"
                "Please configure the 'IsolationForest' sheet, then run this macro again."
            )
            return

        raw = read_settings(settings_path)
        cfg = {key: raw.get(key, default) for key, _, _, _, default in SETTINGS_SCHEMA}
        results = run_isolation_forest_configured(
            wb,
            cfg['sheets'],
            cfg['contamination'],
            cfg['monte_carlo_samples'],
            cfg['ascending'],
            cfg['epsilon'],
            cfg['output_sheet'],
        )

        wb.app.alert(
            f"Anomaly Detection Complete!\n\n"
            f"Sheets: {results['n_sheets']}\n"
            f"Samples: {results['n_samples']}\n"
            f"Features: {results['n_features']}\n"
            f"Anomalies: {results['n_anomalies']}\n\n"
            f"See '{results['output_sheet']}' sheet for details."
        )
    except Exception as e:
        xw.Book.caller().app.alert(f"Error: {e}")


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

def main():
    """CLI entry point. python isolation_forest.py <workbook.xlsx> [flags]"""
    parser = argparse.ArgumentParser(description='Isolation Forest Anomaly Detection (CLI mode)')
    parser.add_argument('excel_path', help='Path to Excel workbook')
    for key, flags, label, kind, default in SETTINGS_SCHEMA:
        help_text = _SETTINGS_HELP.get(key, f'{label} (default: {default})')
        if kind == 'bool':
            parser.add_argument(*flags, dest=key, action='store_true', help=help_text)
        elif kind == 'int':
            parser.add_argument(*flags, dest=key, type=int, default=None, help=help_text)
        elif kind == 'float':
            parser.add_argument(*flags, dest=key, type=float, default=None, help=help_text)
        else:  # 'str' or 'list'
            parser.add_argument(*flags, dest=key, default=None, help=help_text)
    parser.add_argument('--settings', default=None, help='Path to settings workbook (default: PyexcelSettings.xlsx beside this script)')

    args = parser.parse_args()

    try:
        if args.settings:
            if not os.path.exists(args.settings):
                print(f"ERROR: settings file not found: {args.settings}", file=sys.stderr)
                return 1
            if not settings_sheet_exists(args.settings):
                print(f"ERROR: settings file '{args.settings}' does not contain the '{SETTINGS_SHEET_NAME}' sheet", file=sys.stderr)
                return 1
            settings = read_settings(args.settings)
        else:
            default_settings_path = get_settings_path()
            settings = read_settings(default_settings_path) if settings_sheet_exists(default_settings_path) else {}

        cfg = {}
        reported = []
        for key, _flags, _label, kind, default in SETTINGS_SCHEMA:
            flag_value = getattr(args, key)
            if kind == 'bool':
                if flag_value:
                    value, from_settings = True, False
                elif key in settings:
                    value, from_settings = settings[key], True
                else:
                    value, from_settings = default, False
            elif kind == 'list':
                if flag_value is not None:
                    value = [s.strip() for s in flag_value.split(',') if s.strip()]
                    from_settings = False
                elif key in settings:
                    value = settings[key]
                    from_settings = True
                else:
                    value = default
                    from_settings = False
            else:
                if flag_value is not None:
                    value, from_settings = flag_value, False
                elif key in settings:
                    value, from_settings = settings[key], True
                else:
                    value, from_settings = default, False
            cfg[key] = value
            if from_settings:
                reported.append((key, value))

        for name, value in reported:
            print(f"settings: {name}={value}")

        wb = xw.Book(args.excel_path)
        results = run_isolation_forest_configured(
            wb, cfg['sheets'], cfg['contamination'], cfg['monte_carlo_samples'],
            cfg['ascending'], cfg['epsilon'], cfg['output_sheet'],
        )

        n_samples = results['n_samples']
        n_anomalies = results['n_anomalies']
        print(f"\n{'='*50}")
        print(f"Anomaly Detection Complete!")
        print(f"Sheets: {results['n_sheets']}")
        print(f"Features: {results['n_features']}")
        print(f"Samples: {n_samples}")
        if n_samples:
            print(f"Anomalies: {n_anomalies} ({n_anomalies/n_samples:.2%})")
        else:
            print(f"Anomalies: {n_anomalies}")
        print(f"See '{results['output_sheet']}' sheet for details.")

        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
