#!/usr/bin/env python3
"""
TabPFN Forecast

Two entry points, both routed through run_forecast_configured() via _run_forecast():
- run_forecast() - macro entry point, invoked from VBA via RunPython
- main()         - CLI entry point (python tabpfn_forecast.py <workbook.xlsx> [flags])

See README.md for configuration details (PyexcelSettings.xlsx, 'Forecast' sheet).
"""

import os
import xlwings as xw  # type: ignore
import pandas as pd
import numpy as np
import sys
import argparse
from typing import Optional, List, Tuple
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error

# Licensing verified 2026-10-26: recent TabPFN weights (2.5 and later) on
# Hugging Face use non-commercial, non-production licenses for research/evaluation.
# HF authentication grants download access, not commercial usage rights. TabPFN-2
# has a different license; check the exact checkpoint's terms before use:
# https://github.com/PriorLabs/TabPFN#license
# https://huggingface.co/Prior-Labs/tabpfn_2_5/blob/main/LICENSE
# Prior Labs' hosted API is separate and permits commercial use under its terms:
# https://priorlabs.ai/pricing
# Optional imports
try:
    from tabpfn import TabPFNRegressor # type: ignore
    HAS_TABPFN = True
except ImportError:
    HAS_TABPFN = False
    TabPFNRegressor = None  # type: ignore



# =============================================================================
# GLOBAL HELPER FUNCTIONS
# =============================================================================

SETTINGS_FILENAME = 'PyexcelSettings.xlsx'
SETTINGS_SHEET_NAME = 'Forecast'

# Single source of truth for every settings key: (key, cli flags, settings-sheet
# label, value kind, built-in default). Drives read_settings, create_default_settings,
# the macro entry's cfg dict, and the CLI's argparse setup + precedence resolution.
SETTINGS_SCHEMA = [
    ('data_sheets', ('--data-sheets',), 'Data Sheets (comma-separated)', 'list', []),
    ('header_row', ('--header-row',), 'Header Row (1-based)', 'int', 1),
    ('data_start_row', ('--data-start-row',), 'Data Start Row (1-based)', 'int', 2),
    ('target_column', ('--target-column',), 'Target Column', 'str', None),
    ('feature_columns', ('--feature-columns',), 'Feature Columns (or "all")', 'str', 'all'),
    ('categorical_columns', ('--categorical-columns',), 'Categorical Columns', 'list', []),
    ('primary_time_index', ('--primary-time-index',), 'Primary Time Index', 'str', None),
    ('secondary_time_index', ('--secondary-time-index',), 'Secondary Time Index', 'str', None),
    ('model_type', ('--model-type',), 'Model Type', 'str', 'TabPFN'),
    ('n_ensemble', ('--ensemble-size',), 'Ensemble Size (TabPFN)', 'int', 50),
    ('n_gb_estimators', ('--gb-estimators',), 'GB Estimators (GradientBoosting)', 'int', 100),
    ('future_rows', ('--future-rows',), 'Future Rows', 'int', 10),
    ('output_sheet', ('--output-sheet',), 'Output Sheet', 'str', 'Forecast'),
]

# CLI help text per key (falls back to a generated string from the label if absent).
_SETTINGS_HELP = {
    'data_sheets': 'Comma-separated sheet names to read',
    'header_row': 'Row containing column headers (1-based, default: 1)',
    'data_start_row': 'Row where data starts (1-based, default: 2)',
    'target_column': 'Name of target variable',
    'feature_columns': 'Comma-separated feature names, or "all" (default: all)',
    'categorical_columns': 'Comma-separated categorical column names',
    'primary_time_index': 'Primary time column name (e.g., "Year")',
    'secondary_time_index': 'Secondary time column name (e.g., "Date", "Quarter")',
    'model_type': 'Model to use: TabPFN, GradientBoosting (default: TabPFN)',
    'n_ensemble': 'TabPFN ensemble size (default: 50)',
    'n_gb_estimators': 'GradientBoosting n_estimators (default: 100)',
    'future_rows': 'Rows to generate in output (default: 10)',
    'output_sheet': 'Output sheet name (default: Forecast)',
}


def _apply_fixed_effects_demean(
    X: np.ndarray,
    y: np.ndarray,
    groups: np.ndarray,
    feature_cols: List[str],
) -> Tuple[np.ndarray, np.ndarray]:
    """Within transformation: subtract group means (vectorized)."""
    gdf = pd.DataFrame(X, columns=feature_cols)
    gdf['__y__'] = y
    gdf['__g__'] = groups
    means = gdf.groupby('__g__', sort=False).transform('mean')
    y_out = (gdf['__y__'] - means['__y__']).to_numpy()
    X_out = (gdf[feature_cols] - means[feature_cols]).to_numpy()
    return X_out, y_out


def _mean_noise_future_features(X_ref: np.ndarray, future_rows: int, n_features: int) -> np.ndarray:
    """Generate future feature rows as historical mean + Gaussian noise scaled by historical std."""
    means = np.mean(X_ref, axis=0)
    stds = np.std(X_ref, axis=0)
    feature_stds = np.where(stds == 0, 1.0, stds)
    return means + np.random.normal(0, 0.5, (future_rows, n_features)) * feature_stds


def _get_or_create_sheet(wb, sheet_name: str, wb_sheet_names: set):
    """Return sheet_name, adding it after the last sheet if absent, else clearing and reusing it."""
    if sheet_name not in wb_sheet_names:
        return wb.sheets.add(sheet_name, after=wb.sheets[-1])
    sheet = wb.sheets[sheet_name]
    sheet.clear()
    return sheet


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


def read_settings(path: str) -> dict:
    """
    Read the 'Forecast' settings sheet (label/value pairs in columns A/B).
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
        elif kind == 'int':
            settings[key] = int(raw_value)
        elif kind == 'float':
            settings[key] = float(raw_value)
        else:  # 'str'
            settings[key] = str(raw_value).strip()
    return settings


def create_default_settings(path: str) -> None:
    """
    Create PyexcelSettings.xlsx with a default 'Forecast' settings sheet.
    If the workbook already exists, open it and add/replace the 'Forecast' sheet rather than
    overwriting the file.
    """
    default_rows = [
        [label, '' if (kind == 'list' or default is None) else default]
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


def run_forecast_configured(
    wb,
    data_sheets: List[str],
    header_row: int,
    data_start_row: int,
    target_column: str,
    feature_columns: str,
    categorical_columns: List[str],
    primary_time_index: Optional[str] = None,
    secondary_time_index: Optional[str] = None,
    model_type: str = 'TabPFN',
    n_ensemble: int = 50,
    n_gb_estimators: int = 100,
    future_rows: int = 10,
    output_sheet: str = 'Forecast'
) -> dict:
    """
    Run forecast with explicit user configuration.

    Parameters:
        wb: xlwings Book object
        data_sheets: List of sheet names to read (must be identically formatted)
        header_row: Row containing column headers (1-based)
        data_start_row: Row where data starts (1-based)
        target_column: Name of target variable
        feature_columns: Comma-separated feature names, or "all" for all except target
        categorical_columns: List of categorical column names
        primary_time_index: Primary time column name (e.g., "Year") - if None, uses mean+noise
        secondary_time_index: Secondary time column name (e.g., "Date", "Quarter") - optional
        model_type: Model to use ('TabPFN', 'GradientBoosting')
        n_ensemble: TabPFN ensemble size
        n_gb_estimators: GradientBoosting n_estimators
        future_rows: Rows to generate in output
        output_sheet: Name of output sheet
    """
    if model_type.lower() not in ('tabpfn', 'gradientboosting'):
        raise ValueError("Choose model_type TabPFN or GradientBoosting")
    if future_rows < 1:
        raise ValueError("future_rows must be positive")

    # Cache sheet names for O(1) lookup (used multiple times)
    wb_sheet_names = {s.name for s in wb.sheets}

    # Read and combine data from all specified sheets
    all_data = []
    columns = None

    for sheet_name in data_sheets:
        sheet = wb.sheets[sheet_name]

        # Read header row (convert to 0-based for indexing)
        header_range = sheet.range((header_row, 1)).expand('right')
        headers = [str(h).strip() if h else f'Col{i}' for i, h in enumerate(header_range.value)]

        if columns is None:
            columns = headers
        elif headers != columns:
            raise ValueError(f"Sheet '{sheet_name}' has different columns than first sheet")

        # Read data starting from data_start_row
        data_range = sheet.range((data_start_row, 1)).expand('table')
        data_values = data_range.value

        if data_values:
            # Handle single row case
            if not isinstance(data_values[0], list):
                data_values = [data_values]
            all_data.extend(data_values)

    print(f"Read {len(all_data)} rows from {len(data_sheets)} sheets")

    # Create DataFrame
    df = pd.DataFrame(all_data, columns=columns)

    # Convert non-categorical columns to numeric
    for col in df.columns:
        if col not in categorical_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        else:
            # Keep categoricals as-is (can be strings or integers)
            df[col] = df[col].astype(str)

    # Determine feature columns (excluding categoricals - they'll be handled as fixed effects)
    if feature_columns.lower().strip() == 'all':
        feature_cols = [c for c in df.columns if c != target_column and c not in categorical_columns]
    else:
        feature_cols = [c.strip() for c in feature_columns.split(',') if c.strip() and c not in categorical_columns]

    print(f"Target: {target_column}")
    print(f"Features ({len(feature_cols)}): {feature_cols[:5]}{'...' if len(feature_cols) > 5 else ''}")
    if categorical_columns:
        print(f"Fixed effects (categorical): {categorical_columns}")

    # Extract arrays
    X = df[feature_cols].values.astype(np.float64)
    y = df[target_column].values.astype(np.float64)

    # Apply fixed effects transformation (within transformation: demean by group)
    if categorical_columns:
        # Create group identifier from categorical columns
        if len(categorical_columns) == 1:
            groups = df[categorical_columns[0]].values
        else:
            # Combine multiple categoricals into single group identifier (vectorized)
            # Convert to string array first, then join - more efficient than per-row conversion
            cat_values = df[categorical_columns].astype(str).values
            groups = np.array(['_'.join(row) for row in cat_values])

        X, y = _apply_fixed_effects_demean(X, y, groups, feature_cols)
        n_groups = len(np.unique(groups))
        print(f"Applied fixed effects transformation for {n_groups} groups")

    # Remove NaN rows
    valid = ~(np.isnan(X).any(axis=1) | np.isnan(y))
    X, y = X[valid], y[valid]
    df_valid = df.iloc[valid].reset_index(drop=True)  # Keep df aligned with filtered X, y
    n_samples, n_features = X.shape

    print(f"Training: {n_samples} samples, {n_features} features")

    # Step 1: Initial fit based on model_type
    model_type_lower = model_type.lower()
    if model_type_lower not in ('tabpfn', 'gradientboosting'):
        raise ValueError(f"Unknown model_type: {model_type}. Choose from: TabPFN, GradientBoosting")
    model = None

    if model_type_lower == 'tabpfn':
        if not HAS_TABPFN:
            print("TabPFN not available. Falling back to GradientBoosting.")
            model_type_lower = 'gradientboosting'  # Fall back
        else:
            try:
                model = TabPFNRegressor(
                    n_estimators=n_ensemble,
                    memory_saving_mode='auto',
                    inference_precision='auto'
                )
                model.fit(X, y)
                print(f"Using TabPFN (n_estimators={n_ensemble})")
            except RuntimeError as e:
                if "Authentication" in str(e) or "gated" in str(e).lower():
                    print("TabPFN requires authentication. Falling back to GradientBoosting.")
                    model = None
                    model_type_lower = 'gradientboosting'  # Fall back
                else:
                    raise

    # Handle fallback to GradientBoosting (either explicit choice or fallback from TabPFN)
    if model_type_lower == 'gradientboosting':
        if model is None:  # Only create if not already created (fallback case)
            model = GradientBoostingRegressor(n_estimators=n_gb_estimators, random_state=42)
            model.fit(X, y)
            print(f"Using GradientBoosting (n_estimators={n_gb_estimators})")

    y_pred = model.predict(X)
    mae = mean_absolute_error(y, y_pred)
    rmse = np.sqrt(mean_squared_error(y, y_pred))

    print(f"MAE: {mae:.6f}, RMSE: {rmse:.6f}")

    # Step 2: Generate future feature values
    # If time indices provided, project features forward in time; otherwise use mean+noise
    future_time_periods = None
    has_secondary_time = secondary_time_index and secondary_time_index in df_valid.columns

    if primary_time_index and primary_time_index in df_valid.columns:
        # Time-based forecasting: project features forward
        print(f"Using time-based forecasting with primary index: {primary_time_index}")
        if has_secondary_time:
            print(f"  Secondary index: {secondary_time_index}")

        # Extract and convert time columns in one pass
        time_cols = [primary_time_index]
        if has_secondary_time:
            time_cols.append(secondary_time_index)
        time_df = df_valid[time_cols].copy()  # Copy needed - we modify columns in place

        # Vectorized conversion: try datetime first, then numeric
        for col in time_cols:
            if time_df[col].dtype == object:
                time_df[col] = pd.to_datetime(time_df[col], errors='coerce')
                if time_df[col].isna().all():  # If datetime failed, try numeric
                    time_df[col] = pd.to_numeric(df_valid[col], errors='coerce')

        # Sort once and get indices - use argsort for efficiency (only need indices, not full sorted copy)
        sort_idx = time_df.sort_values(by=time_cols).index.values
        time_values = pd.to_numeric(time_df.loc[sort_idx, primary_time_index], errors='coerce').values
        X_sorted = X[sort_idx]

        # Pre-compute valid time mask (same for all features)
        time_valid = ~np.isnan(time_values)

        if np.sum(time_valid) > 1:
            # Compute future times once
            future_times = np.arange(1, future_rows + 1, dtype=np.float64) * np.mean(np.diff(time_values[time_valid])) + time_values[time_valid][-1]
            future_time_periods = future_times

            # Vectorized feature projection: fit linear trends for all features
            X_future = np.zeros((future_rows, n_features))

            # Process all features in vectorized manner where possible
            for feat_idx in range(n_features):
                y_feat = X_sorted[:, feat_idx]
                feat_valid = ~np.isnan(y_feat) & time_valid
                n_feat_valid = np.sum(feat_valid)

                if n_feat_valid > 1:
                    # Linear fit: y = a*t + b (use aligned time and feature values)
                    coeffs = np.polyfit(time_values[feat_valid], y_feat[feat_valid], 1)
                    X_future[:, feat_idx] = np.polyval(coeffs, future_times)
                elif n_feat_valid == 1:
                    # Single valid point: use that value
                    X_future[:, feat_idx] = y_feat[feat_valid][0]
                else:
                    # No valid points: use mean of all non-NaN values (vectorized)
                    valid_vals = y_feat[~np.isnan(y_feat)]
                    X_future[:, feat_idx] = np.mean(valid_vals) if len(valid_vals) > 0 else 0.0
        else:
            # Not enough time data: fall back to mean+noise
            X_future = _mean_noise_future_features(X_sorted, future_rows, n_features)
    else:
        # No time index: use mean + noise
        print("No time index specified, using mean+noise for future features")
        X_future = _mean_noise_future_features(X, future_rows, n_features)

    # Heuristic training-residual bounds, not calibrated prediction intervals.
    y_future = model.predict(X_future)
    y_future_lower = y_future - 1.96 * rmse
    y_future_upper = y_future + 1.96 * rmse

    # Prepare categorical columns for Forecast sheet (sample from historical distribution)
    categorical_future = None
    if categorical_columns:
        if len(categorical_columns) == 1:
            # Sample categorical values from historical data (cache unique values)
            cat_unique = df[categorical_columns[0]].dropna().unique()
            if len(cat_unique) > 0:
                categorical_future = np.random.choice(cat_unique, size=future_rows, replace=True)
        else:
            # Sample combinations from historical data (cache unique combinations)
            cat_combos = df[categorical_columns].drop_duplicates().values
            if len(cat_combos) > 0:
                categorical_future = cat_combos[np.random.randint(0, len(cat_combos), size=future_rows)]

    # Write to output sheet (use cached wb_sheet_names)
    out_sht = _get_or_create_sheet(wb, output_sheet, wb_sheet_names)

    # Build header: Time columns (if available), Index, Categoricals, Features, Prediction, Lower Bound, Upper Bound
    header = []
    if future_time_periods is not None:
        header.append(primary_time_index)
        if has_secondary_time:
            header.append(secondary_time_index)
    header.extend(['Index'] + (categorical_columns if categorical_columns else []) +
                  feature_cols + ['Prediction', 'Lower (approx.)', 'Upper (approx.)'])

    # Build data rows: vectorized where possible, but use list comprehension for mixed types
    # Pre-compute predictions array once (vectorized column stack)
    predictions_arr = np.column_stack([y_future, y_future_lower, y_future_upper])

    # Build rows efficiently: use list comprehension but minimize per-row work
    data_rows = []
    for i in range(future_rows):
        row = []
        # Time columns (if available)
        if future_time_periods is not None:
            row.append(future_time_periods[i])
            if has_secondary_time:
                row.append(future_time_periods[i])  # Placeholder
        # Index (1-based)
        row.append(i + 1)
        # Categorical columns
        if categorical_future is not None:
            if len(categorical_columns) == 1:
                row.append(categorical_future[i])
            else:
                row.extend(categorical_future[i].tolist() if hasattr(categorical_future[i], 'tolist') else list(categorical_future[i]))
        # Features (vectorized slice)
        row.extend(X_future[i].tolist())
        # Predictions (vectorized slice from pre-computed array)
        row.extend(predictions_arr[i].tolist())
        data_rows.append(row)

    out_sht.range('A1').value = [header]
    out_sht.range('A2').value = data_rows

    return {
        'historical_mae': mae,
        'historical_rmse': rmse,
        'n_features': n_features,
        'n_samples': n_samples,
        'feature_columns': feature_cols,
        'target_column': target_column,
        'model_type': 'TabPFN' if model_type_lower == 'tabpfn' else 'GradientBoosting',
    }


def _run_forecast(wb, cfg: dict) -> dict:
    """
    Validate a resolved config dict, run the forecast, and build the Diagnostics
    sheet. Shared by the macro entry point and the CLI so the two surfaces
    cannot drift apart.
    """
    if not cfg.get('data_sheets'):
        raise ValueError("No data sheets specified")
    if not cfg.get('target_column'):
        raise ValueError("No target column specified")

    results = run_forecast_configured(
        wb=wb,
        data_sheets=cfg['data_sheets'],
        header_row=cfg['header_row'],
        data_start_row=cfg['data_start_row'],
        target_column=cfg['target_column'],
        feature_columns=cfg['feature_columns'],
        categorical_columns=cfg['categorical_columns'],
        primary_time_index=cfg['primary_time_index'],
        secondary_time_index=cfg['secondary_time_index'],
        model_type=cfg['model_type'],
        n_ensemble=cfg['n_ensemble'],
        n_gb_estimators=cfg['n_gb_estimators'],
        future_rows=cfg['future_rows'],
        output_sheet=cfg['output_sheet'],
    )

    # Build the Diagnostics sheet (shared by both entry points)
    diag_sheet = _get_or_create_sheet(wb, 'Diagnostics', {s.name for s in wb.sheets})
    diag_content = [
        ['--- MODEL PERFORMANCE ---', ''],
        ['Model Type', results['model_type']],
        ['Target Column', cfg['target_column']],
        ['Number of Features', results['n_features']],
        ['Number of Samples', results['n_samples']],
        ['', ''],
        ['--- RESIDUAL DIAGNOSTICS ---', ''],
        ['Mean Absolute Error (MAE)', f"{results['historical_mae']:.6f}"],
        ['Root Mean Squared Error (RMSE)', f"{results['historical_rmse']:.6f}"],
    ]

    diag_content.extend([
        ['', ''],
        ['--- MODEL CONFIGURATION ---', ''],
        ['Ensemble Size (TabPFN)', cfg['n_ensemble']],
        ['GB Estimators (GradientBoosting)', cfg['n_gb_estimators']],
        ['Sheets Processed', ', '.join(cfg['data_sheets'])]
    ])

    diag_sheet.range('A1').value = diag_content
    diag_sheet.range('A:A').column_width = 35
    diag_sheet.range('B:B').column_width = 25

    # Single save point: every sheet this run writes is committed once.
    # Section markers above use '---', not '==='. Excel parses a cell string
    # starting with '=' as a formula, and one invalid formula silently discards
    # the whole bulk assignment (measured: used_range stayed $A$1, no exception).
    wb.save()
    print(f"Saved to '{cfg['output_sheet']}' and 'Diagnostics'")

    return results


# =============================================================================
# MACRO ENTRY POINT (call via Alt+F8 / RunPython)
# =============================================================================

def run_forecast():
    """Macro entry point. Invoked from VBA via RunPython."""
    try:
        wb = xw.Book.caller()
        settings_path = get_settings_path()

        if not settings_sheet_exists(settings_path):
            create_default_settings(settings_path)
            wb.app.alert(
                f"Created '{SETTINGS_FILENAME}' with default configuration.\n\n"
                f"Location: {settings_path}\n\n"
                f"Please configure the '{SETTINGS_SHEET_NAME}' sheet, then run this macro again."
            )
            return

        raw = read_settings(settings_path)
        cfg = {key: raw.get(key, default) for key, _, _, _, default in SETTINGS_SCHEMA}

        results = _run_forecast(wb, cfg)

        wb.sheets[cfg['output_sheet']].activate()

        wb.app.alert(
            f"Forecast Complete!\n\n"
            f"Sheets processed: {len(cfg['data_sheets'])}\n"
            f"Target: {cfg['target_column']}\n"
            f"Features: {results['n_features']}\n"
            f"Samples: {results['n_samples']}\n"
            f"MAE: {results['historical_mae']:.4f}\n"
            f"RMSE: {results['historical_rmse']:.4f}\n\n"
            f"See 'Diagnostics' sheet for full details."
        )
    except Exception as e:
        xw.Book.caller().app.alert(f"Error: {e}")


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

def main():
    """
    CLI entry point. python tabpfn_forecast.py <workbook.xlsx> [flags]

    Configuration comes from CLI flags, falling back to the settings workbook
    (default: PyexcelSettings.xlsx beside this script, 'Forecast' sheet),
    falling back to built-in defaults.
    """
    parser = argparse.ArgumentParser(description='TabPFN / GradientBoosting Forecast (CLI mode)')
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
    parser.add_argument('--settings', default=None,
                         help=f'Path to settings workbook (default: {SETTINGS_FILENAME} beside this script)')

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
        results = _run_forecast(wb, cfg)

        print(f"\n{'='*50}")
        print(f"Forecast Complete!")
        print(f"Target: {cfg['target_column']}")
        print(f"Features: {results['n_features']}")
        print(f"Samples: {results['n_samples']}")
        print(f"MAE: {results['historical_mae']:.6f}")
        print(f"RMSE: {results['historical_rmse']:.6f}")
        print(f"Predictions: {cfg['future_rows']}")
        print(f"See '{cfg['output_sheet']}' sheet for details.")

        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
