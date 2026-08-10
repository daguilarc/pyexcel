#!/usr/bin/env python3
"""
TabPFN Forecast

Two entry points, both routed through run_forecast_configured() via _run_forecast():
- run_forecast() - macro entry point, invoked from VBA via RunPython
- main()         - CLI entry point (python tabpfn_forecast.py <workbook.xlsx> [flags])

See README.md for configuration details (PyexcelSettings.xlsx, 'Forecast' sheet).
"""

import time
import os
import xlwings as xw  # type: ignore
import pandas as pd
import numpy as np
import sys
import argparse
from typing import Optional, List, Tuple
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error
from scipy import stats

# Optional imports for linear models
try:
    import statsmodels.api as sm
    from statsmodels.regression.linear_model import OLS
    from statsmodels.genmod.generalized_linear_model import GLM
    from statsmodels.genmod.families import Gamma
    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False
    OLS = None  # type: ignore
    GLM = None  # type: ignore
    Gamma = None  # type: ignore

# Optional imports for Bayesian MCMC
try:
    import bambi as bmb  # type: ignore
    import pymc as pm  # type: ignore
    HAS_BAMBI = True
    HAS_PYMC = True
except ImportError:
    HAS_BAMBI = False
    HAS_PYMC = False
    bmb = None  # type: ignore
    pm = None  # type: ignore

# Optional imports
try:
    from tabpfn import TabPFNRegressor # type: ignore
    HAS_TABPFN = True
except ImportError:
    HAS_TABPFN = False
    TabPFNRegressor = None  # type: ignore


try:
    from huggingface_hub import login  # type: ignore
    HAS_HUGGINGFACE = True
except ImportError:
    HAS_HUGGINGFACE = False
    login = None  # type: ignore


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
    ('n_mc_samples', ('--mc-samples',), 'MC Samples (linear only)', 'int', 5000),
    ('n_ensemble', ('--ensemble-size',), 'Ensemble Size (TabPFN)', 'int', 50),
    ('n_gb_estimators', ('--gb-estimators',), 'GB Estimators (GradientBoosting)', 'int', 100),
    ('uncertainty_mode', ('--uncertainty-mode',), 'Uncertainty Mode', 'str', 'Frequentist'),
    ('prior_confidence', ('--prior-confidence',), 'Prior Confidence (1-10, Bayesian only)', 'int', 5),
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
    'model_type': 'Model to use: TabPFN, GradientBoosting, OLS, Gamma (default: TabPFN)',
    'n_mc_samples': 'Monte Carlo samples, OLS/Gamma only (default: 5000)',
    'n_ensemble': 'TabPFN ensemble size (default: 50)',
    'n_gb_estimators': 'GradientBoosting n_estimators (default: 100)',
    'uncertainty_mode': 'Frequentist or Bayesian, OLS/Gamma only (default: Frequentist)',
    'prior_confidence': 'Prior confidence 1-10, Bayesian only (default: 5)',
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


def _gamma_offset(y: np.ndarray) -> float:
    """Small positive offset ensuring y is usable with the Gamma family (requires y > 0)."""
    return float(np.abs(y).min() * 0.01 + 1e-6)


def _add_constant(X: np.ndarray, n_rows: int) -> np.ndarray:
    """Add an intercept column, or an all-ones column when there are no features."""
    return sm.add_constant(X) if X.shape[1] > 0 else np.ones((n_rows, 1))


def _bootstrap_ols_predict(
    X_boot: np.ndarray,
    y_boot: np.ndarray,
    X_future_const: np.ndarray,
    n_samples: int,
    _gamma_offset: Optional[float],
) -> np.ndarray:
    X_boot_const = _add_constant(X_boot, n_samples)
    return OLS(y_boot, X_boot_const).fit().predict(X_future_const)


def _bootstrap_gamma_predict(
    X_boot: np.ndarray,
    y_boot: np.ndarray,
    X_future_const: np.ndarray,
    n_samples: int,
    gamma_offset: Optional[float],
) -> np.ndarray:
    y_boot_gamma = np.maximum(y_boot, gamma_offset)
    X_boot_const = _add_constant(X_boot, n_samples)
    return GLM(y_boot_gamma, X_boot_const, family=Gamma()).fit().predict(X_future_const)


BOOTSTRAP_FITS = {
    'ols': _bootstrap_ols_predict,
    'gamma': _bootstrap_gamma_predict,
}


def _aggregate_mc_scenarios(y_future_scenarios: np.ndarray) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Aggregate Monte Carlo scenario draws into a mean forecast and 95% interval."""
    y_future = np.mean(y_future_scenarios, axis=0)
    y_future_lower = np.percentile(y_future_scenarios, 2.5, axis=0)
    y_future_upper = np.percentile(y_future_scenarios, 97.5, axis=0)
    return y_future, y_future_lower, y_future_upper


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
    If the workbook already exists (e.g. isolation_forest.py created it with its
    own sheet), open it and add/replace the 'Forecast' sheet rather than
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
    n_mc_samples: int = 5000,
    n_ensemble: int = 50,
    n_gb_estimators: int = 100,
    uncertainty_mode: str = 'Frequentist',
    prior_confidence: int = 5,
    future_rows: int = 10,
    output_sheet: str = 'Forecast',
    batch_size: int = 50,
    delay: float = 0.5
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
        model_type: Model to use ('TabPFN', 'GradientBoosting', 'OLS', 'Gamma')
        n_mc_samples: Monte Carlo samples (only for OLS/Gamma models)
        n_ensemble: TabPFN ensemble size
        n_gb_estimators: GradientBoosting n_estimators
        uncertainty_mode: 'Frequentist' or 'Bayesian' (only for OLS/Gamma)
        prior_confidence: Prior confidence level 1-10 (1=low confidence/high variance, 10=high confidence/low variance)
        future_rows: Rows to generate in output
        output_sheet: Name of output sheet
        batch_size: Samples per batch (throttling always enabled)
        delay: Seconds between batches
    """
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
    if model_type_lower not in ('tabpfn', 'gradientboosting', 'ols', 'gamma'):
        raise ValueError(f"Unknown model_type: {model_type}. Choose from: TabPFN, GradientBoosting, OLS, Gamma")
    model = None
    model_diagnostics = {}
    use_mc = False  # Monte Carlo only for OLS and Gamma

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
                    model_type_lower = 'gradientboosting'  # Fall back
                else:
                    raise

    # Handle fallback to GradientBoosting (either explicit choice or fallback from TabPFN)
    if model_type_lower == 'gradientboosting':
        if model is None:  # Only create if not already created (fallback case)
            model = GradientBoostingRegressor(n_estimators=n_gb_estimators, random_state=42)
            model.fit(X, y)
            print(f"Using GradientBoosting (n_estimators={n_gb_estimators})")

    elif model_type_lower == 'ols':
        if not HAS_STATSMODELS:
            raise ValueError("OLS requires statsmodels. Install statsmodels package.")
        # Add constant for intercept
        X_with_const = _add_constant(X, n_samples)
        model = OLS(y, X_with_const).fit()
        use_mc = True
        print("Using OLS (Ordinary Least Squares)")
        # Store OLS diagnostics
        model_diagnostics['r_squared'] = model.rsquared
        model_diagnostics['adj_r_squared'] = model.rsquared_adj
        model_diagnostics['f_statistic'] = model.fvalue
        model_diagnostics['f_pvalue'] = model.f_pvalue

    elif model_type_lower == 'gamma':
        if not HAS_STATSMODELS:
            raise ValueError("Gamma GLM requires statsmodels. Install statsmodels package.")
        # Add constant for intercept
        X_with_const = _add_constant(X, n_samples)
        # Ensure y is positive for Gamma (pre-compute offset once)
        y_gamma = np.maximum(y, _gamma_offset(y))
        model = GLM(y_gamma, X_with_const, family=Gamma()).fit()
        use_mc = True
        print("Using Gamma GLM (Gamma-likelihood)")
        # Store Gamma diagnostics
        model_diagnostics['pseudo_r_squared'] = model.pseudo_rsquared()
        model_diagnostics['aic'] = model.aic
        model_diagnostics['bic'] = model.bic

    # Extract coefficients and p-values for linear models (pre-compute param_names once)
    if model_type_lower in ['ols', 'gamma']:
        param_names = ['const'] + feature_cols
        model_diagnostics['coefficients'] = model.params.to_dict() if hasattr(model.params, 'to_dict') else dict(zip(param_names, model.params))
        model_diagnostics['pvalues'] = model.pvalues.to_dict() if hasattr(model.pvalues, 'to_dict') else dict(zip(param_names, model.pvalues))

    # Get predictions (different methods for different models)
    if model_type_lower in ['ols', 'gamma']:
        y_pred = model.fittedvalues
    else:
        y_pred = model.predict(X)

    residuals = y - y_pred
    mae = mean_absolute_error(y, y_pred)
    rmse = np.sqrt(mean_squared_error(y, y_pred))

    # Calculate MAD (Median Absolute Deviation) for linear models (OLS/Gamma)
    if model_type_lower in ['ols', 'gamma']:
        mad = stats.median_abs_deviation(residuals, scale=1)  # Raw MAD (not scaled for normal distribution)
        model_diagnostics['mad'] = mad
        print(f"MAE: {mae:.6f}, RMSE: {rmse:.6f}, MAD: {mad:.6f}")
    else:
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

    # Step 3: Generate future predictions
    # For OLS/Gamma: Monte Carlo (Frequentist bootstrap or Bayesian MCMC)
    # For TabPFN/GradientBoosting: Single prediction
    if use_mc:
        # Pre-compute X_future with constant for linear models
        X_future_with_const = _add_constant(X_future, future_rows)

        # Check if Bayesian mode (inline one-time variables)
        if (uncertainty_mode.lower() if uncertainty_mode else 'frequentist') == 'bayesian':
            # Bayesian MCMC using Bambi (built on PyMC)
            if not HAS_BAMBI:
                raise ValueError("Bayesian mode requires Bambi (which depends on PyMC). Install with: pip install bambi")

            print(f"Bayesian MCMC: {n_mc_samples} samples with prior confidence {prior_confidence}/10")

            # Convert prior confidence (1-10) to prior scale (standard deviation)
            # 1 = low confidence = high variance (e.g., scale = 10)
            # 10 = high confidence = low variance (e.g., scale = 0.1)
            # Use exponential scale: scale = 10^(2 - prior_confidence)
            prior_scale = 10 ** (2 - prior_confidence)

            # Prepare data DataFrame for Bambi (formula-based interface) - compute once
            data_dict = {feat_name: X[:, i] for i, feat_name in enumerate(feature_cols)}
            data_dict['target'] = y
            data_df = pd.DataFrame(data_dict)

            # Build formula: target ~ feature1 + feature2 + ... (compute once)
            formula = 'target ~ ' + ' + '.join(feature_cols) if feature_cols else 'target ~ 1'

            # Build priors dict (common structure for both models)
            priors = {'Intercept': bmb.Prior('Normal', mu=0, sigma=prior_scale)}
            for feat_name in feature_cols:
                priors[feat_name] = bmb.Prior('Normal', mu=0, sigma=prior_scale)

            # Prepare future data for prediction (compute once, reused for both models)
            future_dict = {feat_name: X_future[:, i] for i, feat_name in enumerate(feature_cols)}
            future_df = pd.DataFrame(future_dict)

            if model_type_lower == 'ols':
                # Bayesian OLS with Bambi (Gaussian family)
                model = bmb.Model(formula, data_df, family='gaussian')
                priors['sigma'] = bmb.Prior('HalfNormal', sigma=10)  # Add sigma prior for OLS
                model.set_priors(priors)

                # Fit and predict
                idata = model.fit(draws=n_mc_samples, progressbar=False)
                posterior = model.predict(idata, data=future_df)

            elif model_type_lower == 'gamma':
                # Bayesian Gamma GLM with Bambi (reuse cached y_min_abs if available)
                y_gamma = np.maximum(y, _gamma_offset(y))
                data_df['target'] = y_gamma  # Update target in existing DataFrame

                model = bmb.Model(formula, data_df, family='gamma', link='log')
                model.set_priors(priors)  # Reuse priors dict (no sigma for Gamma)

                # Fit and predict
                idata = model.fit(draws=n_mc_samples, progressbar=False)
                posterior = model.predict(idata, data=future_df)

            # Extract posterior predictive samples (common for both models)
            if hasattr(posterior, 'posterior_predictive'):
                y_future_scenarios = posterior.posterior_predictive['target'].values.reshape(-1, future_rows)
            else:
                # Fallback: extract from trace directly
                y_future_scenarios = posterior['target'].values.reshape(-1, future_rows)

            # Aggregate MCMC samples
            y_future_scenarios = np.array(y_future_scenarios)  # Shape: (n_mc_samples, future_rows)
            y_future, y_future_lower, y_future_upper = _aggregate_mc_scenarios(y_future_scenarios)

        else:
            # Frequentist: Monte Carlo pairs bootstrap (case resampling)
            y_future_scenarios = []
            n_batches = (n_mc_samples + batch_size - 1) // batch_size
            print(f"MC Pairs Bootstrap: {n_mc_samples} simulations in {n_batches} batches of {batch_size}")

            # Pre-compute Gamma offset if needed (O(1) instead of per-bootstrap)
            gamma_offset = _gamma_offset(y) if model_type_lower == 'gamma' else None

            for batch_idx in range(n_batches):
                start_idx = batch_idx * batch_size
                end_idx = min(start_idx + batch_size, n_mc_samples)
                batch_size_actual = end_idx - start_idx

                # Generate predictions for this batch of scenarios
                bootstrap_fn = BOOTSTRAP_FITS[model_type_lower]
                for _ in range(batch_size_actual):
                    bootstrap_indices = np.random.choice(n_samples, size=n_samples, replace=True)
                    y_future_boot = bootstrap_fn(
                        X[bootstrap_indices],
                        y[bootstrap_indices],
                        X_future_with_const,
                        n_samples,
                        gamma_offset,
                    )
                    y_future_scenarios.append(y_future_boot)

                # Progress bar and throttling (pre-compute percentage once)
                pct = (batch_idx + 1) / n_batches
                bar_len = int(pct * 30)
                bar = '█' * bar_len + '░' * (30 - bar_len)
                print(f"\r  [{bar}] {batch_idx+1}/{n_batches} ({pct*100:.0f}%)", end='', flush=True)
                if batch_idx + 1 < n_batches:
                    time.sleep(delay)
            print()

            # Aggregate Monte Carlo scenarios
            y_future_scenarios = np.array(y_future_scenarios)  # Shape: (n_mc_samples, future_rows)
            y_future, y_future_lower, y_future_upper = _aggregate_mc_scenarios(y_future_scenarios)
    else:
        # Single prediction for TabPFN and GradientBoosting (no Monte Carlo)
        print("Single model prediction (no Monte Carlo)")
        y_future = model.predict(X_future)
        # For non-MC models, use prediction intervals based on residuals (use rmse directly)
        y_future_lower = y_future - 1.96 * rmse  # Approximate 95% CI
        y_future_upper = y_future + 1.96 * rmse
        # Create dummy scenarios array for compatibility (same shape as MC: n_samples x future_rows)
        y_future_scenarios = y_future.reshape(1, -1)  # Shape: (1, future_rows) - more efficient than tile

    # Compute sensitivity analysis statistics from Monte Carlo scenarios (only meaningful for MC models)
    # MC models: shape is (n_mc_samples, future_rows)
    # Non-MC models: shape is (1, future_rows) - no meaningful sensitivity stats
    if use_mc and y_future_scenarios.shape[0] > 1:
        # Aggregate across all future periods for overall sensitivity metrics
        y_future_flat = y_future_scenarios.flatten()
        # Pre-compute statistics once
        flat_mean = np.mean(y_future_flat)
        flat_std = np.std(y_future_flat)
        sensitivity_stats = {
            'mean': flat_mean,
            'std': flat_std,
            'min': np.min(y_future_flat),
            'max': np.max(y_future_flat),
            'p10': np.percentile(y_future_flat, 10),
            'p25': np.percentile(y_future_flat, 25),
            'p75': np.percentile(y_future_flat, 75),
            'p90': np.percentile(y_future_flat, 90),
            'p2_5': np.percentile(y_future_flat, 2.5),
            'p97_5': np.percentile(y_future_flat, 97.5),
            'cv': flat_std / np.abs(flat_mean) if flat_mean != 0 else np.nan
        }
    else:
        # For non-MC models, sensitivity stats are not meaningful
        sensitivity_stats = None

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
                  feature_cols + ['Prediction', 'Lower (95%)', 'Upper (95%)'])

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
        'model_type': model_type_lower.capitalize(),  # Use actual model used (may differ from requested if fallback)
        'n_mc_samples': n_mc_samples if use_mc else 0,
        'sensitivity_stats': sensitivity_stats,
        'model_diagnostics': model_diagnostics,
        'y_future_scenarios': y_future_scenarios  # Full array for detailed analysis if needed
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
        n_mc_samples=cfg['n_mc_samples'],
        n_ensemble=cfg['n_ensemble'],
        n_gb_estimators=cfg['n_gb_estimators'],
        uncertainty_mode=cfg['uncertainty_mode'],
        prior_confidence=cfg['prior_confidence'],
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

    if results.get('model_diagnostics'):
        diag_content.append(['', ''])
        diag_content.append(['--- MODEL-SPECIFIC DIAGNOSTICS ---', ''])
        diags = results['model_diagnostics']
        if 'r_squared' in diags:
            diag_content.append(['R-squared', f"{diags['r_squared']:.6f}"])
            diag_content.append(['Adjusted R-squared', f"{diags['adj_r_squared']:.6f}"])
            diag_content.append(['F-statistic', f"{diags['f_statistic']:.6f}"])
            diag_content.append(['F p-value', f"{diags['f_pvalue']:.6e}"])
            if 'mad' in diags:
                diag_content.append(['MAD (Median Absolute Deviation)', f"{diags['mad']:.6f}"])
        elif 'pseudo_r_squared' in diags:
            diag_content.append(['Pseudo R-squared', f"{diags['pseudo_r_squared']:.6f}"])
            diag_content.append(['AIC', f"{diags['aic']:.2f}"])
            diag_content.append(['BIC', f"{diags['bic']:.2f}"])
            if 'mad' in diags:
                diag_content.append(['MAD (Median Absolute Deviation)', f"{diags['mad']:.6f}"])

    if results.get('sensitivity_stats'):
        sens = results['sensitivity_stats']
        diag_content.extend([
            ['', ''],
            ['--- MONTE CARLO SENSITIVITY ANALYSIS ---', ''],
            ['MC Samples', results['n_mc_samples']],
            ['', ''],
            ['Overall Distribution (All Periods)', ''],
            ['Mean Forecast', f"{sens['mean']:.6f}"],
            ['Standard Deviation', f"{sens['std']:.6f}"],
            ['Coefficient of Variation', f"{sens['cv']:.4f}" if not np.isnan(sens['cv']) else 'N/A'],
            ['Minimum', f"{sens['min']:.6f}"],
            ['Maximum', f"{sens['max']:.6f}"],
            ['', ''],
            ['Percentiles', ''],
            ['10th Percentile', f"{sens['p10']:.6f}"],
            ['25th Percentile', f"{sens['p25']:.6f}"],
            ['75th Percentile', f"{sens['p75']:.6f}"],
            ['90th Percentile', f"{sens['p90']:.6f}"],
            ['2.5th Percentile (Lower 95% CI)', f"{sens['p2_5']:.6f}"],
            ['97.5th Percentile (Upper 95% CI)', f"{sens['p97_5']:.6f}"],
        ])

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
    parser = argparse.ArgumentParser(description='TabPFN + Monte Carlo Forecast (CLI mode)')
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
