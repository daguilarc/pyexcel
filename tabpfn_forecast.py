#!/usr/bin/env python3
"""
TabPFN Forecast - Excel Macro Tool

REQUIRES Settings sheet. Call via Excel macro (Alt+F8 → RunForecastWithSettings).

Workflow:
1. TabPFN fit on historical data
2. Monte Carlo simulation to augment data
3. TabPFN refit on combined data
4. Predict future values
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



# CORE LOGIC


def get_settings_file_path() -> str:
    """Get the path to ForecastSettings.xlsx in the same directory as this script."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ForecastSettings.xlsx')


def detect_header_row(df: pd.DataFrame, max_scan: int = 10) -> int:
    """Auto-detect header row (first row with string column names)."""
    for i in range(min(max_scan, len(df))):
        non_null = df.iloc[i].dropna()
        if len(non_null) == 0:
            continue
        if all(isinstance(v, str) and not v.replace('.', '').replace('-', '').isdigit() 
               for v in non_null):
            return i
    return 0


def clean_dataframe(df: pd.DataFrame, header_row: Optional[int] = None) -> pd.DataFrame:
    """
    Clean DataFrame: set headers, convert to numeric, handle NaN.
    
    Parameters:
        df: DataFrame to clean
        header_row: 
            - None: Auto-detect header row
            - > 0: Use that row (0-based) as headers
            - 0: Headers already set, skip header detection (just convert to numeric)
    """
    df = df.copy()
    
    # Handle header row specification
    if header_row is not None and header_row > 0:
        # User specified header row
        df.columns = df.iloc[header_row].values
        df = df.iloc[header_row + 1:].reset_index(drop=True)
    elif header_row is None and len(df) > 0:
        # Auto-detect: check if first row looks like headers
        non_null = df.iloc[0].dropna()
        if len(non_null) > 0 and all(
            isinstance(v, str) and not _is_numeric_string(v) for v in non_null
        ):
            df.columns = df.iloc[0].values
            df = df.iloc[1:].reset_index(drop=True)
    # If header_row == 0, headers already set, skip header detection
    
    # Convert object columns to numeric
    for col in df.columns:
        if hasattr(df[col], 'dtype') and df[col].dtype == object:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Remove entirely NaN rows
    df = df.dropna(how='all').reset_index(drop=True)
    
    return df


def _is_numeric_string(val) -> bool:
    """Check if a string represents a number."""
    if not isinstance(val, str):
        return isinstance(val, (int, float))
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


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


def _bootstrap_ols_predict(
    X_boot: np.ndarray,
    y_boot: np.ndarray,
    X_future_const: np.ndarray,
    n_features: int,
    n_samples: int,
    _gamma_offset: Optional[float],
) -> np.ndarray:
    X_boot_const = sm.add_constant(X_boot) if n_features > 0 else np.ones((n_samples, 1))
    return OLS(y_boot, X_boot_const).fit().predict(X_future_const)


def _bootstrap_gamma_predict(
    X_boot: np.ndarray,
    y_boot: np.ndarray,
    X_future_const: np.ndarray,
    n_features: int,
    n_samples: int,
    gamma_offset: Optional[float],
) -> np.ndarray:
    y_boot_gamma = np.maximum(y_boot, gamma_offset)
    X_boot_const = sm.add_constant(X_boot) if n_features > 0 else np.ones((n_samples, 1))
    return GLM(y_boot_gamma, X_boot_const, family=Gamma()).fit().predict(X_future_const)


BOOTSTRAP_FITS = {
    'ols': _bootstrap_ols_predict,
    'gamma': _bootstrap_gamma_predict,
}


def _read_settings_workbook(settings_wb) -> dict:
    """Read ForecastSettings.xlsx Settings sheet into a configuration dict."""
    s = settings_wb.sheets['Settings']
    return {
        'data_sheets': [x.strip() for x in str(s.range('B2').value or '').split(',') if x.strip()],
        'header_row': int(s.range('B3').value or 1),
        'data_start_row': int(s.range('B4').value or 2),
        'target_col': s.range('B7').value,
        'feature_cols_raw': s.range('B8').value or 'all',
        'categorical_cols': [x.strip() for x in str(s.range('B9').value or '').split(',') if x.strip()],
        'primary_time_index': s.range('B12').value,
        'secondary_time_index': s.range('B13').value,
        'model_type': s.range('B16').value or 'TabPFN',
        'n_mc_samples': int(s.range('B17').value or 5000),
        'n_ensemble': int(s.range('B18').value or 50),
        'n_gb_estimators': int(s.range('B19').value or 100),
        'uncertainty_mode': s.range('B21').value or 'Frequentist',
        'prior_confidence': int(s.range('B22').value or 5),
        'future_rows': int(s.range('B24').value or 10),
        'output_sheet': s.range('B25').value or 'Forecast',
    }


# CONFIGURED FORECAST (user-specified settings)


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
        X_with_const = sm.add_constant(X) if n_features > 0 else np.ones((n_samples, 1))
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
        X_with_const = sm.add_constant(X) if n_features > 0 else np.ones((n_samples, 1))
        # Ensure y is positive for Gamma (pre-compute offset once)
        y_gamma = np.maximum(y, np.abs(y).min() * 0.01 + 1e-6)
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
    
    else:
        raise ValueError(f"Unknown model_type: {model_type}. Choose from: TabPFN, GradientBoosting, OLS, Gamma")
    
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
            # Not enough time data: fall back to mean+noise (vectorized, compute std once)
            means = np.mean(X_sorted, axis=0)
            stds = np.std(X_sorted, axis=0)
            feature_stds = np.where(stds == 0, 1.0, stds)
            X_future = means + np.random.normal(0, 0.5, (future_rows, n_features)) * feature_stds
    else:
        # No time index: use mean + noise (vectorized, compute std once)
        print("No time index specified, using mean+noise for future features")
        means = np.mean(X, axis=0)
        stds = np.std(X, axis=0)
        feature_stds = np.where(stds == 0, 1.0, stds)
        X_future = means + np.random.normal(0, 0.5, (future_rows, n_features)) * feature_stds
    
    # Step 3: Generate future predictions
    # For OLS/Gamma: Monte Carlo (Frequentist bootstrap or Bayesian MCMC)
    # For TabPFN/GradientBoosting: Single prediction
    if use_mc:
        # Pre-compute X_future with constant for linear models
        X_future_with_const = sm.add_constant(X_future) if n_features > 0 else np.ones((future_rows, 1))
        
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
                y_gamma = np.maximum(y, (np.abs(y).min() * 0.01 + 1e-6))
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
            y_future = np.mean(y_future_scenarios, axis=0)  # Mean forecast
            y_future_lower = np.percentile(y_future_scenarios, 2.5, axis=0)  # Lower bound (2.5th percentile)
            y_future_upper = np.percentile(y_future_scenarios, 97.5, axis=0)  # Upper bound (97.5th percentile)
            
        else:
            # Frequentist: Monte Carlo pairs bootstrap (case resampling)
            y_future_scenarios = []
            n_batches = (n_mc_samples + batch_size - 1) // batch_size
            print(f"MC Pairs Bootstrap: {n_mc_samples} simulations in {n_batches} batches of {batch_size}")
            
            # Pre-compute Gamma offset if needed (O(1) instead of per-bootstrap)
            gamma_offset = (np.abs(y).min() * 0.01 + 1e-6) if model_type_lower == 'gamma' else None
            
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
                        n_features,
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
            y_future = np.mean(y_future_scenarios, axis=0)  # Mean forecast
            y_future_lower = np.percentile(y_future_scenarios, 2.5, axis=0)  # Lower bound (2.5th percentile)
            y_future_upper = np.percentile(y_future_scenarios, 97.5, axis=0)  # Upper bound (97.5th percentile)
    else:
        # Single prediction for TabPFN and GradientBoosting (no Monte Carlo)
        print("Single model prediction (no Monte Carlo)")
        if model_type_lower == 'tabpfn' or model_type_lower == 'gradientboosting':
            y_future = model.predict(X_future)
            # For non-MC models, use prediction intervals based on residuals (use rmse directly)
            y_future_lower = y_future - 1.96 * rmse  # Approximate 95% CI
            y_future_upper = y_future + 1.96 * rmse
            # Create dummy scenarios array for compatibility (same shape as MC: n_samples x future_rows)
            y_future_scenarios = y_future.reshape(1, -1)  # Shape: (1, future_rows) - more efficient than tile
        else:
            raise ValueError(f"Unexpected model_type for non-MC path: {model_type}")
    
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
    if output_sheet not in wb_sheet_names:
        out_sht = wb.sheets.add(output_sheet, after=wb.sheets[-1])
    else:
        out_sht = wb.sheets[output_sheet]
        out_sht.clear()
    
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
    
    wb.save()
    print(f"Saved to '{output_sheet}'")
    
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



# EXCEL MACROS (call via Alt+F8)


@xw.sub
def RunForecastWithSettings():
    """
    Forecast using 'Settings' sheet configuration. Call via Alt+F8.
    
    User specifies:
    - Data Sheets: comma-separated list of identically-formatted sheets
    - Header Row: which row contains column headers (1-based)
    - Target Column: name of the target variable
    - Categorical Columns: comma-separated list of categorical variable names
    - Feature Columns: comma-separated list, or "all" for all except target
    """
    try:
        wb = xw.Book.caller()
        
        # Get or create ForecastSettings.xlsx file
        settings_path = get_settings_file_path()
        
        # Check if ForecastSettings.xlsx exists, create if not
        try:
            settings_wb = xw.Book(settings_path)
        except:
            # Create new ForecastSettings.xlsx with default Settings sheet
            settings_wb = xw.Book()
            s = settings_wb.sheets[0]
            s.name = 'Settings'
            s.range('A1:B24').value = [
                [' DATA CONFIGURATION ', ''],
                ['Data Sheets (comma-separated)', wb.sheets[0].name],
                ['Header Row (1-based)', 1],
                ['Data Start Row (1-based)', 2],
                ['', ''],
                [' COLUMN CONFIGURATION ', ''],
                ['Target Column', ''],
                ['Feature Columns (or "all")', 'all'],
                ['Categorical Columns', ''],
                ['', ''],
                [' TIME CONFIGURATION ', ''],
                ['Primary Time Index (e.g., Year)', ''],
                ['Secondary Time Index (e.g., Date/Quarter)', ''],
                ['', ''],
                [' MODEL CONFIGURATION ', ''],
                ['Model Type', 'TabPFN'],
                ['MC Samples (linear only)', 5000],
                ['Ensemble Size (TabPFN)', 50],
                ['GB Estimators (GradientBoosting)', 100],
                ['', ''],
                [' UNCERTAINTY CONFIGURATION (linear only) ', ''],
                ['Uncertainty Mode', 'Frequentist'],
                ['Prior Confidence (1-10, Bayesian only)', 5],
                ['', ''],
                [' OUTPUT ', ''],
                ['Future Rows', 10],
                ['Output Sheet', 'Forecast']
            ]
            s.range('A:A').column_width = 30
            s.range('B:B').column_width = 40
            settings_wb.save(settings_path)
            wb.app.alert(
                f"Created 'ForecastSettings.xlsx' with default configuration.\n\n"
                f"Location: {settings_path}\n\n"
                "Please configure:\n"
                "1. Data Sheets - which sheets to read (comma-separated)\n"
                "2. Header Row - row with column names (1-based)\n"
                "3. Target Column - what to predict\n"
                "4. Categorical Columns - for fixed effects (comma-separated)\n"
                "5. Primary Time Index - time column for forecasting (e.g., 'Year')\n"
                "6. Secondary Time Index - optional sub-period (e.g., 'Date', 'Quarter')\n\n"
                "Time indices enable time-based forecasting. Leave blank to use mean+noise.\n\n"
                "Then run this macro again to execute the forecast."
            )
            settings_wb.close()
            return
        
        cfg = _read_settings_workbook(settings_wb)

        if not cfg['data_sheets']:
            wb.app.alert("Error: No data sheets specified in Settings!")
            return
        
        if not cfg['target_col']:
            wb.app.alert("Error: No target column specified in Settings!")
            settings_wb.close()
            return

        settings_wb.close()

        results = run_forecast_configured(
            wb=wb,
            data_sheets=cfg['data_sheets'],
            header_row=cfg['header_row'],
            data_start_row=cfg['data_start_row'],
            target_column=cfg['target_col'],
            feature_columns=cfg['feature_cols_raw'],
            categorical_columns=cfg['categorical_cols'],
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
        target_col = cfg['target_col']
        data_sheets = cfg['data_sheets']
        n_ensemble = cfg['n_ensemble']
        n_gb_estimators = cfg['n_gb_estimators']
        output_sheet = cfg['output_sheet']
        
        # Create Diagnostics sheet (cache sheet names for O(1) lookup)
        wb_sheet_names = {s.name for s in wb.sheets}
        if 'Diagnostics' not in wb_sheet_names:
            diag_sheet = wb.sheets.add('Diagnostics', after=wb.sheets[-1])
        else:
            diag_sheet = wb.sheets['Diagnostics']
            diag_sheet.clear()
        
        # Build diagnostics content
        diag_content = [
            ['=== MODEL PERFORMANCE ===', ''],
            ['Model Type', results['model_type']],
            ['Target Column', target_col],
            ['Number of Features', results['n_features']],
            ['Number of Samples', results['n_samples']],
            ['', ''],
            ['=== RESIDUAL DIAGNOSTICS ===', ''],
            ['Mean Absolute Error (MAE)', f"{results['historical_mae']:.6f}"],
            ['Root Mean Squared Error (RMSE)', f"{results['historical_rmse']:.6f}"],
        ]
        
        # Add model-specific diagnostics
        if results.get('model_diagnostics'):
            diag_content.append(['', ''])
            diag_content.append(['=== MODEL-SPECIFIC DIAGNOSTICS ===', ''])
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
        
        # Add Monte Carlo sensitivity analysis (only for OLS/Gamma)
        if results.get('sensitivity_stats'):
            sens = results['sensitivity_stats']
            diag_content.extend([
                ['', ''],
                ['=== MONTE CARLO SENSITIVITY ANALYSIS ===', ''],
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
            ['=== MODEL CONFIGURATION ===', ''],
            ['Ensemble Size (TabPFN)', n_ensemble],
            ['GB Estimators (GradientBoosting)', n_gb_estimators],
            ['Sheets Processed', ', '.join(data_sheets)]
        ])
        
        diag_sheet.range('A1').value = diag_content
        diag_sheet.range('A:A').column_width = 35
        diag_sheet.range('B:B').column_width = 25
        
        wb.sheets[output_sheet].activate()
        wb.app.alert(
            f"Forecast Complete!\n\n"
            f"Sheets processed: {len(data_sheets)}\n"
            f"Target: {target_col}\n"
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
    """CLI entry point - reads Settings sheet from Excel file."""
    parser = argparse.ArgumentParser(description='TabPFN + Monte Carlo Forecast (CLI mode)')
    parser.add_argument('excel_path', help='Path to Excel workbook with Settings sheet')
    
    args = parser.parse_args()
    
    try:
        wb = xw.Book(args.excel_path)
        
        # Read Settings from ForecastSettings.xlsx
        settings_path = get_settings_file_path()
        
        try:
            settings_wb = xw.Book(settings_path)
        except:
            print(f"ERROR: 'ForecastSettings.xlsx' not found!")
            print(f"Expected location: {settings_path}")
            print("Please run the macro once to create the settings file, or create it manually.")
            return 1
        
        cfg = _read_settings_workbook(settings_wb)

        if not cfg['data_sheets']:
            print("ERROR: No data sheets specified in Settings!")
            return 1

        if not cfg['target_col']:
            print("ERROR: No target column specified in Settings!")
            settings_wb.close()
            return 1

        settings_wb.close()

        results = run_forecast_configured(
            wb=wb,
            data_sheets=cfg['data_sheets'],
            header_row=cfg['header_row'],
            data_start_row=cfg['data_start_row'],
            target_column=cfg['target_col'],
            feature_columns=cfg['feature_cols_raw'],
            categorical_columns=cfg['categorical_cols'],
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

        print(f"\n{'='*50}")
        print(f"Forecast Complete!")
        print(f"Target: {cfg['target_col']}")
        print(f"Features: {results['n_features']}")
        print(f"Samples: {results['n_samples']}")
        print(f"MAE: {results['historical_mae']:.6f}")
        print(f"RMSE: {results['historical_rmse']:.6f}")
        print(f"Predictions: {cfg['future_rows']}")
        print(f"See '{cfg['output_sheet']}' and 'Diagnostics' sheets for details.")
        
        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())

