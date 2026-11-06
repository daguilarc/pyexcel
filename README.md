# pyexcel

## Forecasting

[Forecasting](forecasting/README.md) forecasts a target from historical Excel rows using TabPFN or gradient boosting. Choose history, target and forecast length; it backtests automatically and writes point predictions in original units through desktop Excel. TabPFN is a pretrained tabular model that can learn patterns from a small history. Gradient boosting builds successive trees that correct earlier errors; both models here learn past target changes to forecast subsequent values. See its README for installation, settings, macros and model licenses.

## Divvy/Bill cleanup

[Divvy cleanup](divvy_xlsx/README.md) removes unwanted transaction-export columns, sorts oldest first, and writes a cleaned copy. It needs Python with pandas/openpyxl and does not require Excel. Its dependencies and environment are separate from forecasting.

## License

Scripts are MIT licensed. TabPFN weights have [separate licensing](forecasting/README.md#tabpfn-licensing).
