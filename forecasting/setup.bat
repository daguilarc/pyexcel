@echo off
setlocal
pushd "%~dp0"
if errorlevel 1 exit /b %errorlevel%
where conda >nul 2>nul
if errorlevel 1 (
  echo Install conda or use the Python/venv instructions in README.md
  popd
  exit /b 1
)
call conda env update --file "%~dp0environment.yml"
if errorlevel 1 goto failed
call conda run -n pyexcel-forecast python -m xlwings addin install
if errorlevel 1 goto failed
echo Configure xlwings to use pyexcel-forecast and import "%~dp0pyexcel_macros.bas"
popd
exit /b 0
:failed
set "FORECAST_EXIT=%errorlevel%"
popd
exit /b %FORECAST_EXIT%
