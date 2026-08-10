@echo off
REM One-click setup script for Excel Isolation Forest UDF
REM This script sets up the conda environment and configures xlwings

echo ========================================
echo Excel Isolation Forest - Setup
echo ========================================
echo.

REM Check if miniconda exists in this directory
if not exist "miniconda\python.exe" (
    echo ERROR: Miniconda not found in miniconda\ folder
    echo.
    echo Please extract Miniconda to the miniconda\ folder:
    echo 1. Download Miniconda from: https://docs.conda.io/en/latest/miniconda.html
    echo 2. Extract to: miniconda\ folder in this directory
    echo 3. Run this script again
    pause
    exit /b 1
)

echo [1/4] Activating Miniconda...
call miniconda\Scripts\activate.bat

echo [2/4] Creating conda environment from environment.yml...
call conda env remove -n excelpy --yes
call conda env create -f environment.yml

echo [3/4] Activating excelpy environment...
call conda activate excelpy

echo [4/4] Configuring xlwings add-in...
python -m xlwings addin install

echo.
echo ========================================
echo Setup Complete!
echo ========================================
echo.
REM Get the full path to the scripts folder
set SCRIPTS_DIR=%~dp0
echo IMPORTANT: Scripts folder location:
echo %SCRIPTS_DIR%
echo.
echo Next steps:
echo 1. Open your workbook in Excel and save it as .xlsm
echo 2. Alt+F11 -^> File -^> Import File -^> select pyexcel_macros.bas
echo    (Full path: %SCRIPTS_DIR%pyexcel_macros.bas)
echo 3. In the VBA editor set SCRIPT_DIR to the scripts folder path shown above
echo 4. Alt+F8 -^> RunIsolationForest or RunForecast
echo 5. First run creates PyexcelSettings.xlsx beside the scripts -
echo    configure it and run again
echo.
echo CLI alternative (no Excel/VBA required):
echo   python isolation_forest.py ^<workbook.xlsx^>
echo   python tabpfn_forecast.py ^<workbook.xlsx^>
echo.
echo TIP: Keep the scripts in this folder for easy access
echo.
pause

