#!/bin/bash
# One-click setup script for Excel Isolation Forest UDF
# This script sets up the conda environment and configures xlwings

echo "========================================"
echo "Excel Isolation Forest - Setup"
echo "========================================"
echo ""

# Check if miniconda exists in this directory
if [ ! -f "miniconda/bin/python" ]; then
    echo "ERROR: Miniconda not found in miniconda/ folder"
    echo ""
    echo "Please extract Miniconda to the miniconda/ folder:"
    echo "1. Download Miniconda from: https://docs.conda.io/en/latest/miniconda.html"
    echo "2. Extract to: miniconda/ folder in this directory"
    echo "3. Run this script again: bash setup.sh"
    exit 1
fi

echo "[1/4] Activating Miniconda..."
source miniconda/bin/activate

echo "[2/4] Creating conda environment from environment.yml..."
conda env remove -n excelpy --yes
conda env create -f environment.yml

echo "[3/4] Activating excelpy environment..."
conda activate excelpy

echo "[4/4] Configuring xlwings add-in..."
python -m xlwings addin install

echo ""
echo "========================================"
echo "Setup Complete!"
echo "========================================"
echo ""

# Get the full path to the scripts folder
SCRIPTS_DIR="$(cd "$(dirname "$0")" && pwd)"
echo "IMPORTANT: Scripts folder location:"
echo "$SCRIPTS_DIR"
echo ""

echo "Next steps:"
echo "1. Open your workbook in Excel and save it as .xlsm"
echo "2. Alt+F11 -> File -> Import File -> select pyexcel_macros.bas"
echo "   (Full path: $SCRIPTS_DIR/pyexcel_macros.bas)"
echo "3. In the VBA editor set SCRIPT_DIR to the scripts folder path shown above"
echo "4. Alt+F8 -> RunIsolationForest or RunForecast"
echo "5. First run creates PyexcelSettings.xlsx beside the scripts -"
echo "   configure it and run again"
echo ""
echo "CLI alternative (no Excel/VBA required):"
echo "  python isolation_forest.py <workbook.xlsx>"
echo "  python tabpfn_forecast.py <workbook.xlsx>"
echo ""
echo "TIP: Keep the scripts in this folder for easy access"
echo ""

