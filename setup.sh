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
conda env create -f environment.yml --force

echo "[3/4] Activating excelpy environment..."
conda activate excelpy

echo "[4/4] Configuring xlwings add-in..."
python -m xlwings addin install

echo ""
echo "========================================"
echo "Setup Complete!"
echo "========================================"
echo ""

# Get the full path to isolation_forest.py
PYTHON_FILE="$(cd "$(dirname "$0")" && pwd)/isolation_forest.py"
echo "IMPORTANT: Python file location:"
echo "$PYTHON_FILE"
echo ""

echo "Next steps:"
echo "1. Open Excel"
echo "2. You should see the xlwings ribbon"
echo "3. Click 'Import Functions'"
echo "4. Navigate to and select: isolation_forest.py"
echo "   (Full path shown above)"
echo "5. Use =ISOLATION_FOREST(...) in your formulas"
echo ""
echo "TIP: Keep isolation_forest.py in this folder for easy access"
echo ""

