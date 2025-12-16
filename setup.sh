#!/bin/bash
# Carta Cap Table Transformer - Setup Script (Mac/Linux)

set -e

echo "=================================="
echo "Carta Cap Table Transformer Setup"
echo "=================================="
echo ""

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Check for Python 3
echo "Checking for Python 3..."
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
    echo "✓ Found Python 3: $(python3 --version)"
elif command -v python &> /dev/null; then
    # Check if python is Python 3
    if python --version 2>&1 | grep -q "Python 3"; then
        PYTHON_CMD="python"
        echo "✓ Found Python 3: $(python --version)"
    else
        echo "✗ Python 3 is required but not found."
        echo "  Install from: https://www.python.org/downloads/"
        echo "  Or via Homebrew: brew install python3"
        exit 1
    fi
else
    echo "✗ Python is not installed."
    echo "  Install from: https://www.python.org/downloads/"
    echo "  Or via Homebrew: brew install python3"
    exit 1
fi

# Create virtual environment
echo ""
echo "Creating virtual environment..."
if [ -d ".venv" ]; then
    echo "  Virtual environment already exists, skipping creation."
else
    $PYTHON_CMD -m venv .venv
    echo "✓ Created .venv directory"
fi

# Activate virtual environment
echo ""
echo "Activating virtual environment..."
source .venv/bin/activate
echo "✓ Activated .venv"

# Upgrade pip
echo ""
echo "Upgrading pip..."
pip install --upgrade pip --quiet
echo "✓ pip upgraded"

# Install dependencies
echo ""
echo "Installing dependencies..."
pip install -r requirements.txt --quiet
echo "✓ Installed: xlwings, pandas, openpyxl, streamlit"

# Install xlwings Excel add-in
echo ""
echo "Installing xlwings Excel add-in..."
echo "  (This adds a ribbon tab to Excel for running Python scripts)"
xlwings addin install || {
    echo "  Note: xlwings addin install may require Excel to be closed."
    echo "  If it failed, close Excel and run: xlwings addin install"
}

echo ""
echo "=================================="
echo "Setup Complete!"
echo "=================================="
echo ""
echo "USAGE:"
echo ""
echo "1. Activate the virtual environment:"
echo "   source .venv/bin/activate"
echo ""
echo "2. Run via command line:"
echo "   python src/carta_to_cap_table.py <carta_export.xlsx> templates/Cap_Table_Template.xlsx"
echo ""
echo "3. Run the web interface:"
echo "   streamlit run src/app.py"
echo ""
echo "4. For Excel integration, see README.md"
echo ""
