# Carta Cap Table Transformer

Transforms Carta cap table exports into your firm's internal cap table format.

## Installation

### 1. Install Python (if not already installed)
- Download from https://www.python.org/downloads/
- **Windows**: Check "Add Python to PATH" during installation
- **Mac**: Use `brew install python3` or download from python.org

### 2. Install required packages
Open Terminal (Mac) or Command Prompt (Windows) and run:

```bash
pip install xlwings pandas openpyxl
```

### 3. Install xlwings Excel Add-in
```bash
xlwings addin install
```

This adds the xlwings ribbon tab to Excel.

### 4. Set up the project folder
Create a folder (e.g., `C:\CapTableTool` or `~/CapTableTool`) containing:
```
CapTableTool/
├── carta_to_cap_table.py      # Main transformation script
├── Cap_Table_Template.xlsx    # Your firm's template
├── xlwings.conf               # Configuration (optional)
└── CartaTransformer.xlsm      # Excel workbook with button
```

## Usage

### Option A: Run from Excel (xlwings button)

1. Open `CartaTransformer.xlsm`
2. Click the "Transform Carta" button
3. Select your Carta export file when prompted
4. Output file opens automatically

### Option B: Run from command line

```bash
python carta_to_cap_table.py <carta_export.xlsx> <template.xlsx>
```

Example:
```bash
python carta_to_cap_table.py "Downloads/Carta_Export.xlsx" "Cap_Table_Template.xlsx"
```

Output: `{CompanyName}_Cap_Table_{Date}.xlsx` in the same folder as the Carta export.

## Creating the Excel Button (CartaTransformer.xlsm)

### Method 1: Using xlwings quickstart
```bash
cd CapTableTool
xlwings quickstart CartaTransformer
```
This creates a `.xlsm` file pre-configured for xlwings.

### Method 2: Manual setup

1. Create a new Excel workbook, save as `.xlsm` (macro-enabled)
2. Press `Alt+F11` to open VBA editor
3. Insert > Module, paste this code:

```vba
Sub TransformCarta()
    RunPython "import carta_to_cap_table; carta_to_cap_table.main()"
End Sub
```

4. Close VBA editor
5. Insert > Shapes > Button, assign `TransformCarta` macro
6. Save

### Configuring Python path (if needed)

If Excel can't find Python:

1. In Excel, go to xlwings ribbon tab
2. Click "Interpreter" and browse to your python.exe location
   - Windows: Usually `C:\Users\YourName\AppData\Local\Programs\Python\Python3x\python.exe`
   - Mac: Usually `/usr/local/bin/python3`

Or edit `xlwings.conf`:
```ini
INTERPRETER_WIN = C:\path\to\python.exe
INTERPRETER_MAC = /usr/local/bin/python3
```

## What the tool does

1. **Parses Carta export**: Extracts company name, date, stakeholders, share counts
2. **Populates template**:
   - Company name → Cell I6
   - Cap table date → Cell I7
   - Share class headers → Row 30 (F-I)
   - Top 9 investors by shares → Rows 31-39
   - Remaining investors → "Other Investors" row 40
3. **Preserves all formulas** in other sheets (Pro Forma, Waterfall, Summary)
4. **Saves output** as `{CompanyName}_Cap_Table_{Date}.xlsx`

## Limitations

- Maximum 9 individual investor rows (+ "Other Investors" bucket)
- Maximum 4 share classes mapped to columns F-I
- Round metadata (rows 10-19) left untouched - edit manually after import
- Options from all plans are combined into single Options column

## Troubleshooting

### "Python not found"
- Ensure Python is installed and in your PATH
- Update interpreter path in xlwings.conf or Excel ribbon

### "Module not found"
- Ensure all .py files are in the same folder as the .xlsm
- Check PYTHONPATH in xlwings.conf points to correct folder

### "Template not found"
- Place `Cap_Table_Template.xlsx` in the same folder as the script

### Formulas show #REF! errors
- Open and save the output file in Excel to recalculate
- Or run: `python -c "from openpyxl import load_workbook; wb = load_workbook('output.xlsx'); wb.calculation.calcMode = 'auto'; wb.save('output.xlsx')"`
