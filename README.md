# Carta Cap Table Transformer

Transforms Carta cap table exports into your firm's internal cap table format.

## Quick Start

### Mac/Linux
```bash
./setup.sh
```

### Windows
Double-click `setup.bat` or run in Command Prompt:
```cmd
setup.bat
```

The setup script will:
1. Check for Python 3
2. Create a virtual environment (`.venv`)
3. Install all dependencies
4. Install the xlwings Excel add-in

---

## Project Structure

```
hercules-app/
├── README.md                 # This file
├── requirements.txt          # Python dependencies
├── setup.sh                  # Mac/Linux setup script
├── setup.bat                 # Windows setup script
├── src/
│   ├── carta_to_cap_table.py # Core transformation logic
│   └── app.py                # Streamlit web interface
├── templates/
│   └── Cap Table Template.xlsx
├── excel/
│   ├── VBA_Module_Code.bas   # VBA code for Excel button
│   └── xlwings.conf          # xlwings configuration
└── examples/
    └── Carta Cap Table_v1.xlsx
```

---

## Usage

### Option 1: Web Interface (Streamlit)

1. Activate the virtual environment:
   ```bash
   # Mac/Linux
   source .venv/bin/activate

   # Windows
   .venv\Scripts\activate
   ```

2. Run the web app:
   ```bash
   streamlit run src/app.py
   ```

3. Open your browser to `http://localhost:8501`
4. Upload your Carta export and click Transform

### Option 2: Command Line

```bash
# Activate venv first (see above)
python src/carta_to_cap_table.py <carta_export.xlsx> templates/"Cap Table Template.xlsx"
```

Example:
```bash
python src/carta_to_cap_table.py ~/Downloads/Carta_Export.xlsx templates/"Cap Table Template.xlsx"
```

Output: `{CompanyName}_Cap_Table_{Date}.xlsx` in the same folder as the Carta export.

### Option 3: Excel Button (xlwings)

1. Create a macro-enabled workbook (`CartaTransformer.xlsm`)
2. Press `Alt+F11` to open VBA editor
3. Insert > Module, paste this code:
   ```vba
   Sub TransformCarta()
       RunPython "import carta_to_cap_table; carta_to_cap_table.main()"
   End Sub
   ```
4. Close VBA editor
5. Insert > Shapes > Button, assign `TransformCarta` macro
6. Save the workbook in the project root folder

See `excel/VBA_Module_Code.bas` for the complete VBA code.

---

## Manual Installation (if setup script fails)

### 1. Install Python
- Download from https://www.python.org/downloads/
- **Windows**: Check "Add Python to PATH" during installation
- **Mac**: Use `brew install python3` or download from python.org

### 2. Create virtual environment
```bash
python3 -m venv .venv

# Activate
source .venv/bin/activate    # Mac/Linux
.venv\Scripts\activate       # Windows
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Install xlwings Excel add-in
```bash
xlwings addin install
```

---

## What the Tool Does

1. **Parses Carta export**: Extracts company name, date, stakeholders, share counts
2. **Populates template**:
   - Company name → Cell I6
   - Cap table date → Cell I7
   - Share class headers → Row 30 (columns F-I)
   - Top 9 investors by shares → Rows 31-39
   - Remaining investors → "Other Investors" row 40
3. **Preserves all formulas** in other sheets (Pro Forma, Waterfall, Summary)
4. **Saves output** as `{CompanyName}_Cap_Table_{Date}.xlsx`

## Limitations

- Maximum 9 individual investor rows (+ "Other Investors" bucket)
- Maximum 4 share classes mapped to columns F-I
- Round metadata (rows 10-19) left untouched - edit manually after import
- Options from all plans are combined into single Options column

---

## Troubleshooting

### "Python not found"
- Ensure Python is installed and in your PATH
- Windows: Reinstall Python and check "Add Python to PATH"
- Mac: Run `brew install python3`

### "Module not found"
- Make sure you've activated the virtual environment
- Run `pip install -r requirements.txt` again

### "Template not found"
- Ensure `Cap Table Template.xlsx` is in the `templates/` folder

### xlwings issues
- Close Excel completely, then run: `xlwings addin install`
- Update interpreter path in `excel/xlwings.conf`

### Formulas show #REF! errors
- Open and save the output file in Excel to recalculate
