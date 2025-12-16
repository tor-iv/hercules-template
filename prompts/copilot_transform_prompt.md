# Copilot 365 Prompt: Transform Carta Export to Cap Table Template

Copy and paste this prompt into Microsoft Copilot when you have both your Carta export and the Cap Table Template open in Excel.

---

## Prompt

```
Transform the data from my Carta export into the Cap Table Template using these rules:

**Source Data (Carta Export - "Detailed Cap" sheet):**
- Company name: First row, first cell (remove "Detailed Capitalization Table" text)
- Cap table date: Look for "As of MM/DD/YYYY" in the first few rows
- Stakeholder data: Rows after the header row containing "Name"
- Share class columns: Any column with "Class" or "Series" AND "Units" in the header
- Options columns: Any column with "Option" or "RSU" in the header
- Ignore summary rows containing: "total", "outstanding", "available", "fully diluted", "percentage", "price per"

**Target Template (Inputs sheet):**
1. Put company name in cell I6
2. Put cap table date in cell I7
3. Put share class names (cleaned up, without "Units" suffix) in row 30, columns F through I (max 4 classes)
4. Sort all stakeholders by total shares (descending)
5. Put the TOP 9 investors in rows 31-39:
   - Investor name in column D
   - Share counts for each class in columns F-I (matching the headers)
   - Options total in column Q
   - Put 0 in column P (Common)
6. For row 40 "Other Investors":
   - Sum all remaining investors (after top 9) for each share class
   - Sum their options in column Q
7. Leave rows 41+ unchanged (Management section)

**Important:**
- Keep all existing formulas in the template
- Only modify the Inputs sheet
- Use 0 for empty numeric cells, not blank
- Clean share class names: "Class A Units (CA)" becomes "Class A"
```

---

## Alternative Shorter Prompt

```
I have a Carta cap table export open. Transform it to my Cap Table Template:

1. Company name (cell A1 without "Detailed Capitalization Table") → Template Inputs!I6
2. Date from "As of" row → Template Inputs!I7
3. Share class headers → Template row 30, columns F-I
4. Top 9 investors by total shares → Template rows 31-39 (Name in D, shares in F-I, options in Q)
5. Remaining investors summed → Template row 40 "Other Investors"

Sort investors by total shares descending. Ignore summary rows (total, outstanding, etc.).
```

---

## Step-by-Step Version (for complex exports)

```
Help me transform this Carta export step by step:

STEP 1: Find the company name in the Carta export (first cell, remove "Detailed Capitalization Table") and put it in the Template's Inputs sheet cell I6.

STEP 2: Find the date (look for "As of MM/DD/YYYY") and put it in Inputs!I7.

STEP 3: Identify all share class columns in Carta (columns containing "Class" or "Series" with "Units"). List the first 4 class names.

STEP 4: Put those class names (without "Units" or parenthetical text) in Template row 30, columns F, G, H, I.

STEP 5: List all stakeholders with their total shares. Exclude any rows containing "total", "outstanding", "available", "fully diluted".

STEP 6: Sort stakeholders by total shares (descending). Put the top 9 in Template rows 31-39:
- Name in column D
- Each share class amount in the corresponding column (F-I)
- Sum of options/RSUs in column Q
- 0 in column P

STEP 7: Sum all remaining stakeholders (after top 9) and put totals in row 40 as "Other Investors".
```

---

## Tips for Best Results

1. **Open both files** in Excel before using Copilot
2. **Name your sheets clearly** - Copilot works better when it can reference sheet names
3. **If Copilot struggles**, try the step-by-step version above
4. **Verify the output** - especially check that:
   - Share class columns align correctly
   - Top 9 investors are truly the largest by total shares
   - "Other Investors" row sums are correct
   - Formulas in other sheets still work
