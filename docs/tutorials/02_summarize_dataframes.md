# Summarizing Multiple DataFrames

This tutorial demonstrates how to create a multi-sheet Excel template with conditional formatting and populate it with multiple DataFrames.

---

## Prerequisites

Install xltemplate with plotnine for demo DataFrames:

```bash
pip install xltemplate[pandas] plotnine
```

---

## Overview

We'll create an Excel workbook with:

- **mtcars** — Motor Trend Car Road Tests (1974)
- **mpg** — Fuel economy data (1999-2008)  
- **diamonds** — Prices of 50,000 diamonds
- **Summary** — Metadata table with row/column counts

Each data sheet will have:
- Formatted headers
- Conditional formatting (green → yellow → red color scale)

---

## Step 1: Create the Template

> **Note:** In typical usage, you'll work with pre-existing Excel templates. This tutorial creates templates programmatically so the example is fully reproducible and self-contained.

First, we create a template with openpyxl that includes conditional formatting:

```python
from openpyxl import Workbook as OpenpyxlWorkbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from plotnine.data import mtcars, mpg, diamonds

# DataFrame metadata
DATAFRAMES = {
    "mtcars": {
        "df": mtcars,
        "description": "Motor Trend Car Road Tests (1974)",
    },
    "mpg": {
        "df": mpg,
        "description": "Fuel economy data (1999-2008)",
    },
    "diamonds": {
        "df": diamonds,
        "description": "Prices of 50,000 round cut diamonds",
    },
}

# Create workbook
wb = OpenpyxlWorkbook()
wb.remove(wb.active)

# Styles
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="2E86AB", fill_type="solid")

# Create sheets for each DataFrame
for sheet_name, meta in DATAFRAMES.items():
    df = meta["df"]
    ws = wb.create_sheet(title=sheet_name)
    
    # Add headers
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        ws.column_dimensions[get_column_letter(col_idx)].width = 12
    
    # Add conditional formatting (color scale)
    data_end_row = len(df) + 1
    for col_idx in range(1, len(df.columns) + 1):
        col_letter = get_column_letter(col_idx)
        cell_range = f"{col_letter}2:{col_letter}{data_end_row}"
        
        color_scale = ColorScaleRule(
            start_type="min", start_color="63BE7B",  # Green
            mid_type="percentile", mid_value=50, mid_color="FFEB84",  # Yellow
            end_type="max", end_color="F8696B",  # Red
        )
        ws.conditional_formatting.add(cell_range, color_scale)

wb.save("dataframes_template.xlsx")
wb.close()
```

---

## Step 2: Create the Summary Sheet

Add a summary sheet with formulas that reference the data sheets:

```python
# Create Summary sheet
ws_summary = wb.create_sheet(title="Summary")

summary_headers = ["Dataset", "Description", "Rows", "Columns"]
for col_idx, header in enumerate(summary_headers, start=1):
    cell = ws_summary.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill

# Add formulas
for row_idx, (sheet_name, meta) in enumerate(DATAFRAMES.items(), start=2):
    ws_summary.cell(row=row_idx, column=1, value=sheet_name)
    ws_summary.cell(row=row_idx, column=2, value=meta["description"])
    # Formula to count non-empty cells in column A minus header
    ws_summary.cell(row=row_idx, column=3, value=f"=COUNTA('{sheet_name}'!A:A)-1")
    # Formula to count columns in header row
    ws_summary.cell(row=row_idx, column=4, value=f"=COUNTA('{sheet_name}'!1:1)")
```

---

## Step 3: Populate with xltemplate

Now use xltemplate to write the DataFrames while preserving all formatting:

```python
from xltemplate import Workbook

with Workbook("dataframes_template.xlsx") as wb:
    for sheet_name, meta in DATAFRAMES.items():
        wb.sheet(sheet_name).write_df(
            meta["df"], 
            row=2,      # Below headers
            col=1, 
            headers=False  # Template has headers
        )
    
    wb.save("dataframes_summary.xlsx")
```

---

## Result

Open `dataframes_summary.xlsx` and you'll see:

| Sheet | Content |
|-------|---------|
| mtcars | 32 rows × 11 columns with color-scaled values |
| mpg | 234 rows × 11 columns with color-scaled values |
| diamonds | 53,940 rows × 10 columns with color-scaled values |
| Summary | Auto-calculated row/column counts via formulas |

The conditional formatting applies automatically—numeric columns show a gradient from green (low) through yellow to red (high).

---

## Key Takeaways

1. **Create formatting in the template** — Conditional formatting, styles, and formulas go in the template
2. **xltemplate preserves everything** — Writing data doesn't destroy your formatting
3. **Formulas auto-calculate** — Summary formulas update when Excel opens the file

---

## Next Steps

- Add more complex conditional formatting rules
- Use named ranges for cleaner formula references
- Explore the [API Reference](../reference/api.md) for all options
