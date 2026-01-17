# Getting Started

This tutorial walks you through your first xltemplate workflow.

---

## Prerequisites

Install xltemplate with DataFrame support:

```bash
pip install xltemplate[pandas]
# or
pip install xltemplate[polars]
```

---

## Step 1: Create a Template

First, create an Excel template with your headers and formatting. You can use Excel, LibreOffice, or openpyxl directly:

```python
from openpyxl import Workbook as OpenpyxlWorkbook
from openpyxl.styles import Font, PatternFill

# Create a template with formatting
wb = OpenpyxlWorkbook()
ws = wb.active
ws.title = "Sales Data"

# Add formatted headers
headers = ["Date", "Product", "Quantity", "Price", "Total"]
for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="2E86AB", fill_type="solid")

# Add a formula in the Total column (row 2 as example)
ws["E2"] = "=C2*D2"

wb.save("sales_template.xlsx")
wb.close()
```

---

## Step 2: Populate the Template

Now use xltemplate to fill in your data:

```python
import pandas as pd
from xltemplate import Workbook

# Sample data
df = pd.DataFrame({
    "Date": ["2024-01-15", "2024-01-16", "2024-01-17"],
    "Product": ["Widget A", "Widget B", "Widget A"],
    "Quantity": [10, 5, 8],
    "Price": [29.99, 49.99, 29.99],
})

# Load template and write data
with Workbook("sales_template.xlsx") as wb:
    # Write DataFrame starting at row 2 (below headers), column 1
    # headers=False since template already has headers
    wb.sheet("Sales Data").write_df(df, row=2, col=1, headers=False)
    
    wb.save("sales_report.xlsx")
```

---

## Step 3: Verify the Result

Open `sales_report.xlsx` and you'll see:

- ✅ Your data in the correct location
- ✅ Header formatting preserved (bold, colored)
- ✅ Formula in column E calculating totals

---

## Method Chaining

Write to multiple locations in a fluent style:

```python
with Workbook("template.xlsx") as wb:
    (wb.sheet("Data")
       .write_df(main_data, row=5, col=1)
       .write_value("Generated: 2024-01-17", row=1, col=1))
    
    wb.save("output.xlsx")
```

---

## Next Steps

- Explore the [API Reference](../reference/api.md) for all available options
- Learn about [preserving formulas](../reference/api.md#preserving-formulas) when writing over cells
