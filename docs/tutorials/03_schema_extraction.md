# Schema Extraction with MultiIndex DataFrames

This tutorial demonstrates how to extract a hierarchical column schema from an Excel template and create a matching DataFrame for data population.

---

## Prerequisites

```bash
pip install xltemplate[pandas] plotnine openpyxl
```

---

## Overview

We'll build a summary table for the `mtcars` dataset with a **3-level hierarchical header**:

| Level | Variable | Description |
|-------|----------|-------------|
| Top | `vs` | Engine shape (0 = V-shaped, 1 = Straight) |
| Middle | `am` | Transmission (0 = Automatic, 1 = Manual) |
| Leaf | `cyl` | Number of cylinders (4, 6, or 8) |

Each cell will contain **Count**, **Mean MPG**, and **Std Dev MPG** for the corresponding group.

---

## Step 1: Explore the Data

First, identify which (vs, am, cyl) combinations exist in mtcars:

```python
from plotnine.data import mtcars
import pandas as pd

combos = (
    mtcars
    .groupby(["vs", "am", "cyl"])
    .agg(count=("mpg", "size"), mean=("mpg", "mean"), std=("mpg", "std"))
    .reset_index()
)
print(combos)
```

Output:

| vs | am | cyl | count | mean | std |
|----|----|-----|-------|------|-----|
| 0 | 0 | 8 | 12 | 15.05 | 2.77 |
| 0 | 1 | 4 | 1 | 26.00 | NaN |
| 0 | 1 | 6 | 3 | 20.57 | 0.75 |
| 0 | 1 | 8 | 2 | 15.40 | 0.57 |
| 1 | 0 | 4 | 3 | 22.90 | 1.45 |
| 1 | 0 | 6 | 4 | 19.12 | 1.63 |
| 1 | 1 | 4 | 7 | 28.37 | 4.76 |

Note: Only 7 of the possible 12 combinations exist—we'll build the template to match.

---

## Step 2: Create the Template

> **Note:** In typical usage, you'll work with pre-existing Excel templates. This tutorial creates the template programmatically so the example is fully reproducible and self-contained.

Build an Excel template with merged cells reflecting the hierarchy:

```python
from openpyxl import Workbook as OpenpyxlWorkbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from plotnine.data import mtcars

combos = (
    mtcars
    .groupby(["vs", "am", "cyl"])
    .size()
    .reset_index(name="n")
    .sort_values(["vs", "am", "cyl"])
)

# Group by vs and am to determine column spans
structure = (
    combos
    .groupby(["vs", "am"])["cyl"]
    .apply(list)
    .reset_index()
)

# Create workbook
wb = OpenpyxlWorkbook()
ws = wb.active
ws.title = "MPG Summary"

# Styles
header_font = Font(bold=True, color="FFFFFF")
header_fill_vs = PatternFill(start_color="2E4057", fill_type="solid")
header_fill_am = PatternFill(start_color="048A81", fill_type="solid")
header_fill_cyl = PatternFill(start_color="54C6EB", fill_type="solid")
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Column A: Row labels (created first to avoid column shift issues)
ws.cell(row=1, column=1, value="Statistic").font = Font(bold=True)
ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1)
ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")
ws.column_dimensions["A"].width = 12

stats = ["Count", "Mean", "Std Dev"]
for i, stat in enumerate(stats, start=4):
    ws.cell(row=i, column=1, value=stat).font = Font(bold=True)

# Build data columns starting at column 2
col_idx = 2
vs_spans = {}  # {vs: [start_col, end_col]}
am_spans = {}  # {(vs, am): [start_col, end_col]}

for _, row in structure.iterrows():
    vs, am, cyls = row["vs"], row["am"], row["cyl"]
    
    if vs not in vs_spans:
        vs_spans[vs] = [col_idx, col_idx]
    
    am_key = (vs, am)
    am_spans[am_key] = [col_idx, col_idx + len(cyls) - 1]
    vs_spans[vs][1] = col_idx + len(cyls) - 1
    
    # Write cyl headers (row 3)
    for cyl in cyls:
        cell = ws.cell(row=3, column=col_idx, value=f"{cyl} cyl")
        cell.font = header_font
        cell.fill = header_fill_cyl
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = 10
        col_idx += 1

# Write vs headers (row 1) with merging
for vs, (start, end) in vs_spans.items():
    label = "V-Shaped (vs=0)" if vs == 0 else "Straight (vs=1)"
    cell = ws.cell(row=1, column=start, value=label)
    cell.font = header_font
    cell.fill = header_fill_vs
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border
    if end > start:
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)

# Write am headers (row 2) with merging
for (vs, am), (start, end) in am_spans.items():
    label = "Automatic (am=0)" if am == 0 else "Manual (am=1)"
    cell = ws.cell(row=2, column=start, value=label)
    cell.font = header_font
    cell.fill = header_fill_am
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border
    if end > start:
        ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=end)

# Apply number formats to data cells
for row in range(4, 7):
    for col in range(2, col_idx):
        cell = ws.cell(row=row, column=col)
        if row == 4:  # Count - integer
            cell.number_format = "0"
        else:  # Mean, Std Dev - one decimal
            cell.number_format = "0.0"

wb.save("mtcars_summary_template.xlsx")
wb.close()
print("Created mtcars_summary_template.xlsx")
```


---

## Step 3: Extract Schema and Create Matching DataFrame

Now use xltemplate to read the template's column structure:

```python
from xltemplate import Workbook

with Workbook("mtcars_summary_template.xlsx") as wb:
    schema = (
        wb.sheet("MPG Summary")
        .extract_header_schema(row=1, col=2, n_cols=7, n_header_rows=3)
    )

print(f"Schema has {schema.n_levels} levels and {len(schema)} columns")
print("\nMultiIndex columns:")
for col in schema.to_multiindex():
    print(f"  {col}")
```

Output:

```
Schema has 3 levels and 7 columns

MultiIndex columns:
  ('V-Shaped (vs=0)', 'Automatic (am=0)', '8 cyl')
  ('V-Shaped (vs=0)', 'Manual (am=1)', '4 cyl')
  ('V-Shaped (vs=0)', 'Manual (am=1)', '6 cyl')
  ('V-Shaped (vs=0)', 'Manual (am=1)', '8 cyl')
  ('Straight (vs=1)', 'Automatic (am=0)', '4 cyl')
  ('Straight (vs=1)', 'Automatic (am=0)', '6 cyl')
  ('Straight (vs=1)', 'Manual (am=1)', '4 cyl')
```

---

## Step 4: Populate Using MultiIndex Access

Create a DataFrame from the schema and fill it by addressing columns hierarchically:

```python
from plotnine.data import mtcars
import numpy as np

# Create empty DataFrame with 3 rows (Count, Mean, Std)
df = schema.empty_df(n_rows=3)

# Calculate statistics for each group
stats = (
    mtcars
    .groupby(["vs", "am", "cyl"])
    .agg(count=("mpg", "size"), mean=("mpg", "mean"), std=("mpg", "std"))
)

# Map (vs, am, cyl) to column labels
label_map = {
    (0, 0): ("V-Shaped (vs=0)", "Automatic (am=0)"),
    (0, 1): ("V-Shaped (vs=0)", "Manual (am=1)"),
    (1, 0): ("Straight (vs=1)", "Automatic (am=0)"),
    (1, 1): ("Straight (vs=1)", "Manual (am=1)"),
}

# Fill DataFrame by MultiIndex
for (vs, am, cyl), row in stats.iterrows():
    vs_label, am_label = label_map[(vs, am)]
    cyl_label = f"{cyl} cyl"
    col = (vs_label, am_label, cyl_label)
    
    df.loc[0, col] = row["count"]
    df.loc[1, col] = row["mean"]
    df.loc[2, col] = np.nan if pd.isna(row["std"]) else row["std"]

print("Populated DataFrame:")
print(df)
```

---

## Step 5: Write to Template

Finally, write the data back to the template:

```python
from xltemplate import Workbook

with Workbook("mtcars_summary_template.xlsx") as wb:
    # Extract schema
    schema = (
        wb.sheet("MPG Summary")
        .extract_header_schema(row=1, col=2, n_cols=7, n_header_rows=3)
    )
    
    # Validate before writing
    if not schema.validate_df(df):
        raise ValueError("DataFrame doesn't match template schema!")
    
    # Write data (row 4 = first data row, col 2 = after row labels)
    wb.sheet("MPG Summary").write_df(df, row=4, col=2, headers=False)
    wb.save("mtcars_summary_filled.xlsx")

print("Saved mtcars_summary_filled.xlsx")
```

---

## Result

Open `mtcars_summary_filled.xlsx`:

| Statistic | V-Shaped Automatic 8cyl | V-Shaped Manual 4cyl | ... | Straight Manual 4cyl |
|-----------|-------------------------|----------------------|-----|----------------------|
| Count | 12 | 1 | ... | 7 |
| Mean | 15.1 | 26.0 | ... | 28.4 |
| Std Dev | 2.8 | — | ... | 4.8 |

The template's number formatting (integers for Count, decimals for Mean/Std) is preserved.

---

## Key Takeaways

1. **`extract_header_schema()`** reads hierarchical headers from templates
2. **`empty_df()`** creates a DataFrame with MultiIndex columns matching the template
3. **MultiIndex access** lets you fill cells by logical path: `df[("Group", "Subgroup", "Column")]`
4. **`validate_df()`** ensures your data matches before writing

---

## Next Steps

- Explore more complex header structures with 4+ levels
- Use `header_rows` to inspect intermediate groupings
- See the [API Reference](../reference/api.md) for all `TableSchema` methods
