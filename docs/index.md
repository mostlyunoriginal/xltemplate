# xltemplate

A clean, OOP interface for populating Excel templates with DataFrames.

---

## What is xltemplate?

`xltemplate` provides a stateful, object-oriented API for loading Excel templates, writing pandas or polars DataFrames to specific locations, and saving the result—all while **preserving existing formatting and formulas**.

### Key Features

- **DataFrame support** — Works with both pandas and polars
- **Preserves formatting** — Cell styles, fonts, and colors remain intact
- **Preserves formulas** — Existing formulas are not overwritten by default
- **Method chaining** — Fluent API for writing multiple DataFrames
- **Context manager** — Auto-close with `with Workbook(...) as wb:`

---

## Quick Example

```python
from xltemplate import Workbook

# Load an existing template
with Workbook("template.xlsx") as wb:
    # Write a DataFrame starting at row 5, column 2
    wb.sheet("Data").write_df(df, row=5, col=2)
    
    # Write a single value
    wb.sheet("Summary").write_value("Report Title", row=1, col=1)
    
    # Save to a new file
    wb.save("output.xlsx")
```

---

## Installation

```bash
pip install xltemplate
```

### Optional Dependencies

```bash
# pandas support
pip install xltemplate[pandas]

# polars support
pip install xltemplate[polars]

# both
pip install xltemplate[all]
```

---

## Next Steps

- **[Tutorials](tutorials/index.md)** — Step-by-step guides
- **[API Reference](reference/api.md)** — Complete documentation

---

## Links

- [GitHub Repository](https://github.com/mostlyunoriginal/xltemplate)
