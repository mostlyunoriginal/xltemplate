# API Reference

Complete documentation for xltemplate classes and methods.

---

## Workbook

::: xltemplate.Workbook
    options:
      show_root_heading: true
      heading_level: 3
      members:
        - __init__
        - sheet
        - sheet_names
        - save
        - close
        - __enter__
        - __exit__

---

## Sheet

::: xltemplate.Sheet
    options:
      show_root_heading: true
      heading_level: 3
      members:
        - name
        - write_df
        - write_value

---

## TableSchema

::: xltemplate.TableSchema
    options:
      show_root_heading: true
      heading_level: 3
      members:
        - column_names
        - header_rows
        - groups
        - n_levels
        - to_multiindex
        - empty_df
        - validate_df

---

## Preserving Formulas

By default, `write_df()` will **skip cells containing formulas**. This prevents accidental overwriting of calculated fields.

```python
# Formula at C5 will NOT be overwritten
wb.sheet("Data").write_df(df, row=1, col=1, preserve_formulas=True)

# Force overwrite formulas
wb.sheet("Data").write_df(df, row=1, col=1, preserve_formulas=False)
```

---

## Preserving Formatting

Cell formatting (fonts, colors, borders, fills) is preserved by default when writing values:

```python
# Formatting preserved (default)
wb.sheet("Data").write_value("New Text", row=1, col=1, preserve_format=True)

# Clear formatting
wb.sheet("Data").write_value("New Text", row=1, col=1, preserve_format=False)
```
