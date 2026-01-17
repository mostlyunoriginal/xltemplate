"""Schema classes for template structure extraction."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class TableSchema:
    """
    Represents the column structure extracted from a template header.
    
    Use this to create DataFrames that match a template's expected structure,
    or to validate that existing DataFrames conform to the template.
    
    Attributes:
        column_names: List of column names in order (leaf-level for multi-row headers)
        groups: Optional list of (group_name, span) tuples for multi-level headers
    
    Example:
        >>> schema = sheet.extract_header_schema(row=3, col=2, n_cols=6)
        >>> df = schema.empty_df()
        >>> # ... populate df ...
        >>> sheet.write_df(df, row=4, col=2, headers=False)
    """
    
    column_names: list[str]
    groups: list[tuple[str, int]] | None = None
    
    def empty_df(self, n_rows: int = 0) -> Any:
        """
        Create an empty DataFrame matching this schema's column structure.
        
        Args:
            n_rows: Number of rows to pre-allocate (default: 0)
            
        Returns:
            A pandas DataFrame with columns matching the schema.
            
        Note:
            Requires pandas to be installed. Returns a pandas DataFrame
            regardless of whether you typically use Polars.
        """
        try:
            import pandas as pd
        except ImportError as e:
            raise ImportError(
                "pandas is required to use empty_df(). "
                "Install it with: pip install pandas"
            ) from e
        
        if n_rows > 0:
            return pd.DataFrame(
                index=range(n_rows),
                columns=self.column_names,
            )
        return pd.DataFrame(columns=self.column_names)
    
    def validate_df(self, df: Any) -> bool:
        """
        Check if a DataFrame's columns match this schema.
        
        Args:
            df: A pandas or Polars DataFrame to validate
            
        Returns:
            True if columns match exactly in name and order, False otherwise
        """
        if hasattr(df, "columns"):
            return list(df.columns) == self.column_names
        return False
    
    def __len__(self) -> int:
        """Return the number of columns in the schema."""
        return len(self.column_names)
