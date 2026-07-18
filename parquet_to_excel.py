"""
Convert a parquet file to an Excel (.xlsx) file.

Usage:
    python parquet_to_excel.py <input.parquet> [output.xlsx]

If no output path is given, the Excel file is written next to the parquet
with the same name (e.g. validation_results.parquet -> validation_results.xlsx).

Requires:
    pip install pandas pyarrow openpyxl
"""

import sys
from pathlib import Path
import pandas as pd


def parquet_to_excel(parquet_path: Path, excel_path: Path = None) -> Path:
    """
    Read a parquet file and write it to an Excel .xlsx file.

    Returns the output Excel path.
    """
    parquet_path = Path(parquet_path)

    if not parquet_path.exists():
        raise FileNotFoundError(f"Parquet file not found: {parquet_path}")

    # Default output: same folder + name, .xlsx extension
    if excel_path is None:
        excel_path = parquet_path.with_suffix(".xlsx")
    else:
        excel_path = Path(excel_path)

    excel_path.parent.mkdir(parents=True, exist_ok=True)

    # Read parquet
    print(f"[READ]  {parquet_path}")
    df = pd.read_parquet(parquet_path)
    print(f"[INFO]  {len(df)} rows, {len(df.columns)} columns")

    # Write Excel
    print(f"[WRITE] {excel_path}")
    df.to_excel(excel_path, index=False, engine="openpyxl")

    print(f"[DONE]  Converted successfully -> {excel_path}")
    return excel_path


if __name__ == "__main__":
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\Category.parquet"
    )
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\CoaDetails.parquet"
    )
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\DataType.parquet"
    )
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\DisplayNameInfo.parquet"
    )
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\MetaData.parquet"
    )
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\RawData.parquet"
    )
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\TemplateType.parquet"
    )
    parquet_to_excel(
        r"C:\S2\Test\DST 2.0_2026\10.PFG_Final\parquet\UnitMaster.parquet"
    )