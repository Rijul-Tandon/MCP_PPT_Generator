"""
tools_query.py — Data-source helpers used by the Excel and PPT MCP servers.

list_data_sources : scans context/<use_case>/ and returns Excel + PPT file paths.
query_excel_data  : executes DuckDB SQL against a Pandas DataFrame loaded from Excel/CSV.

Note: extract_ppt_text was removed. The LLM now reads raw OOXML directly via
      ppt_server.get_slide_xml, which is lossless compared to the old text-only approach.
"""

import duckdb
import pandas as pd
from pathlib import Path


def list_data_sources(base_dir: Path, use_case: str) -> str:
    """
    Scan context/<use_case>/excel_context/ and context/<use_case>/reference_decks/
    and return all usable file paths.
    Returns absolute paths so MCP tools can pass them directly to get_deck_manifest
    or get_slide_xml without guessing.
    """
    folder = base_dir / "context" / use_case
    if not folder.exists():
        return f"Use case '{use_case}' not found in context/."

    lines: list[str] = []

    # Excel / CSV sources
    excel_dir = folder / "excel_context"
    lines.append("=== Excel Sources ===")
    if excel_dir.exists():
        found = [f for f in excel_dir.glob("*.*") if f.suffix in (".xlsx", ".csv", ".tsv")]
        lines.extend(str(f.resolve()) for f in found) if found else lines.append("  (none)")
    else:
        lines.append("  (no excel_context folder)")

    # PPT reference decks
    ref_dir = folder / "reference_decks"
    lines.append("\n=== PPT Reference Decks ===")
    if ref_dir.exists():
        found = list(ref_dir.glob("*.pptx"))
        lines.extend(str(f.resolve()) for f in found) if found else lines.append("  (none)")
    else:
        lines.append("  (no reference_decks folder)")

    return "\n".join(lines)


def query_excel_data(file_path: str, sql_query: str) -> str:
    """
    Execute a SQL query against an Excel or CSV file using DuckDB.

    The loaded file is registered as the table 'df'.
    Column names are normalised (spaces → underscores) for SQL compatibility.
    Example: SELECT HCP_Name, SUM(Referral_Count) FROM df GROUP BY HCP_Name LIMIT 10
    """
    path = Path(file_path)
    if not path.exists():
        return f"File not found: {file_path}"

    try:
        if path.suffix == ".xlsx":
            df = pd.read_excel(path)
        elif path.suffix == ".csv":
            df = pd.read_csv(path)
        elif path.suffix == ".tsv":
            df = pd.read_csv(path, sep="\t")
        else:
            return f"Unsupported file type: {path.suffix}"

        # Normalise column names so SQL identifiers are safe
        df.columns = [
            str(c).strip().replace(" ", "_").replace(".", "_").replace("-", "_")
            for c in df.columns
        ]

        result_df = duckdb.query(sql_query).df()
        return result_df.to_markdown(index=False)

    except Exception as e:
        return f"Query Error: {e}"
