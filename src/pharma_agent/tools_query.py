"""Tools for dynamically querying data sources (Excel, PPT)."""
import pandas as pd
import duckdb
from pathlib import Path
from .pptx_reference import PptxReferenceLibrary

def list_data_sources(base_dir: Path, use_case: str) -> str:
    """List available Excel and PPT files for a given use case."""
    folder = base_dir / "context" / use_case
    if not folder.exists():
        return f"Use case '{use_case}' not found."
    
    excel_dir = folder / "excel_context"
    ref_dir = folder / "reference_decks"
    
    output = []
    output.append("=== Excel Sources ===")
    if excel_dir.exists():
        for f in excel_dir.glob("*.*"):
            if f.suffix in ['.xlsx', '.csv', '.tsv']:
                output.append(str(f.relative_to(base_dir)))
    else:
        output.append("No excel_context folder.")
        
    output.append("\n=== PPT Reference Sources ===")
    if ref_dir.exists():
        for f in ref_dir.glob("*.pptx"):
            output.append(str(f.relative_to(base_dir)))
    else:
        output.append("No reference_decks folder.")
        
    return "\n".join(output)

def query_excel_data(file_path: str, sql_query: str) -> str:
    """Execute a SQL query against an Excel/CSV file."""
    path = Path(file_path)
    if not path.exists():
        return f"File not found: {file_path}"
    
    try:
        # Load the file into a pandas dataframe
        if path.suffix == '.xlsx':
            df = pd.read_excel(path)
        elif path.suffix == '.csv':
            df = pd.read_csv(path)
        elif path.suffix == '.tsv':
            df = pd.read_csv(path, sep='\t')
        else:
            return f"Unsupported file type: {path.suffix}"
            
        # Clean column names to make SQL easier (remove spaces/special chars)
        df.columns = [str(c).strip().replace(' ', '_').replace('.', '_').replace('-', '_') for c in df.columns]
            
        # Execute DuckDB query against the dataframe 'df'
        result_df = duckdb.query(sql_query).df()
        
        # Convert result to a markdown table or string
        return result_df.to_markdown(index=False)
        
    except Exception as e:
        return f"Query Error: {e}"

def extract_ppt_text(file_path: str) -> str:
    """Extract all text chunks from a specific PPTX file."""
    path = Path(file_path)
    if not path.exists() or path.suffix != '.pptx':
        return f"Invalid or missing PPTX file: {file_path}"
        
    try:
        library = PptxReferenceLibrary(path.parent)
        chunks = library.extract_pptx_text(path)
        
        output = []
        for c in chunks:
            output.append(f"[Slide {c.slideNumber}]: {c.text}")
            
        return "\n".join(output)
    except Exception as e:
        return f"Extraction Error: {e}"
