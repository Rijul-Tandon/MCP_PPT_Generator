"""MCP Server for Excel data extraction."""
from mcp.server.fastmcp import FastMCP
from pathlib import Path
from .tools_query import list_data_sources as list_ds, query_excel_data as query_excel

mcp = FastMCP("ExcelServer")

@mcp.tool()
def list_data_sources(use_case: str) -> str:
    """List available Excel and PPT files for a given use case. Use this to find exact file paths for querying."""
    return list_ds(Path.cwd(), use_case)

@mcp.tool()
def extract_crisp_insights_from_excel(file_path: str, sql_query: str) -> str:
    """
    Execute a SQL query against an Excel/CSV file to extract crisp business facts and data for charts. 
    The file acts as a table named 'df'. 
    NOTE: Column names are automatically stripped of spaces and hyphens (e.g., 'Total Sales' becomes 'Total_Sales').
    Example SQL: "SELECT column_name, SUM(sales) FROM df GROUP BY column_name ORDER BY 2 DESC LIMIT 5"
    """
    return query_excel(file_path, sql_query)
