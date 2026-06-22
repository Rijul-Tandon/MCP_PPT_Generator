# Pharma Presentation Agent (MCP Server)

This repository provides an automated PowerPoint Presentation Generator for pharmaceutical business-development tasks.

**Important Architectural Note:** 
This system was recently refactored into a **Model Context Protocol (MCP) Server**. Instead of running static CLI commands, this repository is designed to be connected to an AI Assistant (like Cursor, Claude Desktop, or custom MCP clients). The AI acts as the orchestrator, dynamically using the tools defined here to talk to the user, gather context, query Excel files, and finally build the PowerPoint.

## Current Capabilities
Currently, the MCP server exposes tools to iteratively build context and query data:

**Phase 1: Context Management**
- `list_use_cases()`: Discovers available project folders inside `context/`.
- `read_context(use_case)`: Reads the current `context.txt` file so the AI can understand the project goals.
- `write_context(use_case, full_content)`: Allows the AI to save new information gathered from the user directly into the `context.txt` file.

**Phase 2: Dynamic Data Querying**
- `list_data_sources(use_case)`: Lists available `.xlsx`, `.csv`, and `.pptx` files for a project.
- `query_excel_data(file_path, sql_query)`: Executes standard SQL queries against Excel files (via DuckDB and Pandas) so the AI can extract precise facts.
- `extract_ppt_text(file_path)`: Pulls text and layout hints from older reference presentations.

**Phase 3: PowerPoint Generation**
- `generate_powerpoint(use_case, slide_definitions, template_path)`: Uses the `python-pptx` library to build the final slides. It strictly follows the layout of the provided master template. If specific charts or visuals are missing, it intelligently injects a "Guided Visual Required" text block to help the human user finish the deck perfectly.

## Repo structure

```text
context/
  Referral_Analysis/
    context.txt
    reference_decks/
    excel_context/
src/pharma_agent/
  mcp_server.py           # The MCP Server definition and tools
  context_manager.py      # Helpers for reading/writing context
  presentation_builder.py # Engine for writing PowerPoint slides (Phase 3 WIP)
main.py                   # Entry point for the MCP Server
requirements.txt
```

## Setup

### 1. Create a virtual environment

```powershell
python -m venv bd_venv
```

### 2. Activate it

```powershell
.\bd_venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

### 4. Run the MCP Server

You can run the server directly (it will listen on standard input/output for MCP communication):

```powershell
python main.py
```

To integrate with an MCP client (e.g., Claude Desktop), configure it with:
```json
{
  "mcpServers": {
    "pharma_agent": {
      "command": "python",
      "args": ["/absolute/path/to/main.py"]
    }
  }
}
```
