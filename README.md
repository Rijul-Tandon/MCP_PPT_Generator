# Pharma Presentation Agent

An AI-powered PowerPoint engine for pharmaceutical analytics teams.
The agent conducts a structured interview with the user, seamlessly clones designer-styled reference decks, and intelligently maps new content into them while guaranteeing 100% aesthetic fidelity.

---

## Architecture: "Clone & Replace"

Instead of generating blank PowerPoint shells and struggling to author raw XML from scratch, the system uses a **Clone & Replace Architecture** to ensure zero design loss.

```
┌───────────────────────────────────────────────────────────────────┐
│  User (Terminal)                                                  │
└────────────────────────┬──────────────────────────────────────────┘
                         │
                         ▼
┌───────────────────────────────────────────────────────────────────┐
│  chat.py  —  MCP Client (Cerebras/Groq)                           │
│                                                                   │
│  - Runs the interactive interview workflow                        │
│  - Uses Cerebras (primary) or Groq (fallback) based on .env       │
│  - Connects to BOTH MCP servers via stdio                         │
│                                                                   │
│  Usage:                                                           │
│    python chat.py            # Full mode (PPT + Excel)            │
│    python chat.py --no-excel # Layout-only mode (PPT only)        │
└────────────┬──────────────────────────────────┬───────────────────┘
             │ JSON-RPC (stdio)                  │ JSON-RPC (stdio)
             ▼                                   ▼
┌────────────────────────┐         ┌─────────────────────────────────┐
│  PPT MCP Server        │         │  Excel MCP Server               │
│  (run_ppt.py)          │         │  (run_excel.py)                 │
│                        │         │                                 │
│  Deck Cloning Tools:   │         │  Data tools:                    │
│  clone_deck_from_ref...│         │  list_data_sources              │
│                        │         │  extract_crisp_insights         │
│  Text Tools:           │         │    _from_excel (DuckDB SQL)     │
│  get_deck_manifest     │         └────────────────────────────────-┘
│  get_slide_text        │
│  replace_slide_text    │
│                        │
│  Context Tools:        │
│  list_use_cases        │
│  read_context          │
│  write_context         │
│  list_data_sources     │
└────────────┬───────────┘
             │ win32com + python-pptx
             ▼
┌───────────────────────────────────────────────────────────────────┐
│  .pptx file  (Output)                                             │
│                                                                   │
│  - Cloned natively via Windows COM (perfect master/theme match)   │
│  - Text replaced directly in the shapes preserving exact styles   │
└───────────────────────────────────────────────────────────────────┘
```

### Why "Clone & Replace"?

Previously, the LLM attempted to read and write raw Office Open XML (`<p:sld>`). This frequently caused **Token Limit Errors (429)** because a single slide could contain 4,000+ tokens of XML. It also caused aesthetic loss if the LLM hallucinated XML structures.

The new workflow:
1. **Clone**: The MCP server uses `win32com` to control PowerPoint natively in the background, making a pristine duplicate of the chosen reference slides. This preserves logos, complex themes, SmartArt, and master layouts exactly.
2. **Read Text**: The LLM uses `get_slide_text` (~100 tokens) to simply view the existing text boxes on the reference slides.
3. **Replace**: The LLM proposes a text mapping (Old Text -> New Text). The MCP server uses `python-pptx` to surgically overwrite the text inside the cloned shapes, perfectly preserving the original designer's paragraph formatting, fonts, and colors.

---

## MCP Tool Reference

### PPT Server (`run_ppt.py`)

| Tool | Explanation |
|---|---|
| `list_use_cases` | Scans the `context/` directory to list available project folders (e.g., `referral_analysis`). |
| `read_context` | Reads `context.txt` inside a use-case folder to give the LLM business background and objectives. |
| `write_context` | Updates the `context.txt` file with new learnings. |
| `list_data_sources` | Returns absolute paths to all Excel data files and PPT reference decks in the use-case folder. |
| `get_deck_manifest` | Returns a highly compact JSON list of slide indices and their titles (~300 tokens/deck). Allows the LLM to scan a 50-slide deck instantly without opening every slide. |
| `clone_deck_from_reference` | Uses Windows COM automation to perfectly duplicate chosen slides from a reference deck into a new output file. Guarantees 100% aesthetic fidelity. |
| `get_slide_text` | Uses `python-pptx` to cleanly extract all text strings from a given slide. Very token-efficient (~100 tokens). |
| `replace_slide_text` | Takes a JSON mapping (`{"Old String": "New String"}`) and surgically updates the shapes in the cloned deck while preserving the original paragraph/run styling. |
| `write_slide_chart_xml` | Specifically targets the raw OOXML of a chart (`<c:chartSpace>`) to inject new data values into graphs while maintaining the chart's theme colors. |

### Excel Server (`run_excel.py`)

| Tool | Explanation |
|---|---|
| `list_data_sources` | Same as above. Lists accessible `.xlsx` and `.csv` files. |
| `extract_crisp_insights_from_excel` | Uses DuckDB to run ultra-fast SQL queries against raw Excel/CSV data files to extract statistical insights for the presentation. |

---

## Repo Structure

```
chat.py                     # MCP client — Cerebras/Groq agent + interview logic
run_ppt.py                  # Entry point for the PPT MCP server
run_excel.py                # Entry point for the Excel MCP server
requirements.txt
.env                        # API keys and provider configuration
context/
  Referral_Analysis/
    context.txt
    reference_decks/        # .pptx files used for aesthetic cloning
    excel_context/          # .xlsx / .csv files queried via DuckDB
src/pharma_agent/
  ppt_server.py             # PPT server tool definitions
  excel_server.py           # Excel server tool definitions
  clone_builder.py          # win32com logic for perfect slide duplication
  context_manager.py        # context.txt read/write helpers
  tools_query.py            # list_data_sources + query_excel_data
output/                     # Generated .pptx files land here
```

---

## Setup & Running

```powershell
# 1. Create and activate virtual environment
python -m venv bd_venv
.\bd_venv\Scripts\Activate.ps1

# 2. Install dependencies (requires win32com and python-pptx)
pip install -r requirements.txt
pip install cerebras-cloud-sdk

# 3. Configure API Keys in .env
# You can use Cerebras (recommended for high token limits) or Groq
echo "LLM_PROVIDER=cerebras" > .env
echo "CEREBRAS_API_KEY=your_cerebras_key" >> .env
echo "CEREBRAS_MODEL=gpt-oss-120b" >> .env
echo "GROQ_API_KEY=your_groq_key" >> .env
echo "GROQ_MODEL=llama-3.3-70b-versatile" >> .env

# 4. Run the Agent
python chat.py              # Full mode (PPT + Excel data)
python chat.py --no-excel   # Layout-only mode (Fast, builds presentation structure only)
```

---

## Agent Workflow (Layout-Only Mode)

1. **Mode**: "New deck or refine existing?"
2. **Use Case**: Select project (e.g., Referral Analysis).
3. **Deck Discovery**: LLM lists reference decks and you choose which aesthetic style to clone.
4. **Slide Selection**: You provide the slide indices you want cloned (e.g., `0, 4, 4, 9` to duplicate slide 4 twice).
5. **Content Read**: LLM uses `get_slide_text` to understand the layout placeholders.
6. **Plan**: LLM drafts a plan of text replacements. Wait for user to type `confirmed`.
7. **Generate**: 
   - `clone_deck_from_reference` perfectly creates the new `.pptx`.
   - `replace_slide_text` injects the new content.
