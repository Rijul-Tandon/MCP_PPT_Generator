"""MCP Server for PowerPoint Generation and Context."""
from mcp.server.fastmcp import FastMCP
from pathlib import Path
from .context_manager import ContextManager
from .tools_query import extract_ppt_text as extract_ppt

mcp = FastMCP("PPTServer")

@mcp.tool()
def list_use_cases() -> str:
    """List all available use cases in the project."""
    base_dir = Path.cwd()
    manager = ContextManager(base_dir)
    cases = manager.list_use_cases()
    if not cases:
        return "No use cases found in the context directory."
    return "\n".join([f"- {c.id}: {c.name} (Context: {c.hasContextFile})" for c in cases])

@mcp.tool()
def read_context(use_case: str) -> str:
    """Read the current context.txt for a given use case. Use this to understand the current project goals."""
    base_dir = Path.cwd()
    manager = ContextManager(base_dir)
    try:
        folder = manager.resolve_use_case(use_case)
        path = manager.resolve_context_path(folder)
        return manager.load_context(path)
    except Exception as e:
        return f"Error reading context: {e}"

@mcp.tool()
def write_context(use_case: str, full_content: str) -> str:
    """Overwrite the entire context.txt for a given use case. Use this after gathering new information from the user."""
    base_dir = Path.cwd()
    manager = ContextManager(base_dir)
    try:
        folder = manager.resolve_use_case(use_case)
        path = manager.resolve_context_path(folder)
        path.write_text(full_content, encoding="utf-8")
        return f"Successfully updated context.txt for '{use_case}'"
    except Exception as e:
        return f"Error writing context: {e}"

@mcp.tool()
def extract_ppt_text(file_path: str) -> str:
    """Extract raw text from a previous PowerPoint presentation to understand its structure, tone, or content."""
    return extract_ppt(file_path)

@mcp.tool()
def generate_powerpoint(use_case: str, slide_definitions: list[dict], template_path: str = "") -> str:
    """
    Generate the final PowerPoint presentation based on the planned slide definitions.
    slide_definitions should be a list of dictionaries, each containing:
    - title: str
    - archetype: str (e.g., 'insight', 'agenda', 'patient_funnel')
    - bullets: list of str
    - visualSuggestion: str (important guided text for missing charts/visuals)
    """
    from .presentation_builder import PresentationBuilder
    import json
    
    base_dir = Path.cwd()
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)
    
    output_file = output_dir / f"{use_case}_generated.pptx"
    
    content = {
        "title": f"{use_case.replace('_', ' ').title()} Presentation",
        "templateDeckPath": template_path,
        "slides": slide_definitions
    }
    
    builder = PresentationBuilder()
    try:
        warnings = builder.build_from_content(content, output_file)
        result = f"Successfully generated presentation at: {output_file}\n"
        if warnings:
            result += "Warnings:\n" + "\n".join(warnings)
        return result
    except Exception as e:
        return f"Failed to generate presentation: {e}"

@mcp.tool()
def read_ppt_structure(file_path: str) -> str:
    """
    Deeply read a PowerPoint deck to understand its exact slide count, placeholder layouts, and visual object types (e.g. charts, tables).
    Returns a JSON string of the deck's structural memory.
    """
    from .presentation_builder import PresentationBuilder
    import json
    from pathlib import Path
    try:
        builder = PresentationBuilder()
        structure = builder.read_structure(Path(file_path))
        return json.dumps(structure, indent=2)
    except Exception as e:
        return f"Error reading ppt structure: {e}"

@mcp.tool()
def update_existing_slide(file_path: str, slide_index: int, title: str, archetype: str, bullets: list, visual_suggestion: str = "") -> str:
    """
    Target and improve a specific slide in an existing deck in-place without overwriting the whole file.
    slide_index is 0-indexed.
    """
    from .presentation_builder import PresentationBuilder
    from pathlib import Path
    try:
        builder = PresentationBuilder()
        slide_data = {
            "title": title,
            "archetype": archetype,
            "bullets": bullets,
            "visualSuggestion": visual_suggestion
        }
        builder.update_slide(Path(file_path), slide_index, slide_data)
        return f"Successfully updated slide {slide_index} in {file_path}"
    except Exception as e:
        return f"Failed to update slide: {e}"

@mcp.tool()
def add_chart_to_slide(file_path: str, slide_index: int, chart_type: str, chart_data: dict) -> str:
    """
    Inject native Column, Line, or Pie charts directly into PowerPoint using data.
    chart_type: 'column', 'line', 'pie', or 'bar'
    chart_data schema: {"categories": ["Q1", "Q2"], "series": [{"name": "Sales", "values": [10, 20]}]}
    """
    from .presentation_builder import PresentationBuilder
    from pathlib import Path
    try:
        builder = PresentationBuilder()
        builder.add_chart(Path(file_path), slide_index, chart_type, chart_data)
        return f"Successfully injected {chart_type} chart into slide {slide_index} in {file_path}"
    except Exception as e:
        return f"Failed to add chart: {e}"
