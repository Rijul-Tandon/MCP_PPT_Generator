from __future__ import annotations

"""Helpers for finding, parsing, and validating `context.txt`.

The context file is the main contract between the human user and the agent.
This module exists so that every other part of the system can work with a clean,
normalized view of that file instead of re-parsing free text repeatedly.
"""

import re # Used for regular expressions to parse sections from markdown text
from pathlib import Path # Used for reliable file and directory path manipulations

from .models import ContextValidationResult, UseCaseDescriptor # Data structures for returning validation state and use case metadata


# `SECTION_LABELS` maps our normalized internal keys back to friendly names shown
# in validation messages. Keeping the labels here avoids hard-coding the same text
# in multiple error messages throughout the codebase.
SECTION_LABELS = {
    "project name": "Project Name",
    "client / brand / indication": "Client / Brand / Indication",
    "use-case type": "Use-Case Type",
    "audience": "Audience",
    "slide count target": "Slide Count Target",
    "template deck": "Template Deck",
    "generate json output": "Generate JSON Output",
    "presentation objective": "Presentation Objective",
    "current business question": "Current Business Question",
    "approved facts and numbers": "Approved Facts and Numbers",
    "patient funnel or workflow summary": "Patient Funnel or Workflow Summary",
    "analytical outputs available": "Analytical Outputs Available",
    "draft recommendations or hypotheses": "Draft Recommendations or Hypotheses",
    "slides to include or avoid": "Slides to Include or Avoid",
    "constraints / caveats": "Constraints / Caveats",
    "additional instructions": "Additional Instructions",
}

# Required sections are the minimum needed for a useful deck plan.
REQUIRED_SECTIONS = {
    "project name",
    "client / brand / indication",
    "use-case type",
    "audience",
    "slide count target",
    "presentation objective",
    "current business question",
}

# Recommended sections are important but not always blockers.
RECOMMENDED_SECTIONS = {
    "approved facts and numbers",
    "patient funnel or workflow summary",
    "analytical outputs available",
    "draft recommendations or hypotheses",
    "slides to include or avoid",
    "constraints / caveats",
}

# These sections are expected to be short labels/values rather than long prose.
SHORT_FORM_SECTIONS = {
    "project name",
    "client / brand / indication",
    "use-case type",
    "audience",
    "template deck",
    "slide count target",
    "generate json output",
}

EXCEL_EXTENSIONS = {".xlsx", ".csv", ".tsv"}

# Alias handling makes the parser tolerant of small wording differences in
# human-written context files.
ALIAS_MAP = {
    "constraints": "constraints / caveats",
    "slides to include": "slides to include or avoid",
    "slides to avoid": "slides to include or avoid",
    "data sources": "analytical outputs available",
    "methodology": "patient funnel or workflow summary",
    "project workflow / methodology": "patient funnel or workflow summary",
    "recommendations": "draft recommendations or hypotheses",
}


class ContextManager:
    """Load and validate use-case folders and `context.txt` files."""

    def __init__(self, root: Path) -> None:
        self.root = root
        self.context_root = root / "context"

    def list_use_cases(self) -> list[UseCaseDescriptor]:
        """Inspect the `context/` directory and summarize each use case.

        This method is used by the CLI to show what is available without the user
        having to inspect folders manually.
        """
        if not self.context_root.exists():
            return []

        use_cases: list[UseCaseDescriptor] = []
        for folder in sorted(path for path in self.context_root.iterdir() if path.is_dir()):
            reference_dir = folder / "reference_decks"
            use_cases.append(
                UseCaseDescriptor(
                    id=folder.name.lower(),
                    name=folder.name.replace("_", " "),
                    path=str(folder),
                    hasContextFile=(folder / "context.txt").exists(),
                    referenceDeckCount=len(list(reference_dir.glob("*.pptx"))) if reference_dir.exists() else 0,
                    excelContextFileCount=len(self.list_excel_context_files(folder)),
                )
            )
        return use_cases

    def resolve_use_case(self, use_case_id: str) -> Path:
        """Convert a CLI-facing use-case id into a real folder path."""
        normalized = use_case_id.strip().lower()
        for descriptor in self.list_use_cases():
            if descriptor.id == normalized:
                return Path(descriptor.path)
        raise FileNotFoundError(f"Use case '{use_case_id}' was not found under {self.context_root}.")

    def resolve_context_path(self, use_case_dir: Path, override_path: Path | None = None) -> Path:
        """Choose either the explicit `--context` file or the default context.txt."""
        return override_path if override_path is not None else use_case_dir / "context.txt"

    def load_context(self, context_path: Path) -> str:
        """Read the context file or raise a user-friendly error."""
        if not context_path.exists():
            raise FileNotFoundError(
                f"context.txt was not found at {context_path}. Create it in the use-case folder and retry."
            )
        return context_path.read_text(encoding="utf-8")

    def list_excel_context_files(self, use_case_dir: Path) -> list[Path]:
        """List spreadsheet-like files that can contribute runtime context."""
        excel_dir = use_case_dir / "excel_context"
        if not excel_dir.exists():
            return []
        return sorted(path for path in excel_dir.iterdir() if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS)

    def validate_context(self, text: str) -> ContextValidationResult:
        """Run all context checks and bundle them into one result object."""
        sections = self.parse_sections(text)
        missing = self._find_missing_sections(sections)
        weak = self._find_weak_sections(sections)
        warnings = self._build_validation_warnings(sections)
        return ContextValidationResult(
            isValid=not missing and len(weak) <= 2,
            missingSections=sorted(missing),
            weakSections=sorted(weak),
            warnings=warnings,
        )

    def parse_sections(self, text: str) -> dict[str, str]:
        """Split markdown-like `## Section Name` blocks into a dictionary.

        The parser keeps things intentionally simple. We rely on the user keeping
        a readable text template rather than inventing a complicated config format.
        """
        matches = list(re.finditer(r"(?m)^##\s*(.+?)\s*$", text))
        sections: dict[str, str] = {}
        for index, match in enumerate(matches):
            raw_name = match.group(1).strip().lower()
            name = ALIAS_MAP.get(raw_name, raw_name)
            start = match.end()
            end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
            sections[name] = text[start:end].strip()
        return sections

    def resolve_generate_json_output(self, sections: dict[str, str], cli_override: bool | None = None) -> bool:
        """Decide whether intermediate JSON/Markdown artifacts should be kept."""
        if cli_override is not None:
            return cli_override
        return sections.get("generate json output", "false").strip().lower() in {"true", "yes", "1"}

    def _find_missing_sections(self, sections: dict[str, str]) -> list[str]:
        """Identify required sections that are missing or completely empty."""
        missing: list[str] = []
        for key in REQUIRED_SECTIONS:
            if not sections.get(key, "").strip():
                missing.append(SECTION_LABELS[key])
        return missing

    def _find_weak_sections(self, sections: dict[str, str]) -> list[str]:
        """Mark long-form sections that exist but are probably too thin to trust."""
        weak: list[str] = []
        for key in REQUIRED_SECTIONS:
            content = sections.get(key, "").strip()
            if content and key not in SHORT_FORM_SECTIONS and len(content.split()) < 8:
                weak.append(SECTION_LABELS[key])
        return weak

    def _build_validation_warnings(self, sections: dict[str, str]) -> list[str]:
        """Collect softer issues that should be shown to the user but not block every run."""
        warnings = self._recommended_section_warnings(sections)
        warnings.extend(self._slide_count_warnings(sections.get("slide count target", "").strip()))
        warnings.extend(self._approved_facts_warnings(sections.get("approved facts and numbers", "")))
        warnings.extend(self._template_warnings(sections.get("template deck", "").strip()))
        warnings.extend(self._json_flag_warnings(sections.get("generate json output", "false").strip().lower()))
        return warnings

    def _recommended_section_warnings(self, sections: dict[str, str]) -> list[str]:
        """Generate warnings if recommended (but non-blocking) sections are missing."""
        warnings: list[str] = []
        for key in RECOMMENDED_SECTIONS:
            if not sections.get(key, "").strip():
                warnings.append(f"Recommended section missing: {SECTION_LABELS[key]}.")
        return warnings

    def _slide_count_warnings(self, slide_count: str) -> list[str]:
        """Validate that the slide count is a properly formatted positive integer."""
        if not slide_count:
            return []
        try:
            if int(slide_count) <= 0:
                return ["Slide Count Target must be a positive integer."]
        except ValueError:
            return ["Slide Count Target must be a whole number such as 8 or 12."]
        return []

    def _approved_facts_warnings(self, approved_facts: str) -> list[str]:
        # Quantitative slides are common in this repo, so we warn if the supposedly
        # factual section contains no numerals at all.
        if approved_facts and not re.search(r"\d", approved_facts):
            return [
                "Approved Facts and Numbers does not include any numerals. Add validated metrics if you expect quantitative slides."
            ]
        return []

    def _template_warnings(self, template_deck: str) -> list[str]:
        """Warn if a template deck is not specified, as the system will use a default."""
        if template_deck:
            return []
        return ["Template Deck is blank. The tool will fall back to the first deck in reference_decks/."]

    def _json_flag_warnings(self, generate_json: str) -> list[str]:
        """Warn if the generate JSON flag is provided in an unexpected format."""
        if generate_json in {"true", "false", "yes", "no", "1", "0", ""}:
            return []
        return ["Generate JSON Output should be true or false."]
