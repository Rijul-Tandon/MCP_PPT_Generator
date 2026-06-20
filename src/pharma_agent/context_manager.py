from __future__ import annotations

import re
from pathlib import Path

from .models import ContextValidationResult, UseCaseDescriptor


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

REQUIRED_SECTIONS = {
    "project name",
    "client / brand / indication",
    "use-case type",
    "audience",
    "slide count target",
    "presentation objective",
    "current business question",
    "approved facts and numbers",
    "patient funnel or workflow summary",
    "analytical outputs available",
    "draft recommendations or hypotheses",
    "slides to include or avoid",
    "constraints / caveats",
}

SHORT_FORM_SECTIONS = {
    "project name",
    "client / brand / indication",
    "use-case type",
    "audience",
    "template deck",
    "slide count target",
    "generate json output",
}


class ContextManager:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.context_root = root / "context"

    def list_use_cases(self) -> list[UseCaseDescriptor]:
        use_cases: list[UseCaseDescriptor] = []
        if not self.context_root.exists():
            return use_cases
        for folder in sorted(path for path in self.context_root.iterdir() if path.is_dir()):
            reference_decks = folder / "reference_decks"
            use_cases.append(
                UseCaseDescriptor(
                    id=folder.name.lower(),
                    name=folder.name.replace("_", " "),
                    path=str(folder),
                    hasContextFile=(folder / "context.txt").exists(),
                    referenceDeckCount=len(list(reference_decks.glob("*.pptx"))) if reference_decks.exists() else 0,
                )
            )
        return use_cases

    def resolve_use_case(self, use_case_id: str) -> Path:
        normalized = use_case_id.strip().lower()
        for descriptor in self.list_use_cases():
            if descriptor.id == normalized:
                return Path(descriptor.path)
        raise FileNotFoundError(f"Use case '{use_case_id}' was not found under {self.context_root}.")

    def resolve_context_path(self, use_case_dir: Path, override_path: Path | None = None) -> Path:
        if override_path is not None:
            return override_path
        return use_case_dir / "context.txt"

    def load_context(self, context_path: Path) -> str:
        if not context_path.exists():
            raise FileNotFoundError(
                f"context.txt was not found at {context_path}. "
                f"Create it in the use-case folder and retry."
            )
        return context_path.read_text(encoding="utf-8")

    def validate_context(self, text: str) -> ContextValidationResult:
        sections = self.parse_sections(text)
        missing: list[str] = []
        weak: list[str] = []
        warnings: list[str] = []

        for key in REQUIRED_SECTIONS:
            label = SECTION_LABELS[key]
            content = sections.get(key, "").strip()
            if not content:
                missing.append(label)
            elif key not in SHORT_FORM_SECTIONS and len(content.split()) < 8:
                weak.append(label)

        slide_count = sections.get("slide count target", "").strip()
        if slide_count:
            try:
                count = int(slide_count)
                if count <= 0:
                    warnings.append("Slide Count Target must be a positive integer.")
            except ValueError:
                warnings.append("Slide Count Target must be a whole number such as 8 or 12.")

        approved_facts = sections.get("approved facts and numbers", "")
        if approved_facts and not re.search(r"\d", approved_facts):
            warnings.append(
                "Approved Facts and Numbers does not include any numerals. "
                "Add validated metrics if you expect quantitative slides."
            )

        template_deck = sections.get("template deck", "").strip()
        if not template_deck:
            warnings.append(
                "Template Deck is blank. The tool will fall back to the first deck in reference_decks/."
            )

        generate_json = sections.get("generate json output", "false").strip().lower()
        if generate_json not in {"true", "false", "yes", "no", "1", "0", ""}:
            warnings.append("Generate JSON Output should be true or false.")

        return ContextValidationResult(
            isValid=not missing and len(weak) <= 2,
            missingSections=sorted(missing),
            weakSections=sorted(weak),
            warnings=warnings,
        )

    def parse_sections(self, text: str) -> dict[str, str]:
        matches = list(re.finditer(r"(?m)^##\s*(.+?)\s*$", text))
        if not matches:
            return {}
        sections: dict[str, str] = {}
        for index, match in enumerate(matches):
            name = match.group(1).strip().lower()
            start = match.end()
            end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
            sections[name] = text[start:end].strip()
        return sections

    def resolve_generate_json_output(self, sections: dict[str, str], cli_override: bool | None = None) -> bool:
        if cli_override is not None:
            return cli_override
        raw = sections.get("generate json output", "false").strip().lower()
        return raw in {"true", "yes", "1"}
