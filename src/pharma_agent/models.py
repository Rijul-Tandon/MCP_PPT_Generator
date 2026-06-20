from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class UseCaseDescriptor:
    id: str
    name: str
    path: str
    hasContextFile: bool
    referenceDeckCount: int


@dataclass
class ContextValidationResult:
    isValid: bool
    missingSections: list[str] = field(default_factory=list)
    weakSections: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class DeckContextChunk:
    sourceFile: str
    slideNumber: int
    text: str
    tags: list[str] = field(default_factory=list)


@dataclass
class StylePattern:
    sectionType: str
    commonTitles: list[str]
    toneNotes: str
    layoutHints: str


@dataclass
class OutlineRequest:
    useCaseId: str
    contextPath: str
    presentationGoal: str = ""
    audience: str = ""
    slideCountTarget: int | None = None
    generateJsonOutput: bool = False


@dataclass
class SlideDraft:
    slideNumber: int
    title: str
    objective: str
    bullets: list[str]
    visualSuggestion: str
    claimSources: list[str] = field(default_factory=list)
    styleSources: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    status: str = "draft"
    archetype: str = "insight"
    layoutHint: str = ""


@dataclass
class OutlineResponse:
    title: str
    slides: list[SlideDraft]
    templateDeckPath: str = ""
    contextFile: str = ""
    generateJsonOutput: bool = False
    referenceDecks: list[str] = field(default_factory=list)
    openQuestions: list[str] = field(default_factory=list)
    validationWarnings: list[str] = field(default_factory=list)


@dataclass
class BuildResult:
    pptxPath: str
    tracePath: str = ""
    warnings: list[str] = field(default_factory=list)


def dataclass_to_dict(value: Any) -> Any:
    if hasattr(value, "__dataclass_fields__"):
        return asdict(value)
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, list):
        return [dataclass_to_dict(item) for item in value]
    if isinstance(value, dict):
        return {key: dataclass_to_dict(item) for key, item in value.items()}
    return value
