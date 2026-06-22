from __future__ import annotations

"""Shared data structures used across the project.

This file is intentionally simple and declarative. The rest of the codebase passes
these dataclasses around so the planning flow stays readable and so JSON output can
be generated in a predictable shape.

When a teammate wants to understand what information moves through the system,
this is the best place to start.
"""

from dataclasses import asdict, dataclass, field # Used to easily define and serialize data-holding classes
from pathlib import Path # Used to handle file path objects when converting them to dictionary/JSON format
from typing import Any # Used to type-hint arbitrary incoming data types for serialization


@dataclass
class UseCaseDescriptor:
    """One folder under `context/` as shown to the CLI.

    We surface a small amount of metadata so the user can understand whether a
    use case is actually ready to run or is still just a placeholder.
    """

    id: str
    name: str
    path: str
    hasContextFile: bool
    referenceDeckCount: int
    excelContextFileCount: int


@dataclass
class ContextValidationResult:
    """Result of checking whether `context.txt` is usable.

    `missingSections` blocks execution.
    `weakSections` means the section exists but may not be detailed enough.
    `warnings` are softer issues that the planner can work around.
    """

    isValid: bool
    missingSections: list[str] = field(default_factory=list)
    weakSections: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class DeckContextChunk:
    """A single text chunk extracted from one slide of a deck.

    We use this shape for both historical reference decks and an existing deck
    being refined. Keeping the extracted text in a normalized format makes it
    easier to build prompts and derive style patterns later.
    """

    sourceFile: str
    slideNumber: int
    text: str
    tags: list[str] = field(default_factory=list)


@dataclass
class StylePattern:
    """A lightweight summary of recurring storytelling/layout behavior.

    This does not try to perfectly model PowerPoint design. It is only meant to
    give the planner enough information to imitate the client's presentation
    rhythm without copying old factual content.
    """

    sectionType: str
    commonTitles: list[str]
    toneNotes: str
    layoutHints: str


@dataclass
class OutlineRequest:
    """All user-controlled inputs for planning/building a deck.

    `existingDeckPath` means "refine this existing deck as input context".
    `outputPptxPath` means "write the final result here".
    Those two paths may point to the same file, but they serve different roles.
    """

    useCaseId: str
    contextPath: str
    presentationGoal: str = ""
    audience: str = ""
    slideCountTarget: int | None = None
    generateJsonOutput: bool = False
    outputPptxPath: str = ""
    existingDeckPath: str = ""


@dataclass
class SlideDraft:
    """One slide in the intermediate `content_v1` representation.

    The PowerPoint writer does not invent structure on its own. It relies on
    these fields to decide what title, body text, and visual guidance should be
    placed on the slide.
    """

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
    """The full planning output before PowerPoint writing begins."""

    title: str
    slides: list[SlideDraft]
    templateDeckPath: str = ""
    contextFile: str = ""
    generateJsonOutput: bool = False
    generationMode: str = "offline"
    generationMessage: str = ""
    referenceDecks: list[str] = field(default_factory=list)
    excelContextFiles: list[str] = field(default_factory=list)
    openQuestions: list[str] = field(default_factory=list)
    validationWarnings: list[str] = field(default_factory=list)


@dataclass
class BuildResult:
    """What the CLI needs to report after PowerPoint writing finishes."""

    pptxPath: str
    tracePath: str = ""
    warnings: list[str] = field(default_factory=list)


def dataclass_to_dict(value: Any) -> Any:
    """Recursively turn dataclasses and Paths into JSON-safe values.

    The project uses dataclasses heavily because they are easy to reason about in
    Python code. JSON output, however, needs plain dictionaries/lists/strings.
    This helper bridges those two representations.
    """

    if hasattr(value, "__dataclass_fields__"):
        return asdict(value)
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, list):
        return [dataclass_to_dict(item) for item in value]
    if isinstance(value, dict):
        return {key: dataclass_to_dict(item) for key, item in value.items()}
    return value
