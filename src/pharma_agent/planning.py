from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path

from .context_manager import ContextManager
from .llm import GeminiClient
from .models import BuildResult, ContextValidationResult, OutlineRequest, OutlineResponse, SlideDraft, dataclass_to_dict
from .pptx_reference import PptxReferenceLibrary


class PlanningService:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.context_manager = ContextManager(root)
        self.llm = GeminiClient()

    def list_use_cases(self) -> list[dict]:
        return [asdict(item) for item in self.context_manager.list_use_cases()]

    def generate_content_plan(self, request: OutlineRequest, output_dir: Path) -> tuple[Path, Path]:
        outline, _context_path = self._prepare_outline(request)
        timestamp = self._timestamp()
        safe_name = request.useCaseId.lower()
        json_path = output_dir / f"{timestamp}_{safe_name}_content_v1.json"
        md_path = output_dir / f"{timestamp}_{safe_name}_content_v1.md"
        json_path.write_text(json.dumps(dataclass_to_dict(outline), indent=2), encoding="utf-8")
        md_path.write_text(self._outline_to_markdown(outline), encoding="utf-8")
        return json_path, md_path

    def build_presentation(self, content_path: Path, output_dir: Path, generate_json_output: bool = False) -> BuildResult:
        data = json.loads(content_path.read_text(encoding="utf-8"))
        return self._build_from_content_dict(data, output_dir, content_path.stem.replace("_content_v1", ""), generate_json_output)

    def run_pipeline(self, request: OutlineRequest, output_dir: Path) -> BuildResult:
        outline, _context_path = self._prepare_outline(request)
        content = dataclass_to_dict(outline)
        base_name = f"{self._timestamp()}_{request.useCaseId.lower()}"
        return self._build_from_content_dict(content, output_dir, base_name, request.generateJsonOutput)

    def _prepare_outline(self, request: OutlineRequest) -> tuple[OutlineResponse, Path]:
        use_case_dir = self.context_manager.resolve_use_case(request.useCaseId)
        override_path = Path(request.contextPath) if request.contextPath else None
        context_path = self.context_manager.resolve_context_path(use_case_dir, override_path)
        context_text = self.context_manager.load_context(context_path)
        validation = self.context_manager.validate_context(context_text)
        if not validation.isValid:
            self._raise_validation_error(context_path, validation)

        sections = self.context_manager.parse_sections(context_text)
        request.generateJsonOutput = self.context_manager.resolve_generate_json_output(
            sections,
            cli_override=True if request.generateJsonOutput else None,
        )
        references = PptxReferenceLibrary(use_case_dir / "reference_decks")
        reference_summary = references.summarize_reference_patterns()
        template_deck_path = self._resolve_template_deck(use_case_dir, sections, reference_summary)
        outline = self._draft_outline(request, sections, context_text, reference_summary, validation, context_path, template_deck_path)
        return outline, context_path

    def _build_from_content_dict(self, data: dict, output_dir: Path, base_name: str, generate_json_output: bool) -> BuildResult:
        from .presentation_builder import PresentationBuilder

        builder = PresentationBuilder()
        pptx_path = output_dir / f"{base_name}_deck.pptx"
        trace_path = output_dir / f"{base_name}_trace.json" if generate_json_output or data.get("generateJsonOutput") else None
        warnings = builder.build_from_content(data, pptx_path, trace_path)
        return BuildResult(pptxPath=str(pptx_path), tracePath=str(trace_path) if trace_path else "", warnings=warnings)

    def _draft_outline(
        self,
        request: OutlineRequest,
        sections: dict[str, str],
        context_text: str,
        reference_summary: dict,
        validation: ContextValidationResult,
        context_path: Path,
        template_deck_path: Path | None,
    ) -> OutlineResponse:
        if self.llm.enabled:
            try:
                return self._draft_outline_with_gemini(
                    request,
                    sections,
                    context_text,
                    reference_summary,
                    validation,
                    context_path,
                    template_deck_path,
                )
            except Exception:
                pass
        return self._draft_outline_offline(
            request,
            sections,
            reference_summary,
            validation,
            context_path,
            template_deck_path,
        )

    def _draft_outline_with_gemini(
        self,
        request: OutlineRequest,
        sections: dict[str, str],
        context_text: str,
        reference_summary: dict,
        validation: ContextValidationResult,
        context_path: Path,
        template_deck_path: Path | None,
    ) -> OutlineResponse:
        slide_count = self._resolve_slide_count(request, sections)
        prompt = f"""
You are creating a pharma business-development presentation content plan that will later be converted to PowerPoint.

Safety rules:
- Reference decks are style and layout references only.
- Never copy or restate factual claims, numbers, dates, patient funnel details, or recommendations from reference decks.
- Current project facts must come only from CONTEXT_TXT.
- If a needed fact is missing, write a bracketed placeholder instead of inventing it.
- Vary the slide archetypes and layout types across the deck rather than repeating one layout.
- Return valid JSON only.

Requested output schema:
{{
  "title": "string",
  "templateDeckPath": "string",
  "contextFile": "string",
  "generateJsonOutput": true,
  "referenceDecks": ["string"],
  "openQuestions": ["string"],
  "validationWarnings": ["string"],
  "slides": [
    {{
      "slideNumber": 1,
      "title": "string",
      "objective": "string",
      "bullets": ["string"],
      "visualSuggestion": "string",
      "claimSources": ["project_context:section name"],
      "styleSources": ["style_reference:pattern name"],
      "notes": ["string"],
      "status": "draft",
      "archetype": "title|agenda|executive_summary|framework|patient_funnel|insight|recommendation|appendix|section_divider",
      "layoutHint": "title|section|exec|two_column|chart|funnel|appendix"
    }}
  ]
}}

REQUEST:
{json.dumps(asdict(request), indent=2)}

SLIDE_COUNT_TARGET: {slide_count}
TEMPLATE_DECK_PATH: {str(template_deck_path) if template_deck_path else ''}
REFERENCE_STYLE_SUMMARY:
{json.dumps(reference_summary, indent=2)}

CONTEXT_TXT:
{context_text}

VALIDATION_WARNINGS:
{json.dumps(validation.warnings, indent=2)}
"""
        raw = self.llm.generate_json(prompt)
        slides = [SlideDraft(**slide) for slide in raw["slides"]]
        return OutlineResponse(
            title=raw["title"],
            slides=slides[:slide_count],
            templateDeckPath=raw.get("templateDeckPath") or str(template_deck_path or ""),
            contextFile=raw.get("contextFile") or str(context_path),
            generateJsonOutput=request.generateJsonOutput,
            referenceDecks=raw.get("referenceDecks") or reference_summary.get("referenceDecks", []),
            openQuestions=raw.get("openQuestions", []),
            validationWarnings=raw.get("validationWarnings", []) + validation.warnings,
        )

    def _draft_outline_offline(
        self,
        request: OutlineRequest,
        sections: dict[str, str],
        reference_summary: dict,
        validation: ContextValidationResult,
        context_path: Path,
        template_deck_path: Path | None,
    ) -> OutlineResponse:
        title = sections.get("project name", "Pharma Analytics Presentation").strip() or "Pharma Analytics Presentation"
        audience = sections.get("audience", request.audience)
        objective = sections.get("presentation objective", request.presentationGoal)
        business_question = sections.get("current business question", "")
        facts = self._to_bullets(sections.get("approved facts and numbers", ""))
        workflow = self._to_bullets(sections.get("patient funnel or workflow summary", ""))
        outputs = self._to_bullets(sections.get("analytical outputs available", ""))
        recommendations = self._to_bullets(sections.get("draft recommendations or hypotheses", ""))
        constraints = self._to_bullets(sections.get("constraints / caveats", ""))
        include_or_avoid = sections.get("slides to include or avoid", "")
        additional_instructions = sections.get("additional instructions", "")
        slide_count = self._resolve_slide_count(request, sections)
        style_patterns = reference_summary.get("stylePatterns", [])
        style_note = style_patterns[0]["toneNotes"] if style_patterns else "Use concise executive storytelling."

        slides = [
            SlideDraft(1, title, "Introduce the project, audience, and deck purpose.", [
                f"Audience: {audience or '[Add target audience]'}",
                f"Objective: {objective or '[Add presentation objective]'}",
                f"Use case: {sections.get('use-case type', request.useCaseId)}",
            ], "Client-style title page with subtitle and minimal setup text.", ["project_context:project name", "project_context:audience", "project_context:presentation objective"], ["style_reference:narrative"], [style_note], "draft", "title", "title"),
            SlideDraft(2, "Agenda", "Set the story flow for the discussion.", [
                "Business context and objective",
                "Current-state findings and workflow implications",
                "Recommendations, risks, and next steps",
            ], "Agenda with 3 clear sections using the client master.", [], ["style_reference:agenda"], ["Structure only. No factual claims should be introduced here."], "draft", "agenda", "section"),
            SlideDraft(3, "Executive Summary", "Summarize the most decision-relevant takeaways.", self._ensure_n([
                facts[0] if facts else "[Add the top approved metric or fact]",
                recommendations[0] if recommendations else "[Add the top recommendation or hypothesis]",
                constraints[0] if constraints else "[Add the primary caveat or decision note]",
            ], 3), "Three takeaway cards or a summary panel.", self._non_empty_sources("approved facts and numbers", facts[:1]) + self._non_empty_sources("draft recommendations or hypotheses", recommendations[:1]) + self._non_empty_sources("constraints / caveats", constraints[:1]), ["style_reference:executive_summary"], ["Use only approved current-project claims."], "draft", "executive_summary", "exec"),
            SlideDraft(4, "Current Business Question", "Define the business problem and why it matters now.", self._ensure_n([
                business_question or "[Describe the business question for this project]",
                outputs[0] if outputs else "[List the primary analytical output available]",
                include_or_avoid or "[List any mandatory or prohibited slide topics]",
            ], 3), "Two-column framing slide with context on the left and implications on the right.", ["project_context:current business question", "project_context:analytical outputs available"], ["style_reference:narrative"], ["Frame the challenge using only current-project context."], "draft", "framework", "two_column"),
            SlideDraft(5, "Project Workflow or Patient Funnel", "Show the key journey or process relevant to this project.", self._ensure_n(workflow[:5] if workflow else ["[Add funnel stage 1]", "[Add funnel stage 2]", "[Add funnel stage 3]"], 3), "A stepped funnel or journey visual using client formatting.", self._non_empty_sources("patient funnel or workflow summary", workflow[:5]), ["style_reference:narrative"], ["Do not infer funnel steps from historical decks."], "draft", "patient_funnel", "funnel"),
            SlideDraft(6, "Analytical Insights", "Translate the current analysis into concrete findings.", self._ensure_n(facts[:2] + outputs[:2], 4), "Insight slide with space for chart, table, or network visual.", self._non_empty_sources("approved facts and numbers", facts[:2]) + self._non_empty_sources("analytical outputs available", outputs[:2]), ["style_reference:narrative"], ["If evidence is missing, leave explicit placeholders."], "draft", "insight", "chart"),
            SlideDraft(7, "Recommendations and Implications", "Recommend practical next steps based on the current project context.", self._ensure_n(recommendations[:3] if recommendations else [
                "[Add recommendation 1 tied to the current evidence]",
                "[Add recommendation 2 tied to the current evidence]",
                "[Add recommendation 3 tied to the current evidence]",
            ], 3), "Priority-based recommendation layout with decision callouts.", self._non_empty_sources("draft recommendations or hypotheses", recommendations[:3]), ["style_reference:recommendation"], ["Recommendation language can be directional, but not fabricated."], "draft", "recommendation", "two_column"),
            SlideDraft(8, "Appendix and Open Questions", "Capture unresolved questions, caveats, and add-on supporting content.", self._ensure_n(validation.warnings[:2] + constraints[:2] + ([additional_instructions] if additional_instructions else []), 4), "Appendix slide with review notes and support requests.", self._non_empty_sources("constraints / caveats", constraints[:2]), ["style_reference:narrative"], ["Use this slide to flag missing inputs before finalization."], "draft", "appendix", "appendix"),
        ]

        while len(slides) < slide_count:
            index = len(slides) + 1
            source_bullets = facts + outputs + recommendations + workflow
            chunk = source_bullets[(index - 1) % len(source_bullets):(index - 1) % len(source_bullets) + 3] if source_bullets else []
            slides.append(SlideDraft(index, f"Supporting Insight {index - 6}", "Add an additional fact pattern or implication that supports the story.", self._ensure_n(chunk or ["[Add another approved insight from context.txt]"], 3), "Additional insight slide with a chart or table placeholder.", self._non_empty_sources("approved facts and numbers", chunk), ["style_reference:narrative"], ["Extra support slide generated because Slide Count Target exceeds the base storyline."], "draft", "insight", "chart"))

        return OutlineResponse(
            title=title,
            slides=slides[:slide_count],
            templateDeckPath=str(template_deck_path or ""),
            contextFile=str(context_path),
            generateJsonOutput=request.generateJsonOutput,
            referenceDecks=reference_summary.get("referenceDecks", []),
            openQuestions=self._build_open_questions(validation, facts, workflow, recommendations),
            validationWarnings=validation.warnings + ([] if self.llm.enabled else ["Gemini API key not found. Used offline drafting mode."]),
        )

    def _outline_to_markdown(self, outline: OutlineResponse) -> str:
        lines = [
            f"# {outline.title}",
            "",
            f"- Context file: `{outline.contextFile}`",
            f"- Template deck: `{outline.templateDeckPath or 'auto-select first reference deck'}`",
            f"- Reference decks analyzed: `{len(outline.referenceDecks)}`",
            f"- Generate JSON output: `{outline.generateJsonOutput}`",
            "",
        ]
        if outline.validationWarnings:
            lines.append("## Validation Warnings")
            lines.extend(f"- {warning}" for warning in outline.validationWarnings)
            lines.append("")
        if outline.openQuestions:
            lines.append("## Open Questions")
            lines.extend(f"- {question}" for question in outline.openQuestions)
            lines.append("")
        lines.append("## Slides")
        for slide in outline.slides:
            lines.append(f"### {slide.slideNumber}. {slide.title}")
            lines.append(f"- Objective: {slide.objective}")
            lines.append(f"- Archetype: {slide.archetype}")
            lines.append(f"- Layout hint: {slide.layoutHint}")
            lines.append(f"- Visual: {slide.visualSuggestion}")
            lines.extend(f"- {bullet}" for bullet in slide.bullets)
            if slide.claimSources:
                lines.append(f"- Claim sources: {', '.join(slide.claimSources)}")
            if slide.styleSources:
                lines.append(f"- Style sources: {', '.join(slide.styleSources)}")
            lines.append("")
        return "\n".join(lines)

    def _resolve_template_deck(self, use_case_dir: Path, sections: dict[str, str], reference_summary: dict) -> Path | None:
        template_value = sections.get("template deck", "").strip()
        reference_dir = use_case_dir / "reference_decks"
        if template_value:
            explicit = Path(template_value)
            for candidate in (explicit, reference_dir / template_value, use_case_dir / template_value):
                if candidate.exists() and candidate.suffix.lower() == ".pptx":
                    return candidate.resolve()
        reference_decks = reference_summary.get("referenceDecks", [])
        if reference_decks:
            first = reference_dir / reference_decks[0]
            if first.exists():
                return first.resolve()
        decks = sorted(reference_dir.glob("*.pptx"))
        return decks[0].resolve() if decks else None

    def _resolve_slide_count(self, request: OutlineRequest, sections: dict[str, str]) -> int:
        if request.slideCountTarget:
            return request.slideCountTarget
        raw = sections.get("slide count target", "8").strip()
        try:
            return max(1, int(raw))
        except ValueError:
            return 8

    def _build_open_questions(self, validation: ContextValidationResult, facts: list[str], workflow: list[str], recommendations: list[str]) -> list[str]:
        questions = []
        if not facts:
            questions.append("What approved metrics or quantified outcomes should appear in the deck?")
        if not workflow:
            questions.append("What is the current-project patient funnel or workflow that should be visualized?")
        if not recommendations:
            questions.append("Which recommendations are approved vs still hypothesis-stage?")
        questions.extend(f"Can you strengthen the section: {item}?" for item in validation.weakSections)
        return questions

    def _raise_validation_error(self, context_path: Path, validation: ContextValidationResult) -> None:
        lines = ["context.txt is incomplete."]
        if validation.missingSections:
            lines.append("Missing sections: " + ", ".join(validation.missingSections))
        if validation.weakSections:
            lines.append("Weak sections: " + ", ".join(validation.weakSections))
        lines.append(f"Update {context_path} with current project facts and retry.")
        raise ValueError(" ".join(lines))

    @staticmethod
    def _to_bullets(text: str) -> list[str]:
        items = []
        for raw_line in text.splitlines():
            line = raw_line.strip().lstrip("-").strip()
            if line and not line.startswith("#"):
                items.append(line)
        return items

    @staticmethod
    def _non_empty_sources(section_name: str, items: list[str]) -> list[str]:
        return [f"project_context:{section_name}" for item in items if item and not item.startswith("[")]

    @staticmethod
    def _ensure_n(items: list[str], count: int) -> list[str]:
        values = [item for item in items if item]
        while len(values) < count:
            values.append("[Add supporting detail from context.txt]")
        return values[:count]

    @staticmethod
    def _timestamp() -> str:
        from datetime import datetime
        return datetime.now().strftime("%Y%m%d_%H%M%S")
