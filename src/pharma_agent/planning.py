from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

from .context_manager import ContextManager
from .excel_context import ExcelContextLibrary
from .llm import LLMClient
from .models import BuildResult, ContextValidationResult, OutlineRequest, OutlineResponse, SlideDraft, dataclass_to_dict
from .pptx_reference import PptxReferenceLibrary


@dataclass
class PreparedInputs:
    """Keep planning inputs explicit so the orchestration stays easy to review."""

    request: OutlineRequest
    context_path: Path
    context_text: str
    runtime_context_text: str
    sections: dict[str, str]
    validation: ContextValidationResult
    reference_summary: dict
    excel_summary: dict
    template_deck_path: Path | None
    existing_deck_path: Path | None
    existing_deck_summary: dict
    existing_deck_excerpt: str


class PlanningService:
    """Coordinate context loading, outline drafting, and PPT generation."""

    def __init__(self, root: Path) -> None:
        self.root = root
        self.context_manager = ContextManager(root)
        self.excel_context = ExcelContextLibrary()
        self.llm = LLMClient()

    def list_use_cases(self) -> list[dict]:
        return [asdict(item) for item in self.context_manager.list_use_cases()]

    def generate_content_plan(self, request: OutlineRequest, output_dir: Path) -> tuple[Path, Path]:
        outline = self._prepare_outline(request)
        timestamp = self._timestamp()
        safe_name = request.useCaseId.lower()
        json_path = output_dir / f"{timestamp}_{safe_name}_content_v1.json"
        md_path = output_dir / f"{timestamp}_{safe_name}_content_v1.md"
        json_path.write_text(json.dumps(dataclass_to_dict(outline), indent=2), encoding="utf-8")
        md_path.write_text(self._outline_to_markdown(outline), encoding="utf-8")
        return json_path, md_path

    def build_presentation(
        self,
        content_path: Path,
        output_dir: Path,
        generate_json_output: bool = False,
        output_pptx_path: str = "",
        existing_pptx_path: str = "",
    ) -> BuildResult:
        data = json.loads(content_path.read_text(encoding="utf-8"))
        base_name = content_path.stem.replace("_content_v1", "")
        return self._build_from_content_dict(data, output_dir, base_name, generate_json_output, output_pptx_path, existing_pptx_path)

    def run_pipeline(self, request: OutlineRequest, output_dir: Path) -> tuple[BuildResult, str, str]:
        outline = self._prepare_outline(request)
        base_name = f"{self._timestamp()}_{request.useCaseId.lower()}"
        result = self._build_from_content_dict(
            dataclass_to_dict(outline),
            output_dir,
            base_name,
            request.generateJsonOutput,
            request.outputPptxPath,
            request.existingDeckPath,
        )
        return result, outline.generationMode, outline.generationMessage

    def _prepare_outline(self, request: OutlineRequest) -> OutlineResponse:
        prepared = self._load_inputs(request)
        return self._draft_outline(prepared)

    def _load_inputs(self, request: OutlineRequest) -> PreparedInputs:
        use_case_dir = self.context_manager.resolve_use_case(request.useCaseId)
        context_path = self.context_manager.resolve_context_path(
            use_case_dir,
            Path(request.contextPath) if request.contextPath else None,
        )
        context_text = self.context_manager.load_context(context_path)
        validation = self.context_manager.validate_context(context_text)
        if not validation.isValid:
            self._raise_validation_error(context_path, validation)

        # Parse the human-maintained context.txt into named sections once so every downstream step uses the same view of the project.
        sections = self.context_manager.parse_sections(context_text)
        request.generateJsonOutput = self.context_manager.resolve_generate_json_output(
            sections,
            cli_override=True if request.generateJsonOutput else None,
        )

        reference_library = PptxReferenceLibrary(use_case_dir / "reference_decks")
        reference_summary = reference_library.summarize_reference_patterns()
        excel_summary = self.excel_context.summarize_files(self.context_manager.list_excel_context_files(use_case_dir))
        # Excel-derived findings are appended at runtime instead of being copied back into context.txt by hand.
        runtime_context_text = self._compose_runtime_context(context_text, excel_summary)
        existing_deck_path = self._resolve_existing_deck_path(request.existingDeckPath)
        existing_deck_summary, existing_deck_excerpt = self._load_existing_deck_context(existing_deck_path)
        template_deck_path = self._resolve_template_deck(use_case_dir, sections, reference_summary, existing_deck_path)

        return PreparedInputs(
            request=request,
            context_path=context_path,
            context_text=context_text,
            runtime_context_text=runtime_context_text,
            sections=sections,
            validation=validation,
            reference_summary=reference_summary,
            excel_summary=excel_summary,
            template_deck_path=template_deck_path,
            existing_deck_path=existing_deck_path,
            existing_deck_summary=existing_deck_summary,
            existing_deck_excerpt=existing_deck_excerpt,
        )

    def _compose_runtime_context(self, context_text: str, excel_summary: dict) -> str:
        runtime_block = self.excel_context.build_runtime_context_block(excel_summary)
        return context_text if not runtime_block else f"{context_text.strip()}\n\n{runtime_block}\n"

    def _resolve_existing_deck_path(self, existing_deck_path: str) -> Path | None:
        if not existing_deck_path:
            return None
        path = Path(existing_deck_path)
        return path.resolve() if path.exists() and path.suffix.lower() == ".pptx" else None

    def _load_existing_deck_context(self, existing_deck_path: Path | None) -> tuple[dict, str]:
        if existing_deck_path is None:
            return {}, ""

        library = PptxReferenceLibrary(existing_deck_path.parent)
        chunks = library.extract_pptx_text(existing_deck_path)
        summary = {
            "existingDeckName": existing_deck_path.name,
            "slideCount": len(chunks),
            "stylePatterns": [pattern.__dict__ for pattern in library.derive_style_patterns(chunks)],
        }

        # Keep only a short excerpt in the prompt so the LLM sees current deck flow
        # without us flooding the context window with full slide XML text.
        lines = []
        for chunk in chunks[:8]:
            excerpt = chunk.text[:240].replace("\n", " ")
            lines.append(f"Slide {chunk.slideNumber}: {excerpt}")
        return summary, "\n".join(lines)

    def _build_from_content_dict(
        self,
        data: dict,
        output_dir: Path,
        base_name: str,
        generate_json_output: bool,
        output_pptx_path: str = "",
        existing_pptx_path: str = "",
    ) -> BuildResult:
        from .presentation_builder import PresentationBuilder

        builder = PresentationBuilder()
        existing_path = self._resolve_existing_deck_path(existing_pptx_path)
        pptx_path = Path(output_pptx_path).resolve() if output_pptx_path else (existing_path or output_dir / f"{base_name}_deck.pptx")
        trace_path = output_dir / f"{base_name}_trace.json" if generate_json_output or data.get("generateJsonOutput") else None
        warnings = builder.build_from_content(
            data,
            pptx_path,
            trace_path,
            existing_presentation_path=existing_path if existing_path and existing_path.exists() else (pptx_path if pptx_path.exists() else None),
        )
        return BuildResult(pptxPath=str(pptx_path), tracePath=str(trace_path) if trace_path else "", warnings=warnings)

    def _draft_outline(self, prepared: PreparedInputs) -> OutlineResponse:
        if not self.llm.enabled:
            return self._draft_outline_offline(prepared, "No supported LLM API key was configured.")

        try:
            outline = self._draft_outline_with_llm(prepared)
            outline.generationMode = self.llm.provider
            outline.generationMessage = f"{self.llm.provider} generation succeeded using model {self.llm.model}."
            return outline
        except Exception as exc:
            return self._draft_outline_offline(prepared, str(exc))

    def _draft_outline_with_llm(self, prepared: PreparedInputs) -> OutlineResponse:
        prompt = self._build_llm_prompt(prepared)
        raw = self.llm.generate_json(prompt)
        slide_count = self._resolve_slide_count(prepared.request, prepared.sections, prepared.existing_deck_summary)
        slides = [SlideDraft(**slide) for slide in raw["slides"]][:slide_count]
        return OutlineResponse(
            title=raw["title"],
            slides=slides,
            templateDeckPath=raw.get("templateDeckPath") or str(prepared.template_deck_path or ""),
            contextFile=raw.get("contextFile") or str(prepared.context_path),
            generateJsonOutput=prepared.request.generateJsonOutput,
            generationMode=self.llm.provider,
            generationMessage=raw.get("generationMessage", f"{self.llm.provider} generation succeeded using model {self.llm.model}."),
            referenceDecks=raw.get("referenceDecks") or prepared.reference_summary.get("referenceDecks", []),
            excelContextFiles=raw.get("excelContextFiles") or [item["fileName"] for item in prepared.excel_summary.get("files", [])],
            openQuestions=raw.get("openQuestions", []),
            validationWarnings=raw.get("validationWarnings", []) + prepared.validation.warnings + prepared.excel_summary.get("warnings", []),
        )

    def _build_llm_prompt(self, prepared: PreparedInputs) -> str:
        slide_count = self._resolve_slide_count(prepared.request, prepared.sections, prepared.existing_deck_summary)
        refine_mode = "true" if prepared.existing_deck_path else "false"
        return f"""
You are creating a pharma business-development presentation content plan that will later be converted to PowerPoint.

Safety rules:
- Reference decks are style and layout references only.
- Never copy or restate factual claims, numbers, dates, patient funnel details, or recommendations from reference decks.
- Current project facts must come only from CONTEXT_TXT and optional EXCEL_CONTEXT.
- Excel context is refreshed dynamically on every run, so use the runtime Excel findings below when they strengthen the current project story.
- Every slide title must be a talking header written as a clear business sentence or clause, typically 8-20 words, not a short label like "Agenda", "Insights", or "Recommendations" unless the slide is a title or section divider.
- Slide titles should explain what the slide is showing, what the result means, or why the point matters in simple client-ready language.
- Use the reference decks to their full potential for story arc, section sequencing, layout variety, density, figure placement, and executive tone.
- The deck should feel client-ready: each slide must have a deliberate role in the story and should not read like scaffolding or a draft outline.
- If a needed fact is missing, use a precise instructional placeholder that says exactly what content is missing and what should replace it.
- Never use vague filler such as "placeholder", "add chart here", "TBD", or generic lorem-style language.
- When a visual is needed, describe the intended visual and the business point it should prove.
- If REFINE_EXISTING_DECK is true, treat the existing deck as the starting point and improve its slide flow, wording, and layout intent rather than inventing a completely unrelated structure.
- Preserve the useful storyline and layout spirit of the existing deck when refining it, but update weak titles, weak copy, and unsupported claims.
- Return valid JSON only.

Requested output schema:
{{
  "title": "string",
  "templateDeckPath": "string",
  "contextFile": "string",
  "generateJsonOutput": true,
  "generationMode": "{self.llm.provider}",
  "generationMessage": "{self.llm.provider} generation succeeded.",
  "referenceDecks": ["string"],
  "excelContextFiles": ["string"],
  "openQuestions": ["string"],
  "validationWarnings": ["string"],
  "slides": [
    {{
      "slideNumber": 1,
      "title": "talking-header sentence that explains the slide point",
      "objective": "string",
      "bullets": ["string"],
      "visualSuggestion": "string",
      "claimSources": ["project_context:section name|excel_context:file name"],
      "styleSources": ["style_reference:pattern name|existing_deck:slide number"],
      "notes": ["string"],
      "status": "draft",
      "archetype": "title|agenda|executive_summary|framework|patient_funnel|insight|recommendation|appendix|section_divider",
      "layoutHint": "title|section|exec|two_column|chart|funnel|appendix"
    }}
  ]
}}

REQUEST:
{json.dumps(asdict(prepared.request), indent=2)}

REFINE_EXISTING_DECK: {refine_mode}
SLIDE_COUNT_TARGET: {slide_count}
TEMPLATE_DECK_PATH: {str(prepared.template_deck_path) if prepared.template_deck_path else ''}
REFERENCE_STYLE_SUMMARY:
{json.dumps(prepared.reference_summary, indent=2)}

EXISTING_DECK_SUMMARY:
{json.dumps(prepared.existing_deck_summary, indent=2)}

EXISTING_DECK_EXCERPT:
{prepared.existing_deck_excerpt}

CONTEXT_TXT:
{prepared.context_text}

RUNTIME_CONTEXT_TXT:
{prepared.runtime_context_text}

EXCEL_CONTEXT:
{json.dumps(prepared.excel_summary, indent=2)}

VALIDATION_WARNINGS:
{json.dumps(prepared.validation.warnings + prepared.excel_summary.get('warnings', []), indent=2)}
"""

    def _draft_outline_offline(self, prepared: PreparedInputs, failure_message: str) -> OutlineResponse:
        sections = prepared.sections
        request = prepared.request
        title = sections.get("project name", "Pharma Analytics Presentation").strip() or "Pharma Analytics Presentation"
        project_short_name = sections.get("client / brand / indication", request.useCaseId).strip() or request.useCaseId.replace("_", " ").title()
        facts = self._to_bullets(sections.get("approved facts and numbers", ""))
        workflow = self._to_bullets(sections.get("patient funnel or workflow summary", ""))
        outputs = self._to_bullets(sections.get("analytical outputs available", ""))
        recommendations = self._to_bullets(sections.get("draft recommendations or hypotheses", ""))
        constraints = self._to_bullets(sections.get("constraints / caveats", ""))
        include_or_avoid = sections.get("slides to include or avoid", "")
        additional_instructions = sections.get("additional instructions", "")
        audience = sections.get("audience", request.audience)
        objective = sections.get("presentation objective", request.presentationGoal)
        business_question = sections.get("current business question", "")
        style_note = (prepared.reference_summary.get("stylePatterns", [{}])[0] or {}).get("toneNotes", "Use concise executive storytelling.")
        excel_numeric, excel_text, excel_files = self._extract_excel_bullets(prepared.excel_summary)
        excel_claim_sources = [f"excel_context:{name}" for name in excel_files]
        insight_bullets = self._ensure_n(facts[:2] + outputs[:2] + excel_numeric[:2], 4)
        slide_count = self._resolve_slide_count(request, sections, prepared.existing_deck_summary)

        # When an existing deck is provided, the offline mode still uses it as the
        # main shell/template and tries to preserve a familiar story arc.
        slides = [
            SlideDraft(1, title, "Introduce the project, audience, and deck purpose.", [f"Audience: {audience or '[Add target audience]'}", f"Objective: {objective or '[Add presentation objective]'}", f"Use case: {sections.get('use-case type', request.useCaseId)}"], "Client-style title page with subtitle and minimal setup text.", ["project_context:project name", "project_context:audience", "project_context:presentation objective"], ["style_reference:narrative"], [style_note], "draft", "title", "title"),
            SlideDraft(2, "This deck moves from referral context to quantified insights and action priorities", "Set the story flow for the discussion.", ["Business context and objective", "Current-state findings and workflow implications", "Recommendations, risks, and next steps"], "Agenda with 3 clear sections using the client master.", [], ["style_reference:agenda"], ["Structure only. No factual claims should be introduced here."], "draft", "agenda", "section"),
            SlideDraft(3, "Cross-state referrals show meaningful scale and clear corridors that warrant focused action", "Summarize the most decision-relevant takeaways.", self._ensure_n([facts[0] if facts else (excel_numeric[0] if excel_numeric else "[Insert the top approved metric and explain why it matters]"), recommendations[0] if recommendations else "[Insert the top approved recommendation tied to the current evidence]", constraints[0] if constraints else "[Insert the main caveat or decision note that leaders should keep in mind]"], 3), "Three takeaway cards or a summary panel.", self._non_empty_sources("approved facts and numbers", facts[:1]) + (excel_claim_sources[:1] if not facts and excel_claim_sources else []) + self._non_empty_sources("draft recommendations or hypotheses", recommendations[:1]) + self._non_empty_sources("constraints / caveats", constraints[:1]), ["style_reference:executive_summary"], ["Use only approved current-project claims."], "draft", "executive_summary", "exec"),
            SlideDraft(4, f"{project_short_name} needs a repeatable referral story grounded in current approved evidence", "Define the business problem and why it matters now.", self._ensure_n([business_question or "[Describe the business question for this project and why it matters now]", outputs[0] if outputs else (excel_text[0] if excel_text else "[Describe the primary analytical output available for this slide]"), include_or_avoid or "[List any mandatory slide topics or prohibited claims for this project]"], 3), "Two-column framing slide with context on the left and implications on the right.", ["project_context:current business question", "project_context:analytical outputs available"] + (excel_claim_sources[:1] if not outputs and excel_claim_sources else []), ["style_reference:narrative"], ["Frame the challenge using only current-project context."], "draft", "framework", "two_column"),
            SlideDraft(5, "The workflow shows how referral-network evidence is built and interpreted step by step", "Show the key journey or process relevant to this project.", self._ensure_n(workflow[:5] if workflow else ["[Describe the first workflow step and why it matters]", "[Describe the next workflow step and what is measured]", "[Describe the final workflow step and the decision output]"], 3), "A stepped funnel or journey visual using client formatting.", self._non_empty_sources("patient funnel or workflow summary", workflow[:5]), ["style_reference:narrative"], ["Do not infer funnel steps from historical decks."], "draft", "patient_funnel", "funnel"),
            SlideDraft(6, "The analysis highlights where cross-state referral activity is concentrated and why it matters", "Translate the current analysis into concrete findings.", insight_bullets, "Ranked corridor chart, state-pair table, or referral network visual that proves the main concentration pattern.", self._non_empty_sources("approved facts and numbers", facts[:2]) + self._non_empty_sources("analytical outputs available", outputs[:2]) + (excel_claim_sources if excel_numeric else []), ["style_reference:narrative"], ["Excel context can supplement approved metrics when files are present in excel_context/."], "draft", "insight", "chart"),
            SlideDraft(7, "The strongest referral corridors should anchor the first wave of business follow-up", "Recommend practical next steps based on the current project context.", self._ensure_n(recommendations[:3] if recommendations else ["[Insert recommendation 1 tied directly to an approved finding]", "[Insert recommendation 2 tied directly to an approved finding]", "[Insert recommendation 3 tied directly to an approved finding]"], 3), "Priority-based recommendation layout with decision callouts.", self._non_empty_sources("draft recommendations or hypotheses", recommendations[:3]), ["style_reference:recommendation"], ["Recommendation language can be directional, but not fabricated."], "draft", "recommendation", "two_column"),
            SlideDraft(8, "The remaining caveats and follow-up asks are clear before final client circulation", "Capture unresolved questions, caveats, and add-on supporting content.", self._ensure_n(prepared.validation.warnings[:2] + prepared.excel_summary.get("warnings", [])[:1] + constraints[:2] + ([additional_instructions] if additional_instructions else []), 4), "Appendix slide with review notes, caveats, and specific follow-up requests.", self._non_empty_sources("constraints / caveats", constraints[:2]), ["style_reference:narrative"], ["Use this slide to flag missing inputs before finalization."], "draft", "appendix", "appendix"),
        ]

        while len(slides) < slide_count:
            source_bullets = facts + outputs + recommendations + workflow + excel_numeric + excel_text
            index = len(slides) + 1
            chunk = source_bullets[(index - 1) % max(1, len(source_bullets)):] if source_bullets else []
            slides.append(
                SlideDraft(
                    index,
                    f"Supporting evidence slide {index - 6} extends the core story with one more approved insight",
                    "Add an additional fact pattern or implication that supports the story.",
                    self._ensure_n(chunk[:3] or ["[Insert one additional approved insight from context.txt or excel_context and explain why it matters]"], 3),
                    "Supporting chart, ranked table, or comparison visual that reinforces the main story with one additional proof point.",
                    self._non_empty_sources("approved facts and numbers", chunk[:3]) + (excel_claim_sources if any(item in excel_numeric + excel_text for item in chunk[:3]) else []),
                    ["style_reference:narrative"],
                    ["Extra support slide generated because Slide Count Target exceeds the base storyline."],
                    "draft",
                    "insight",
                    "chart",
                )
            )

        warnings = prepared.validation.warnings + prepared.excel_summary.get("warnings", [])
        if not self.llm.enabled:
            warnings.append("No supported LLM API key found. Used offline drafting mode.")
        if prepared.existing_deck_path is not None:
            warnings.append(f"Existing deck refinement used template shell: {prepared.existing_deck_path.name}")

        return OutlineResponse(
            title=title,
            slides=slides[:slide_count],
            templateDeckPath=str(prepared.template_deck_path or ""),
            contextFile=str(prepared.context_path),
            generateJsonOutput=request.generateJsonOutput,
            generationMode="offline",
            generationMessage=f"Offline fallback used because {self.llm.provider or 'configured LLM'} was unavailable: {failure_message}",
            referenceDecks=prepared.reference_summary.get("referenceDecks", []),
            excelContextFiles=excel_files,
            openQuestions=self._build_open_questions(prepared.validation, facts + excel_numeric, workflow, recommendations),
            validationWarnings=warnings,
        )

    def _outline_to_markdown(self, outline: OutlineResponse) -> str:
        lines = [
            f"# {outline.title}",
            "",
            f"- Context file: `{outline.contextFile}`",
            f"- Template deck: `{outline.templateDeckPath or 'auto-select first reference deck'}`",
            f"- Reference decks analyzed: `{len(outline.referenceDecks)}`",
            f"- Excel context files: `{len(outline.excelContextFiles)}`",
            f"- Generate JSON output: `{outline.generateJsonOutput}`",
            f"- Generation mode: `{outline.generationMode}`",
            f"- Generation message: `{outline.generationMessage}`",
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

    def _resolve_template_deck(
        self,
        use_case_dir: Path,
        sections: dict[str, str],
        reference_summary: dict,
        existing_deck_path: Path | None,
    ) -> Path | None:
        # Existing deck wins because refinement should preserve its layout system.
        if existing_deck_path is not None:
            return existing_deck_path

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

    def _resolve_slide_count(self, request: OutlineRequest, sections: dict[str, str], existing_deck_summary: dict) -> int:
        if request.slideCountTarget:
            return request.slideCountTarget
        if existing_deck_summary.get("slideCount"):
            return max(1, int(existing_deck_summary["slideCount"]))
        try:
            return max(1, int(sections.get("slide count target", "8").strip()))
        except ValueError:
            return 8

    def _extract_excel_bullets(self, excel_summary: dict) -> tuple[list[str], list[str], list[str]]:
        numeric: list[str] = []
        text: list[str] = []
        files: list[str] = []
        for file_summary in excel_summary.get("files", []):
            files.append(file_summary.get("fileName", ""))
            numeric.extend(file_summary.get("numericSamples", [])[:4])
            text.extend(file_summary.get("textSamples", [])[:4])
        return numeric, text, [name for name in files if name]

    def _build_open_questions(self, validation: ContextValidationResult, facts: list[str], workflow: list[str], recommendations: list[str]) -> list[str]:
        questions: list[str] = []
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
        bullets: list[str] = []
        for raw_line in text.splitlines():
            line = raw_line.strip().lstrip("-").strip()
            if line and not line.startswith("#"):
                bullets.append(line)
        return bullets

    @staticmethod
    def _non_empty_sources(section_name: str, items: list[str]) -> list[str]:
        return [f"project_context:{section_name}" for item in items if item and not item.startswith("[")]

    @staticmethod
    def _ensure_n(items: list[str], count: int) -> list[str]:
        values = [item for item in items if item]
        while len(values) < count:
            values.append("[Insert a specific approved fact, result, or implication from context.txt or excel_context here]")
        return values[:count]

    @staticmethod
    def _timestamp() -> str:
        from datetime import datetime

        return datetime.now().strftime("%Y%m%d_%H%M%S")

