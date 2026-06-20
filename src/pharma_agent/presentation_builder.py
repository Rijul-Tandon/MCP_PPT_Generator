from __future__ import annotations

import json
from pathlib import Path


class PresentationBuilder:
    def __init__(self) -> None:
        try:
            from pptx import Presentation  # type: ignore
            from pptx.dml.color import RGBColor  # type: ignore
            from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER  # type: ignore
            from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # type: ignore
            from pptx.util import Inches, Pt  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "python-pptx is required for the build command. Install it with `pip install -r requirements.txt`."
            ) from exc

        self.Presentation = Presentation
        self.RGBColor = RGBColor
        self.MSO_AUTO_SHAPE_TYPE = MSO_AUTO_SHAPE_TYPE
        self.PP_PLACEHOLDER = PP_PLACEHOLDER
        self.MSO_ANCHOR = MSO_ANCHOR
        self.PP_ALIGN = PP_ALIGN
        self.Inches = Inches
        self.Pt = Pt

    def build_from_content(self, content: dict, pptx_path: Path, trace_path: Path | None = None) -> list[str]:
        template_path = Path(content.get("templateDeckPath", "")) if content.get("templateDeckPath") else None
        warnings: list[str] = []
        presentation = self._load_presentation(template_path, warnings)
        trace = {
            "title": content.get("title", "Presentation"),
            "templateDeckPath": str(template_path) if template_path else "",
            "slides": [],
        }

        self._reset_to_template_shell(presentation)
        self._apply_metadata(presentation, content)

        for slide_data in content.get("slides", []):
            slide = self._create_slide(presentation, slide_data)
            trace["slides"].append(
                {
                    "slideNumber": slide_data.get("slideNumber"),
                    "title": slide_data.get("title"),
                    "archetype": slide_data.get("archetype"),
                    "layoutHint": slide_data.get("layoutHint"),
                    "claimSources": slide_data.get("claimSources", []),
                    "styleSources": slide_data.get("styleSources", []),
                    "notes": slide_data.get("notes", []),
                }
            )
            if not slide_data.get("claimSources") and slide_data.get("archetype") not in {"agenda", "section_divider"}:
                warnings.append(f"Slide {slide_data.get('slideNumber')} has no claim-bearing sources.")
            self._attach_notes(slide, slide_data)

        pptx_path.parent.mkdir(parents=True, exist_ok=True)
        presentation.save(str(pptx_path))
        if trace_path is not None:
            trace_path.write_text(json.dumps(trace, indent=2), encoding="utf-8")
        return warnings

    def _load_presentation(self, template_path: Path | None, warnings: list[str]):
        if template_path and template_path.exists():
            return self.Presentation(str(template_path))
        warnings.append("Template deck could not be resolved. Falling back to a blank presentation shell.")
        presentation = self.Presentation()
        presentation.slide_width = self.Inches(13.333)
        presentation.slide_height = self.Inches(7.5)
        return presentation

    def _reset_to_template_shell(self, presentation) -> None:
        slide_id_list = list(presentation.slides._sldIdLst)
        for slide_id in slide_id_list:
            rel_id = slide_id.rId
            presentation.part.drop_rel(rel_id)
            presentation.slides._sldIdLst.remove(slide_id)

    def _apply_metadata(self, presentation, content: dict) -> None:
        presentation.core_properties.author = "Pharma Presentation Agent"
        presentation.core_properties.title = content.get("title", "Pharma Presentation Agent Output")
        presentation.core_properties.subject = "Business-development presentation draft"

    def _create_slide(self, presentation, slide_data: dict):
        layout = self._choose_layout(presentation, slide_data)
        slide = presentation.slides.add_slide(layout)
        self._clear_text(slide)
        archetype = slide_data.get("archetype", "insight")
        if archetype == "title":
            self._render_title_slide(slide, slide_data)
        elif archetype == "agenda":
            self._render_agenda_slide(slide, slide_data)
        elif archetype == "executive_summary":
            self._render_exec_summary_slide(slide, slide_data)
        elif archetype == "framework":
            self._render_framework_slide(slide, slide_data)
        elif archetype == "patient_funnel":
            self._render_funnel_slide(slide, slide_data)
        elif archetype == "recommendation":
            self._render_recommendation_slide(slide, slide_data)
        elif archetype == "appendix":
            self._render_appendix_slide(slide, slide_data)
        else:
            self._render_insight_slide(slide, slide_data)
        self._add_footer(slide, slide_data)
        return slide

    def _choose_layout(self, presentation, slide_data: dict):
        archetype = slide_data.get("archetype", "insight")
        preferred_names = {
            "title": ["Title Page 1", "Title Slide", "Title Page 2"],
            "agenda": ["Section Title Page", "2_Standard 1-Column Text", "Standard 1-Column Text"],
            "executive_summary": ["1_Exec Sum", "Exec Sum", "Exec Sum2", "2_Exec Sum"],
            "framework": ["Standard 2-Column Text", "2_Standard 1-Column Text", "Standard"],
            "patient_funnel": ["Patient_Flow", "Advanced Chart Full Width", "Standard 2-Column Text"],
            "insight": ["Advanced Chart Full Width", "Advanced Chart 2/3", "Standard 2-Column Text"],
            "recommendation": ["Standard 2-Column Text", "2_Standard 1-Column Text", "Exec Sum"],
            "appendix": ["2_Header Only", "Standard 1-Column Text", "Blank slide"],
        }
        names = preferred_names.get(archetype, ["Standard 2-Column Text", "Blank slide"])
        for name in names:
            for layout in presentation.slide_layouts:
                if layout.name.strip().lower() == name.strip().lower():
                    return layout
        return presentation.slide_layouts[0 if len(presentation.slide_layouts) else 0]

    def _clear_text(self, slide) -> None:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                shape.text_frame.clear()

    def _render_title_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Untitled"), top=0.9, size=28)
        self._add_textbox(slide, 0.95, 1.9, 10.5, 0.8, slide_data.get("objective", ""), 14, italic=True)
        for index, bullet in enumerate(slide_data.get("bullets", [])[:3]):
            self._add_textbox(slide, 1.0, 2.7 + index * 0.45, 9.0, 0.35, bullet, 14)

    def _render_agenda_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Agenda"), top=0.6, size=24)
        card_positions = [0.9, 4.45, 8.0]
        for index, bullet in enumerate(slide_data.get("bullets", [])[:3]):
            self._add_card(slide, card_positions[index], 1.7, 2.8, 2.3, f"0{index + 1}", bullet)

    def _render_exec_summary_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Executive Summary"), top=0.65, size=24)
        bullets = slide_data.get("bullets", [])[:3]
        positions = [0.9, 4.5, 8.1]
        for index, bullet in enumerate(bullets):
            self._add_summary_panel(slide, positions[index], 1.6, 2.9, 3.5, f"Takeaway {index + 1}", bullet)

    def _render_framework_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Framework"), top=0.65, size=24)
        self._add_text_block(slide, 0.85, 1.65, 5.7, 4.8, slide_data.get("bullets", []), 16)
        self._add_callout(slide, 7.1, 1.8, 4.7, 3.6, slide_data.get("visualSuggestion", ""))

    def _render_funnel_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Patient Funnel"), top=0.65, size=24)
        bullets = slide_data.get("bullets", [])[:5]
        start_left = 0.9
        width = 2.1
        for index, bullet in enumerate(bullets):
            left = start_left + index * 2.35
            height = 3.0 - index * 0.22 if len(bullets) > 3 else 2.5
            top = 2.0 + index * 0.08
            shape = slide.shapes.add_shape(self.MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, self.Inches(left), self.Inches(top), self.Inches(width), self.Inches(height))
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.RGBColor(230 - index * 10, 239 - index * 6, 252 - index * 6)
            shape.line.color.rgb = self.RGBColor(30, 64, 175)
            self._set_shape_text(shape, bullet, 13)
        self._add_textbox(slide, 0.95, 5.7, 10.6, 0.45, slide_data.get("visualSuggestion", ""), 12, color=(75, 85, 99), italic=True)

    def _render_insight_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Insight"), top=0.65, size=24)
        self._add_text_block(slide, 0.85, 1.65, 5.3, 4.8, slide_data.get("bullets", []), 16)
        self._add_visual_placeholder(slide, 6.7, 1.65, 5.2, 4.2, slide_data.get("visualSuggestion", ""))

    def _render_recommendation_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Recommendations"), top=0.65, size=24)
        labels = ["Priority 1", "Priority 2", "Priority 3"]
        for index, bullet in enumerate(slide_data.get("bullets", [])[:3]):
            self._add_summary_panel(slide, 0.95, 1.7 + index * 1.4, 10.9, 1.0, labels[index], bullet)

    def _render_appendix_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", "Appendix"), top=0.65, size=24)
        self._add_text_block(slide, 0.95, 1.7, 10.8, 4.9, slide_data.get("bullets", []), 15)

    def _set_title(self, slide, text: str, top: float, size: int) -> None:
        title_placeholder = None
        for shape in slide.placeholders:
            try:
                if shape.placeholder_format.type == self.PP_PLACEHOLDER.TITLE:
                    title_placeholder = shape
                    break
            except Exception:
                continue
        if title_placeholder is not None:
            title_placeholder.text = text
            for paragraph in title_placeholder.text_frame.paragraphs:
                paragraph.font.size = self.Pt(size)
                paragraph.font.bold = True
            return
        self._add_textbox(slide, 0.85, top, 10.5, 0.7, text, size, bold=True)

    def _add_text_block(self, slide, left: float, top: float, width: float, height: float, bullets: list[str], size: int) -> None:
        box = slide.shapes.add_textbox(self.Inches(left), self.Inches(top), self.Inches(width), self.Inches(height))
        frame = box.text_frame
        frame.word_wrap = True
        for index, bullet in enumerate(bullets):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.text = bullet
            paragraph.level = 0
            paragraph.font.size = self.Pt(size)
            paragraph.font.color.rgb = self.RGBColor(31, 41, 55)
            paragraph.space_after = self.Pt(10)

    def _add_visual_placeholder(self, slide, left: float, top: float, width: float, height: float, hint: str) -> None:
        shape = slide.shapes.add_shape(self.MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, self.Inches(left), self.Inches(top), self.Inches(width), self.Inches(height))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.RGBColor(255, 255, 255)
        shape.line.color.rgb = self.RGBColor(191, 219, 254)
        self._set_shape_text(shape, f"Visual placeholder\n\n{hint or 'Add chart, network, table, or patient-flow visual here.'}", 14)

    def _add_callout(self, slide, left: float, top: float, width: float, height: float, text: str) -> None:
        shape = slide.shapes.add_shape(self.MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, self.Inches(left), self.Inches(top), self.Inches(width), self.Inches(height))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.RGBColor(239, 246, 255)
        shape.line.color.rgb = self.RGBColor(30, 64, 175)
        self._set_shape_text(shape, text or "Add implication, definition, or supporting visual note here.", 14)

    def _add_card(self, slide, left: float, top: float, width: float, height: float, label: str, body: str) -> None:
        shape = slide.shapes.add_shape(self.MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, self.Inches(left), self.Inches(top), self.Inches(width), self.Inches(height))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.RGBColor(250, 250, 252)
        shape.line.color.rgb = self.RGBColor(203, 213, 225)
        tf = shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = self.MSO_ANCHOR.MIDDLE
        p1 = tf.paragraphs[0]
        p1.text = label
        p1.font.size = self.Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = self.RGBColor(30, 64, 175)
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.size = self.Pt(14)
        p2.font.color.rgb = self.RGBColor(31, 41, 55)

    def _add_summary_panel(self, slide, left: float, top: float, width: float, height: float, heading: str, body: str) -> None:
        shape = slide.shapes.add_shape(self.MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, self.Inches(left), self.Inches(top), self.Inches(width), self.Inches(height))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.RGBColor(255, 255, 255)
        shape.line.color.rgb = self.RGBColor(203, 213, 225)
        tf = shape.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = heading
        p1.font.size = self.Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = self.RGBColor(30, 64, 175)
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.size = self.Pt(14)
        p2.font.color.rgb = self.RGBColor(31, 41, 55)

    def _add_textbox(self, slide, left: float, top: float, width: float, height: float, text: str, size: int, bold: bool = False, italic: bool = False, color=(31, 41, 55)) -> None:
        box = slide.shapes.add_textbox(self.Inches(left), self.Inches(top), self.Inches(width), self.Inches(height))
        paragraph = box.text_frame.paragraphs[0]
        paragraph.text = text
        paragraph.font.size = self.Pt(size)
        paragraph.font.bold = bold
        paragraph.font.italic = italic
        paragraph.font.color.rgb = self.RGBColor(*color)

    def _set_shape_text(self, shape, text: str, size: int) -> None:
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        paragraph = tf.paragraphs[0]
        paragraph.text = text
        paragraph.font.size = self.Pt(size)
        paragraph.font.color.rgb = self.RGBColor(31, 41, 55)
        paragraph.alignment = self.PP_ALIGN.LEFT

    def _add_footer(self, slide, slide_data: dict) -> None:
        footer = slide.shapes.add_textbox(self.Inches(0.7), self.Inches(6.75), self.Inches(12.0), self.Inches(0.3))
        paragraph = footer.text_frame.paragraphs[0]
        paragraph.text = (
            f"Claim sources: {', '.join(slide_data.get('claimSources', [])) or 'None'} | "
            f"Style sources: {', '.join(slide_data.get('styleSources', [])) or 'None'}"
        )
        paragraph.font.size = self.Pt(9)
        paragraph.font.color.rgb = self.RGBColor(107, 114, 128)

    def _attach_notes(self, slide, slide_data: dict) -> None:
        notes = slide.notes_slide.notes_text_frame
        notes.clear()
        text = [
            f"Objective: {slide_data.get('objective', '')}",
            f"Visual suggestion: {slide_data.get('visualSuggestion', '')}",
        ]
        claim_sources = slide_data.get('claimSources', [])
        style_sources = slide_data.get('styleSources', [])
        if claim_sources:
            text.append("Claim sources: " + ", ".join(claim_sources))
        if style_sources:
            text.append("Style sources: " + ", ".join(style_sources))
        for note in slide_data.get('notes', []):
            text.append(note)
        notes.text = "\n".join(text)
