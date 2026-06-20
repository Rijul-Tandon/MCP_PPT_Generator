from __future__ import annotations

import re
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

from .models import DeckContextChunk, StylePattern


class PptxReferenceLibrary:
    def __init__(self, reference_dir: Path) -> None:
        self.reference_dir = reference_dir

    def list_decks(self) -> list[Path]:
        if not self.reference_dir.exists():
            return []
        return sorted(self.reference_dir.glob("*.pptx"))

    def extract_pptx_text(self, deck_path: Path) -> list[DeckContextChunk]:
        chunks: list[DeckContextChunk] = []
        with ZipFile(deck_path) as archive:
            slide_names = sorted(
                (name for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")),
                key=self._slide_sort_key,
            )
            for slide_name in slide_names:
                xml = archive.read(slide_name).decode("utf-8", errors="ignore")
                text_runs = re.findall(r"<a:t>(.*?)</a:t>", xml)
                clean_text = " ".join(self._clean_xml_text(text) for text in text_runs).strip()
                if not clean_text:
                    continue
                slide_number = self._slide_sort_key(slide_name)
                chunks.append(
                    DeckContextChunk(
                        sourceFile=deck_path.name,
                        slideNumber=slide_number,
                        text=clean_text,
                        tags=self._infer_tags(clean_text),
                    )
                )
        return chunks

    def derive_style_patterns(self, chunks: list[DeckContextChunk]) -> list[StylePattern]:
        title_counter: Counter[str] = Counter()
        has_exec_summary = False
        has_agenda = False
        has_geo = False
        has_persona = False
        has_recommendation = False

        for chunk in chunks:
            first_sentence = re.split(r"[.!?]", chunk.text)[0].strip()
            if first_sentence:
                title_counter[first_sentence[:100]] += 1
            lowered = chunk.text.lower()
            has_exec_summary |= "executive summary" in lowered
            has_agenda |= "focus of the conversation" in lowered or "agenda" in lowered
            has_geo |= "geograph" in lowered or "regional" in lowered
            has_persona |= "persona" in lowered or "hcp" in lowered
            has_recommendation |= "recommendation" in lowered or "engagement" in lowered

        common_titles = [title for title, _ in title_counter.most_common(6)]
        patterns = [
            StylePattern(
                sectionType="narrative",
                commonTitles=common_titles,
                toneNotes="Use concise executive headlines followed by evidence-based support.",
                layoutHints="Favor one takeaway headline with 3-4 supporting bullets and a visual callout.",
            )
        ]
        if has_agenda:
            patterns.append(
                StylePattern(
                    sectionType="agenda",
                    commonTitles=["Focus of the conversation today", "Agenda"],
                    toneNotes="Set expectations early and group the story into clear modules.",
                    layoutHints="Use an agenda slide near the beginning with short section names.",
                )
            )
        if has_exec_summary:
            patterns.append(
                StylePattern(
                    sectionType="executive_summary",
                    commonTitles=["Executive Summary"],
                    toneNotes="Lead with the decision and summarize supporting themes.",
                    layoutHints="Open each section with a summary slide before the evidence slides.",
                )
            )
        if has_geo:
            patterns.append(
                StylePattern(
                    sectionType="geography",
                    commonTitles=["Regional Care Patterns", "Geographic Hotspots"],
                    toneNotes="Compare regions in terms of concentration, maturity, and opportunity.",
                    layoutHints="Use a map or ranked table placeholder with 3 insight bullets.",
                )
            )
        if has_persona:
            patterns.append(
                StylePattern(
                    sectionType="persona",
                    commonTitles=["Persona Landscape", "HCP Personas"],
                    toneNotes="Frame provider or patient cohorts as operationally distinct segments.",
                    layoutHints="Use a matrix or grouped chart placeholder with segment-specific implications.",
                )
            )
        if has_recommendation:
            patterns.append(
                StylePattern(
                    sectionType="recommendation",
                    commonTitles=["Recommendations", "Strategic Implications"],
                    toneNotes="Translate analytics into action and prioritize by business leverage.",
                    layoutHints="Close with action-oriented bullets and explicit next steps.",
                )
            )
        return patterns

    def summarize_reference_patterns(self) -> dict[str, object]:
        decks = self.list_decks()
        chunks = [chunk for deck in decks for chunk in self.extract_pptx_text(deck)]
        patterns = self.derive_style_patterns(chunks)
        return {
            "referenceDecks": [deck.name for deck in decks],
            "slideCount": len(chunks),
            "stylePatterns": [pattern.__dict__ for pattern in patterns],
        }

    @staticmethod
    def _infer_tags(text: str) -> list[str]:
        lowered = text.lower()
        tags: list[str] = []
        if "executive summary" in lowered:
            tags.append("executive_summary")
        if "referral" in lowered:
            tags.append("referral")
        if "patient" in lowered:
            tags.append("patient")
        if "network" in lowered:
            tags.append("network")
        if "persona" in lowered or "hcp" in lowered:
            tags.append("persona")
        if "recommendation" in lowered or "insight" in lowered:
            tags.append("recommendation")
        return tags

    @staticmethod
    def _clean_xml_text(text: str) -> str:
        return (
            text.replace("&amp;", "&")
            .replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&#10;", " ")
            .strip()
        )

    @staticmethod
    def _slide_sort_key(name: str) -> int:
        match = re.search(r"slide(\d+)\.xml$", name)
        return int(match.group(1)) if match else 0
