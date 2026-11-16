"""
AutoDesigner Semantic Content Analyzer
Parses input PowerPoint slides into structured semantic representations:
Titles, Subtitles, Card Clusters, Lists, KPIs, and Steps.
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class ItemData:
    subheader: str = ""
    body: str = ""
    badge: str = ""
    is_kpi: bool = False
    kpi_metric: str = ""
    kpi_label: str = ""


@dataclass
class CardData:
    header: str = ""
    category: str = ""
    items: List[ItemData] = field(default_factory=list)
    kpi_metric: Optional[str] = None
    kpi_label: Optional[str] = None


@dataclass
class SlideData:
    slide_index: int = 0
    title: str = ""
    category: str = ""
    subtitle: str = ""
    cards: List[CardData] = field(default_factory=list)
    raw_texts: List[str] = field(default_factory=list)


class ContentAnalyzer:
    """Analyzes raw slides and decomposes shapes into structured card data."""

    KPI_REGEX = re.compile(
        r"^([\$€£¥]?\d+[\d,\.]*[%xXkKmMbB\+]?|\+\d+[\d,\.]*%?)(?:\s+|$)"
    )

    @classmethod
    def analyze_slide(cls, slide, slide_index: int = 0) -> SlideData:
        slide_data = SlideData(slide_index=slide_index)

        # Collect all text-bearing shapes
        text_shapes = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    text_shapes.append((shape, text))

        if not text_shapes:
            return slide_data

        # Determine if there is a dedicated slide title
        title_shape = None
        remaining_shapes = []

        # Check placeholder first
        for shape, text in text_shapes:
            if shape.is_placeholder and shape.placeholder_format.type in (1, 3):  # Title or Center Title
                title_shape = shape
                slide_data.title = text
                break

        # If no placeholder title, inspect topmost shapes
        if not title_shape:
            # Sort by top coordinate
            sorted_by_top = sorted(text_shapes, key=lambda s: s[0].top)
            min_top = sorted_by_top[0][0].top
            # Check how many shapes are at approximately this topmost Y level (within 25pt = 25 * 12700 EMUs)
            top_tier = [s for s in sorted_by_top if abs(s[0].top - min_top) < 25 * 12700]

            if len(top_tier) == 1:
                # Single prominent topmost shape: check if it qualifies as title
                candidate_shape, candidate_text = top_tier[0]
                if candidate_shape.top < 120 * 12700 and (len(candidate_text) < 100 or candidate_shape.width > 350 * 12700):
                    title_shape = candidate_shape
                    slide_data.title = candidate_text
            # If len(top_tier) > 1, these are multi-column headers (e.g. "The Problem", "The Bridge", "The Future State")!
            # Do NOT treat them as slide title.

        # Remaining shapes to deconstruct into cards
        for shape, text in text_shapes:
            if shape != title_shape:
                remaining_shapes.append((shape, text))

        # Check if title has category or subtitle
        if slide_data.title and "\n" in slide_data.title:
            lines = [l.strip() for l in slide_data.title.split("\n") if l.strip()]
            if len(lines) >= 2:
                slide_data.title = lines[0]
                slide_data.subtitle = " - ".join(lines[1:])

        # Deconstruct remaining shapes into cards
        cls._deconstruct_into_cards(remaining_shapes, slide_data)
        return slide_data

    @classmethod
    def _deconstruct_into_cards(cls, remaining_shapes, slide_data: SlideData):
        if not remaining_shapes:
            return

        # Case 1: Multiple shapes spaced horizontally (each is a card or column)
        shapes_by_x = {}
        for shape, text in remaining_shapes:
            x_col = round(shape.left / (914400 / 72 * 120))  # cluster by ~120pt horizontal buckets
            if x_col not in shapes_by_x:
                shapes_by_x[x_col] = []
            shapes_by_x[x_col].append((shape, text))

        # If clustered horizontally into 2 to 4 columns:
        if 2 <= len(shapes_by_x) <= 4:
            sorted_cols = sorted(shapes_by_x.keys())
            for col_idx, col_key in enumerate(sorted_cols):
                col_shapes = shapes_by_x[col_key]
                col_shapes.sort(key=lambda s: s[0].top)
                card = cls._build_card_from_shapes(col_shapes, default_title=f"Section {col_idx + 1}")
                slide_data.cards.append(card)
            return

        # Case 2: Shapes in a single column or single text shape with multiple paragraphs
        all_paragraphs = []
        for shape, text in remaining_shapes:
            for p in shape.text_frame.paragraphs:
                p_text = p.text.strip()
                if p_text:
                    all_paragraphs.append(p_text)

        if not all_paragraphs:
            return

        # Parse paragraphs into logical cards or single multi-item card
        cards = cls._partition_paragraphs_into_cards(all_paragraphs)
        slide_data.cards.extend(cards)

    @classmethod
    def _build_card_from_shapes(cls, shapes, default_title: str) -> CardData:
        card = CardData(header=default_title)
        items = []

        # If top shape has short text, treat as card header
        first_shape, first_text = shapes[0]
        first_lines = [l.strip() for l in first_text.split("\n") if l.strip()]

        if len(first_lines) == 1 and len(first_text) < 60:
            card.header = first_lines[0]
            shapes = shapes[1:]
        elif len(first_lines) > 1 and len(first_lines[0]) < 60:
            card.header = first_lines[0]
            # Remaining lines are items
            for line in first_lines[1:]:
                items.append(cls._parse_item_line(line))
            shapes = shapes[1:]

        for shape, text in shapes:
            for p in shape.text_frame.paragraphs:
                p_text = p.text.strip()
                if p_text:
                    items.append(cls._parse_item_line(p_text))

        card.items = items
        cls._detect_kpi_on_card(card)
        return card

    @classmethod
    def _partition_paragraphs_into_cards(cls, paragraphs: List[str]) -> List[CardData]:
        cards = []

        # Check if paragraphs naturally divide into 2, 3, or 4 groups
        # Look for headers (short lines < 40 chars without bullet/punctuation endings)
        header_indices = []
        for idx, p in enumerate(paragraphs):
            is_header_candidate = (
                len(p) <= 45
                and not p.startswith(("-", "•", "*"))
                and (p.isupper() or p.istitle() or p.endswith(":") or idx == 0)
            )
            if is_header_candidate and idx < len(paragraphs) - 1:
                header_indices.append(idx)

        # If we have 2 to 4 distinct headers, partition by header
        if 2 <= len(header_indices) <= 4:
            for i, h_idx in enumerate(header_indices):
                next_idx = header_indices[i + 1] if i + 1 < len(header_indices) else len(paragraphs)
                header_text = paragraphs[h_idx].rstrip(":")
                card = CardData(header=header_text)
                for item_text in paragraphs[h_idx + 1 : next_idx]:
                    card.items.append(cls._parse_item_line(item_text))
                cls._detect_kpi_on_card(card)
                cards.append(card)
            return cards

        # Fallback: create cards grouped by count or a single structured card
        if len(paragraphs) <= 4:
            # Single card with items
            card = CardData(header="Key Highlights")
            for p in paragraphs:
                card.items.append(cls._parse_item_line(p))
            cls._detect_kpi_on_card(card)
            cards.append(card)
        elif 5 <= len(paragraphs) <= 8:
            # Split into 2 cards
            mid = (len(paragraphs) + 1) // 2
            c1 = CardData(header="Focus Area 1", items=[cls._parse_item_line(p) for p in paragraphs[:mid]])
            c2 = CardData(header="Focus Area 2", items=[cls._parse_item_line(p) for p in paragraphs[mid:]])
            cards.extend([c1, c2])
        else:
            # Split into 3 cards
            chunk_size = (len(paragraphs) + 2) // 3
            for i in range(0, len(paragraphs), chunk_size):
                chunk = paragraphs[i : i + chunk_size]
                c = CardData(header=f"Phase {(i // chunk_size) + 1}", items=[cls._parse_item_line(p) for p in chunk])
                cards.append(c)

        return cards

    @classmethod
    def _parse_item_line(cls, text: str) -> ItemData:
        # Strip leading bullet points or numbers (e.g. "• ", "- ", "1. ", "1) ")
        # but preserve figures like "500+ activities"
        cleaned = re.sub(r"^[\s•\-\*]+|^\d+[\.\)]\s*", "", text).strip()
        item = ItemData()

        # Check for subheader pattern: "Key Concept: Detailed explanation here"
        if ":" in cleaned:
            parts = cleaned.split(":", 1)
            if len(parts[0]) <= 35 and len(parts[1]) > 0:
                item.subheader = parts[0].strip()
                item.body = parts[1].strip()
                return item

        # Check for dash separator: "Concept - Explanation"
        if " - " in cleaned:
            parts = cleaned.split(" - ", 1)
            if len(parts[0]) <= 30 and len(parts[1]) > 0:
                item.subheader = parts[0].strip()
                item.body = parts[1].strip()
                return item

        item.body = cleaned
        return item

    @classmethod
    def _detect_kpi_on_card(cls, card: CardData):
        if not card.items:
            return

        # Check if first item looks like a metric
        first_item = card.items[0]
        text_to_check = first_item.subheader or first_item.body
        match = cls.KPI_REGEX.match(text_to_check)
        if match:
            card.kpi_metric = match.group(1)
            remainder = text_to_check[match.end() :].strip()
            card.kpi_label = remainder if remainder else (first_item.body if first_item.subheader else "")
            # Remove KPI from regular items
            card.items = card.items[1:]
