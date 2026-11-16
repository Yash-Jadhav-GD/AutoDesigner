"""
Semantic Slide Composer
Analyzes extracted SlideModel elements from any PowerPoint slide and
automatically maps them into a Composable SlideComposition tree (VStack, HStack, Grid).
"""

from typing import Any, Dict, List, Optional, Tuple
import re

from ..slide_model import SlideModel, Element
from .primitives import (
    SlideComposition,
    LayoutNode,
    HStack,
    VStack,
    Grid,
    CardNode,
    PillHeaderNode,
    BulletListNode,
    SectionListNode,
    SectionItem,
    KPINode,
    ProcessRowNode,
    TableNode,
    DividerNode,
)


class SemanticComposer:
    """Transforms raw SlideModel elements into a clean SlideComposition tree."""

    @classmethod
    def compose(cls, model: SlideModel) -> SlideComposition:
        title_text, subtitle_text, body_elements = cls._extract_titles(model)

        # Check for Tables
        table_elem = next((e for e in body_elements if e.kind == "table" and e.table), None)
        if table_elem is not None:
            root = cls._compose_table(table_elem)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # Check for Multi-Tier Cards + Bottom KPIs (Slide 4 pattern)
        cards, kpis = cls._detect_cards_and_kpis(body_elements)
        if cards and kpis and len(cards) >= 2 and len(kpis) >= 2:
            root = cls._compose_cards_and_kpis(cards, kpis)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # Check for Columnar Cards (Slide 1 pattern, 2-4 columns)
        column_clusters = cls._cluster_columns(body_elements)
        if len(column_clusters) >= 2:
            root = cls._compose_column_cards(column_clusters, title_text)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # Fallback: Process Rows (vertical stack)
        row_clusters = cls._cluster_rows(body_elements)
        if len(row_clusters) >= 2:
            root = cls._compose_process_rows(row_clusters)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # General Card Fallback
        single_card = CardNode(
            header=PillHeaderNode(title="Overview", style="flush", fill_color="BRAND_INDIGO"),
            body=BulletListNode(items=[e.text for e in body_elements if e.text])
        )
        return SlideComposition(title=title_text, subtitle=subtitle_text, root=single_card)

    @classmethod
    def _extract_titles(cls, model: SlideModel) -> Tuple[str, Optional[str], List[Element]]:
        title_text = ""
        subtitle_text = None
        remaining: List[Element] = []

        # Find title
        title_elem = model.title() if callable(getattr(model, "title", None)) else getattr(model, "title", None)
        if title_elem is None:
            # Heuristic: top-most text element with largest font size
            text_elems = [e for e in model.elements if e.kind == "text" and e.text.strip()]
            if text_elems:
                text_elems.sort(key=lambda e: (e.y, -(e.size or 0)))
                title_elem = text_elems[0]

        if title_elem:
            # Split multiline title if present (e.g. Title\nSubtitle)
            lines = [p.text.strip() for p in title_elem.paras if p.text.strip()]
            if len(lines) >= 2:
                title_text = lines[0]
                subtitle_text = lines[1]
            elif lines:
                title_text = lines[0]

        # Find subtitle if not extracted yet
        sub_elem = model.subtitle() if callable(getattr(model, "subtitle", None)) else getattr(model, "subtitle", None)
        if not subtitle_text and sub_elem:
            sub_lines = [p.text.strip() for p in sub_elem.paras if p.text.strip()]
            if sub_lines:
                subtitle_text = " ".join(sub_lines)

        for e in model.elements:
            if title_elem and e.id == title_elem.id:
                continue
            if sub_elem and e.id == sub_elem.id:
                continue
            # Ignore decorative background rectangles
            if e.w >= 900.0 and e.h >= 500.0:
                continue
            remaining.append(e)

        return title_text, subtitle_text, remaining

    @classmethod
    def _compose_table(cls, table_elem: Element) -> TableNode:
        grid = table_elem.table or []
        headers = grid[0] if grid else []
        rows = grid[1:] if len(grid) > 1 else []
        return TableNode(
            headers=headers,
            rows=rows,
            header_fill="BRAND_INDIGO",
            zebra_fill="NEUTRAL_SURFACE",
            totals_row=True,
            totals_fill="BRAND_NAVY"
        )

    @classmethod
    def _detect_cards_and_kpis(cls, elements: List[Element]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """Detects if slide content separates into a top card band and bottom KPI summary tiles."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if len(text_elems) < 5:
            return [], []

        # Find bottom band (e.g. y > 400 or lower 30% of slide)
        bottom_cutoff = 380.0
        top_elems = [e for e in text_elems if e.y < bottom_cutoff]
        bottom_elems = [e for e in text_elems if e.y >= bottom_cutoff]

        # Check if bottom elems look like KPIs (short numbers/labels)
        if len(bottom_elems) >= 3:
            kpis = []
            # Cluster bottom elements by X
            sorted_bottom = sorted(bottom_elems, key=lambda e: e.x)
            for e in sorted_bottom:
                lines = [p.text.strip() for p in e.paras if p.text.strip()]
                if lines:
                    num = lines[0]
                    lbl = lines[1] if len(lines) > 1 else ""
                    sub = lines[2] if len(lines) > 2 else None
                    kpis.append({"number": num, "label": lbl, "sublabel": sub})

            # Cluster top elements into 2-4 columns
            top_cols = cls._cluster_columns(top_elems)
            return top_cols, kpis

        return [], []

    @classmethod
    def _compose_cards_and_kpis(cls, cards_data: List[Dict[str, Any]], kpis_data: List[Dict[str, Any]]) -> VStack:
        # Build Top Cards
        card_nodes = []
        for idx, col in enumerate(cards_data):
            h_color = "BRAND_TANGERINE" if idx == len(cards_data) - 1 else "BRAND_INDIGO"
            header_node = PillHeaderNode(title=col["header"], style="flush", fill_color=h_color)
            card_nodes.append(CardNode(
                header=header_node,
                body=BulletListNode(items=col["bullets"], font_size=10.5)
            ))

        top_hstack = HStack(height_weight=3.2, gap=18.0, children=card_nodes)

        # Build Bottom KPIs
        kpi_nodes = []
        for idx, k in enumerate(kpis_data):
            n_color = "BRAND_TANGERINE" if (idx == 0 or idx == len(kpis_data) - 1) else "BRAND_INDIGO"
            kpi_nodes.append(KPINode(
                number=k["number"],
                label=k["label"],
                sublabel=k.get("sublabel"),
                number_color=n_color,
                number_size=20.0
            ))

        bottom_hstack = HStack(height_weight=1.0, gap=14.0, children=kpi_nodes)

        return VStack(gap=14.0, children=[top_hstack, bottom_hstack])

    @classmethod
    def _cluster_columns(cls, elements: List[Element]) -> List[Dict[str, Any]]:
        """Clusters elements into columns based on horizontal positions (X-coordinates)."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if not text_elems:
            return []

        # Sort by X
        sorted_x = sorted(text_elems, key=lambda e: e.x)
        columns = []
        cur_col = [sorted_x[0]]

        for e in sorted_x[1:]:
            prev = cur_col[-1]
            if abs(e.x - prev.x) > 60.0:  # distinct column threshold
                columns.append(cur_col)
                cur_col = [e]
            else:
                cur_col.append(e)
        if cur_col:
            columns.append(cur_col)

        if len(columns) < 2 or len(columns) > 4:
            return []

        result = []
        for col_elems in columns:
            col_elems.sort(key=lambda e: e.y)
            top_elem = col_elems[0]
            paras = [p.text.strip() for p in top_elem.paras if p.text.strip()]
            header = paras[0] if paras else "Section"
            bullets = []
            if len(paras) > 1:
                bullets.extend(paras[1:])
            for other in col_elems[1:]:
                bullets.extend([p.text.strip() for p in other.paras if p.text.strip()])

            result.append({"header": header, "bullets": bullets})

        return result

    @classmethod
    def _compose_column_cards(cls, columns_data: List[Dict[str, Any]], title_text: str) -> HStack:
        is_strategic_pillars = len(columns_data) == 3 and any(
            k in title_text.lower() for k in ("pillar", "future", "state", "transformation", "strategic")
        )
        gradients = ["linear_indigo", "linear_navy", "linear_tangerine"]

        card_nodes = []
        for idx, col in enumerate(columns_data):
            if is_strategic_pillars:
                h = PillHeaderNode(
                    title=col["header"],
                    style="floating_pill",
                    gradient=gradients[idx % 3],
                    protrusion_pt=16.0
                )
            else:
                h_color = "BRAND_TANGERINE" if idx == len(columns_data) - 1 else "BRAND_INDIGO"
                h = PillHeaderNode(title=col["header"], style="flush", fill_color=h_color)

            card_nodes.append(CardNode(
                header=h,
                body=BulletListNode(items=col["bullets"], font_size=11.0)
            ))

        return HStack(gap=18.0, children=card_nodes)

    @classmethod
    def _cluster_rows(cls, elements: List[Element]) -> List[Dict[str, Any]]:
        """Clusters elements into vertical rows."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if not text_elems:
            return []

        sorted_y = sorted(text_elems, key=lambda e: e.y)
        rows = []
        cur_row = [sorted_y[0]]

        for e in sorted_y[1:]:
            prev = cur_row[-1]
            if abs(e.y - prev.y) > 25.0:
                rows.append(cur_row)
                cur_row = [e]
            else:
                cur_row.append(e)
        if cur_row:
            rows.append(cur_row)

        result = []
        for row_elems in rows:
            text = "\n".join(e.text for e in row_elems if e.text.strip())
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            title = lines[0] if lines else ""
            body = " ".join(lines[1:]) if len(lines) > 1 else ""
            result.append({"title": title, "body": body})

        return result

    @classmethod
    def _compose_process_rows(cls, rows_data: List[Dict[str, Any]]) -> VStack:
        nodes = []
        for r in rows_data:
            nodes.append(ProcessRowNode(
                title=r["title"],
                body=r["body"],
                divider="side_fade_r"
            ))
        return VStack(gap=10.0, children=nodes)
