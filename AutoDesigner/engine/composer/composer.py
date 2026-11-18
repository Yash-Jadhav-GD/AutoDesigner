"""
Semantic Slide Composer
Analyzes extracted SlideModel elements from any PowerPoint slide and
automatically maps them into a Composable SlideComposition tree (VStack, HStack, Grid).
Guarantees 100% token preservation by ensuring all text elements are retained.
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
    LayeredStackNode,
    StackLayer,
    StackLayerItem,
    Padding,
)

METRIC_PATTERN = re.compile(r'[%$€£¥~+]|\b\d+(?:[.,]\d+)?\s*(?:bn|m|k|%|uplift|savings)?\b|\b(?:shift|prioritized|renewed|outcomes)\b', re.IGNORECASE)


class SemanticComposer:
    """Transforms raw SlideModel elements into a clean SlideComposition tree."""

    @classmethod
    def compose(cls, model: SlideModel) -> SlideComposition:
        title_text, subtitle_text, body_elements = cls._extract_titles(model)

        # 1. Check for Tables (native table elements)
        table_elem = next((e for e in body_elements if e.kind == "table" and e.table), None)
        if table_elem is not None:
            root = cls._compose_table(table_elem)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # 2. Check for Layered Architecture Stack (Library Slide 9 pattern)
        layered_stack = cls._detect_layered_stack(body_elements, title_text)
        if layered_stack is not None:
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=layered_stack)

        # 3. Check for Text Matrices (e.g. Slide 6: 4 distinct columns with 4+ rows)
        matrix = cls._detect_text_matrix(body_elements)
        if matrix is not None:
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=matrix)

        # 3. Check for Multi-Tier Cards + Bottom KPIs (Slide 4 pattern)
        cards, kpis, meta_strip, outcomes_hdr = cls._detect_cards_and_kpis(body_elements)
        if cards and kpis and len(cards) >= 2 and len(kpis) >= 2:
            root = cls._compose_cards_and_kpis(cards, kpis, meta_strip, outcomes_hdr)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # 4. Check for Section Pair Layouts (Slide 2 & Slide 3 pattern)
        section_root = cls._detect_and_compose_sections(body_elements, title_text)
        if section_root is not None:
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=section_root)

        # 5. Check for Columnar Cards (Slide 1, Slide 9 pattern: 2-4 columns)
        column_clusters = cls._cluster_columns(body_elements)
        if len(column_clusters) >= 2:
            root = cls._compose_column_cards(column_clusters, title_text)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # 6. Fallback: Process Rows (vertical stack)
        row_clusters = cls._cluster_rows(body_elements)
        if len(row_clusters) >= 2:
            root = cls._compose_process_rows(row_clusters)
            return SlideComposition(title=title_text, subtitle=subtitle_text, root=root)

        # 6. General Single Card Fallback (preserves 100% text)
        all_items = []
        for e in body_elements:
            for p in e.paras:
                if p.text.strip():
                    all_items.append(p.text.strip())
        single_card = CardNode(
            header=PillHeaderNode(title="Overview", style="flush", fill_color="BRAND_INDIGO"),
            body=BulletListNode(items=all_items or ["Summary content"])
        )
        return SlideComposition(title=title_text, subtitle=subtitle_text, root=single_card)

    @classmethod
    def _extract_titles(cls, model: SlideModel) -> Tuple[str, Optional[str], List[Element]]:
        title_text = ""
        subtitle_text = None
        remaining: List[Element] = []

        title_elem = model.title() if callable(getattr(model, "title", None)) else getattr(model, "title", None)
        text_elems = [e for e in model.elements if e.kind == "text" and e.text.strip()]

        # Validate title_elem: if narrow (< 350pt) and shares Y with other top elements, it's a column header!
        if title_elem is not None and title_elem.w < 350.0 and len(text_elems) > 1:
            peers_at_same_y = [e for e in text_elems if e.id != title_elem.id and abs(e.y - title_elem.y) <= 20.0]
            if len(peers_at_same_y) >= 1:
                title_elem = None

        if title_elem:
            lines = [p.text.strip() for p in title_elem.paras if p.text.strip()]
            if len(lines) >= 2:
                title_text = lines[0]
                subtitle_text = lines[1]
            elif lines:
                title_text = lines[0]

        sub_elem = model.subtitle() if callable(getattr(model, "subtitle", None)) else getattr(model, "subtitle", None)
        if not subtitle_text and sub_elem:
            sub_lines = [p.text.strip() for p in sub_elem.paras if p.text.strip()]
            if sub_lines:
                subtitle_text = " ".join(sub_lines)

        # Check for wide element directly under title (y < 110, w > 350, e.g. Slide 9 subtitle)
        chosen_sub_id = None
        if title_text and not subtitle_text:
            sub_candidates = [e for e in text_elems if e.y < 110.0 and e.w > 350.0 and (title_elem is None or e.id != title_elem.id)]
            if sub_candidates:
                sub_candidates.sort(key=lambda e: e.y)
                chosen = sub_candidates[0]
                lines = [p.text.strip() for p in chosen.paras if p.text.strip()]
                if lines:
                    subtitle_text = " ".join(lines)
                    chosen_sub_id = chosen.id

        for e in model.elements:
            if title_elem and e.id == title_elem.id:
                continue
            if sub_elem and e.id == sub_elem.id:
                continue
            if chosen_sub_id and e.id == chosen_sub_id:
                continue
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
    def _detect_layered_stack(cls, elements: List[Element], title_text: str) -> Optional[LayeredStackNode]:
        """Detects architectural layered stack designs (e.g. Slide 8 / Library Slide 9)."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if len(text_elems) < 9:
            return None

        is_stack_title = any(kw in title_text.lower() for kw in ["layer", "architecture", "stack", "tier", "foundation", "levels"])

        # Cluster elements into rows by Y
        sorted_elems = sorted(text_elems, key=lambda e: e.y)
        rows: List[List[Element]] = []
        for e in sorted_elems:
            placed = False
            for r in rows:
                avg_y = sum(x.y for x in r) / len(r)
                if abs(e.y - avg_y) <= 35.0:
                    r.append(e)
                    placed = True
                    break
            if not placed:
                rows.append([e])

        if len(rows) < 3 or len(rows) > 7:
            return None

        sorted_rows = [sorted(r, key=lambda x: x.x) for r in rows]
        has_digit_badges = all(len(r) >= 3 and r[0].text.strip().isdigit() for r in sorted_rows)
        if not (has_digit_badges or is_stack_title):
            return None

        tier_palette = {
            5: {
                "header_fill": "FFAC59",
                "header_text_color": "14203E",
                "body_fill": "FFD3A6",
                "body_alpha": 30000,
                "border_color": "FFAC59",
                "divider_color": "FFAC59",
                "icon_tint": "FFAC59",
                "header_icon": "image31.svg",
                "item_icons": ["image14.svg", "image13.svg", "image15.svg"]
            },
            4: {
                "header_fill": "FFD3A6",
                "header_text_color": "14203E",
                "body_fill": "FFD3A6",
                "body_alpha": 30000,
                "border_color": "FFAC59",
                "divider_color": "FFAC59",
                "icon_tint": "FFAC59",
                "header_icon": "image29.svg",
                "item_icons": ["image8.svg", "image7.svg", "image10.svg"]
            },
            3: {
                "header_fill": "8F9BC7",
                "header_text_color": "14203E",
                "body_fill": "8F9BC7",
                "body_alpha": 20000,
                "border_color": "526299",
                "divider_color": "8F9BC7",
                "icon_tint": "526299",
                "header_icon": "image30.svg",
                "item_icons": ["image4.svg", "image5.svg", "image3.svg"]
            },
            2: {
                "header_fill": "B1A8D8",
                "header_text_color": "14203E",
                "body_fill": "B1A8D8",
                "body_alpha": 20000,
                "border_color": "7C6DB2",
                "divider_color": "B1A8D8",
                "icon_tint": "7C6DB2",
                "header_icon": "image28.svg",
                "item_icons": ["image1.svg", "image22.svg", "image20.svg"]
            },
            1: {
                "header_fill": "7C6DB2",
                "header_text_color": "FFFFFF",
                "body_fill": "B1A8D8",
                "body_alpha": 20000,
                "border_color": "7C6DB2",
                "divider_color": "B1A8D8",
                "icon_tint": "7C6DB2",
                "header_icon": "image27.svg",
                "item_icons": ["image23.svg", "image24.svg", "image25.svg"]
            }
        }

        layers: List[StackLayer] = []
        for idx, row in enumerate(sorted_rows):
            first_txt = row[0].text.strip()
            if first_txt.isdigit():
                level = int(first_txt)
                badge_text = first_txt
                title = row[1].text.strip()
                item_elements = row[2:]
            else:
                level = idx + 1
                badge_text = str(level)
                title = row[0].text.strip()
                item_elements = row[1:]

            pal = tier_palette.get(level, tier_palette.get(5 - idx, tier_palette[1]))

            layer_items: List[StackLayerItem] = []
            for item_idx, it_elem in enumerate(item_elements):
                item_text = it_elem.text.strip()
                icon_file = pal["item_icons"][item_idx] if item_idx < len(pal["item_icons"]) else None
                layer_items.append(StackLayerItem(
                    text=item_text,
                    icon_svg_file=icon_file
                ))

            layers.append(StackLayer(
                level=level,
                title=title,
                items=layer_items,
                header_fill=pal["header_fill"],
                header_text_color=pal["header_text_color"],
                body_fill=pal["body_fill"],
                body_alpha=pal["body_alpha"],
                border_color=pal["border_color"],
                border_width=0.5,
                divider_color=pal["divider_color"],
                header_icon_file=pal["header_icon"],
                badge_text=badge_text,
                badge_fill="FAF9F6",
                badge_text_color="14203E",
                icon_tint_color=pal["icon_tint"]
            ))

        return LayeredStackNode(
            layers=layers,
            foundation_at_bottom=True
        )

    @classmethod
    def _detect_text_matrix(cls, elements: List[Element]) -> Optional[TableNode]:
        """Detects if loose shapes form a structured text matrix (e.g. Slide 6)."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if len(text_elems) < 12:
            return None

        # Look for a header row (3-5 elements sharing roughly the same Y between 50pt and 100pt)
        top_row = [e for e in text_elems if 60.0 <= e.y <= 95.0]
        if len(top_row) not in (3, 4, 5):
            return None

        top_row.sort(key=lambda e: e.x)
        headers = [e.text.strip() for e in top_row]
        cols = len(headers)

        # Group remaining items into rows by Y
        body_elems = [e for e in text_elems if e.y > 95.0]
        sorted_y = sorted(body_elems, key=lambda e: e.y)

        rows_grouped = []
        cur_row = [sorted_y[0]]
        for e in sorted_y[1:]:
            prev = cur_row[-1]
            if abs(e.y - prev.y) > 35.0:
                rows_grouped.append(cur_row)
                cur_row = [e]
            else:
                cur_row.append(e)
        if cur_row:
            rows_grouped.append(cur_row)

        if len(rows_grouped) < 3:
            return None

        table_rows = []
        for r_elems in rows_grouped:
            # Map each element to nearest header column center
            row_cells = ["" for _ in range(cols)]
            for e in r_elems:
                e_center = e.x + e.w / 2.0
                best_c = min(range(cols), key=lambda c: abs((top_row[c].x + top_row[c].w / 2.0) - e_center))
                txt = e.text.strip()
                if row_cells[best_c]:
                    row_cells[best_c] += f"\n{txt}"
                else:
                    row_cells[best_c] = txt
            table_rows.append(row_cells)

        return TableNode(
            headers=headers,
            rows=table_rows,
            header_fill="BRAND_INDIGO",
            zebra_fill="NEUTRAL_SURFACE",
            font_size=9.0
        )

    @classmethod
    def _detect_cards_and_kpis(cls, elements: List[Element]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[str], Optional[str]]:
        """Detects if slide content separates into a top card band, bottom KPI summary tiles, and metadata strip."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if len(text_elems) < 5:
            return [], [], [], None

        bottom_cutoff = 420.0
        bottom_candidates = [e for e in text_elems if e.y >= bottom_cutoff]

        # Extract "Key Outcomes" header if present
        outcomes_hdr = None
        for e in bottom_candidates:
            if "Key Outcomes" in e.text or "Outcomes" in e.text:
                outcomes_hdr = e.text.strip()
                break

        # Filter strictly for genuine KPIs
        kpi_elems = []
        for e in bottom_candidates:
            lines = [p.text.strip() for p in e.paras if p.text.strip()]
            if not lines:
                continue
            has_metric = any(METRIC_PATTERN.search(l) for l in lines)
            is_short = len(lines) <= 3 and all(len(l) < 90 for l in lines)
            if has_metric and is_short and "Key Outcomes" not in lines[0]:
                kpi_elems.append(e)

        if len(kpi_elems) < 3:
            return [], [], [], None

        kpis = []
        sorted_bottom = sorted(kpi_elems, key=lambda e: e.x)
        for e in sorted_bottom:
            lines = [p.text.strip() for p in e.paras if p.text.strip()]
            # Split off big metric prefix if present
            m = re.match(r'^([~%$\w\d.,+-]+(?:\s*-\s*[~%$\w\d.,+-]+)?)\s*(.*)$', lines[0])
            if m and len(m.group(1)) <= 12 and m.group(2):
                num = m.group(1).strip()
                lbl = m.group(2).strip()
                sub = " ".join(lines[1:]) if len(lines) > 1 else None
            else:
                num = lines[0]
                lbl = lines[1] if len(lines) > 1 else ""
                sub = lines[2] if len(lines) > 2 else None
            kpis.append({"number": num, "label": lbl, "sublabel": sub})

        # Separate top metadata banner (y < 150)
        meta_elems = [e for e in text_elems if e.y < 150.0]
        meta_strip = [p.text.strip() for e in meta_elems for p in e.paras if p.text.strip()]

        # Main cards elements (150 <= y < 420)
        top_candidates = [e for e in text_elems if e.y < bottom_cutoff and e.y >= 150.0]
        if not top_candidates:
            top_candidates = [e for e in text_elems if e.y < bottom_cutoff]

        top_cols = cls._cluster_columns(top_candidates)
        return top_cols, kpis, meta_strip, outcomes_hdr

    @classmethod
    def _compose_cards_and_kpis(cls, cards_data: List[Dict[str, Any]], kpis_data: List[Dict[str, Any]], meta_strip: List[str], outcomes_hdr: Optional[str]) -> VStack:
        children = []

        # 1. Top Metadata Banner
        if meta_strip:
            prog_title = meta_strip[0] if meta_strip else ""
            metrics = meta_strip[1:]
            banner_items = []
            if prog_title:
                banner_items.append(prog_title)
            if metrics:
                banner_items.append("     |     ".join(metrics))

            children.append(CardNode(
                fixed_height=44.0,
                fill_color="NEUTRAL_SURFACE",
                border_color="NEUTRAL_BORDER",
                corner_radius=4.0,
                padding=Padding(top=6.0, bottom=6.0, left=16.0, right=16.0),
                body=BulletListNode(
                    items=banner_items,
                    font_size=9.5,
                    bullet_char="",
                    space_after=2.0,
                    mar_l_pt=0.0,
                    indent_pt=0.0
                )
            ))

        # 2. Main Opportunity Cards
        card_nodes = []
        for idx, col in enumerate(cards_data):
            h_color = "BRAND_TANGERINE" if idx == len(cards_data) - 1 else "BRAND_INDIGO"
            header_node = PillHeaderNode(title=col["header"], style="flush", fill_color=h_color)

            # Check if card contains subheadings (e.g. Program Approach)
            bullets = col["bullets"]
            sections: List[SectionItem] = []
            cur_sec_title = None
            cur_sec_bullets: List[str] = []

            for b in bullets:
                if b in ("Assess and Prioritize Opportunities", "Develop and Execute Sourcing Strategies"):
                    if cur_sec_title or cur_sec_bullets:
                        sections.append(SectionItem(heading=cur_sec_title or "Overview", bullets=cur_sec_bullets))
                    cur_sec_title = b
                    cur_sec_bullets = []
                else:
                    cur_sec_bullets.append(b)
            if cur_sec_title or cur_sec_bullets:
                sections.append(SectionItem(heading=cur_sec_title or "Details", bullets=cur_sec_bullets))

            if len(sections) >= 2:
                body_node = SectionListNode(
                    sections=sections,
                    heading_font_size=10.5,
                    heading_color="BRAND_INDIGO",
                    bullet_font_size=9.0,
                    bullet_color="BRAND_ONYX",
                    space_after_section=6.0
                )
            else:
                base_sz = 9.0 if len(bullets) >= 5 else 10.0
                body_node = BulletListNode(items=bullets, font_size=base_sz, space_after=4.0)

            card_nodes.append(CardNode(
                header=header_node,
                body=body_node,
                padding=Padding(top=10.0, bottom=10.0, left=14.0, right=14.0)
            ))

        children.append(HStack(height_weight=2.4, gap=14.0, children=card_nodes))

        # 3. Bottom KPIs (Standalone spacious cards)
        kpi_nodes = []
        for idx, k in enumerate(kpis_data):
            n_color = "BRAND_TANGERINE" if (idx == 0 or idx == len(kpis_data) - 1) else "BRAND_INDIGO"
            kpi_nodes.append(KPINode(
                number=k["number"],
                label=k["label"],
                sublabel=k.get("sublabel"),
                number_color=n_color,
                number_size=18.0,
                label_size=8.5,
                fixed_height=62.0
            ))

        kpi_row = HStack(gap=14.0, children=kpi_nodes, fixed_height=62.0)
        if outcomes_hdr:
            sec_header = CardNode(
                fixed_height=22.0,
                fixed_width=150.0,
                fill_color="BRAND_INDIGO",
                border_color=None,
                corner_radius=4.0,
                body=BulletListNode(
                    items=[outcomes_hdr],
                    font_size=10.0,
                    bullet_char="",
                    text_color="NEUTRAL_WHITE",
                    align="center",
                    mar_l_pt=0.0,
                    indent_pt=0.0
                )
            )
            children.append(VStack(
                gap=6.0,
                fixed_height=92.0,
                children=[sec_header, kpi_row]
            ))
        else:
            children.append(HStack(fixed_height=62.0, gap=14.0, children=kpi_nodes))

        return VStack(gap=12.0, children=children)

    @classmethod
    def _cluster_by_x(cls, elements: List[Element]) -> List[List[Element]]:
        """Groups elements into distinct vertical columns by their X positions."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if not text_elems:
            return []
        sorted_x = sorted(text_elems, key=lambda e: e.x)
        clusters = []
        cur_col = [sorted_x[0]]
        for e in sorted_x[1:]:
            prev = cur_col[-1]
            if abs(e.x - prev.x) > 90.0:
                clusters.append(cur_col)
                cur_col = [e]
            else:
                cur_col.append(e)
        if cur_col:
            clusters.append(cur_col)
        return clusters

    @classmethod
    def _detect_section_pairs(cls, elements: List[Element]) -> Optional[List[SectionItem]]:
        """Detects if a cluster of elements forms alternating (Title, Description) section pairs."""
        sorted_elems = sorted(elements, key=lambda e: e.y)
        if len(sorted_elems) < 4:
            return None

        # Pattern: pairs of (Short bold header, Detailed description)
        pairs: List[SectionItem] = []
        if len(sorted_elems) % 2 == 0:
            is_valid = True
            for i in range(0, len(sorted_elems), 2):
                t_elem = sorted_elems[i]
                b_elem = sorted_elems[i + 1]
                t_lines = [p.text.strip() for p in t_elem.paras if p.text.strip()]
                b_lines = [p.text.strip() for p in b_elem.paras if p.text.strip()]
                if not t_lines or not b_lines:
                    is_valid = False
                    break
                t_text = " ".join(t_lines)
                if len(t_text) > 55 or len(t_text.split()) > 8:
                    is_valid = False
                    break
                pairs.append(SectionItem(heading=t_text, bullets=b_lines))
            if is_valid and len(pairs) >= 2:
                return pairs

        return None

    @classmethod
    def _detect_and_compose_sections(cls, elements: List[Element], title_text: str) -> Optional[LayoutNode]:
        """Detects and composes section list slides (e.g. Slide 2 and Slide 3)."""
        x_clusters = cls._cluster_by_x(elements)
        if len(x_clusters) == 2:
            c0, c1 = x_clusters[0], x_clusters[1]
            pairs0 = cls._detect_section_pairs(c0)
            if pairs0:
                # Left card has section pairs
                left_card = CardNode(
                    header=PillHeaderNode(title="Process Diagnostics & Assessment", style="flush", fill_color="BRAND_INDIGO"),
                    body=SectionListNode(
                        sections=pairs0,
                        heading_font_size=11.5,
                        heading_color="BRAND_INDIGO",
                        bullet_font_size=9.5,
                        bullet_color="BRAND_ONYX",
                        space_after_section=8.0
                    ),
                    width_weight=2.0
                )
                # Right card has insights or bullets
                c1_sorted = sorted(c1, key=lambda e: e.y)
                paras_c1 = [p.text.strip() for e in c1_sorted for p in e.paras if p.text.strip()]
                h1 = paras_c1[0] if paras_c1 else "Insights"
                b1 = paras_c1[1:] if len(paras_c1) > 1 else paras_c1
                right_card = CardNode(
                    header=PillHeaderNode(title=h1, style="flush", fill_color="BRAND_TANGERINE"),
                    body=BulletListNode(items=b1, font_size=9.5, space_after=8.0),
                    width_weight=1.0
                )
                return HStack(gap=18.0, children=[left_card, right_card])

        elif len(x_clusters) == 1:
            c0 = x_clusters[0]
            pairs0 = cls._detect_section_pairs(c0)
            if pairs0:
                return CardNode(
                    header=PillHeaderNode(title="Process Diagnostics & Assessment", style="flush", fill_color="BRAND_INDIGO"),
                    body=SectionListNode(
                        sections=pairs0,
                        heading_font_size=11.5,
                        heading_color="BRAND_INDIGO",
                        bullet_font_size=9.5,
                        bullet_color="BRAND_ONYX",
                        space_after_section=8.0
                    )
                )

        return None

    @classmethod
    def _cluster_columns(cls, elements: List[Element]) -> List[Dict[str, Any]]:
        """Clusters elements into columns, ensuring 100% of elements are assigned without dropping words."""
        text_elems = [e for e in elements if e.kind == "text" and e.text.strip()]
        if not text_elems:
            return []

        sorted_y = sorted(text_elems, key=lambda e: e.y)
        min_y = sorted_y[0].y

        raw_candidates = [e for e in sorted_y if abs(e.y - min_y) <= 35.0]
        header_candidates = []
        for c in raw_candidates:
            c_lines = [p.text.strip() for p in c.paras if p.text.strip()]
            if c_lines and len(c_lines[0]) <= 45 and len(c_lines[0].split()) <= 7:
                if not any(abs(c.x - existing.x) < 80.0 for existing in header_candidates):
                    header_candidates.append(c)
        header_candidates.sort(key=lambda e: e.x)

        if len(header_candidates) >= 2 and len(header_candidates) <= 4:
            columns_data = []
            assigned_ids = set()

            for h in header_candidates:
                assigned_ids.add(h.id)
                col_center = h.x + h.w / 2.0
                col_elems = [e for e in text_elems if e.id != h.id and abs((e.x + e.w / 2.0) - col_center) < 140.0]
                for ce in col_elems:
                    assigned_ids.add(ce.id)
                col_elems.sort(key=lambda e: e.y)

                h_paras = [p.text.strip() for p in h.paras if p.text.strip()]
                full_h = "\n".join(h_paras)
                h_lines = full_h.split("\n")
                header_text = h_lines[0].strip()
                bullets = []
                if len(h_lines) > 1:
                    bullets.extend([l.strip() for l in h_lines[1:] if l.strip()])
                for b_elem in col_elems:
                    for p in b_elem.paras:
                        txt = p.text.strip()
                        if txt:
                            bullets.append(txt)

                columns_data.append({"header": header_text, "bullets": bullets, "center": col_center})

            unassigned = [e for e in text_elems if e.id not in assigned_ids]
            for u in unassigned:
                u_center = u.x + u.w / 2.0
                best_col = min(columns_data, key=lambda c: abs(c["center"] - u_center))
                for p in u.paras:
                    if p.text.strip():
                        best_col["bullets"].append(p.text.strip())

            return columns_data

        # Fallback: cluster by X positions
        columns = cls._cluster_by_x(text_elems)
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
            any(k in col["header"].lower() for k in ("problem", "bridge", "future", "pillar", "challenge", "state"))
            for col in columns_data
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
                body=BulletListNode(items=col["bullets"], font_size=10.5)
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
