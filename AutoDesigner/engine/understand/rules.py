"""
Rules mapper: scores every archetype against a SlideModel and fills the winner's schema.

Each matcher returns None (does not apply) or (score, data, reasons). Scores are 0..1.
A matcher whose data does not cover every source word is penalised, and a generic
card_grid built from reading order is always available as a fallback, so the chosen
result never drops text when any candidate keeps it all.
"""

import re
from typing import Any, Dict, List, Optional, Tuple

from ..analyzer import ContentAnalyzer
from ..slide_model import (
    Element,
    SlideModel,
    columns_by_gap,
    coverage,
    rows_by_midpoint,
    split_by_largest_gap,
)
from ..builders.schemas import validate_data

Match = Optional[Tuple[float, Any, List[str]]]

KPI_VALUE_RE = re.compile(r"^[~<>≈+\-]?[\$€£¥]?\d[\d,.]*\s*(?:[–\-]\s*\d[\d,.]*)?\s*[%xXkKmMbB+]?$")
FTE_LABEL_RE = re.compile(r"^(FTEs?|people|roles?|heads?|members?)$", re.I)
FTE_VALUE_RE = re.compile(r"^\d+\s*FTEs?$", re.I)
RANGE_BADGE_RE = re.compile(r"^\d+\s*[-–]\s*\d+\b")
FALLBACK_THRESHOLD = 0.5


# ------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------

class Ctx:
    """Per-slide derived roles shared by matchers."""

    def __init__(self, m: SlideModel):
        self.m = m
        self.med = m.median_size() or 12.0
        self.title_el = m.title()
        self.title, self.subtitle = "", ""
        self.subtitle_el = None
        if self.title_el is not None:
            lines = [p.text for p in self.title_el.paras if p.text] or self.title_el.text.split("\n")
            self.title = lines[0] if lines else ""
            if len(lines) > 1 and self.title_el.placeholder:
                self.subtitle = " ".join(lines[1:])
            elif len(lines) > 1:
                self.title = self.title_el.text
            t = self.title_el
            below = [e for e in m.text_elements
                     if e is not t and e.kind == "text" and not e.bold
                     and t.y <= e.y < t.y + t.h + 20 and e.y < m.height * 0.2
                     and (e.size or 0) < (t.size or 99)]
            if below and not self.subtitle:
                self.subtitle_el = min(below, key=lambda e: e.y)
                self.subtitle = self.subtitle_el.text
        self.body = [e for e in m.text_elements if e is not self.title_el and e is not self.subtitle_el]

    def heading(self, e: Element) -> bool:
        return e.bold or bool(e.size and e.size > self.med * 1.1)

    def big(self, e: Element, factor: float = 1.6) -> bool:
        return bool(e.size and e.size >= self.med * factor)


def heading_pairs(ctx: Ctx, els: List[Element]) -> List[Dict[str, str]]:
    """Bold/larger element starts a pair; following plain elements form its body."""
    pairs: List[Dict[str, str]] = []
    for e in sorted(els, key=lambda e: (e.y, e.x)):
        if ctx.heading(e) or not pairs:
            pairs.append({"subheader": e.text if ctx.heading(e) else "", "body": "" if ctx.heading(e) else e.text})
        else:
            pairs[-1]["body"] = (pairs[-1]["body"] + "\n" + e.text).strip("\n")
    return pairs


def _cells(table: List[List[str]]) -> List[List[List[str]]]:
    return [[[ln for ln in c.split("\n") if ln.strip()] for c in row] for row in table]


def _nearest_above(target: Element, cands: List[Element]) -> Optional[Element]:
    above = [c for c in cands if c.y < target.y]
    return min(above, key=lambda c: (abs(c.cx - target.cx), target.y - c.y)) if above else None


# ------------------------------------------------------------------------------
# Matchers
# ------------------------------------------------------------------------------

def match_table(ctx: Ctx) -> Match:
    tables = [e for e in ctx.body if e.kind == "table"]
    if len(tables) != 1:
        return None
    t = tables[0]
    rest = [e for e in ctx.body if e is not t]
    note = [e for e in rest if e.y + e.h <= t.y + 5]
    footer = [e for e in rest if e not in note]
    reasons = [f"one table {len(t.table)}x{len(t.table[0]) if t.table else 0}"]
    if note or footer:
        data = {
            "title": ctx.title or None,
            "note": "\n".join(e.text for e in note) or None,
            "footer": "\n".join(e.text for e in footer) or None,
            "table_data": _cells(t.table),
        }
        return 0.95, ("allocation_table", data), reasons + ["note/footer text around table"]
    return 0.95, ("sizing_table", {"title": ctx.title or None, "table_data": _cells(t.table)}), reasons


def match_org(ctx: Ctx) -> Match:
    m = ctx.m
    ftes = sorted([e for e in ctx.body if FTE_VALUE_RE.match(e.text) and ctx.big(e, 1.8)], key=lambda e: e.x)
    badges = sorted([e for e in ctx.body if RANGE_BADGE_RE.match(e.text) and e.bold], key=lambda e: e.x)
    if len(ftes) < 2 or len(badges) < 2:
        return None
    used = set(id(e) for e in ftes + badges)
    bolds = [e for e in ctx.body if e.bold and id(e) not in used]

    # Banner: filled bold text between title and leadership row
    banner = next((e for e in sorted(ctx.body, key=lambda e: e.y)
                   if e.fill and e.y < ftes[0].y and id(e) not in used), None)
    if banner:
        used.add(id(banner))

    # Swatch legend: bold texts below everything else, each with a filled square to its left
    lowest_badge = max(b.y for b in badges)
    pools_bottom = max(e.y + e.h for e in ctx.body if e.y > lowest_badge and not e.bold) if ctx.body else 0
    legend = sorted([e for e in bolds if e.y >= pools_bottom - 2], key=lambda e: e.x)
    for e in legend:
        used.add(id(e))
    swatches = []
    for e in legend:
        sw = [s for s in m.elements if s.kind == "shape" and s.fill and abs(s.cy - e.cy) < 10 and s.x < e.x]
        swatches.append(max(sw, key=lambda s: s.x).fill if sw else None)

    leaders = []
    for f in ftes:
        head = _nearest_above(f, [b for b in bolds if id(b) not in used and b.y < f.y])
        used.add(id(f))
        if head is not None:
            used.add(id(head))
        lines = (head.text if head else "").split("\n")
        box = [s for s in m.elements if s.kind == "shape" and s.fill and s.contains(f)]
        box_fill = min(box, key=lambda s: s.w * s.h).fill if box else None
        kind = "span_breaker"
        if box_fill and box_fill in swatches:
            kind = "spend_value" if swatches.index(box_fill) == 0 else "span_breaker"
        leaders.append({"title": lines[0], "sub": "\n".join(lines[1:]) or None, "ftes": f.text, "type": kind})

    pools = []
    titles = []
    for b in badges:
        row = [e for e in bolds if id(e) not in used and abs(e.y - b.y) < 8 and e.x < b.x]
        t = max(row, key=lambda e: e.x) if row else None
        titles.append(t)
        if t is not None:
            used.add(id(t))
    for i, (b, t) in enumerate(zip(badges, titles)):
        x0 = (t or b).x - 5
        x1 = ((titles[i + 1] or badges[i + 1]).x - 5) if i + 1 < len(badges) else m.width + 1
        anchor_y = (t or b).y
        region = [e for e in ctx.body if id(e) not in used and x0 <= e.x < x1 and e.y > anchor_y]
        region.sort(key=lambda e: e.y)
        sub = region[0] if region and not region[0].bold else None
        cards = [e for e in region if e is not sub]
        rows = rows_by_midpoint(cards)
        pools.append({"title": t.text if t else "", "badge": b.text, "sub": sub.text if sub else "",
                      "items": [e.text for r in rows for e in r]})
        used.update(id(e) for e in region)
    if len(pools) != 3 or len(legend) < 2:
        return None
    data = {
        "title": ctx.title,
        "top_banner": banner.text if banner else "",
        "leadership": leaders,
        "pool1": {k: v for k, v in pools[0].items() if k != "items"} | {"cards": pools[0]["items"]},
        "pool2": {k: v for k, v in pools[1].items() if k != "items"} | {"regions": pools[1]["items"]},
        "pool3": {k: v for k, v in pools[2].items() if k != "items"} | {"cards": pools[2]["items"]},
        "legend": {"l1": legend[0].text, "l2": legend[1].text},
    }
    return 0.92, ("org_structure", data), [f"{len(ftes)} leader FTE figures", f"{len(badges)} pool range badges",
                                         "leader type from legend swatch colours"]


def match_persona(ctx: Ctx) -> Match:
    anchors = [e for e in ctx.body if e.text.isdigit() and ctx.big(e, 2.0)]
    labels = [e for e in ctx.body if FTE_LABEL_RE.match(e.text)]
    if len(anchors) < 4 or len(labels) < len(anchors) // 2:
        return None
    rows = rows_by_midpoint(anchors)
    row_ys = [r[0].y for r in rows]
    pitch = min((b - a for a, b in zip(row_ys, row_ys[1:])), default=200)
    ordered = [a for r in rows for a in r]
    cards = {id(a): [] for a in ordered}
    outside = []
    for e in ctx.body:
        if id(e) in cards:
            continue
        valid = [a for a in ordered if a.x - 5 <= e.x and a.y - 10 <= e.y < a.y + 0.85 * pitch]
        if valid:
            a = min(valid, key=lambda a: (e.x - a.x) + (e.y - a.y))
            cards[id(a)].append(e)
        else:
            outside.append(e)
    persona = []
    for a in ordered:
        items = cards[id(a)]
        label = next((e for e in items if FTE_LABEL_RE.match(e.text)), None)
        rest = [e for e in items if e is not label]
        name = next((e for e in rest if e.bold), None)
        regions = [e for e in rest if e is not name]
        persona.append({"count": a.text, "label": label.text if label else "",
                        "name": name.text if name else "", "regions": "\n".join(e.text for e in regions)})
    top = min(a.y for a in anchors)
    above = [e for e in outside if e.y < top]
    below = [e for e in outside if e.y >= top]
    data = {"title": ctx.title, "subtitle": ctx.subtitle,
            "lead_badge": "\n".join(e.text for e in above), "total_badge": "\n".join(e.text for e in below),
            "persona_cards": persona}
    return 0.9, ("persona_grid", data), [f"{len(anchors)} large count badges in {len(rows)} rows", "FTE labels"]


def match_workflow(ctx: Ctx) -> Match:
    letters = sorted([e for e in ctx.body if len(e.text) == 1 and e.text.isalpha() and e.text.isupper()
                      and (e.size or 0) > ctx.med], key=lambda e: e.y)
    if len(letters) < 2:
        return None
    groups = {id(l): [] for l in letters}
    for e in ctx.body:
        if id(e) in groups:
            continue
        nearest = min(letters, key=lambda l: abs(e.cy - l.cy))
        groups[id(nearest)].append(e)
    flows = []
    for l in letters:
        f = {"letter": l.text, "title": "", "desc": "", "in": "", "out": "", "autonomy": ""}
        for e in sorted(groups[id(l)], key=lambda e: (e.y, e.x)):
            t = e.text
            if re.match(r"^in\s*:", t, re.I):
                f["in"] = t
            elif re.match(r"^out\s*:", t, re.I):
                f["out"] = t
            elif "%" in t and len(t) <= 30:
                f["autonomy"] = t
            elif ctx.heading(e) and not f["title"]:
                f["title"] = t
            else:
                f["desc"] = (f["desc"] + "\n" + t).strip("\n")
        flows.append(f)
    return 0.9, ("workflow_cards", {"title": ctx.title, "subtitle": ctx.subtitle, "workflows": flows}), \
        [f"{len(letters)} letter badges starting rows", "In:/Out: lines"]


def match_assessment(ctx: Ctx) -> Match:
    split = split_by_largest_gap(ctx.body)
    if not split:
        return None
    left, right = split
    kpis = [e for e in right if ctx.big(e, 1.6) and KPI_VALUE_RE.match(e.text.replace(" ", ""))]
    if len(kpis) < 2:
        return None
    used = set(id(k) for k in kpis)
    rows = []
    for k in sorted(kpis, key=lambda e: e.y):
        beside = sorted([e for e in right if id(e) not in used and e.x > k.x + k.w - 5
                         and k.y - 8 <= e.y < k.y + k.h + 10], key=lambda e: e.y)
        label = next((e for e in beside if e.bold), None)
        sub = next((e for e in beside if e is not label), None)
        used.update(id(e) for e in (label, sub) if e)
        rows.append({"val": k.text, "label": label.text if label else "", "sub": sub.text if sub else ""})
    headers = sorted([e for e in right if id(e) not in used and e.bold], key=lambda e: e.y)
    if len(headers) < 2:
        return None
    first, second = headers[0], headers[1]
    deliv = [e for e in right if id(e) not in used and e not in headers and first.y < e.y < second.y]
    data = {"title": ctx.title, "subtitle": ctx.subtitle, "left_pairs": heading_pairs(ctx, left),
            "deliv_header": first.text, "deliv_items": [e.text for e in sorted(deliv, key=lambda e: e.y)],
            "insights_header": second.text, "kpi_rows": rows}
    return 0.88, ("assessment_kpis", data), ["left/right split", f"{len(kpis)} KPI figures in right panel",
                                            "two right-panel headers"]


def match_split_process(ctx: Ctx) -> Match:
    split = split_by_largest_gap(ctx.body)
    if not split:
        return None
    left, right = split
    pairs = heading_pairs(ctx, left)
    if len(pairs) < 2 or len(right) < 2 or not all(p["subheader"] for p in pairs):
        return None
    if sum(e.w for e in left) / len(left) < 1.3 * sum(e.w for e in right) / len(right):
        return None
    right = sorted(right, key=lambda e: e.y)
    return 0.8, ("split_process_insights", {"left_pairs": pairs, "right_header": right[0].text,
                                            "right_items": [e.text for e in right[1:]]}), \
        ["wide left column of heading/body pairs", "narrow right panel"]


def match_process_rows(ctx: Ctx) -> Match:
    if split_by_largest_gap(ctx.body) or len(ctx.body) < 4:
        return None
    pairs = heading_pairs(ctx, ctx.body)
    if len(pairs) < 2 or not all(p["subheader"] and p["body"] for p in pairs):
        return None
    return 0.75, ("process_rows", pairs), [f"{len(pairs)} full-width heading/body rows"]


def match_columns(ctx: Ctx) -> Match:
    if ctx.title_el is not None or any(e.kind == "table" for e in ctx.body):
        return None
    cols = columns_by_gap(ctx.body, 25)
    if not 2 <= len(cols) <= 4 or any(len(c) < 2 for c in cols):
        return None
    widths = [max(e.w for e in c) for c in cols]
    if max(widths) > 1.5 * min(widths):
        return None
    data = [{"header": c[0].text, "items": [e.text for e in c[1:]]} for c in cols]
    return 0.75, ("three_column_pills", data), [f"{len(cols)} equal-width columns, header on top of each"]


def _paragraph_cards(ctx: Ctx) -> Match:
    """Single text body with 'Header:' paragraphs (sample-style drafts)."""
    if len(ctx.body) != 1 or ctx.body[0].kind != "text":
        return None
    paras = [p.text for p in ctx.body[0].paras if p.text]
    headers = [i for i, p in enumerate(paras) if p.endswith(":")]
    if headers:
        cards = []
        for n, h in enumerate(headers):
            end = headers[n + 1] if n + 1 < len(headers) else len(paras)
            cards.append({"header": paras[h].rstrip(":"), "lines": paras[h + 1:end]})
        lead = paras[:headers[0]]
        if lead:
            cards.insert(0, {"header": "", "lines": lead})
        if all(c["lines"] and ContentAnalyzer.KPI_REGEX.match(c["lines"][0]) for c in cards):
            out = []
            for c in cards:
                mt = ContentAnalyzer.KPI_REGEX.match(c["lines"][0])
                out.append({"header": c["header"], "kpi_metric": mt.group(1),
                            "kpi_label": c["lines"][0][mt.end():].strip(),
                            "items": [{"subheader": "", "body": l} for l in c["lines"][1:]]})
            return 0.8, ("kpi_cards", {"title": ctx.title, "subtitle": ctx.subtitle, "cards": out}), \
                ["header paragraphs each leading with a metric"]
        out = [{"header": c["header"], "items": [{"subheader": "", "body": l} for l in c["lines"]]} for c in cards]
        return 0.7, ("card_grid", {"title": ctx.title, "subtitle": ctx.subtitle, "cards": out}), \
            [f"{len(cards)} 'Header:' paragraph groups"]
    items = []
    for p in paras:
        sub, _, body = p.partition(":")
        if body.strip() and len(sub) <= 40:
            items.append({"subheader": sub.strip(), "body": body.strip()})
        else:
            items.append({"subheader": "", "body": p})
    if sum(1 for i in items if i["subheader"]) >= max(2, len(items) * 0.75):
        return 0.75, ("subheader_list", {"title": ctx.title, "subtitle": ctx.subtitle,
                                         "cards": [{"header": ctx.title, "items": items}]}), \
            ["'Subheader: body' paragraphs"]
    return None


def fallback(ctx: Ctx) -> Match:
    """Generic card grid from layout columns; tables become rows. Keeps every word."""
    cards = []
    flat = []
    for e in ctx.body:
        if e.kind == "table":
            flat.append(("table", e))
    texts = [e for e in ctx.body if e.kind != "table"]
    cols = columns_by_gap(texts, 25) if texts else []
    if len(cols) > 4:
        cols = [texts]
    for c in cols:
        c = sorted(c, key=lambda e: (e.y, e.x))
        head = c[0] if len(c) > 1 and len(c[0].text) <= 60 else None
        lines = [p.text for e in c if e is not head for p in e.paras if p.text]
        cards.append({"header": head.text if head else "", "items": [{"subheader": "", "body": l} for l in lines]})
    for _, t in flat:
        cards.append({"header": "", "items": [{"subheader": "", "body": " | ".join(r)} for r in t.table]})
    if not cards:
        cards = [{"header": "", "items": []}]
    return FALLBACK_THRESHOLD - 0.2, ("card_grid", {"title": ctx.title, "subtitle": ctx.subtitle, "cards": cards}), \
        ["fallback: generic cards from reading order"]


IMPORTANT_KEYWORDS = re.compile(r"\b(important|key takeaways?|critical|note|notice|highlight|takeaways?)\b", re.I)


def match_important_box(ctx: Ctx) -> Match:
    """Matches slides that emphasize critical takeaways, important notes, or key highlights."""
    if not any(IMPORTANT_KEYWORDS.search(t) for t in (ctx.title, ctx.subtitle)):
        headings = [e.text for e in ctx.body if ctx.heading(e)]
        if not any(IMPORTANT_KEYWORDS.search(h) for h in headings):
            return None
    pairs = heading_pairs(ctx, ctx.body)
    if pairs and 1 <= len(pairs) <= 6:
        items = [{"subheader": p["subheader"], "body": p["body"]} for p in pairs if p["subheader"] or p["body"]]
        data = {"title": ctx.title, "subtitle": ctx.subtitle, "cards": [{"header": "Important", "items": items}]}
        return 0.78, ("important_box", data), ["Keywords indicate an important callout or key takeaways box"]
    return None


MATCHERS = [match_table, match_org, match_persona, match_workflow, match_assessment,
            match_split_process, match_process_rows, match_important_box, match_columns, _paragraph_cards, fallback]


def candidates(m: SlideModel) -> List[Dict[str, Any]]:
    ctx = Ctx(m)
    out = []
    for fn in MATCHERS:
        try:
            r = fn(ctx)
        except Exception as e:  # a matcher bug must never stop the deck
            r = None
            out.append({"archetype": None, "score": 0.0, "reasons": [f"{fn.__name__} failed: {e}"]})
        if not r:
            continue
        score, (arch, data), reasons = r
        try:
            data = validate_data(arch, data)
        except Exception as e:
            out.append({"archetype": arch, "score": 0.0, "reasons": reasons + [f"schema: {e.__class__.__name__}"]})
            continue
        cov = coverage(m, data)
        full = not cov["missing"]
        if not full:
            score *= 0.5
            reasons = reasons + [f"coverage {cov['coverage']:.3f}, missing {len(cov['missing'])} words: penalised"]
        out.append({"archetype": arch, "data": data, "score": round(score, 3), "reasons": reasons, "full": full})
    # Never drop content: any candidate that keeps every word outranks one that does not.
    return sorted([c for c in out if "data" in c], key=lambda c: (not c["full"], -c["score"]))


def map_slide(m: SlideModel) -> Dict[str, Any]:
    ranked = candidates(m)
    best = ranked[0]
    return {
        "archetype": best["archetype"],
        "data": best["data"],
        "confidence": best["score"],
        "reasons": best["reasons"],
        "alternatives": [{"archetype": c["archetype"], "score": c["score"]} for c in ranked[1:3]],
        "source": "rules",
    }
