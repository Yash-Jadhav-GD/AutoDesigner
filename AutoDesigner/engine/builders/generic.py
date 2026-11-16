"""
Generic card archetypes, rendered by LayoutEngine + SlideStyler from the Library_copy.pptx
prototype shapes (AUTOFORMAT_SPECIFICATION.md sections 3, 5 and 7).
Data shape: schemas.CardSlide.
"""

import json
from pathlib import Path

from ..analyzer import SlideData, CardData, ItemData
from ..color_tokens import brand_palette
from ..layout_engine import LayoutEngine, Prototypes
from ..styler import SlideStyler
from .common import load_library

PALETTES_PATH = Path(__file__).resolve().parents[2] / "templates" / "palettes.json"

# Spec archetype -> styler archetype (layout_engine.DESIGNS)
STYLER_ARCHETYPE = {
    "card_grid": "design1_floating_pill",
    "kpi_cards": "kpi_metric_card",
    "subheader_list": "crd_sub_lst",
    "data_box": "crd_icn_v1",
    "crd_icn_v1": "crd_icn_v1",
    "important_box": "imp_box_v1",
    "imp_box_v1": "imp_box_v1",
}
# card_grid rows get keyword-matched SVG icons on Library backplates when built with icons=True
ICON_ARCHETYPE = "icon_badge_card"


def load_palette(name: str = "gep_purple") -> dict:
    if name == "gep_brand":
        return brand_palette()
    palettes = json.loads(PALETTES_PATH.read_text(encoding="utf-8")).get("palettes", {})
    return palettes.get(name) or palettes["gep_purple"]


def _slide_data(data: dict) -> SlideData:
    cards = [
        CardData(
            header=c["header"],
            items=[ItemData(subheader=i.get("subheader", ""), body=i.get("body", "")) for i in c.get("items", [])],
            kpi_metric=c.get("kpi_metric"),
            kpi_label=c.get("kpi_label"),
        )
        for c in data["cards"]
    ]
    return SlideData(title=data.get("title", ""), subtitle=data.get("subtitle", ""), cards=cards)


def build_card_slide(prs, data: dict, archetype: str, lib):
    """`lib` is a registry.LibraryAssets; icons placed on the slide are recorded in lib.svg_icons."""
    slide_data = _slide_data(data)
    styler_archetype = ICON_ARCHETYPE if lib.icons and archetype == "card_grid" else STYLER_ARCHETYPE[archetype]
    prototypes = Prototypes(lib.prs if lib.prs is not None else load_library()[0])
    layout = LayoutEngine.compute_slide_layout(slide_data, styler_archetype, prototypes)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    icons = [] if lib.icons else None
    SlideStyler.apply_style(slide=slide, slide_data=slide_data, layout=layout,
                            palette=lib.palette, icons=icons)
    if icons:
        lib.svg_icons[len(prs.slides)] = icons
    return slide
