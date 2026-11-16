import json
from pathlib import Path

path = Path("AutoDesigner/templates/archetypes.json")
data = json.loads(path.read_text(encoding="utf-8"))

data["archetypes"]["crd_icn_v1"] = {
    "name": "CRD-ICN-V1 Data Box Design",
    "description": "Card with an icon circle container beside the header text, a horizontal divider, and body items below",
    "corner_radius": 0.04,
    "header_style": "header_icon",
    "header_font_size": 14.0,
    "body_font_size": 12.0,
    "divider_name": "Straight Connector 8",
    "include_dividers": True
}

data["archetypes"]["imp_box_v1"] = {
    "name": "IMP-BOX-V1 Important Box Design",
    "description": "High-contrast callout card with top header tab, white text, and white fading dividers between body items",
    "corner_radius": 0.04,
    "header_style": "top_tab",
    "header_font_size": 14.0,
    "body_font_size": 13.33,
    "divider_name": "White Side Fade R Divider",
    "include_dividers": True
}

path.write_text(json.dumps(data, indent=2), encoding="utf-8")
print("Successfully updated archetypes.json!")
