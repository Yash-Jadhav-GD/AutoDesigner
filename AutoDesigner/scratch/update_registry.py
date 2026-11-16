from pathlib import Path

path = Path("AutoDesigner/engine/builders/registry.py")
text = path.read_text(encoding="utf-8")

old_arch = """    # Generic archetypes rendered by LayoutEngine + SlideStyler from Library prototypes
    "card_grid":               lambda prs, d, lib: build_card_slide(prs, d, "card_grid", lib),
    "kpi_cards":               lambda prs, d, lib: build_card_slide(prs, d, "kpi_cards", lib),
    "subheader_list":          lambda prs, d, lib: build_card_slide(prs, d, "subheader_list", lib),
}"""

new_arch = """    # Generic archetypes rendered by LayoutEngine + SlideStyler from Library prototypes
    "card_grid":               lambda prs, d, lib: build_card_slide(prs, d, "card_grid", lib),
    "kpi_cards":               lambda prs, d, lib: build_card_slide(prs, d, "kpi_cards", lib),
    "subheader_list":          lambda prs, d, lib: build_card_slide(prs, d, "subheader_list", lib),
    "data_box":                lambda prs, d, lib: build_card_slide(prs, d, "data_box", lib),
    "crd_icn_v1":              lambda prs, d, lib: build_card_slide(prs, d, "crd_icn_v1", lib),
    "important_box":           lambda prs, d, lib: build_card_slide(prs, d, "important_box", lib),
    "imp_box_v1":              lambda prs, d, lib: build_card_slide(prs, d, "imp_box_v1", lib),
}"""

old_desc = """    "card_grid": "Generic: title/subtitle and 1-4 cards, each a header over items (optional subheader + body).",
    "kpi_cards": "Generic: title/subtitle and cards that each lead with a metric (kpi_metric + kpi_label), optional items.",
    "subheader_list": "Generic: title/subtitle and one card listing subheader + body items.",
}"""

new_desc = """    "card_grid": "Generic: title/subtitle and 1-4 cards, each a header over items (optional subheader + body).",
    "kpi_cards": "Generic: title/subtitle and cards that each lead with a metric (kpi_metric + kpi_label), optional items.",
    "subheader_list": "Generic: title/subtitle and one card listing subheader + body items.",
    "data_box": "Generic: title/subtitle and 1-4 data cards with an icon circle in the card header, divider, and body text (CRD-ICN-V1).",
    "crd_icn_v1": "Generic: title/subtitle and data cards with an icon circle in the card header (CRD-ICN-V1).",
    "important_box": "Generic: title/subtitle and high-contrast callout card with top header tab and white fading dividers (IMP-BOX-V1).",
    "imp_box_v1": "Generic: title/subtitle and high-contrast callout card with top header tab and white fading dividers (IMP-BOX-V1).",
}"""

assert old_arch in text, "old_arch not found!"
assert old_desc in text, "old_desc not found!"

text = text.replace(old_arch, new_arch, 1)
text = text.replace(old_desc, new_desc, 1)

path.write_text(text, encoding="utf-8")
print("Successfully updated registry.py!")
