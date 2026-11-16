from pathlib import Path

path = Path("AutoDesigner/engine/builders/schemas.py")
text = path.read_text(encoding="utf-8")

old = """    "card_grid": CardSlide,
    "kpi_cards": CardSlide,
    "subheader_list": CardSlide,
}"""

new = """    "card_grid": CardSlide,
    "kpi_cards": CardSlide,
    "subheader_list": CardSlide,
    "data_box": CardSlide,
    "crd_icn_v1": CardSlide,
    "important_box": CardSlide,
    "imp_box_v1": CardSlide,
}"""

assert old in text, "old not found!"
text = text.replace(old, new, 1)

path.write_text(text, encoding="utf-8")
print("Successfully updated schemas.py!")
