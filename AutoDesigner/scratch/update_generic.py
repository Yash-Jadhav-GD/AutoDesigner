from pathlib import Path

path = Path("AutoDesigner/engine/builders/generic.py")
text = path.read_text(encoding="utf-8")

old = """STYLER_ARCHETYPE = {
    "card_grid": "design1_floating_pill",
    "kpi_cards": "kpi_metric_card",
    "subheader_list": "crd_sub_lst",
}"""

new = """STYLER_ARCHETYPE = {
    "card_grid": "design1_floating_pill",
    "kpi_cards": "kpi_metric_card",
    "subheader_list": "crd_sub_lst",
    "data_box": "crd_icn_v1",
    "crd_icn_v1": "crd_icn_v1",
    "important_box": "imp_box_v1",
    "imp_box_v1": "imp_box_v1",
}"""

assert old in text, "old not found!"
text = text.replace(old, new, 1)
path.write_text(text, encoding="utf-8")
print("Successfully updated generic.py!")
