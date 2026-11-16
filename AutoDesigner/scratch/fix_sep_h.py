from pathlib import Path

path = Path("AutoDesigner/engine/layout_engine.py")
text = path.read_text(encoding="utf-8")

old = """        sep_h = divider.height if has_row_divs else 0.0"""

new = """        # For line prototypes that are horizontal or 90-deg rotated in XML, thickness is ~1.0 pt
        div_thick = min(divider.width, divider.height) if min(divider.width, divider.height) > 0 else 1.0
        sep_h = div_thick if has_row_divs else 0.0"""

assert old in text, "old not found!"
text = text.replace(old, new, 1)

path.write_text(text, encoding="utf-8")
print("Successfully updated layout_engine.py sep_h!")
