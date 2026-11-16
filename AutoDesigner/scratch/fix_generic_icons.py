from pathlib import Path

path = Path("AutoDesigner/engine/builders/generic.py")
text = path.read_text(encoding="utf-8")

old = """    slide = prs.slides.add_slide(prs.slide_layouts[6])
    icons = []
    SlideStyler.apply_style(slide=slide, slide_data=slide_data, layout=layout,
                            palette=lib.palette, icons=icons)"""

new = """    slide = prs.slides.add_slide(prs.slide_layouts[6])
    icons = [] if lib.icons else None
    SlideStyler.apply_style(slide=slide, slide_data=slide_data, layout=layout,
                            palette=lib.palette, icons=icons)"""

assert old in text, "old not found!"
text = text.replace(old, new, 1)

path.write_text(text, encoding="utf-8")
print("Successfully updated generic.py!")
