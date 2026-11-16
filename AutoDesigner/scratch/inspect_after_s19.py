import pptx
prs = pptx.Presentation("AutoDesigner/After_copy.pptx")
s19 = prs.slides[18]
print(f"=== After_copy.pptx Slide 19: {len(s19.shapes)} shapes ===")
for sh in s19.shapes:
    t = sh.text.strip().replace("\n", " ") if sh.has_text_frame else ""
    line_info = ""
    if sh._element.tag.endswith("cxnSp") or (hasattr(sh, "line") and sh.line.color.type is not None):
        line_info = " [LINE]"
    print(f"[{sh.name}]{line_info} ({sh.width.pt:.1f}x{sh.height.pt:.1f} at {sh.left.pt:.1f},{sh.top.pt:.1f}): {t[:70]}")
