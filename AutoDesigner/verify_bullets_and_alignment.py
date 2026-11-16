import sys
from pathlib import Path
from pptx import Presentation

PPTX_PATH = Path(__file__).resolve().parent / "output" / "Design_tester_new_formatted.pptx"


def verify():
    prs = Presentation(PPTX_PATH)
    print(f"Loaded {PPTX_PATH.name} with {len(prs.slides)} slides.")

    total_text_boxes = 0
    non_middle_boxes = []
    literal_bullets_found = []
    native_bullet_count = 0
    bad_bullet_schema = []
    small_fonts = []

    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    for s_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            total_text_boxes += 1
            bodyPr = shape.text_frame._txBody.bodyPr
            anchor = bodyPr.get("anchor")
            if anchor != "ctr":
                non_middle_boxes.append((s_idx, shape.name, anchor))

            for p_idx, p in enumerate(shape.text_frame.paragraphs):
                txt = p.text or ""
                if "\u2022" in txt or "•" in txt:
                    literal_bullets_found.append((s_idx, shape.name, txt[:40]))

                pPr = p._p.find(f"{ns_a}pPr")
                if pPr is not None:
                    bu_char = pPr.find(f"{ns_a}buChar")
                    if bu_char is not None:
                        native_bullet_count += 1
                        marL = pPr.get("marL")
                        indent = pPr.get("indent")
                        bu_font = pPr.find(f"{ns_a}buFont")
                        if not marL or not indent or int(indent) >= 0 or bu_font is None:
                            bad_bullet_schema.append((s_idx, shape.name, marL, indent))
                        # Verify schema order: buFont before buChar before defRPr
                        tags = [c.tag.replace(ns_a, "") for c in pPr]
                        if "buFont" in tags and "buChar" in tags:
                            if tags.index("buFont") > tags.index("buChar"):
                                bad_bullet_schema.append((s_idx, shape.name, "buFont after buChar", tags))
                        if "defRPr" in tags and "buChar" in tags:
                            if tags.index("buChar") > tags.index("defRPr"):
                                bad_bullet_schema.append((s_idx, shape.name, "buChar after defRPr", tags))

                p_sz = p.font.size.pt if p.font.size else None
                for r in p.runs:
                    r_sz = r.font.size.pt if r.font.size else p_sz
                    if r_sz is not None and r_sz < 9.0:
                        small_fonts.append((s_idx, shape.name, r.text[:25], r_sz))

    print("\n--- Verification Report ---")
    print(f"Total text boxes checked: {total_text_boxes}")
    print(f"Non-Middle aligned text boxes: {len(non_middle_boxes)}")
    if non_middle_boxes:
        print("  Violations:", non_middle_boxes[:10])

    print(f"Literal '•' bullet characters in text runs: {len(literal_bullets_found)}")
    if literal_bullets_found:
        print("  Violations:", literal_bullets_found[:10])

    print(f"Native PowerPoint bullet paragraphs: {native_bullet_count}")
    print(f"Bullet DrawingML schema/indent violations: {len(bad_bullet_schema)}")
    if bad_bullet_schema:
        print("  Violations:", bad_bullet_schema[:10])

    print(f"Font sizes < 9.0pt: {len(small_fonts)}")
    if small_fonts:
        print("  Violations:", small_fonts[:10])

    all_ok = (
        len(non_middle_boxes) == 0
        and len(literal_bullets_found) == 0
        and native_bullet_count > 0
        and len(bad_bullet_schema) == 0
        and len(small_fonts) == 0
    )
    if all_ok:
        print("\nALL BULLET, ALIGNMENT, AND FONT CHECKS PASSED 100%!")
        sys.exit(0)
    else:
        print("\nVERIFICATION FAILED!")
        sys.exit(1)


if __name__ == "__main__":
    verify()
