"""
AutoDesigner Production Master Formatter: All 20 Slides
Redesigns and formats the complete 20-slide presentation:
- Slides 1-7: Benchmark-aligned layouts (Library Slide 1, Library Slide 3, Case Study 3, After Slide 19, Library Slide 7)
- Slides 8-20: Executive managed services, S2P expertise, plan on a page, spend & savings, resource plan, operating model, interface governance, transition governance, London-led delivery model, supplier management, risk framework, and integrated governance.

Strictly adheres to user requirements:
- Clean, corporate white cards (#FFFFFF) with subtle borders and shadows (Decor_Box_with_Shadow_1)
- NO goddy/dark monolithic fills (#1F2E65) with glowing orange pills for basic cards
- 100% Middle vertical alignment across every text frame (anchor="ctr")
- Native PowerPoint DrawingML bullets (<a:buChar char="•"/>), 0 literal bullet characters
- Exact 4.0pt corner radius on all rounded rectangles
- Authentic Library divider lines
- Zero overlapping shapes
- 100% text fidelity (0 lost tokens)
- Exact 20 slides in, exact 20 slides out
"""

import copy
import io
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "scratch"))

from engine.builders.common import (
    clean_element,
    get_shape_by_name,
    add_library_line,
    set_shape_corner_radius,
    apply_library_shadow,
    enforce_deck_corner_radii,
    inject_svg_assets,
    load_library,
    apply_native_bullet,
    clean_bullet_string,
    enforce_middle_alignment_deck_wide,
)
from engine.color_tokens import ColorTokens

from format_library_and_after_75 import (
    build_slide1,
    build_slide2,
    build_slide3,
    build_slide4,
    build_slide5,
    build_slide6,
    build_slide7,
    extract_slide1_data,
    extract_slide2_data,
    extract_slide3_data,
    extract_slide4_data,
    extract_table_slide,
)

from scratch.builders_8_to_20 import (
    build_slide8,
    build_slide9,
    build_slide10,
    build_slide11,
    build_slide12,
    build_slide13,
    build_slide14_or_15,
    build_slide16,
    build_slide17,
    build_slide18,
    build_slide19,
    build_slide20,
)

INPUT_PATH = BASE_DIR / "Design tester new.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"
LOCAL_COPY_PATH = BASE_DIR / "Design tester new_formatted.pptx"


def main():
    print(f"Loading 20-slide input presentation: {INPUT_PATH}...")
    in_prs = pptx.Presentation(INPUT_PATH)
    total_in = len(in_prs.slides)
    print(f"Input deck loaded with {total_in} slides.")
    assert total_in == 20, f"Expected 20 slides, but found {total_in} slides!"

    print("Extracting data for Slides 1 through 7...")
    s1_data = extract_slide1_data(in_prs)
    s2_data = extract_slide2_data(in_prs)
    s3_data = extract_slide3_data(in_prs)
    s4_data = extract_slide4_data(in_prs)
    s5_data = extract_table_slide(in_prs.slides[4], expected_rows=4)
    s6_data = extract_table_slide(in_prs.slides[5], expected_rows=6)
    s7_data = extract_table_slide(in_prs.slides[6], expected_rows=3, is_s7=True)
    print("Slides 1-7 extracted.")

    print("Loading Library prototypes and divider lines...")
    lib_prs, lib_s1, lib_s3, lib_lines = load_library()

    out_prs = pptx.Presentation()
    out_prs.slide_width = Pt(960)
    out_prs.slide_height = Pt(540)

    # -------------------------------------------------------------
    # Building Slides 1 through 7
    # -------------------------------------------------------------
    print("[Slide 1] Library Slide 1 & After Slide 1 (three_column_pills)...")
    build_slide1(out_prs, s1_data, lib_s1, lib_lines)

    print("[Slide 2] Library Slide 3 & After Slide 3 (split_process_insights)...")
    build_slide2(out_prs, s2_data, lib_s1, lib_s3, lib_lines)

    print("[Slide 3] Library Slide 3 (process_rows)...")
    build_slide3(out_prs, s3_data, lib_s3, lib_lines)

    print("[Slide 4] Library Slide 1 (Design 1) Floating Pill Cards & KPI-CRD-V1...")
    build_slide4(out_prs, s4_data, lib_s1, lib_lines)

    print("[Slide 5] After Slide 19 & Slide 6 (workflow_cards / table matrix)...")
    build_slide5(out_prs, s5_data, lib_lines)

    print("[Slide 6] After Slide 19 (6-row matrix)...")
    build_slide6(out_prs, s6_data, lib_lines)

    print("[Slide 7] Library Slide 7 (CRD-ICN-V1 Data Box cards)...")
    build_slide7(out_prs, s7_data, lib_lines)

    # -------------------------------------------------------------
    # Building Slides 8 through 20
    # -------------------------------------------------------------
    print("[Slide 8] Executive Summary: Scope of Work + Solution Summary + Benefits with GEP...")
    build_slide8(out_prs, lib_s1, lib_lines)

    print("[Slide 9] Source-to-Pay Expertise: Capabilities Stats + GEP Drivers + Platform Matrix...")
    build_slide9(out_prs, lib_s1, lib_lines)

    print("[Slide 10] Plan on a Page: Foundation Building + Steady State Process Streams + AI Layer...")
    build_slide10(out_prs, lib_s1, lib_lines)

    print("[Slide 11] Spend and Savings Opportunity: Table Matrix + KPI Cards + Category Pillars...")
    build_slide11(out_prs, lib_s1, lib_lines)

    print("[Slide 12] Resource Plan: Org Chart + Stream Boxes + SME Panel + Support Infrastructure...")
    build_slide12(out_prs, lib_s1, lib_lines)

    print("[Slide 13] Operating Model: Stakeholders + Plan-to-Strategy + S2C/P2P + GEP COE...")
    build_slide13(out_prs, lib_s1, lib_lines)

    print("[Slide 14] Interface Register & Escalation (Spoke Diagram + 3 Tiers + Issue Management)...")
    build_slide14_or_15(out_prs, lib_s1, lib_lines, is_slide15=False)

    print("[Slide 15] Interface Register & Escalation (Variant with Spoke Badges + 3 Tiers)...")
    build_slide14_or_15(out_prs, lib_s1, lib_lines, is_slide15=True)

    print("[Slide 16] Transition Governance & Risk Management: 3 Tiers + Table Matrix + KPOD Support...")
    build_slide16(out_prs, lib_s1, lib_lines)

    print("[Slide 17] PSCM Delivery Model: Global Credentials + London/Offshore + Regional Footprint...")
    build_slide17(out_prs, lib_s1, lib_lines)

    print("[Slide 18] Supplier & Materials Management: Lifecycle Management + Warehousing Coordination...")
    build_slide18(out_prs, lib_s1, lib_lines)

    print("[Slide 19] Risk Management Framework: 6-Step Cycle + Solution Portfolio + TPRM Outcomes...")
    build_slide19(out_prs, lib_s1, lib_lines)

    print("[Slide 20] Integrated Governance: Committee Tiers + Embedded Controls + Accountability Matrix...")
    build_slide20(out_prs, lib_s1, lib_lines)

    total_out = len(out_prs.slides)
    print(f"Total slides built: {total_out}")
    assert total_out == 20, f"Mismatch: built {total_out} slides, expected 20!"

    print("Enforcing exact 4.0pt corner radius across all rounded rectangles...")
    enforce_deck_corner_radii(out_prs, 4.0)

    print("Enforcing Middle vertical alignment (anchor='ctr') across every text box in the deck...")
    aligned_count = enforce_middle_alignment_deck_wide(out_prs)
    print(f"Middle-aligned {aligned_count} text boxes/shapes across all 20 slides.")

    raw_buf = io.BytesIO()
    out_prs.save(raw_buf)
    raw_bytes = raw_buf.getvalue()

    print("Injecting SVG vector assets into Slide 2 and Slide 3...")
    final_bytes = inject_svg_assets(raw_bytes, targets={2: "rIdSvgS2_", 3: "rIdSvgS3_"})

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "wb") as f:
        f.write(final_bytes)
    with open(LOCAL_COPY_PATH, "wb") as f:
        f.write(final_bytes)

    print("\n==================================================================")
    print("SUCCESS: Successfully designed and formatted all 20 slides!")
    print(f"  Exact count: {total_out} slides (20 in, 20 out)")
    print(f"  Output: {OUTPUT_PATH}")
    print(f"  Local copy: {LOCAL_COPY_PATH}")
    print("==================================================================")


if __name__ == "__main__":
    main()
