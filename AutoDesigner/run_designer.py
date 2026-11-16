"""
AutoDesigner CLI Runner
Run automated PPT design transformations on presentations.

Usage:
    python run_designer.py --create-sample
    python run_designer.py --input input/my_presentation.pptx --palette gep_purple
    python run_designer.py --input input/my_presentation.pptx --mapper rules --spec-only
"""

import sys
import argparse
from pathlib import Path
import pptx
from pptx.util import Inches, Pt

# Ensure local AutoDesigner directory is in path
CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from designer import PPTDesigner


def create_sample_presentation(sample_path: Path):
    """Creates a sample unformatted draft presentation to demonstrate transformation."""
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    title_and_content_layout = prs.slide_layouts[1]

    # Slide 1: 3-column Strategic Pillars
    s1 = prs.slides.add_slide(title_and_content_layout)
    s1.shapes.title.text = "Digital Transformation Strategic Pillars\nFY2026 Core Execution Priorities"
    tf1 = s1.placeholders[1].text_frame
    tf1.text = "Cloud Acceleration:"
    p = tf1.add_paragraph()
    p.text = "Migrate 80% legacy workloads to managed Kubernetes"
    p = tf1.add_paragraph()
    p.text = "Establish multi-region disaster recovery automation"
    p = tf1.add_paragraph()
    p.text = "Reduce infrastructure overhead by 35%"
    p = tf1.add_paragraph()
    p.text = "Data & AI Modernization:"
    p = tf1.add_paragraph()
    p.text = "Deploy unified enterprise data lakehouse"
    p = tf1.add_paragraph()
    p.text = "Embed LLM workflow copilot across operations"
    p = tf1.add_paragraph()
    p.text = "Achieve real-time executive telemetry"
    p = tf1.add_paragraph()
    p.text = "Cyber Resilience:"
    p = tf1.add_paragraph()
    p.text = "Implement Zero Trust identity boundary across all endpoints"
    p = tf1.add_paragraph()
    p.text = "Automated threat intelligence with continuous compliance"

    # Slide 2: KPI Metrics Slide
    s2 = prs.slides.add_slide(title_and_content_layout)
    s2.shapes.title.text = "Enterprise Performance Metrics\nQuarterly Operational Milestones"
    tf2 = s2.placeholders[1].text_frame
    tf2.text = "Cost Efficiency:"
    p = tf2.add_paragraph()
    p.text = "$4.2M Annualized Savings from automated supplier rationalization"
    p = tf2.add_paragraph()
    p.text = "Cycle Time Reduction:"
    p = tf2.add_paragraph()
    p.text = "68% Faster Sourcing cycle times enabled by digital negotiations"
    p = tf2.add_paragraph()
    p.text = "Adoption Velocity:"
    p = tf2.add_paragraph()
    p.text = "94.8% Active Platform Engagement across global business units"

    # Slide 3: Hierarchical List with Concepts & Definitions
    s3 = prs.slides.add_slide(title_and_content_layout)
    s3.shapes.title.text = "Solution Architecture Blueprint"
    tf3 = s3.placeholders[1].text_frame
    tf3.text = "Orchestration Layer: Event-driven microservices coordinating distributed requests"
    p = tf3.add_paragraph()
    p.text = "Security Boundary: End-to-end TLS 1.3 encryption with hardware security modules"
    p = tf3.add_paragraph()
    p.text = "Analytics Pipeline: Real-time streaming telemetry with automated anomaly detection"
    p = tf3.add_paragraph()
    p.text = "Scalability Engine: Auto-scaling clusters dynamically handling traffic surges"

    sample_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(sample_path))
    print(f"Sample draft presentation created at: {sample_path}")


def main():
    parser = argparse.ArgumentParser(description="AutoDesigner PPT Designer")
    parser.add_argument("--input", "-i", type=str, help="Path to input presentation (.pptx)")
    parser.add_argument("--output", "-o", type=str, help="Path to output presentation (.pptx)")
    parser.add_argument(
        "--mapper",
        "-m",
        type=str,
        default="auto",
        choices=["auto", "rules", "llm", "baseline"],
        help="How slides are understood: auto = Claude when credentials exist, else rules",
    )
    parser.add_argument(
        "--spec-only",
        action="store_true",
        help="Write output/specs/<deck>.json for review and stop (build later with build_from_spec.py)",
    )
    parser.add_argument(
        "--palette",
        "-p",
        type=str,
        default="gep_purple",
        choices=["gep_purple", "gep_brand", "corporate_blue", "emerald_slate", "slate_onyx"],
        help="Executive color palette",
    )
    parser.add_argument(
        "--icons",
        action="store_true",
        help="Add keyword-matched SVG icons to generic card_grid rows (engine/svg_icons.py)",
    )
    parser.add_argument(
        "--create-sample",
        action="store_true",
        help="Generate a draft presentation in input/ and design it immediately",
    )

    args = parser.parse_args()
    designer = PPTDesigner()

    input_file = args.input
    if args.create_sample:
        sample_path = CURRENT_DIR / "input" / "sample_draft.pptx"
        create_sample_presentation(sample_path)
        input_file = str(sample_path)

    if not input_file:
        # Check if there are any files in input/
        input_dir = CURRENT_DIR / "input"
        pptx_files = list(input_dir.glob("*.pptx"))
        if pptx_files:
            input_file = str(pptx_files[0])
            print(f"Found presentation in input/: {pptx_files[0].name}")
        else:
            print("No input presentation provided.")
            print("Tip: Drop a presentation into AutoDesigner/input/ or run:")
            print("   python run_designer.py --create-sample")
            sys.exit(1)

    print(f"\n[PPT Designer] Designing: {input_file}")
    print(f"[PPT Designer] Mapper: {args.mapper} | Palette: {args.palette}")

    result = designer.design_presentation(
        input_path=input_file,
        output_path=args.output,
        mapper=args.mapper,
        palette_name=args.palette,
        spec_only=args.spec_only,
        icons=args.icons,
    )

    print(f"\nSpec: {result['spec_file']}  (mapper used: {result['mapper']})")
    for i, s in enumerate(result["slides"], start=1):
        flags = []
        if (s.get("confidence") or 0) < 0.6:
            flags.append("LOW CONFIDENCE - review spec")
        if s.get("coverage", 1) < 1:
            flags.append(f"MISSING TEXT: {' '.join(s.get('missing_words', [])[:8])}")
        print(f" - Slide {i:>2}: {s['archetype']:<22} conf={s.get('confidence', 0):.2f} "
              f"[{s.get('source', '?')}] {s['reasons'][0] if s.get('reasons') else ''}"
              + (f"   <- {'; '.join(flags)}" if flags else ""))

    if args.spec_only:
        print("\nEdit the spec if needed, then build with:")
        print(f'   python build_from_spec.py --spec "{result["spec_file"]}" --output output/deck.pptx')
        return

    fid = result["fidelity"]
    print(f"\nOutput File: {result['output_file']}")
    print(f"Fidelity: slides in={fid['slides_in']} out={fid['slides_out']}, "
          + ("all text preserved, no dangling relationships" if fid["ok"] else "FAILED"))
    if not fid["ok"]:
        for s in fid["slides"]:
            if s["lost"] or s["dangling_rels"]:
                print(f"   slide {s['slide']}: lost {s['lost'][:10]} dangling {s['dangling_rels']}")
        sys.exit(2)


if __name__ == "__main__":
    main()
