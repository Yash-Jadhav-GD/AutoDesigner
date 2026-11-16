"""
End-to-End Test for the Composable Layout Engine.
Takes any draft presentation, deconstructs every slide via slide_model.py,
composes it into a Composable Layout Tree via SemanticComposer,
solves geometry via LayoutSolver, and renders to PPTX via ComposableRenderer.
"""

import sys
from pathlib import Path
import pptx
from pptx.util import Inches

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

from engine.slide_model import read_slide
from engine.composer import SemanticComposer, LayoutSolver, ComposableRenderer
from engine.builders.common import load_library, get_decor_shadow_proto

def run_composer_deck(input_path: Path, output_path: Path):
    print(f"Reading input deck: {input_path.name}")
    in_prs = pptx.Presentation(input_path)
    out_prs = pptx.Presentation()
    out_prs.slide_width = in_prs.slide_width
    out_prs.slide_height = in_prs.slide_height

    lib_prs, _, _, lib_lines = load_library()
    lib_shadow = get_decor_shadow_proto(lib_prs)
    solver = LayoutSolver()

    total_shapes = 0
    for s_idx, slide in enumerate(in_prs.slides, start=1):
        model = read_slide(slide, s_idx, in_prs.slide_width, in_prs.slide_height)
        composition = SemanticComposer.compose(model)
        resolved_shapes, meta = solver.solve(composition)
        total_shapes += len(resolved_shapes)

        out_slide = out_prs.slides.add_slide(out_prs.slide_layouts[6])
        ComposableRenderer.render(out_slide, resolved_shapes, lib_lines, lib_shadow)
        print(f"  Slide {s_idx}: Composed into {composition.root.node_type} with {len(resolved_shapes)} resolved shapes.")

    out_prs.save(output_path)
    print(f"\nSuccessfully generated composable presentation: {output_path.name} ({len(out_prs.slides)} slides, {total_shapes} shapes)")

if __name__ == "__main__":
    sample_path = BASE_DIR / "input" / "sample_draft.pptx"
    out_path = BASE_DIR / "output" / "sample_draft_composable.pptx"
    run_composer_deck(sample_path, out_path)
