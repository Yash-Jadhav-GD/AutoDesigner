"""
AutoDesigner PPT Designer Orchestrator
Analyzes a draft deck into a reviewable spec (engine.understand), builds the styled deck
from it (engine.builders) and verifies text fidelity (eval/verify_output.py).
"""

import json
import sys
from pathlib import Path
from typing import Optional, Dict, Any

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from engine.color_tokens import brand_palette
from engine.understand import analyze_deck, load
from engine.builders.deck import build_deck, save_spec
from eval.verify_output import verify as verify_output


class PPTDesigner:
    """Autonomous presentation designer that transforms raw PPTs into styled slide decks."""

    BASE_DIR = Path(__file__).resolve().parent
    PALETTES_PATH = BASE_DIR / "templates" / "palettes.json"
    ARCHETYPES_PATH = BASE_DIR / "templates" / "archetypes.json"

    def __init__(self):
        self.palettes = self._load_json(self.PALETTES_PATH).get("palettes", {})
        self.archetypes = self._load_json(self.ARCHETYPES_PATH).get("archetypes", {})
        self.palettes.setdefault("gep_brand", brand_palette())

    def _load_json(self, path: Path) -> Dict[str, Any]:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def get_palette(self, name: Optional[str] = None) -> Dict[str, Any]:
        if name and name in self.palettes:
            return self.palettes[name]
        return self.palettes.get("gep_purple", {})

    def check_file_accessible(self, file_path: str) -> bool:
        """Verifies if the file exists and is not locked by PowerPoint."""
        p = Path(file_path)
        if not p.exists():
            raise FileNotFoundError(f"Presentation not found at: {file_path}")
        try:
            with open(p, "rb") as f:
                f.read(1024)
            return True
        except PermissionError:
            raise PermissionError(
                f"The file '{p.name}' is currently open and locked by PowerPoint. "
                "Please save and close the file in PowerPoint or provide a copy."
            )

    def design_presentation(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        mapper: str = "auto",
        palette_name: str = "gep_purple",
        spec_only: bool = False,
        icons: bool = False,
        log=print,
    ) -> Dict[str, Any]:
        """Draft deck -> slide model -> spec (output/specs/<deck>.json) -> styled deck -> fidelity check.
        specs/ holds hand-checked gold specs and is never written by a run."""
        self.check_file_accessible(input_path)
        stem = Path(input_path).stem.replace(" ", "_")

        spec = analyze_deck(load(input_path), mapper=mapper, log=log)
        spec_path = self.BASE_DIR / "output" / "specs" / f"{stem}.json"
        save_spec(spec, str(spec_path))
        result: Dict[str, Any] = {
            "input_file": str(input_path),
            "spec_file": str(spec_path),
            "mapper": spec["mapper"],
            "slides": spec["slides"],
        }
        if spec_only:
            return result

        if not output_path:
            output_dir = self.BASE_DIR / "output"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = str(output_dir / f"{stem}_designed.pptx")
        Path(output_path).write_bytes(build_deck(spec, palette=palette_name, icons=icons))
        result["output_file"] = str(output_path)
        result["fidelity"] = verify_output(input_path, output_path)
        return result

    def design_with_composer(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        log=print,
    ) -> Dict[str, Any]:
        """Designs any presentation using the Composable Layout Engine (primitives, box-model solver, DrawingML renderer)."""
        import pptx
        from engine.composer import SemanticComposer, LayoutSolver, ComposableRenderer
        from engine.builders.common import load_library, get_decor_shadow_proto
        from engine.slide_model import read_slide

        self.check_file_accessible(input_path)
        stem = Path(input_path).stem.replace(" ", "_")

        if not output_path:
            output_dir = self.BASE_DIR / "output"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = str(output_dir / f"{stem}_composable.pptx")

        in_prs = pptx.Presentation(input_path)
        out_prs = pptx.Presentation()
        out_prs.slide_width = in_prs.slide_width
        out_prs.slide_height = in_prs.slide_height

        lib_prs, _, _, lib_lines = load_library()
        lib_shadow = get_decor_shadow_proto(lib_prs)
        solver = LayoutSolver()

        slides_info = []
        for s_idx, slide in enumerate(in_prs.slides, start=1):
            model = read_slide(slide, s_idx, in_prs.slide_width, in_prs.slide_height)
            composition = SemanticComposer.compose(model)
            resolved_shapes, meta = solver.solve(composition)
            out_slide = out_prs.slides.add_slide(out_prs.slide_layouts[6])
            ComposableRenderer.render(out_slide, resolved_shapes, lib_lines, lib_shadow)
            slides_info.append({
                "slide": s_idx,
                "node_type": composition.root.node_type,
                "shapes": len(resolved_shapes)
            })

        out_prs.save(output_path)
        return {
            "input_file": str(input_path),
            "output_file": str(output_path),
            "engine": "composable",
            "slides": slides_info,
            "fidelity": verify_output(input_path, output_path)
        }
