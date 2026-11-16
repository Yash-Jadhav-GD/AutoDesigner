"""
Test Composable Layout Engine on Slide 4
Builds Slide 4 (3 Library Cards + 4 Bottom KPIs) purely via Composable Primitives:
VStack(
    HStack(Card1, Card2, Card3),
    HStack(KPI1, KPI2, KPI3, KPI4)
)
"""

import sys
from pathlib import Path
import pptx
from pptx.util import Inches

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

from engine.composer.primitives import (
    SlideComposition,
    VStack,
    HStack,
    CardNode,
    PillHeaderNode,
    BulletListNode,
    SectionListNode,
    SectionItem,
    KPINode
)
from engine.composer.layout_solver import LayoutSolver
from engine.composer.renderer import ComposableRenderer
from engine.builders.common import load_library, get_decor_shadow_proto

# Define Slide 4 Composition
s4_comp = SlideComposition(
    title="Case Study 3: Opportunity Assessment and Sourcing Enablement",
    subtitle="Life-Safety, Construction and Industrial Services Organization",
    root=VStack(
        gap=14.0,
        children=[
            # Top row: 3 Opportunity Cards
            HStack(
                height_weight=3.2,
                gap=18.0,
                children=[
                    CardNode(
                        header=PillHeaderNode(title="Pain Points / Challenges", style="flush", fill_color="BRAND_INDIGO"),
                        body=BulletListNode(
                            items=[
                                "Decentralized sourcing practices across OpCos limited enterprise-wide leverage",
                                "Fragmented spend and supplier information constrained enterprise visibility",
                                "Predominantly reactive sourcing with limited forward-looking category pipelines",
                                "Inconsistent supplier arrangements and sourcing practices across the organization",
                                "Limited use of category intelligence and market insights in sourcing decisions",
                                "Recent acquisitions created targeted contract harmonization opportunities"
                            ],
                            font_size=10.5
                        )
                    ),
                    CardNode(
                        header=PillHeaderNode(title="Program Approach", style="flush", fill_color="BRAND_INDIGO"),
                        body=SectionListNode(
                            sections=[
                                SectionItem(
                                    heading="Assess and Prioritize Opportunities",
                                    bullets=[
                                        "Analyzed spend, supplier arrangements and category requirements",
                                        "Validated opportunities with procurement and OpCo stakeholders",
                                        "Prioritized initiatives based on value, feasibility and business readiness"
                                    ]
                                ),
                                SectionItem(
                                    heading="Develop and Execute Sourcing Strategies",
                                    bullets=[
                                        "Developed category-specific sourcing and negotiation strategies",
                                        "Mobilized priority initiatives with business stakeholders",
                                        "Tracked savings through negotiation and implementation"
                                    ]
                                )
                            ]
                        )
                    ),
                    CardNode(
                        header=PillHeaderNode(title="Key Opportunities", style="flush", fill_color="BRAND_TANGERINE"),
                        body=BulletListNode(
                            items=[
                                "IT contract harmonization: Reviewed supplier contracts and identified opportunities to move acquired businesses to enterprise agreements",
                                "Acquisition synergies: Compared commercial arrangements across the firm and identified pricing and consolidation opportunities post acquisitions",
                                "Fleet transformation: Developed a broader operating-model and transformation perspective for global fleet leveraging scale",
                                "Strategic sourcing pipeline: Established a prioritized portfolio of category initiatives aligned with business requirements",
                                "Identified and partnered with critical stakeholders to enable successful execution against identified opportunities"
                            ],
                            font_size=10.0
                        )
                    )
                ]
            ),
            # Bottom row: 4 KPI Cards
            HStack(
                height_weight=1.0,
                gap=14.0,
                children=[
                    KPINode(
                        number="5%-6%",
                        label="annual uplift in savings",
                        sublabel="14% - 15% savings via multi-year strategies",
                        number_color="BRAND_TANGERINE",
                        number_size=20.0
                    ),
                    KPINode(
                        number="Shift",
                        label="from Reactive to Proactive Sourcing",
                        number_color="BRAND_INDIGO",
                        number_size=20.0
                    ),
                    KPINode(
                        number="Prioritized",
                        label="Sourcing Pipeline Developed",
                        number_color="BRAND_INDIGO",
                        number_size=20.0
                    ),
                    KPINode(
                        number="~700",
                        label="Supplier Contracts Renewed",
                        number_color="BRAND_TANGERINE",
                        number_size=20.0
                    )
                ]
            )
        ]
    )
)

def run_test():
    solver = LayoutSolver()
    resolved_shapes, meta = solver.solve(s4_comp)
    print(f"LayoutSolver resolved {len(resolved_shapes)} shapes.")

    # Check for overlaps among resolved shapes
    overlaps = []
    for i in range(len(resolved_shapes)):
        for j in range(i + 1, len(resolved_shapes)):
            a = resolved_shapes[i].rect
            b = resolved_shapes[j].rect
            
            # Allow parent container to contain child
            is_contained_ab = (a.left <= b.left + 1 and a.right >= b.right - 1 and a.top <= b.top + 1 and a.bottom >= b.bottom - 1)
            is_contained_ba = (b.left <= a.left + 1 and b.right >= a.right - 1 and b.top <= a.top + 1 and b.bottom >= a.bottom - 1)
            if is_contained_ab or is_contained_ba:
                continue

            ox = max(0.0, min(a.right, b.right) - max(a.left, b.left))
            oy = max(0.0, min(a.bottom, b.bottom) - max(a.top, b.top))
            if ox > 1.0 and oy > 1.0:
                overlaps.append((resolved_shapes[i].shape_id, resolved_shapes[j].shape_id, ox, oy))

    print(f"Overlaps detected by LayoutSolver: {len(overlaps)}")
    if overlaps:
        for o in overlaps:
            print(f"  Overlap: {o}")
        sys.exit(1)

    # Render into Presentation
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    lib_prs, _, _, lib_lines = load_library()
    lib_shadow = get_decor_shadow_proto(lib_prs)

    ComposableRenderer.render(slide, resolved_shapes, lib_lines, lib_shadow)

    out_path = BASE_DIR / "output" / "test_composable_s4.pptx"
    prs.save(out_path)
    print(f"Successfully saved test composable deck: {out_path.name}")

if __name__ == "__main__":
    run_test()
