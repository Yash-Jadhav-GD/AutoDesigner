import sys
from pathlib import Path

# Setup paths
BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import pptx
from pptx.util import Inches, Pt
from engine.builders.registry import ARCHETYPES, LibraryAssets
from engine.builders.common import load_library, enforce_deck_corner_radii

print("Loading library presentation...")
# Avoid locked file by using inspect copy
lib_prs, lib_s1, lib_s3, lib_lines = load_library()
print(f"Library slides loaded: {len(lib_prs.slides)}")

lib_assets = LibraryAssets(
    s1=lib_s1,
    s3=lib_s3,
    lines=lib_lines,
    prs=lib_prs,
    icons=False,
)

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Test Slide 1: data_box (crd_icn_v1)
data_box_data = {
    "title": "Procurement Strategy & Market Insights",
    "subtitle": "Overview of strategic sourcing pillars and operational intelligence",
    "cards": [
        {
            "header": "Category Management",
            "items": [
                {"subheader": "Cross-Functional Sourcing", "body": "Aligning direct & indirect spend priorities across business units."},
                {"subheader": "Supplier Performance", "body": "Quarterly business reviews and SLA governance metrics."},
                {"subheader": "Contract Optimization", "body": "Consolidation of expiring agreements for volume discounts."}
            ]
        },
        {
            "header": "Market Intelligence",
            "items": [
                {"subheader": "Cost Driver Analysis", "body": "Real-time index monitoring across raw materials and logistics."},
                {"subheader": "Supply Risk Mitigation", "body": "Dual-sourcing strategies for single-origin high-risk components."},
                {"subheader": "Inflation Hedging", "body": "Indexed tiered pricing models to reduce spot market volatility."}
            ]
        },
        {
            "header": "Digital Enablement",
            "items": [
                {"subheader": "Automated PO Processing", "body": "90% straight-through processing for standardized catalog requisitions."},
                {"subheader": "AI Spend Analytics", "body": "Autonomous tail spend classification and anomaly detection."},
                {"subheader": "Supplier Portal Integration", "body": "Seamless punchout and e-invoicing workflows across ERPs."}
            ]
        }
    ]
}

print("Building Slide 1: data_box (crd_icn_v1)...")
ARCHETYPES["data_box"](prs, data_box_data, lib_assets)

# Test Slide 2: important_box (imp_box_v1)
important_box_data = {
    "title": "Program Milestones & Governance",
    "subtitle": "Critical operational guardrails and mandatory compliance notices",
    "cards": [
        {
            "header": "Critical Requirements",
            "items": [
                {"subheader": "", "body": "All cross-border data transfers must comply with GDPR & local data sovereignty laws."},
                {"subheader": "", "body": "Dual approval is mandatory for any purchase requisition exceeding $100,000 threshold."},
                {"subheader": "", "body": "Supplier security audits must be renewed annually before contract extension."},
                {"subheader": "", "body": "Escalation to steering committee is required within 24 hours of any delivery breach."},
                {"subheader": "", "body": "All invoice reconciliations must settle within standard Net-30 payment terms."}
            ]
        }
    ]
}

print("Building Slide 2: important_box (imp_box_v1)...")
ARCHETYPES["important_box"](prs, important_box_data, lib_assets)

# Also test important_box with 2 cards
important_box_data_2 = {
    "title": "Key Takeaways & Immediate Next Steps",
    "subtitle": "Summary of executive findings and strategic execution roadmap",
    "cards": [
        {
            "header": "Executive Takeaways",
            "items": [
                {"subheader": "", "body": "Baseline operating costs have decreased by 18% following regional consolidation."},
                {"subheader": "", "body": "Supplier lead times have stabilized across tier-1 manufacturing hubs."},
                {"subheader": "", "body": "High compliance adherence observed across EMEA and APAC business units."}
            ]
        },
        {
            "header": "Immediate Next Steps",
            "items": [
                {"subheader": "", "body": "Finalize RFP documentation and circulate to shortlisted vendor panel."},
                {"subheader": "", "body": "Initiate legal review on master services agreement and data terms."},
                {"subheader": "", "body": "Schedule kickoff alignment meeting with regional steering sponsors."}
            ]
        }
    ]
}

print("Building Slide 3: important_box (2 cards)...")
ARCHETYPES["important_box"](prs, important_box_data_2, lib_assets)

# Enforce corner radii
enforce_deck_corner_radii(prs, 4.0)

out_path = BASE_DIR / "output" / "test_new_designs.pptx"
out_path.parent.mkdir(exist_ok=True, parents=True)
prs.save(str(out_path))
print(f"Presentation saved successfully to: {out_path}")
