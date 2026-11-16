import copy
import io
import os
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from engine.builders.common import (
    clean_element,
    get_shape_by_name,
    add_library_line,
    load_library,
    set_cell_border,
    add_connector_line,
    inject_svg_assets,
)
from engine.color_tokens import (
    ColorTokens,
    hex_to_rgb,
    apply_linear_gradient,
    apply_solid_fill,
)

# Import legacy builders 1, 2, 3
from scratch.legacy.format_all_6_slides import (
    extract_slide1_data,
    build_slide1,
    extract_slide2_data,
    build_slide2,
    extract_slide3_data,
    build_slide3,
)

INPUT_PATH = BASE_DIR / "Design_tester_new_inspect.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"
LOCAL_COPY_PATH = BASE_DIR / "Design tester new_formatted.pptx"


def extract_slide4_data(input_prs):
    s = input_prs.slides[3]
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Case Study 3" in p.text][0]
    
    top_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Accelerated Sourcing Program" in p.text][0]
    
    scope_items = [
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Spend Assessed" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Categories:" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Regions-in-scope" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Opportunities Identified" in p.text][0],
    ]
    
    pain_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Pain Points" in p.text][0]
    pain_bullets = [
        "Decentralized sourcing practices across OpCos limited enterprise-wide leverage",
        "Fragmented spend and supplier information constrained enterprise visibility",
        "Predominantly reactive sourcing with limited forward-looking category pipelines",
        "Inconsistent supplier arrangements and sourcing practices across the organization",
        "Limited use of category intelligence and market insights in sourcing decisions",
        "Recent acquisitions created targeted contract harmonization opportunities",
    ]
    
    approach_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip() == "Program Approach"][0]
    app_sec1_title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Assess and Prioritize" in p.text][0]
    app_sec1_bullets = [
        "Analyzed spend, supplier arrangements and category requirements",
        "Validated opportunities with procurement and OpCo stakeholders",
        "Prioritized initiatives based on value, feasibility and business readiness",
    ]
    app_sec2_title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Develop and Execute" in p.text][0]
    app_sec2_bullets = [
        "Developed category-specific sourcing and negotiation strategies",
        "Mobilized priority initiatives with business stakeholders",
        "Tracked savings through negotiation and implementation",
    ]
    
    opps_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Key Opportunities" in p.text][0]
    opps_items = [
        "IT contract harmonization: Reviewed supplier contracts and identified opportunities to move acquired businesses to more advantageous enterprise agreements",
        "Acquisition synergies: Compared commercial arrangements across the firm and identified pricing and consolidation opportunities post acquisitions",
        "Fleet transformation: Developed a broader operating-model and transformation perspective for global fleet leveraging full scale of combined companies",
        "Strategic sourcing pipeline: Established a prioritized portfolio of category initiatives aligned with business requirements identified through synergy opportunity analysis as well as broader opportunity assessment",
        "Identified and partnered with critical stakeholders to enable successful execution against identified opportunities",
    ]
    
    outcomes_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Key Outcomes" in p.text][0]
    outcomes_metrics = [
        "5%-6% annual uplift in savings, 14% - 15% savings via multi-year strategies",
        "Shift from Reactive to Proactive Sourcing",
        "Prioritized Sourcing Pipeline Developed",
        "~700 Supplier Contracts Renewed",
    ]
    
    return {
        "title": title,
        "top_header": top_header,
        "scope_items": scope_items,
        "pain_header": pain_header,
        "pain_bullets": pain_bullets,
        "approach_header": approach_header,
        "app_sec1_title": app_sec1_title,
        "app_sec1_bullets": app_sec1_bullets,
        "app_sec2_title": app_sec2_title,
        "app_sec2_bullets": app_sec2_bullets,
        "opps_header": opps_header,
        "opps_items": opps_items,
        "outcomes_header": outcomes_header,
        "outcomes_metrics": outcomes_metrics,
    }


def extract_slide5_data(input_prs):
    s = input_prs.slides[4]
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "QXO" in p.text][0]
    
    headers = [
        "Category",
        "Spend Managed by GEP",
        "Subject Matter Experts",
        "Transformative Value Levers (more detail in appendix)",
    ]
    
    table_shape = [sh for sh in s.shapes if sh.has_table][0]
    t = table_shape.table
    rows = []
    for r in t.rows:
        row_vals = []
        for c in r.cells:
            cell_paras = [p.text.strip() for p in c.text_frame.paragraphs if p.text.strip()]
            row_vals.append(cell_paras)
        rows.append(row_vals)
        
    return {
        "title": title,
        "headers": headers,
        "rows": rows,
    }


def extract_slide6_data(input_prs):
    s = input_prs.slides[5]
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Opportunity Perspective" in p.text][0]
    headers = ["Workstreams", "Spend  and Suppliers", "Planned Interventions", "Estimated Impact"]
    
    tracks = [
        {
            "letter": "A",
            "name": "Cross- Entity Supplier Synergies",
            "spend_sup": ["Overlapping Suppliers", "$166M | ~104 suppliers"],
            "interventions": [
                "Consolidate spend & commercial terms across acquired businesses",
                "Renegotiate enterprise-wide agreements leveraging combined scale",
                "Harmonize pricing, rebates, payment terms, and service levels",
            ],
            "impact": ["$20  25M", "(12-15%)"],
        },
        {
            "letter": "B",
            "name": "Strategic Sourcing & Optimization",
            "spend_sup": ["Suppliers > $5M spend", "$450M  ~20 suppliers"],
            "interventions": [
                "Reassess category strategies using  combined QXO spend footprint",
                "Run targeted sourcing waves for high-value categories and suppliers",
                "Introduce competitive tension and marketplace benchmarking",
            ],
            "impact": ["$55  60M", "(12-15%)"],
        },
        {
            "letter": "C",
            "name": "Accelerated RFX Program",
            "spend_sup": ["Suppliers $500k - $5M spend", "$305M | ~230 suppliers"],
            "interventions": [
                "Stand up a centralized RFx execution engine across New QXO",
                "Execute sourcing waves at scale using AI-enabled sourcing infrastructure",
                "Provide dedicated sourcing capacity to address opportunities that would otherwise remain untouched",
            ],
            "impact": ["$15  20M", "(5-6%)"],
        },
        {
            "letter": "D",
            "name": "Digital Tail Spend Management",
            "spend_sup": ["Suppliers < $500K spend", "$310M | ~15K suppliers"],
            "interventions": [
                "Create a centralized procurement channel for  spend across New QXO",
                "Free up QXO Indirect leader time to focus on integration priorities",
                "Enable scalable execution through a combination of managed services, digital workflows and AI",
            ],
            "impact": ["~$10-15M", "(3-4%)"],
        },
    ]
    
    return {
        "title": title,
        "headers": headers,
        "tracks": tracks,
    }

print("Extractor definitions ready.")
