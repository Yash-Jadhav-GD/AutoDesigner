import pptx, zipfile, sys
from xml.etree import ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

inp = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design_tester_new_inspect.pptx')
out = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx')

print(f'Input slides: {len(inp.slides)}, Output slides: {len(out.slides)}')
assert len(inp.slides) == len(out.slides) == 9, "Slide counts must be exactly 9!"

def get_texts(slide):
    texts = []
    for s in slide.shapes:
        if s.has_text_frame:
            for p in s.text_frame.paragraphs:
                t = p.text.strip()
                if t: texts.append(t)
        elif s.has_table:
            for r in s.table.rows:
                for c in r.cells:
                    for p in c.text_frame.paragraphs:
                        t = p.text.strip()
                        if t: texts.append(t)
        elif s.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
            for gs in s.shapes:
                if gs.has_text_frame:
                    for gp in gs.text_frame.paragraphs:
                        t = gp.text.strip()
                        if t: texts.append(t)
    return texts

total_in = 0
total_out = 0
all_perfect = True

for i, (s_in, s_out) in enumerate(zip(inp.slides, out.slides)):
    tin = get_texts(s_in)
    tout = get_texts(s_out)
    total_in += len(tin)
    total_out += len(tout)
    sin, sout = set(tin), set(tout)
    missing = sin - sout
    extra = sout - sin
    status = "PERFECT 100% MATCH" if (not missing and not extra) else "MISMATCH"
    if missing or extra:
        all_perfect = False
    print(f'Slide {i+1}: in={len(tin)}, out={len(tout)} -> {status}')
    if missing:
        print(f'   MISSING ({len(missing)}): {list(missing)[:3]}')
    if extra:
        print(f'   EXTRA ({len(extra)}): {list(extra)[:3]}')

print(f'\nTOTAL TEXT: in={total_in}, out={total_out}, All Perfect={all_perfect}')

# Check OpenXML rels
with zipfile.ZipFile(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx') as z:
    for i in range(1, len(out.slides)+1):
        rel_name = f'ppt/slides/_rels/slide{i}.xml.rels'
        slide_name = f'ppt/slides/slide{i}.xml'
        if rel_name in z.namelist():
            rel_xml = z.read(rel_name)
            slide_xml = z.read(slide_name)
            root = ET.fromstring(rel_xml)
            rel_ids = {el.attrib['Id'] for el in root}
            slide_root = ET.fromstring(slide_xml)
            used_ids = set()
            for elem in slide_root.iter():
                for k, v in elem.attrib.items():
                    if 'id' in k.lower() and v.startswith('rId'):
                        used_ids.add(v)
            diff = used_ids - rel_ids
            if diff:
                print(f'Slide {i} missing rel IDs: {diff}')
            else:
                print(f'Slide {i} rel integrity: OK')
