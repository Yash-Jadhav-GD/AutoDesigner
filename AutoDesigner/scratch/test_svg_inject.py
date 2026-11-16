import io
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import pptx
from pptx.util import Pt
from pptx.oxml import parse_xml

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Add circle
badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.OVAL, Pt(100), Pt(100), Pt(50), Pt(50))

# Add XML pic
rel_id = "rIdSvgTest_1"
pic_l = int(110 * 12700)
pic_t = int(110 * 12700)
pic_sz = int(30 * 12700)
pic_xml = f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvPicPr>
    <p:cNvPr id="9991" name="Graphic Test"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip>
      <a:extLst>
        <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">
          <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{rel_id}"/>
        </a:ext>
      </a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{pic_l}" y="{pic_t}"/><a:ext cx="{pic_sz}" cy="{pic_sz}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>'''

slide.shapes._spTree.append(parse_xml(pic_xml))

buf = io.BytesIO()
prs.save(buf)
raw_bytes = buf.getvalue()

# Inject SVG into package
svg_path = Path(r"c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\icon_finder\assets\icons\laptop-fb9eb9.svg")
svg_data = svg_path.read_bytes()

out_buf = io.BytesIO()
with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zin, zipfile.ZipFile(out_buf, "w") as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "ppt/slides/_rels/slide1.xml.rels":
            root = ET.fromstring(data)
            ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship", {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target": "../media/laptop-fb9eb9.svg"
            })
            data = ET.tostring(root, xml_declaration=True, encoding="utf-8")
        elif item.filename == "[Content_Types].xml":
            root = ET.fromstring(data)
            if not any(c.attrib.get("Extension") == "svg" for c in root):
                ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/content-types}Default", {
                    "Extension": "svg",
                    "ContentType": "image/svg+xml"
                })
            data = ET.tostring(root, xml_declaration=True, encoding="utf-8")
        zout.writestr(item, data)
    zout.writestr("ppt/media/laptop-fb9eb9.svg", svg_data)

final_pptx = Path(r"c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\test_svg_injected.pptx")
final_pptx.write_bytes(out_buf.getvalue())
print("Successfully generated test_svg_injected.pptx!")
