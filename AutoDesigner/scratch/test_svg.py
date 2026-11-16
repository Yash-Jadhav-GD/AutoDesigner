import zipfile, io, os, pptx
from xml.etree import ElementTree as ET
from pptx.oxml import parse_xml

# Test creating a slide with SVG pic shape and checking in PowerPoint COM
src_path = r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx'
test_path = r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\test_svg_result.pptx'

prs = pptx.Presentation(src_path)
slide = prs.slides[3]

pic_xml = '''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvPicPr>
    <p:cNvPr id="9991" name="Graphic Test 1"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip>
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg1"/>
        </a:ext>
      </a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="693144" y="1569139"/><a:ext cx="361152" cy="361152"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>'''

slide.shapes._spTree.append(parse_xml(pic_xml))

buf = io.BytesIO()
prs.save(buf)
buf.seek(0)

# update zip with media and rels
with zipfile.ZipFile(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\After_copy.pptx') as z_in:
    svg_data = z_in.read('ppt/media/image19.svg')

out_buf = io.BytesIO()
with zipfile.ZipFile(buf, 'r') as zin, zipfile.ZipFile(out_buf, 'w') as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'ppt/slides/_rels/slide4.xml.rels':
            root = ET.fromstring(data)
            ET.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', {
                'Id': 'rIdSvg1',
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                'Target': '../media/image19.svg'
            })
            data = ET.tostring(root, xml_declaration=True, encoding='utf-8')
        elif item.filename == '[Content_Types].xml':
            root = ET.fromstring(data)
            has_svg = any(c.attrib.get('Extension') == 'svg' for c in root)
            if not has_svg:
                ET.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/content-types}Default', {
                    'Extension': 'svg',
                    'ContentType': 'image/svg+xml'
                })
            data = ET.tostring(root, xml_declaration=True, encoding='utf-8')
        zout.writestr(item, data)
    zout.writestr('ppt/media/image19.svg', svg_data)

with open(test_path, 'wb') as f:
    f.write(out_buf.getvalue())

print('Successfully created test_svg_result.pptx')
