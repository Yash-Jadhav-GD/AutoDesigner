from pptx import Presentation

prs = Presentation("AutoDesigner/Design tester new.pptx")
s12 = prs.slides[11]
for sh in s12.shapes:
    if "Graphic" in sh.name:
        print(f"Shape {sh.name} XML:")
        xml_str = sh._element.xml
        print(xml_str[:500])
        break
