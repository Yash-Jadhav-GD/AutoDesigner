import pptx
from xml.etree import ElementTree as ET

prs = pptx.Presentation(r"C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Library_inspect.pptx")
s6 = prs.slides[5] # Slide 6

print("==================================================")
print("LINE PROTOTYPES ON LIBRARY SLIDE 6")
print("==================================================")
for sh in s6.shapes:
    print(f"\n--- Shape: {sh.name} (type={sh.shape_type}) ---")
    pos = f"left={sh.left/12700:.1f}, top={sh.top/12700:.1f}, w={sh.width/12700:.1f}, h={sh.height/12700:.1f}"
    print("Position:", pos)
    print(ET.tostring(sh._element, encoding="utf-8").decode("utf-8"))
