import pptx

prs = pptx.Presentation("AutoDesigner/Library_inspect.pptx")
print(f"Slide count: {len(prs.slides)}")
s6 = prs.slides[5] # Slide 6 is index 5
print(f"Shapes on Slide 6: {len(s6.shapes)}")
for sh in s6.shapes:
    print(f"Shape name: '{sh.name}', type: {sh.shape_type}, left: {sh.left.pt:.1f}, top: {sh.top.pt:.1f}, width: {sh.width.pt:.1f}, height: {sh.height.pt:.1f}")
