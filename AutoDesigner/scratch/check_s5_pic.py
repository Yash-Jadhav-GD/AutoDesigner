import pptx

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
s5 = prs.slides[4]
print("Slide 5 shapes:")
for sh in s5.shapes:
    print(sh.name, sh.shape_type)
    if sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
        try:
            print("  Image content_type:", sh.image.content_type, "size:", sh.image.size)
        except Exception as e:
            print("  Image error:", e)
