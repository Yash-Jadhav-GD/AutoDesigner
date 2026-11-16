import pptx

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
s7 = prs.slides[6]
print("Slide 7 shapes:")
for sh in s7.shapes:
    print(f"[{sh.name}] type={sh.shape_type}")
    if sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
        try:
            print("  Image content_type:", sh.image.content_type, "size:", sh.image.size)
            # save
            with open(f"AutoDesigner/scratch/s7_{sh.name}.png", "wb") as f:
                f.write(sh.image.blob)
            print("  Saved image")
        except Exception as e:
            print("  Image error:", e)
    elif sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
        print(f"  Group with {len(sh.shapes)} shapes:")
        for g in sh.shapes:
            print(f"    sub [{g.name}] type={g.shape_type}")
            if g.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                try:
                    print("      sub Image content_type:", g.image.content_type, "size:", g.image.size)
                except Exception as e:
                    print("      sub Image error:", e)
