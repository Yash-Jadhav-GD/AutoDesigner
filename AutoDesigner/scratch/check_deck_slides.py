from pptx import Presentation

formatted = Presentation("AutoDesigner/output/Design_tester_new_formatted.pptx")
print(f"Design_tester_new_formatted.pptx has {len(formatted.slides)} slides.")

raw = Presentation("AutoDesigner/Design tester new.pptx")
print(f"Design tester new.pptx has {len(raw.slides)} slides.")
