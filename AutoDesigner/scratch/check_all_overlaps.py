import pptx

prs = pptx.Presentation('AutoDesigner/output/Design_tester_new_composable.pptx')
print(f'Checking all {len(prs.slides)} slides for potential bounding box overlaps or negative gaps...')

found_issues = []
for s_idx, slide in enumerate(prs.slides):
    boxes = []
    for sp in slide.shapes:
        boxes.append({
            'name': sp.name,
            'left': sp.left.pt,
            'top': sp.top.pt,
            'right': sp.left.pt + sp.width.pt,
            'bottom': sp.top.pt + sp.height.pt,
            'has_text': sp.has_text_frame,
            'text': sp.text_frame.text[:30].replace('\n', ' ') if sp.has_text_frame else ''
        })
    
    text_boxes = [b for b in boxes if b['has_text'] and b['text'].strip()]
    for i in range(len(text_boxes)):
        for j in range(i+1, len(text_boxes)):
            b1 = text_boxes[i]
            b2 = text_boxes[j]
            x_ov = max(0, min(b1['right'], b2['right']) - max(b1['left'], b2['left']))
            y_ov = max(0, min(b1['bottom'], b2['bottom']) - max(b1['top'], b2['top']))
            if x_ov > 5 and y_ov > 5:
                issue = f"Slide {s_idx+1}: TEXT OVERLAP between '{b1['name']}' ({b1['text']}) and '{b2['name']}' ({b2['text']}) -> ov_x={x_ov:.1f}, ov_y={y_ov:.1f}"
                print(issue)
                found_issues.append(issue)

if not found_issues:
    print("Zero text bounding box overlaps detected by coordinates!")
else:
    print(f"Total coordinate overlap issues: {len(found_issues)}")
