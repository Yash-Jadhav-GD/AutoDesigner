import pptx

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
print("Presentation slides:", len(prs.slides))

# Slide 1
s1 = prs.slides[0]
s1_title = [p.text.strip() for sh in s1.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "The Problem" in p.text][0]
print("S1 Title:", s1_title)

# Slide 2
s2 = prs.slides[1]
s2_title = [p.text.strip() for sh in s2.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "supported O&G" in p.text][0]
print("S2 Title:", s2_title[:40])

# Slide 3
s3 = prs.slides[2]
s3_title = [p.text.strip() for sh in s3.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "fabrication clusters" in p.text][0]
print("S3 Title:", s3_title[:40])

# Slide 4
s4 = prs.slides[3]
s4_title = [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Client Requirement" in p.text][0]
print("S4 Title:", s4_title[:40])

# Slide 5
s5 = prs.slides[4]
s5_title = [p.text.strip() for sh in s5.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Benchmarking Parameters" in p.text][0]
print("S5 Title:", s5_title[:40])

# Slide 6
s6 = prs.slides[5]
s6_title = [p.text.strip() for sh in s6.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Higher down payments" in p.text][0]
print("S6 Title:", s6_title[:40])

# Slide 7
s7 = prs.slides[6]
s7_title = [p.text.strip() for sh in s7.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "China has multiple ports" in p.text][0]
print("S7 Title:", s7_title[:40])
print("All 7 titles extracted cleanly!")
