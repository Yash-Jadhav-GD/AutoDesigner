from pathlib import Path

path = Path("AutoDesigner/engine/understand/rules.py")
text = path.read_text(encoding="utf-8")

old = """MATCHERS = [match_table, match_org, match_persona, match_workflow, match_assessment,
            match_split_process, match_process_rows, match_columns, _paragraph_cards, fallback]"""

matcher_code = """IMPORTANT_KEYWORDS = re.compile(r"\\b(important|key takeaways?|critical|note|notice|highlight|takeaways?)\\b", re.I)


def match_important_box(ctx: Ctx) -> Match:
    \"\"\"Matches slides that emphasize critical takeaways, important notes, or key highlights.\"\"\"
    if not any(IMPORTANT_KEYWORDS.search(t) for t in (ctx.title, ctx.subtitle)):
        headings = [e.text for e in ctx.body if ctx.heading(e)]
        if not any(IMPORTANT_KEYWORDS.search(h) for h in headings):
            return None
    pairs = heading_pairs(ctx, ctx.body)
    if pairs and 1 <= len(pairs) <= 6:
        items = [{"subheader": p["subheader"], "body": p["body"]} for p in pairs if p["subheader"] or p["body"]]
        data = {"title": ctx.title, "subtitle": ctx.subtitle, "cards": [{"header": "Important", "items": items}]}
        return 0.78, ("important_box", data), ["Keywords indicate an important callout or key takeaways box"]
    return None


MATCHERS = [match_table, match_org, match_persona, match_workflow, match_assessment,
            match_split_process, match_process_rows, match_important_box, match_columns, _paragraph_cards, fallback]"""

assert old in text, "old not found!"
text = text.replace(old, matcher_code, 1)

path.write_text(text, encoding="utf-8")
print("Successfully updated rules.py!")
