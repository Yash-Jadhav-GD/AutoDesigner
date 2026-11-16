# AutoDesigner: Autonomous PowerPoint Design Engine & Stylist

**AutoDesigner** turns raw, unformatted, or draft PowerPoint presentations into executive-ready, polished decks. It reads each draft slide, decides which layout archetype fits it, fills that archetype with the slide's text (every word kept), and builds the styled deck.

```
draft.pptx ──▶ slide model ──▶ spec JSON (reviewable) ──▶ styled deck ──▶ fidelity check
             engine/slide_model   engine/understand        engine/builders    eval/verify_output
```

---

## Directory Structure

```
AutoDesigner/
├── input/                     # Draft presentations
├── output/                    # Built decks; output/specs/ holds generated specs
├── specs/                     # Hand-checked gold specs (used by the eval; never written by a run)
├── templates/                 # palettes.json, archetypes.json (legacy; generic geometry now comes from Library_copy.pptx)
├── engine/
│   ├── slide_model.py         # Reads every shape (groups, tables, notes) with geometry + typography;
│   │                          # role-clustering helpers from AUTOFORMAT_SPECIFICATION.md section 4
│   ├── understand/            # Slide -> archetype + data
│   │   ├── rules.py           # Local, feature-scored matchers (offline)
│   │   └── llm.py             # Claude mapper (optional), validated + cached, falls back to rules
│   ├── builders/
│   │   ├── registry.py        # Archetype names, descriptions, builder dispatch
│   │   ├── schemas.py         # Pydantic data schema per archetype
│   │   ├── archetypes.py      # Branded builders (pills, SVG icons, org chart, tables)
│   │   ├── generic.py         # Generic card archetypes via layout_engine + styler
│   │   ├── common.py          # Library prototypes, divider lines, SVG injection
│   │   └── deck.py            # validate_spec / build_deck / extract_spec
│   ├── color_tokens.py        # ColorTokens, TypographyTokens, brand_palette()
│   ├── layout_engine.py       # Generic card geometry from Library prototypes (spec section 5)
│   ├── styler.py              # Clones Library prototypes for generic cards (spec sections 3, 7)
│   ├── svg_icons.py           # Keyword -> SVG icon matching for icon_badge rows (spec section 6)
│   ├── analyzer.py, icon_service.py   # original generic engine
├── eval/
│   ├── run_eval.py            # Accuracy against gold specs
│   ├── verify_output.py       # Output deck fidelity: slide count, every word, relationships
│   └── results/               # Saved eval summaries (baseline.json, rules.json)
├── designer.py                # PPTDesigner: analyze -> spec -> build -> verify
├── run_designer.py            # CLI
├── build_from_spec.py         # Build from (or write) a spec
└── scratch/                   # Old exploration scripts; scratch/legacy holds the original per-deck format scripts
```

---

## Quick Start

```powershell
pip install -r requirements.txt

# Design a deck end to end (writes output/specs/<deck>.json and output/<deck>_designed.pptx)
python run_designer.py --input input/your_deck.pptx

# Review first: write the spec only, edit it, then build
python run_designer.py --input input/your_deck.pptx --spec-only
python build_from_spec.py --spec output/specs/your_deck.json --output output/your_deck.pptx

# Force the offline rules mapper, or pick a palette for the generic card archetypes
python run_designer.py --input input/your_deck.pptx --mapper rules --palette gep_brand

# Add keyword-matched SVG icons to generic card_grid rows (Library slide 4 backplates)
python run_designer.py --input input/your_deck.pptx --icons
```

The run prints, per slide, the chosen archetype, its confidence and why. Slides under 0.6 confidence are flagged for review. It finishes with the fidelity check and exits with code 2 if any word was lost.

### Mappers (`--mapper`)
| Mapper | What it does |
| :--- | :--- |
| `auto` (default) | `llm` when Claude credentials exist (`ANTHROPIC_API_KEY` / `ANTHROPIC_AUTH_TOKEN` / `ant auth login`), else `rules` |
| `rules` | Offline. Scores every archetype from geometry, typography, fills and patterns; picks the best that keeps all text |
| `llm` | Sends a compact JSON slide model (text, positions, font sizes, fills; **no images**) to `claude-opus-5-5`; the answer is schema-validated and coverage-checked, retried once, else the rules result is used. Cached in `.cache/understand/` |
| `baseline` | The original analyzer, kept for comparison only |

> Using `llm` sends slide text to the Anthropic API. Use `--mapper rules` for content that must stay on the machine.

---

## Archetypes

`python build_from_spec.py --list-archetypes` lists them; descriptions are in `engine/builders/registry.py`.

| Archetype | Layout |
| :--- | :--- |
| `three_column_pills` | 2–4 parallel columns, gradient pill headers |
| `split_process_insights` | Heading/body process rows + side insights panel (SVG icons) |
| `process_rows` | Full-width heading/body rows |
| `assessment_kpis` | Process rows, deliverables checklist, KPI rows (SVG icons) |
| `workflow_cards` | Lettered workflow rows with In/Out and autonomy |
| `persona_grid` | Grid of team cards with counts and regions |
| `org_structure` | Banner, leadership boxes (type from legend colour), three pools |
| `allocation_table` / `sizing_table` | Styled data table, with or without note/footer |
| `card_grid` / `kpi_cards` / `subheader_list` | Generic cards for any other slide, cloned from the `Library_copy.pptx` prototypes (Design 1 / KPI-CRD-V1 / CRD-SUB-LST); dense cards flow rows into up to 3 columns |

A spec is `{"slides": [{"archetype": "<name>", "data": {...}}]}`; each `data` must match `engine/builders/schemas.py` (checked before building). Branded builders need `Library_copy.pptx` and `After_copy.pptx` in this folder.

---

## Measuring accuracy

```powershell
python eval/run_eval.py --mapper rules          # archetype / field / coverage / schema scores vs specs/
python eval/run_eval.py --mapper baseline --save
python eval/verify_output.py input/deck.pptx output/deck_designed.pptx
```

Current results on the 12 gold slides (`specs/design_tester_new.json`, `specs/sample_draft.json`):

| Mapper | Archetype | Fields | Full text coverage |
| :--- | :---: | :---: | :---: |
| baseline (old analyzer) | 1/12 | 0.08 | 8/12 |
| rules | 12/12 | 1.00 | 12/12 |

The rules were developed against these slides, so treat the score as a regression check, not as accuracy on new decks. On the unseen 20-slide `Before_copy.pptx`, every slide keeps 100% of its text, but 18 of 20 fall back to generic cards. Add more gold specs to grow the eval.

---

## Executive Color Palettes

- **`gep_purple`** (default for generic cards): Primary `#48397E`, Secondary `#302653`, Accent `#7E6BAE`
- **`gep_brand`**: built from `ColorTokens` (Indigo `#48397E`, Navy `#1F2E65`, Tangerine `#FE860E`)
- **`corporate_blue`**, **`emerald_slate`**, **`slate_onyx`**: see `templates/palettes.json`

Branded archetypes always use `ColorTokens` (see `COLOR_THEORY_SPECIFICATION.md`).
