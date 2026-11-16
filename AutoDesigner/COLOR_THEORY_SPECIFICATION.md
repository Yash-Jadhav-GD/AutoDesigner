# GEP AutoDesigner: Color Theory Specification

Companion to `AUTOFORMAT_SPECIFICATION.md`. Source material: `ColorPalette_GEP.pptx` (brand palette and ramps) and `After_copy.pptx` (21-slide reference deck; color usage measured from the slide XML).

**Principle:** the engine never invents a color. Every fill, line, text run and gradient stop must resolve to a token in this file. Anything else is snapped to the nearest token or flagged.

---

## 1. Core brand palette (Layer 1: canonical)

| Token | Name | Hex | RGB | Character |
|---|---|---|---|---|
| `brand.indigo` | Indigo | `#48397E` | 72, 57, 126 | Primary: structure, headers, emphasis |
| `brand.tangerine` | Tangerine | `#FE860E` | 254, 134, 14 | Accent: the one thing to look at |
| `brand.navy` | Navy | `#1F2E65` | 31, 46, 101 | Secondary: depth, "systems" content |
| `brand.onyx` | Onyx | `#221E1F` | 34, 30, 31 | Text and dark neutral |

### Tint and shade ramps (from ColorPalette_GEP.pptx)

| Step | Indigo | Tangerine | Navy | Onyx |
|---|---|---|---|---|
| 90% tint | `#EDEBF2` | `#FFF3E7` | `#E9EAF0` | `#E9E8E9` |
| 70% tint | `#C8C4D8` | `#FFDBB7` | `#BCC0D1` | `#BDBCBC` |
| 50% tint | `#A49CBE` | `#FEC286` | `#8F96B2` | `#908E8F` |
| 30% tint | `#7F74A5` | `#FEAA56` | `#626D93` | `#646262` |
| **Base** | `#48397E` | `#FE860E` | `#1F2E65` | `#221E1F` |
| 20% shade | `#3A2E65` | `#CB6B0B` | `#192551` | `#1B1819` |
| 40% shade | `#2B224C` | `#985008` | `#131C3D` | `#141213` |

---

## 2. Tokens observed in the reference deck (Layer 2: approved "vivid tints")

The reference deck does not use the computed ramp values above for most fills. It uses slightly more saturated tints. Because After_copy is the design benchmark, these are approved as Layer 2 tokens. Counts are shape fills, lines and text runs across all 21 slides.

| Token | Hex | Observed use | Count | Nearest ramp step |
|---|---|---|---|---|
| `vivid.indigo.mid` | `#7C6DB2` | Number badges, pills, icon rings, timeline bars | 34 fills, 29 lines | Indigo 30% `#7F74A5` |
| `vivid.indigo.soft` | `#B1A8D8` | Banners, icon backplates, KPI tiles | 25 fills, 21 lines | Between Indigo 50% and 70% |
| `vivid.indigo.wash` | `#E2DEEE` | Card and panel gradient start | 20 | Indigo 90% `#EDEBF2` |
| `vivid.indigo.light` | `#C1BAE0` | Secondary tiles | 4 | Indigo 70% `#C8C4D8` |
| `vivid.navy.mid` | `#526299` | Navy variant headers, KPI text | 8 fills, 8 text | Navy 30% `#626D93` |
| `vivid.navy.soft` | `#8F9BC7` | Hairline borders, phase headers, gradient ends | 11 fills, 40 lines | Navy 50% `#8F96B2` |
| `vivid.tangerine.mid` | `#FFAC59` | Highlight tiles, orange-family bars | 17 fills, 14 lines | Tangerine 30% `#FEAA56` |
| `vivid.tangerine.soft` | `#FFD3A6` | Commercial and callout fills | 18 fills, 12 lines | Tangerine 70% `#FFDBB7` |
| `vivid.tangerine.wash` | `#FFECD9` | Pale callouts | 3 | Tangerine 90% `#FFF3E7` |

**Decision rule:** when both a ramp step and its Layer 2 counterpart fit, use Layer 2 for fills (matches the benchmark look). Use ramp steps when you need a step Layer 2 does not cover (shades, 90% washes).

### Neutrals observed

| Token | Hex | Use |
|---|---|---|
| `neutral.white` | `#FFFFFF` | Slide background and card fill (76 fills, the most used color in the deck) |
| `neutral.surface` | `#F6F8FB` / `#F9F8FC` | Panel backgrounds behind white cards |
| `neutral.border` | `#DFE3E6` | Card outlines (matches the Library's grey outline) |
| `neutral.line` | `#A2ABB0` / `#8A9BB0` | Dividers, connector lines |

### Text colors

| Token | Hex | Use | Contrast on white |
|---|---|---|---|
| `text.primary` | `#262626` | Body copy (Library standard) | 15.1:1 |
| `text.heading` | `#221E1F` (Onyx) | Slide titles | 16.5:1 |
| `text.secondary` | `#475569` | Captions, sub-labels | 7.6:1 |
| `text.on-dark` | `#FFFFFF` | On Indigo, Navy, Indigo-mid | see §6 |
| `text.accent` | `#48397E` | Highlighted phrases, labels | 9.7:1 |

The deck also uses `#2D3E50`, `#151514`, `#000000`, `#212225` for text. These are drift. Snap them to `text.primary` (`#262626`), or `text.heading` for titles.

---

## 3. Role map (what color goes where)

| Element | Fill | Line | Text |
|---|---|---|---|
| Slide background | `neutral.white` | none | n/a |
| Panel behind cards (optional) | `neutral.surface` | none | n/a |
| Card / Decor box | `neutral.white` at 22% alpha (Library) | `neutral.border`, 0.5 pt | n/a |
| Header pill (default) | `brand.indigo` | none | `#FFFFFF`, bold |
| Header pill (secondary column) | `brand.navy` | none | `#FFFFFF`, bold |
| Header pill (accent column) | `brand.tangerine` | none | **`brand.onyx` or `#FFFFFF` bold ≥18 pt only** (see §6) |
| Header pill shadow | `#302653` at 16% alpha, 15 pt blur | n/a | n/a |
| Number badge / step marker | `vivid.indigo.mid` | none | `#FFFFFF`, bold |
| Icon backplate | `vivid.indigo.soft` | none | icon in `brand.indigo` |
| Divider | `neutral.line` fading to transparent (Library "Grey Side Fade R") | n/a | n/a |
| Section banner (key message) | `vivid.indigo.soft` | none | `text.accent` or Onyx, bold |
| Callout / highlight | `vivid.tangerine.soft` | `vivid.tangerine.mid` | Onyx |
| KPI figure | tile `vivid.indigo.soft` or `vivid.tangerine.soft` | none | figure in `brand.indigo` or `brand.tangerine` shade |
| Footnote text | none | none | `text.secondary` |
| Panel / card gradient | `vivid.indigo.wash` fading to white, top to bottom | `vivid.navy.soft` hairline | n/a |

---

## 4. Proportion rule (60 / 30 / 10)

Measured on the reference deck, slides are mostly white and neutral, with indigo as the structural color and tangerine as a sparing accent.

| Share | Role | Colors |
|---|---|---|
| ~60% | Canvas | White, `neutral.surface`, card fills |
| ~30% | Structure | Indigo family (headers, badges, soft banners); Navy only as a second voice |
| ≤10% | Accent | Tangerine family (one focal element per slide) |

Engine limits:
- Tangerine family covers **at most one focal group per slide** (a header, a callout or a KPI row), never several competing ones.
- Navy and Tangerine headers may appear **together only in a progression** (see §5).
- No more than **three hues** (Indigo, Navy, Tangerine) plus neutrals on one slide.

---

## 5. Sequencing rules (how colors map to content order)

| Content shape | Rule | Reference |
|---|---|---|
| 3-part narrative (Problem → Bridge → Future) | Indigo → Navy → Tangerine pills | Slide 1 (the Design tester's hidden fills match this order) |
| 2-state comparison (Current vs Future) | Indigo (current) vs Navy (future), or Indigo vs Tangerine when the future state is the hero | Slides 5, 15 |
| N-step process (phases) | Step 1 `tangerine.mid`, step 2 `tangerine.soft`, steps 3–4 indigo (`mid`, `soft`), step 5 navy soft, light to heavy emphasis | Slide 20 |
| N parallel items (no order) | All the same color: Indigo, with Tangerine for at most one item to highlight | Slides 7, 12 |
| Value ramp (e.g. % autonomy) | One hue, tint steps by magnitude (higher value = darker step) | Slide 8 |
| Status or risk | Use §7 status colors, never brand colors | Slide 10 |

---

## 6. Contrast rules (deterministic checks)

Measured WCAG contrast ratios. The engine must pass these before output.

| Foreground | Background | Ratio | Verdict |
|---|---|---|---|
| White | Indigo `#48397E` | 9.7 | OK, any size |
| White | Navy `#1F2E65` | 12.8 | OK, any size |
| White | `vivid.indigo.mid` `#7C6DB2` | 4.5 | OK for body (borderline); prefer bold |
| White | `vivid.navy.mid` `#526299` | 5.9 | OK |
| White | Tangerine `#FE860E` | **2.4** | **Fail.** Do not use for text of any size |
| White | `vivid.tangerine.mid` `#FFAC59` | **1.9** | **Fail** |
| White | `vivid.indigo.soft` `#B1A8D8` | **2.2** | **Fail** |
| White | `vivid.navy.soft` `#8F9BC7` | **2.7** | **Fail** |
| Onyx / `#262626` | Tangerine `#FE860E` | 6.8 | OK |
| `#262626` | `vivid.tangerine.mid` / `.soft` | 8.1 / 10.9 | OK |
| `#262626` | `vivid.indigo.soft` `#B1A8D8` | 6.8 | OK |
| `#262626` | `vivid.indigo.mid` `#7C6DB2` | 3.4 | **Fail** for body (use white instead) |
| `brand.indigo` | `vivid.tangerine.soft` | 7.0 | OK |
| `brand.navy` | `vivid.tangerine.soft` | 9.2 | OK |
| Tangerine `#FE860E` as text | White | **2.4** | **Fail.** Use `#CB6B0B` (3.7, large text only) or `#985008` |

**Rules**
1. **Light text only on dark fills** (Indigo, Navy, `vivid.indigo.mid`, `vivid.navy.mid`).
2. **Dark text only on light fills** (all soft/wash tints and Tangerine).
3. **Tangerine fill takes dark text.** The reference deck puts white bold text on tangerine headers (for example "The Future State"); that pairing measures 2.4:1 and the engine should avoid it. Use Onyx on `brand.tangerine`.
4. Minimum contrast: **4.5:1** for text under 18 pt (or 14 pt bold); **3:1** for larger text and for icons and graphics.
5. Never place colored text on a colored fill from a *different* hue family unless it appears in the table above as OK.

---

## 7. Status and semantic colors (exception set)

Used **only** for risk, impact, RAG status and Gantt categories, never for decoration.

| Token | Hex | Observed use |
|---|---|---|
| `status.high` | `#E8503A` | "Impact: HIGH" tags |
| `status.med` | the deck's amber/orange `#F5A623` | Impact and probability MED tags, KPI figures |
| `status.ok` | green used in the roadmap bars | Positive deltas and completed items |

Rules: at most one status color per tag; status color never appears as a card or header fill; keep chart and Gantt colors on the brand palette first, adding green or red only where meaning requires it.

The funnel graphic on slide 21 (blues `#4FB7F6`, `#338DCD`) is a legacy off-palette image. It is not a token.

---

## 8. Gradients and effects

| Effect | Spec |
|---|---|
| Panel gradient | `#E2DEEE` → `#FFFFFF`, vertical, 90° |
| Header pill gradient (optional variant) | Base color → its `soft` tint, horizontal, left to right |
| Divider | `neutral.line` → transparent, left to right (Library fade) |
| Shadow | `#302653` at 16% alpha, 15 pt blur, no offset (header pills only) |
| Rule | Gradients stay within **one hue family**; never run Indigo into Tangerine |

---

## 9. Selection algorithm (for the engine)

```
1. Identify archetype from content shape (§5).
2. Assign column / step colors from the sequencing table.
3. Fill every shape from the role map (§3).
4. Pick text color by fill luminance:
      if fill in {indigo, navy, vivid.*.mid}  -> text = #FFFFFF
      else                                    -> text = text.primary / Onyx
5. Enforce proportion limits (§4): max 1 tangerine focal group, max 3 hues.
6. Run contrast checks (§6). On failure: swap text color, else step the fill one tint lighter or darker.
7. Snap any remaining off-palette color to nearest token (Lab distance); log it.
```

---

## 10. Off-palette snap table (drift found in After_copy)

| Found | Snap to |
|---|---|
| `#2D3E50`, `#151514`, `#000000`, `#212225`, `#0F0F0F`, `#242528` | `text.primary` `#262626` (titles: Onyx `#221E1F`) |
| `#2B3245` | `text.primary` |
| `#1F2F66` | `brand.navy` `#1F2E65` (1-unit difference) |
| `#F5A623` | `status.med` if a status tag, else `brand.tangerine` |
| `#8A9BB0`, `#A2ABB0`, `#66737C` | `neutral.line` |
| `#EDF1F6`, `#FCFDFE` | `neutral.surface` |
| `#B2E0FA`, `#7ECAF8`, `#4FB7F6`, `#338DCD`, `#20303D` | Off-palette. Keep only inside legacy graphics; flag in new slides |

---

## 11. Validation checklist (checker agent)

- [ ] Every color resolves to a token in §1, §2, §3 neutrals or §7 status
- [ ] Hues per slide ≤ 3 (excluding neutrals and status)
- [ ] Tangerine family used in ≤ 1 focal group
- [ ] No white or light text on Tangerine, `vivid.tangerine.*`, `vivid.indigo.soft`, `vivid.navy.soft`
- [ ] No dark body text on `vivid.indigo.mid`
- [ ] All text meets §6 minimum contrast
- [ ] Sequencing matches content shape (§5)
- [ ] Gradients stay in one hue family
- [ ] Status colors appear only on tags, never on structure

---

## 12. Open decisions for you

1. **Layer 2 vs ramp values.** The reference deck uses vivid tints (`#7C6DB2`, `#B1A8D8`, `#FFAC59`, `#FFD3A6`) rather than the computed ramp. This spec treats the deck as the benchmark. If the brand team wants strict ramp values, swap Layer 2 for the nearest ramp step in §2.
2. **White on tangerine headers.** The reference deck uses it, but it fails contrast (2.4:1). This spec says Onyx text. If you prefer to keep white, restrict it to bold text of 18 pt or more and use the `#CB6B0B` shade as the fill.
3. **Text color.** The deck's most common text color is `#2D3E50`; the Library standard is `#262626`. This spec follows the Library.
