# Morph Animation Concepts

## 1. How PowerPoint Morph pairs elements

Morph is a slide transition (`p159:morph`) that animates elements between two slides. Pairing rules, from Microsoft:

- **Same `!!name` on both slides → forced 1:1 pair.** This is the only reliable way to control pairing.
- **Same name without `!!` on both slides → also paired** (older behavior), but avoid relying on it.
- **Different `!!names` → never paired**, each element fades in/out independently.
- **No name (default `Shape 12`, `Text 5`, etc.) → smart-matched by visual similarity.** This is dangerous: a horizontal line can morph from a thick block, a small dot can pair with a large rectangle, an ellipse can pair with a red banner. Always name every shape deterministically.

Naming scheme used in production (all with `!!` prefix so Morph treats them explicitly):

| Prefix | Meaning | Cross-slide behavior |
|---|---|---|
| `!!BG_n` | full-screen background | same name → stable anchor |
| `!!NAV_ITEM_n` | top navigation element | area-tiered: keep if ≥20% page, else page-unique |
| `!!NAV_Pn_ITEM_k` | page-unique nav element | no cross-slide pair |
| `!!FOOTER_n` | footer | same name → stable |
| `!!HL_n` / `!!VL_n` | horizontal / vertical line | page-unique (`!!HL_Pn_k`) so horizontal never pairs with vertical |
| `!!TITLE_Pn` | page title | unique per page; next-page title cloned off-screen to previous slide for fly-in |
| `!!C_Pn_k` | content text/shape | unique per page |
| `!!IMG_Pn_k` | picture | unique per page |

## 2. Off-screen duplicates: the core trick

Morph only animates a pair if the element exists on both slides. To make an element "fly in" from outside the canvas:

1. Copy the element's XML block from slide B to slide A.
2. Keep the same `!!name`.
3. Change its `<a:off x=... y=.../>` to a position outside the canvas.
4. Assign a new unique `cNvPr id`.

On the transition, Morph sees the same name on A (off-screen) and B (on-screen) and animates from A's position to B's. Similarly, copy A's element to B off-screen to make it fly out.

Direction tips:
- Use a per-page main direction, rotated across pages (right/up/left/down/four diagonals), so every page enters differently.
- Vary the off-screen distance in 3 tiers by row (`(row + page_no) % 3`) to create speed differences — elements arrive together but not mechanically.
- Group horizontally adjacent elements in the same row (e.g. number `01` + its label) and give the whole group one direction, otherwise a boundary at mid-page splits a logical unit.
- Merge crowd directions: for complex pages, use only left/right halves rather than 8 directions, to avoid visual chaos.

## 3. Morph XML

```xml
<mc:AlternateContent xmlns:mc="..." xmlns:p14="..." xmlns:p159="...">
  <mc:Choice Requires="p159">
    <p:transition spd="slow" p14:dur="2000">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback xmlns="">
    <p:transition spd="slow"><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

- Insert after `</p:cSld>` in `slideN.xml`; declare `mc`, `p14`, `p159` on the root.
- `p14:dur="2000"` is PowerPoint's default "Slow" Morph duration.
- Page 1 typically has no Morph: remove the entire `<mc:AlternateContent>...</mc:AlternateContent>` block and animate all its on-screen elements instead.

## 4. Animation presetIDs

Extracted from real PowerPoint files (do not guess — incorrect IDs produce wrong effects):

| presetID | Effect |
|---|---|
| 1 | Appear |
| 2 | Fly In |
| 10 | Fade |
| 42 | Float Up |
| 47 | Float Down |
| 53 | Zoom / Grow |
| 55 | Expand |

To obtain templates: apply each effect to a shape with the native tool, save, then regex-extract `<p:par><p:cTn ... presetID="N" ... </p:cTn></p:par>` and replace `spid` and `id` with placeholders (`__SPID__`, `__ID__`). In the generated `<p:timing>`, rewrite `nodeType` to `afterEffect` for the first element of a group and `withEffect` for the rest; set `dur="400"`.

Do not use `presetID=45` (Rotate) — it reads as gimmicky and users consistently ask to remove it.

## 5. Timing structure

```
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn nodeType="tmRoot">
        <p:childTnLst>
          <p:seq>
            <p:cTn nodeType="mainSeq">
              ... one clickPar containing a withGroup of all per-element effects ...
```

All effects on a page live in a single `clickPar`/`withGroup`. Per-group delay comes from the outer `<p:cond delay="..."/>`; per-element delay adds to it. Rhythms (rotate non-mechanically across pages):

- `uniform` — all at once
- `seq` — groups top-to-bottom (or left-to-right), `delay = order × 230ms`
- `reverse` — reversed order
- `wave` — odd/even groups in two waves
- `relay` — group stagger + intra-group relay (`110ms` per element)

Short rating words (`高/中/低`, `01`) use Zoom (53) with 110 ms intra-group step; description sentences use Fade (10) with 90 ms step.

**First delay must be 0.** Do not add a global base delay to "wait for Morph" — the audience perceives a pause after the transition. Keep the page total under 4 s via group stagger only.

## 6. Classifying elements: reverse recognition

Rather than trying to define "tertiary text" positively, first identify **necessary (level-2-and-above) elements** that must Morph; everything else with text is tertiary.

Necessary if any:
- Name starts with `!!TITLE_`
- Font size ≥ 18 pt
- Text is purely digits (scores, indices) — keeps them with their neighboring text
- Short text (≤8 chars, no `，。；`) inside or adjacent to a small graphic (`p:pic`, or an empty `p:sp` ≤140×90 pt)
- Short title (≤8 chars, no punctuation)

Exceptions that must stay tertiary:
- Short rating words inside a genuine table cell (an empty `p:sp` with w≥150, 40≤h≤130) — e.g. 高/中/低
- A short text in a repeated card row whose font size matches the long description in sibling rows — otherwise the row becomes inconsistent
- A long description merely having a small decorative dot to its left — the dot is decoration, not a semantic label

## 7. Layer order (z-order)

`<p:spTree>` child order determines stacking. For every same-name pair across adjacent slides, the percentile position in `spTree` should be close. Algorithm that works:

1. Fix on-screen elements in their current relative order.
2. For each off-screen clone, find the same-name element on either adjacent slide; use its percentile among **on-screen-only** elements as target (fallback: global percentile if that element is also off-screen).
3. Insert the clone just before the on-screen element at that percentile slot.

Putting all clones at the end of `spTree` (top layer) causes flicker when a paired element was at the bottom on the other slide.
