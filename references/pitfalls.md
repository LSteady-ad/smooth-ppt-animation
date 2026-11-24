# Pitfalls and Exact Fixes

Each entry is a common defect observed in real-world decks, with the root cause and the fix that resolved it. Consult this before changing build logic.

## P1 — `r:embed` is slide-local (the biggest one)

**Symptom:** A large rounded-corner banner morphs from an *ellipse* on the previous slide. The middle frame looks badly distorted.

**Root cause:** `<p:pic>` references an image via `r:embed="rId4"`. Relationship IDs are local to each slide's `.rels`. Copying the XML block to another slide keeps `rId4`, but on the target slide `rId4` may point to a completely different image. Cross-slide picture copies then show the wrong picture and force a wrong shape interpolation.

**Fix:** Before copying, build `RELS[slide] = {rId → Target}` from `slideN.xml.rels`. On each copy, resolve the source `rId` to its media Target, then on the destination either reuse an existing rId with the same Target or add a new relationship, and rewrite `r:embed` in the copied block. Flush all `.rels` before zipping.

## P2 — Morph smart-matching a decorative ellipse to a large rectangle

**Symptom:** On two adjacent slides, a large rectangle arrives as a rounded/elliptical shape.

**Root cause:** The earlier slide has small decorative ellipses with no same-name copy on the next slide, so Morph smart-matches one of them to the large rectangle, forcing rect↔ellipse interpolation.

**Fix:** In the fly-out pass, add an exception: small cover-page decorative ellipses must get off-screen same-name copies on the next slide even though their area is below the threshold. Once they have their own pair, they fly away instead of pairing with the large block. Check adjacent slides for any same-name pair with mismatched `prstGeom` (rect vs ellipse) — there must be zero.

## P3 — Wrong picture morph from un-named elements

**Symptom:** A thick block morphs into a thin table border on the next slide.

**Root cause:** Un-named shapes were smart-matched across slides.

**Fix:** Name every shape with a deterministic `!!` name; keep horizontal and vertical lines separately named (`!!HL_Pn_k` / `!!VL_Pn_k`) so a horizontal line never pairs with a vertical one.

## P4 — Same `!!` name used for visually different decorations

**Symptom:** A top navigation bar flies horizontally across the screen.

**Root cause:** The same `!!NAV_ITEM_n` name was used for a large block on one slide and a thin top bar on the next — the same name forced a pair.

**Fix:** Area-tiered naming. Keep `!!NAV_ITEM_n` only if its area ≥ 20% of the page (two large blocks that are allowed to morph regardless of exact appearance). Everything else becomes page-unique `!!NAV_Pn_ITEM_k`.

## P5 — z-order flicker

**Symptom:** At the middle frame, overlapping elements suddenly blink / reorder.

**Root cause:** Paired elements were at layer 1 on one slide and the last layer on the other; off-screen clones were all appended to the end of `spTree`.

**Fix:** See `concepts.md` §7. Do not mix on-screen and off-screen elements on one sorting axis — when many clones compete for percentile slots, targets can drift far apart in layer order.

## P6 — Layer chain for large-block fly-out copies

**Symptom:** A large block's fly-out copy landed at the very top layer while its source was near the bottom.

**Fix:** In the z-order pass, if the same-name element on the adjacent slide is *also* off-screen (a chain of fly-out clones), inherit its global percentile, not "append to end".

## P7 — Flying text lands outside its color block

**Symptom:** White text on a color block becomes invisible mid-transition, then pops into view when it reaches the block.

**Root cause:** Text and block used different off-screen trajectories, and their layer order reversed.

**Fix:** Detect local color blocks (20–85% page area, excluding full-screen `!!BG_`). For text that sits on such a block in slide B, remember the relationship in `ANCHORED`. Do not anchor the text *position* inside slide A's block (that can overlap existing cover text) — keep it off-screen as usual, but in the layer pass force the text above its block on both slides and exclude it from page A's animation.

## P8 — Text appearing before its card

**Symptom:** Description sentences appear while the card/title is still flying in from off-screen.

**Root cause:** In-page animations started at t=0 while Morph takes 2000 ms.

**Attempted fix (wrong):** Add a global `base_delay=1600ms` to all in-page animations.

**Real fix:** Set `base_delay = 0`. The first animated element must move immediately when the transition ends; perception of a "beat of silence" is worse. Ordering comes from group stagger (230 ms) and intra-group relay (90–110 ms) only.

## P9 — Smart-match promoting a description to level-2

**Symptom:** Descriptions on one side of a comparison page did not animate, while the other side did.

**Root cause:** A small decorative dot pattern next to a description was treated as a "small graphic + short label" pair, marking the description as necessary.

**Fix:** Apply the small-graphic rule only to short texts (≤8 chars, no sentence punctuation). Long descriptions stay tertiary regardless of nearby dots.

## P10 — Table rating words swallowed by the graphic rule

**Symptom:** Short rating words had no animation.

**Root cause:** The "short text on a graphic" rule matched the table row container (a large shape) and classified ratings as level-2.

**Fix:** Table-cell exception must only consider empty `p:sp` rectangles with w≥150 and 40≤h≤130. Image banners (e.g. a `p:pic` conclusion bar) must NOT count as table cells, otherwise a short title is wrongly made tertiary.

## P11 — Short descriptions misclassified as short titles

**Symptom:** A short description sentence morphed while its sibling sentences animated.

**Fix:** If ≥3 same-size card containers exist on a page, and a short text inside one has a font size within 1 pt of the long (>8 char) descriptions in sibling containers, classify it as tertiary. Labels whose font size is close to sibling descriptions stay level-2.

## P12 — Mechanical direction alternation

**Symptom:** Every page flew in from the right or top; or directions were so varied the page looked chaotic.

**Fix:** Rotate an 8-direction ring per page, but for complex pages collapse to left/right halves. Group horizontally adjacent elements in a row so a logical unit does not split at the mid-page boundary.

## P13 — Monotonous or mechanical rhythm

**Symptom:** Every page used one animation; or uniform/sequence alternated by chapter too regularly.

**Fix:** Use 5 rhythms in a non-mechanical array (uniform / seq / reverse / wave / relay), check neighboring pages differ. Rating words use Zoom, descriptions use Fade, with different intra-group steps.

## P14 — Some editor SDKs report zero animations on Morph slides

**Symptom:** Some page-info APIs return `anim_count=0` even though `<p:timing>` exists.

**Fix:** Do not trust that result for Morph slides. Verify at the XML level. Also note ElementTree reserialization corrupts `mc/p14/p159` prefixes — use pure string/regex manipulation on slide XML.

## P15 — Accidentally overwriting a hand-edited file

**Symptom:** A patch command failed silently and an unpatched copied script was executed, regenerating over a hand-edited deck. No recovery copy existed.

**Fix (mandatory workflow):**
1. On any "re-do based on the latest version" request, copy the edited file to `_source_latest.pptx` as a read-only baseline FIRST.
2. After copying a build script, verify `SRC` and `OUT` before running it. If a patch fails, STOP — do not run.
3. Add `assert os.path.abspath(SRC) != os.path.abspath(OUT)` to every build script.
4. Always output a new version number. Never overwrite a file the user may have edited.
5. Full backup of the workspace to TEMP before large cleanup operations.
