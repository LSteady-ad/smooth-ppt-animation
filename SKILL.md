---
name: smooth-ppt-animation
description: 为已有 PowerPoint（.pptx）添加高级感动画：Morph 平滑切换、同名元素配对、跨页屏外飞入飞出、页内入场动画、z-order 层级修正和自动校验。当用户提出“给 PPT 加动画 / 平滑切换 / Morph / 炫酷切换 / 元素飞入飞出 / 每页精美动画 / 基于最新版重新上动画”等需求时使用；支持大面积内容随平滑移动、细节文字分批进入、每页 4 秒内完成、图片跨页引用修复，并可安全重做用户已修改的动画稿。
agent_created: true
---

# Morph PPT Animation

## Overview

Produce a polished animated `.pptx` from an existing deck using a hybrid of **Morph transitions + off-screen same-name duplicates + selective in-page animations**. The result: large elements smoothly travel across slides, tertiary text fades/scales in after the transition, and nothing flickers or morphs into the wrong shape.

This toolkit encodes rules learned from repeated iteration on real-world decks. The dominant failure modes (wrong image pairing, ellipse-to-rectangle morph, z-order flicker, text appearing too early) are all handled by specific passes in `scripts/build_morph_anim.py`.

## When to Use

- User asks to add animation / 动画 / 平滑切换 / Morph / 炫酷切换 to an existing PPTX
- User wants cross-slide "fly in from off-screen / fly out" motion with same-name elements
- User wants a mix of smooth morph for major content and entrance animations for detail text
- User edited an already-animated deck and asks to "re-apply animation on the latest version" — the build script is idempotent (strips old animation and old off-screen clones first)

## Core Concepts (read `references/concepts.md` for full detail)

1. **Morph needs same-name pairs.** Elements with the same `!!name` across two slides are force-paired by PowerPoint. Names without `!!` are smart-matched by visual similarity — which is dangerous and causes wrong morphs.
2. **Off-screen duplicates create travel.** Copy a next-slide element onto the previous slide, placed outside the canvas, with the same `!!name`. During Morph it flies into place. Copy current-slide elements onto the next slide off-screen to fly out.
3. **Two layers of motion.** Major content (titles, images, cards, labels) moves via Morph. Tertiary text (descriptions, ratings) uses in-page entrance animations. Mixing them correctly is what makes a deck feel alive without chaos.
4. **`r:embed` is slide-local.** Copying a `<p:pic>` block to another slide without remapping its relationship ID makes it display a *different image*. This alone caused every large image to morph from the wrong shape.
5. **z-order must match across paired elements.** If an element is layer 1 on one slide and last on the next, overlapping elements flicker at the midpoint.
6. **First animated element must have zero delay.** Do not add a global delay to wait for Morph; users perceive it as "text appears a beat after the transition".

## Workflow

### Step 0 — Protect the user's file (MANDATORY)

Never overwrite a file the user may have edited. Before any script runs:

1. Copy the input deck to a read-only baseline, e.g. `_source_latest.pptx`.
2. Verify `SRC` and `OUT` are different absolute paths (the build script asserts this).
3. Always write to a NEW versioned output (`_animated_v2.pptx`, etc.).

**Why this matters:** a patch script can fail silently and an unpatched copy can overwrite a hand-edited file. Confirm the values of `SRC`/`OUT` before executing any copied script, and stop if a patch fails.

### Step 1 — Inspect, name, prepare templates

Run `scripts/inspect_deck.py <input.pptx>` to dump per-slide shape counts, geometry, font sizes, names, and existing animation targets. Determine:
- Slide size `SW`, `SH` — auto-detected by all scripts, no manual config needed
- Slide count `N`
- Whether slides are already named with `!!` prefixes
- Whether there are existing animations/off-screen clones (needs stripping)

If `!!named=0`, run `scripts/apply_naming.py <baseline.pptx> <named.pptx>` first. This assigns deterministic page-unique names (`!!TITLE_Pn`, `!!IMG_Pn_k`, `!!BG_Pn_k`, `!!HL_Pn_k`, `!!VL_Pn_k`, `!!C_Pn_k`) so Morph never smart-matches the wrong shapes. Use the named deck as `SRC`.

The build script needs real animation XML for presetIDs 10/42/55/2/53. If the source lacks them, run `scripts/prepare_templates.py --src <baseline.pptx> --out <tpl.pptx> --donors <deckA.pptx> <deckB.pptx> ...` and pass `--tpl <tpl.pptx>` to the build. Donor decks can be any PPTX on the machine containing those effects. The tool classifies effects by **real visual behavior** (`fade` / `scale` / `motion`) and synthesizes a genuine `p:animScale` Zoom when donors only contain fade blocks — never just rename `presetID` on a fade effect.

If the deck is already animated or re-animating, the build's strip pass removes old `<p:timing>` and generated off-screen clones while preserving on-screen edits.

### Step 2 — Build

`scripts/build_morph_anim.py` is the production script. Configure at the top:

```
SRC          = read-only baseline deck
REF          = clean naming reference (for strip pass; use the pre-animation deck)
TPL_SRC      = a deck containing presetID 10/42/55/2/53 animation templates
OUT          = new output file (never SRC)
SW, SH       = slide size in pt
PAGE_MAIN    = per-page main fly-in direction
```

The build runs these passes in order; each addresses a specific observed defect:

| Pass | Purpose | Defect it prevents |
|---|---|---|
| -1 | Strip old `<p:timing>` and old off-screen clones; keep all on-screen (user-edited) shapes | Re-animating without duplicates/pileup |
| 0 | Area-tiered naming: keep `!!NAV_ITEM_n` only if area ≥ 20% of page; page-uniquify small decorations | Small decorations cross-morphing into big blocks |
| rels | Build `RELS[]` map; `fix_picture_rels()` remaps `r:embed` on every copied `<p:pic>` | Images showing wrong picture and morphing from wrong shape |
| 1 | Fly-in: for each adjacent pair, copy next-slide non-tertiary content to previous slide off-screen, same name. Neighbor elements in the same row are grouped to one direction. | Elements suddenly appearing with no transition |
| 2 | Fly-out: copy major current-slide elements to next slide off-screen. Exception: page-1 decorative ellipses must always fly out even if small. | Ellipse↔rectangle forced morph (rounded-corner artifacts) |
| 3 | In-page animations: page 1 gets all on-screen elements animated and no Morph; other pages animate only tertiary text. Build `<p:timing>` with 5 rotating rhythms, first delay = 0, normalize every effect duration to 400 ms, and keep a real visual behavior (`animEffect` / `animScale` / `animMotion`) for each effect. | Monotonous timing, text appearing before its card, text "popping in" from donor fade templates |
| 4 | z-order alignment: fix on-screen order, insert off-screen clones at the paired element's percentile on the adjacent page (fallback to global percentile if pair is also off-screen) | Layer flicker at transition midpoint |
| 5 | Anchored-text layer fix: text sitting on a color block must be above it | White text hidden or "popping" during morph |

### Step 3 — Verify (MANDATORY; run both scripts)

1. `scripts/verify.py <animated.pptx> --src <baseline.pptx>` — text/picture/z-order/timing checks.
2. `scripts/validate_openxml.py <animated.pptx>` — strict OOXML schema validation using the DocumentFormat.OpenXml assembly shipped with Microsoft Office (requires `pip install pythonnet`). This catches issues that zip/XML well-formedness checks and python-pptx miss. If Office is not installed, report that this check was skipped.
3. `scripts/check_powerpoint_open.py <animated.pptx>` — open the file in a real PowerPoint instance (read-only) to catch repair prompts caused by bad relationship targets. Schema validators alone will not catch these. If PowerPoint is not installed, report that this check was skipped.

Check all of the following and do not deliver if any fail:

- Every on-screen text string matches the source deck exactly (user's edits preserved)
- All cross-slide picture references point to the same media target as their source
- XML parses; no duplicate `cNvPr id` or `cNvPr name`; zip integrity OK
- Page 1 has no Morph; all other pages have Morph
- First animation delay = 0; page total < 4 s
- No same-name pair has z-order percentile difference > 25%
- No rect↔ellipse geometry conflict among same-name pairs on adjacent slides

Note: some page-info APIs report `anim_count=0` on Morph slides even when animations exist in XML. Verify at the XML level, not through those tools. Editing XML must use string/regex operations — ElementTree breaks the `mc/p14/p159` namespace prefixes.

### Step 4 — Deliver

Present the NEW output file with `present_files`. Tell the user which version to review. Only overwrite the "official" filename after the user explicitly confirms.

## Re-application on a user-edited deck

This is a common request ("I edited the deck, redo the animation"). Procedure:
1. Copy the edited deck to `_source_latest.pptx` / `_source_vNN.pptx` — this is the read-only baseline.
2. Point `SRC` at the baseline, `OUT` at a new version.
3. The strip pass removes old animation and off-screen clones while preserving every on-screen shape and font change. Do **not** regenerate from the pristine original.
4. Verify text equality against the baseline, then deliver.

## Compatibility notes

- The Morph wrapper uses PowerPoint's native MCE structure: `p159` is declared on `mc:Choice`, `p`/`p14` locally on `p:transition`, `mc:Fallback` uses a local `p` declaration, and no `mc:Ignorable` or `xmlns=""` is added to the slide root. This is accepted by strict Office/WPS parsers.
- Never fabricate a preset by replacing only `presetID="53"` on a fade block; short rating words will still fade or appear to pop in. Use a real `p:animScale` effect.
- WPS may ignore Morph and show a plain fade; that is expected fallback behavior, not a repair condition.
- Media/audio pictures have `r:link`/`r:id` in addition to `r:embed`. The build script remaps all of them and preserves relationship Type/TargetMode (see `references/pitfalls.md` P16). Never revert to embed-only remapping.
- If PowerPoint still reports "content was removed / needs repair", make isolation builds: (a) animations removed, (b) Morph transitions removed, (c) off-screen clones removed, then (d) shape clones only vs picture clones only. Open each with PowerPoint COM (`Presentations.Open`) to find the failing feature automatically; schema validators alone do not catch bad relationship targets.
- Never overwrite the user's hand-edited file; always deliver a new versioned output.

## Resources

- `scripts/build_morph_anim.py` — production build script (all passes above; slide size auto-detected)
- `scripts/apply_naming.py` — deterministic `!!` naming for un-named decks
- `scripts/prepare_templates.py` — build the `--tpl` deck from source + donor decks
- `scripts/validate_openxml.py` — strict OOXML validation via the Office assembly
- `scripts/check_powerpoint_open.py` — real PowerPoint open test (detects repair prompts)
- `scripts/verify.py` — completeness/consistency verifier
- `scripts/inspect_deck.py` — shape/text/animation dump for planning
- `references/concepts.md` — detailed explanation of Morph mechanics, naming rules, animation presetIDs, and timing rhythms
- `references/pitfalls.md` — common hard-won failure modes and their exact fixes
- `references/config.md` — parameter reference (thresholds, directions, rhythm array, durations)

## Notes

- Cannot render actual PowerPoint playback in this environment. Always ask the user to confirm the visual result on specific slide transitions.
- The `!!` convention is documented at https://support.microsoft.com/en-us/office/morph-transition-tips-and-tricks
