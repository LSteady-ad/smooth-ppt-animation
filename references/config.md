# Build Configuration Reference

All constants live at the top of `scripts/build_morph_anim.py`. Tune them together, then re-verify.

## Paths

| Variable | Meaning |
|---|---|
| `SRC` | Read-only baseline deck. NEVER equal to `OUT`. |
| `REF` | Clean (pre-animation) deck used by the strip pass to decide which off-screen shapes are original vs generated clones. |
| `TPL_SRC` | Any deck containing animation templates with presetID 10, 42, 55, 2, 53. Templates are extracted by regex at startup. |
| `WORK` | Temporary unpack directory, wiped on each run. |
| `OUT` | New versioned output file. |

## Slide geometry

```
SW, SH = 1152, 648   # points, 16:9
EMU_PER_PT = 12700
```

Change `SW`/`SH` to match the deck. All off-screen calculations and the on-screen test derive from these.

## Direction ring

```python
RING = ['R','RD','D','LD','L','LU','U','RU']   # clockwise

PAGE_MAIN = {
  1:'R', 2:'U', 3:'L', 4:'D', 5:'RU', 6:'LD', 7:'LU', 8:'RD',
  9:'R', 10:'L', 11:'U', 12:'D', 13:'RU', 14:'LD',
}
```

`PAGE_MAIN[n]` is the fly-in direction for slide n+1 coming from slide n. Fly-out uses the opposite: `RING[(index(PAGE_MAIN[A])+4) % 8]`. Arrays longer than the deck are fine; extend them for larger decks.

`off_pos(direction, shape, gap)` returns the off-screen `(x, y)` in points. Gap values: fly-in `80 + ((row + page_no) % 3) * 90`; fly-out `80 + idx * 95`.

## Thresholds

| Constant | Value | Role |
|---|---|---|
| Large-area pairing | 0.20 × page | `!!NAV_ITEM_n` kept cross-slide if area ≥ 20% |
| Color block detection | 0.20–0.85 × page | Local color blocks for anchoring (excludes full-screen BG) |
| Fly-out area | ≥ 0.15 × page | Major elements that get a fly-out clone |
| Card container | 300≤w≤700, 40≤h≤120 | Same-size card rows for the tertiary-text consistency rule |
| Table cell | w≥150, 40≤h≤130, empty `p:sp` | Rating-word exception |
| Small graphic | ≤140×90, picture or empty shape | "Small graphic + short label" level-2 rule |
| Short label | ≤8 chars, no `，。；` | Level-2 short-title / graphic-label rule |
| Line-grouping | y diff <28 pt, gap −10..100 pt | Same-row neighbor grouping for one fly-in direction |
| Description column cluster | x-center diff ≤90 pt | Group descriptions by card column |

## Tertiary animation

| Constant | Value |
|---|---|
| `dur` | 400 ms per effect |
| `base` (group stagger) | 230 ms |
| Description intra-delay | 90 ms per element |
| Rating-word intra-delay | 110 ms per element |
| `base_delay` | 0 (first effect immediate) |
| Max groups per page | 4 |

Rhythm array (rotate across pages; index `(page-1) % len`):

```python
RHYTHM = ['uniform','seq','reverse','wave','relay','seq','seq','reverse',
          'relay','wave','seq','uniform','reverse','relay','uniform']
```

Chapter-based animation types (when used): chapter 01 → Fade (10), 02 → Float Up (42), 03 → Expand (55), 04 → Fly In (2). The current production path uses Fade for descriptions and Zoom for short ratings, which readers found clearer than chapter-based types.

## PresetID extraction

`extract_tpls()` scans `TPL_SRC` for the first `<p:par><p:cTn presetID="N" ...>` longer than 500 chars for each ID in `[10, 42, 55, 2, 53]`. If a template is missing, `KeyError` occurs — add a deck with that effect to `TPL_SRC` or drop the ID from the list.

## verification thresholds

See `scripts/verify.py`:

- z-order percentile difference > 0.25 → fail
- Animation end (max delay + 400 ms) > 4000 ms → fail
- Any cross-slide picture target mismatch → fail
- Any on-screen text string differing from SRC → fail
- Same-name `prstGeom` mismatch (rect vs ellipse) on adjacent slides → fail
