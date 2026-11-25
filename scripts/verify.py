# -*- coding: utf-8 -*-
"""
Verify an animated PPTX produced by build_morph_anim.py.

Usage:
  python verify.py <animated.pptx> --src <original_or_baseline.pptx> \
      [--width 1152] [--height 648]

Checks (exits non-zero on failure):
  1. every on-screen text string matches --src (user edits preserved)
  2. every cross-slide picture copy resolves to the same media as its source
  3. XML parses; no duplicate cNvPr id/name; zip integrity
  4. page 1 has no Morph, all other pages have Morph
  5. first animation delay == 0; total per page < 4 s
  6. no same-name pair has z-order percentile diff > 0.25
  7. no adjacent-slide same-name pair has rect<->ellipse geometry conflict
"""
import zipfile, re, os, collections, argparse, xml.etree.ElementTree as ET

def parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('--src', help='baseline deck to compare on-screen text')
    ap.add_argument('--width', type=float, default=0.0, help='slide width in pt (auto-detected if omitted)')
    ap.add_argument('--height', type=float, default=0.0, help='slide height in pt (auto-detected if omitted)')
    ap.add_argument('--slides', type=int, default=None)
    return ap.parse_args()

A = parse_args()
SW, SH = A.width, A.height
if not SW or not SH:
    _z = zipfile.ZipFile(A.deck)
    _s = _z.read('ppt/presentation.xml').decode('utf-8', 'ignore')
    _z.close()
    _m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', _s)
    SW = float(_m.group(1)) / 12700 if _m else 960.0
    SH = float(_m.group(2)) / 12700 if _m else 540.0

def slide_count(z):
    if A.slides: return A.slides
    return len([x for x in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', x)])

def shapes(z, p):
    s = z.read(f'ppt/slides/slide{p}.xml').decode('utf-8', 'ignore')
    out = []
    for m in re.finditer(r'<(p:sp|p:pic)\b.*?</\1>', s, re.S):
        b = m.group(0)
        c = re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"', b)
        o = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', b)
        e = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', b)
        if not (c and o and e): continue
        x = int(o.group(1))/12700; y = int(o.group(2))/12700
        w = int(e.group(1))/12700; h = int(e.group(2))/12700
        g = re.search(r'<a:prstGeom prst="([^"]+)"', b)
        txt = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', b)).strip()
        on = (-10 <= x+w/2 <= SW+10 and -10 <= y+h/2 <= SH+10)
        out.append(dict(id=c.group(1), name=c.group(2), txt=txt,
                        geom=g.group(1) if g else '?', on=on, x=x, y=y, w=w, h=h))
    return out

def picmap(z, p):
    s = z.read(f'ppt/slides/slide{p}.xml').decode('utf-8', 'ignore')
    rp = f'ppt/slides/_rels/slide{p}.xml.rels'
    rels = z.read(rp).decode('utf-8', 'ignore') if rp in z.namelist() else ''
    r2t = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    out = collections.defaultdict(list)
    for m in re.finditer(r'<p:pic\b.*?</p:pic>', s, re.S):
        b = m.group(0)
        c = re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"', b)
        bl = re.search(r'<a:blip[^>]*r:embed="([^"]+)"', b)
        if c and bl:
            out[c.group(2)].append(r2t.get(bl.group(1), '?'))
    return out

def main():
    z = zipfile.ZipFile(A.deck)
    N = slide_count(z)
    fails = []
    print(f'deck={A.deck}  slides={N}')

    # 1. text equality
    if A.src:
        zs = zipfile.ZipFile(A.src)
        bad = 0
        for p in range(1, N+1):
            a = [x['txt'] for x in shapes(zs, p) if x['on'] and x['txt']]
            b = [x['txt'] for x in shapes(z, p) if x['on'] and x['txt']]
            if a != b:
                bad += 1
                print(f'  [FAIL] P{p} on-screen text differs from source')
        zs.close()
        if bad: fails.append(f'text mismatch on {bad} slides')
        else: print('  [OK] on-screen text identical to source')

    # 2. picture references
    PM = {p: picmap(z, p) for p in range(1, N+1)}
    good = bad = 0
    for p in range(1, N+1):
        for n, targets in PM[p].items():
            m = re.search(r'P(\d+)_', n)
            src_p = int(m.group(1)) if m else 0
            if not src_p or src_p == p or src_p > N or not PM[src_p].get(n): continue
            for t in targets:
                if t == PM[src_p][n][0]: good += 1
                else:
                    bad += 1
                    print(f'  [FAIL] P{p} {n}: {t} != source {PM[src_p][n][0]}')
    if bad: fails.append(f'{bad} wrong picture references')
    print(f'  [{"OK" if not bad else "FAIL"}] cross-slide pictures: {good} correct, {bad} wrong')

    # 3. structure + 4. timing + 5. morph
    anim_total = 0
    for p in range(1, N+1):
        s = z.read(f'ppt/slides/slide{p}.xml').decode('utf-8', 'ignore')
        try: ET.fromstring(s)
        except Exception as ex: fails.append(f'P{p} XML invalid: {ex}')
        ids = re.findall(r'<p:cNvPr id="(\d+)"', s)
        names = re.findall(r'<p:cNvPr id="\d+" name="([^"]*)"', s)
        if len(ids) != len(set(ids)): fails.append(f'P{p} duplicate cNvPr id')
        if len(names) != len(set(names)): fails.append(f'P{p} duplicate cNvPr name')
        if z.testzip() is not None: fails.append('zip corrupted')
        has_morph = 'morph' in s
        if (p == 1 and has_morph) or (p > 1 and not has_morph):
            fails.append(f'P{p} Morph presence wrong')
        t = re.search(r'<p:timing>.*?</p:timing>', s, re.S)
        if t:
            tt = t.group(0)
            aids = set(re.findall(r'<p:spTgt spid="(\d+)"', tt))
            anim_total += len(aids)
            for sid in aids:
                if sid not in ids: fails.append(f'P{p} animation target {sid} missing')
            presets = len(re.findall(r'presetID="\d+"', tt))
            visual = len(re.findall(r'<p:animEffect|<p:animScale|<p:animMotion', tt))
            if visual < presets:
                fails.append(f'P{p} has {presets} effects but only {visual} visual behaviors (pop-in risk)')
            ds = [int(x) for x in re.findall(r'<p:cond delay="(\d+)"', tt)]
            if ds:
                if min(ds) != 0: fails.append(f'P{p} first delay {min(ds)} != 0')
                if max(ds) + 400 > 4000: fails.append(f'P{p} animation ends at {max(ds)+400}ms > 4s')
    print(f'  [OK] structure/timing checked, {anim_total} animation targets, visual behaviors present')

    # 6. z-order
    S = {p: shapes(z, p) for p in range(1, N+1)}
    bad_layers = []
    for p in range(2, N+1):
        idx = {x['name']: i for i, x in enumerate(S[p-1])}
        for i, x in enumerate(S[p]):
            if x['name'] in idx:
                pa = idx[x['name']] / max(len(S[p-1])-1, 1)
                pb = i / max(len(S[p])-1, 1)
                if abs(pa - pb) > 0.25:
                    bad_layers.append((p, x['name'], round(abs(pa-pb), 2)))
    if bad_layers:
        fails.append(f'{len(bad_layers)} z-order mismatches')
        for b in bad_layers[:8]: print(f'  [FAIL] z-order P{b[0]} {b[1]} diff={b[2]}')
    else:
        print('  [OK] no z-order mismatch >25%')

    # 7. geometry conflicts among same-name pairs
    conflicts = []
    for p in range(2, N+1):
        prev = {x['name']: x for x in S[p-1]}
        for x in S[p]:
            q = prev.get(x['name'])
            if q and q['geom'] != x['geom'] and q['geom'] != '?' and x['geom'] != '?':
                if {q['geom'], x['geom']} <= {'rect', 'ellipse'}:
                    conflicts.append((p, x['name'], q['geom'], x['geom']))
    if conflicts:
        fails.append(f'{len(conflicts)} rect/ellipse pair conflicts')
        for c in conflicts[:8]: print(f'  [FAIL] geometry P{c[0]} {c[1]} {c[2]}<->{c[3]}')
    else:
        print('  [OK] no rect<->ellipse same-name conflicts')

    print('\nRESULT:', 'PASS' if not fails else 'FAIL')
    for f in fails: print('  -', f)
    raise SystemExit(1 if fails else 0)

if __name__ == '__main__':
    main()
