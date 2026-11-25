# -*- coding: utf-8 -*-
"""
Inspect a PPTX deck before applying animation.

Usage:
  python inspect_deck.py <deck.pptx> [--width 1152] [--height 648] [--page N]

Prints per-slide: shape count, on/off-screen split, !! naming coverage,
animation target count, Morph presence; optional detailed shape dump for one page.
"""
import zipfile, re, argparse

def parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('--width', type=float, default=0.0, help='slide width in pt (auto-detected if omitted)')
    ap.add_argument('--height', type=float, default=0.0, help='slide height in pt (auto-detected if omitted)')
    ap.add_argument('--page', type=int, default=None, help='dump all shapes on this page')
    return ap.parse_args()

A = parse_args()
z = zipfile.ZipFile(A.deck)
SW, SH = A.width, A.height
if not SW or not SH:
    _s = z.read('ppt/presentation.xml').decode('utf-8', 'ignore')
    _m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', _s)
    SW = float(_m.group(1)) / 12700 if _m else 960.0
    SH = float(_m.group(2)) / 12700 if _m else 540.0
slides = sorted([x for x in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', x)],
                key=lambda x: int(re.search(r'\d+', x).group()))
print(f'deck={A.deck}  slides={len(slides)}  size={SW}x{SH} pt')

for path in slides:
    p = int(re.search(r'\d+', path).group())
    s = z.read(path).decode('utf-8', 'ignore')
    rows = []
    for m in re.finditer(r'<(p:sp|p:pic)\b.*?</\1>', s, re.S):
        b = m.group(0)
        c = re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"', b)
        o = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', b)
        e = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', b)
        if not c: continue
        x = int(o.group(1))/12700 if o else 0; y = int(o.group(2))/12700 if o else 0
        w = int(e.group(1))/12700 if e else 0; h = int(e.group(2))/12700 if e else 0
        g = re.search(r'<a:prstGeom prst="([^"]+)"', b)
        fss = [int(t)/100 for t in re.findall(r'<a:rPr[^>]*sz="(\d+)"', b)]
        txt = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', b)).strip()
        on = -10 <= x+w/2 <= SW+10 and -10 <= y+h/2 <= SH+10
        rows.append(dict(id=c.group(1), name=c.group(2), x=x, y=y, w=w, h=h,
                         geom=g.group(1) if g else '-', fs=max(fss) if fss else 0,
                         txt=txt, on=on, tag=m.group(1)))
    on_n = sum(1 for r in rows if r['on'])
    named = sum(1 for r in rows if r['name'].startswith('!!'))
    t = re.search(r'<p:timing>.*?</p:timing>', s, re.S)
    aids = len(set(re.findall(r'<p:spTgt spid="(\d+)"', t.group(0)))) if t else 0
    print(f'P{p:2} shapes={len(rows):3} on-screen={on_n:3} off={len(rows)-on_n:3} '
          f'!!named={named:3} animations={aids:3} morph={"yes" if "morph" in s else "no"}')
    if A.page == p:
        for r in sorted(rows, key=lambda a:(a['y'], a['x'])):
            flag = 'ON ' if r['on'] else 'OFF'
            print(f'   {flag} {r["tag"][2:]} id={r["id"]:>4} '
                  f'x={r["x"]:6.0f} y={r["y"]:6.0f} w={r["w"]:5.0f} h={r["h"]:5.0f} '
                  f'fs={r["fs"]:5.1f} geom={r["geom"]:10} {r["name"][:22]:22} | {r["txt"][:36]}')
z.close()
