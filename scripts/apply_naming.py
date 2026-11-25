# -*- coding: utf-8 -*-
"""Give every shape in a deck a deterministic !! name for reliable Morph pairing.

Usage:
  python apply_naming.py <input.pptx> <output.pptx> [--width N] [--height N]

Slide size is auto-detected from ppt/presentation.xml unless --width/--height
are given. The source is never modified; SRC and OUT must differ.

Naming scheme (all page-unique unless noted):
  !!TITLE_Pn      largest eligible text on the page
  !!IMG_Pn_k      pictures
  !!BG_Pn_k       near-fullscreen shapes
  !!HL_Pn_k       thin horizontal lines
  !!VL_Pn_k       thin vertical lines
  !!C_Pn_k        other content shapes
"""
import zipfile, re, os, argparse, tempfile

EMU_PER_PT = 12700
SHAPE_RE = re.compile(r'<(p:sp|p:pic)\b.*?</\1>', re.S)
CNV_RE = re.compile(r'<p:cNvPr id="(\d+)" name="([^"]*)"')


def detect_size(z):
    s = z.read('ppt/presentation.xml').decode('utf-8', 'ignore')
    m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', s)
    if not m:
        return 960.0, 540.0
    return int(m.group(1)) / EMU_PER_PT, int(m.group(2)) / EMU_PER_PT


def parse(block):
    m = CNV_RE.search(block)
    if not m:
        return None
    o = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', block)
    e = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', block)
    x = int(o.group(1)) / EMU_PER_PT if o else 0.0
    y = int(o.group(2)) / EMU_PER_PT if o else 0.0
    w = int(e.group(1)) / EMU_PER_PT if e else 0.0
    h = int(e.group(2)) / EMU_PER_PT if e else 0.0
    txt = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', block)).strip()
    fss = [int(t) / 100 for t in re.findall(r'<a:rPr[^>]*sz="(\d+)"', block)]
    return dict(cnv=m, x=x, y=y, w=w, h=h, txt=txt,
                fs=max(fss) if fss else 0.0, pic='<p:pic' in block)


def name_shapes(xml, page, sw, sh):
    infos = []
    for m in SHAPE_RE.finditer(xml):
        d = parse(m.group(0))
        if d:
            infos.append((m, d))
    candidates = [d for _, d in infos
                  if d['txt'] and len(d['txt']) > 3 and d['y'] < sh * 0.55 and d['fs'] >= 14]
    title = max(candidates, key=lambda d: d['fs']) if candidates else None

    counters = {}
    out = xml
    for m, d in sorted(infos, key=lambda pair: pair[0].start(), reverse=True):
        if title is not None and d is title:
            name = f'!!TITLE_P{page}'
        else:
            if d['pic']:
                prefix = 'IMG'
            elif d['w'] >= sw * 0.94 and d['h'] >= sh * 0.94:
                prefix = 'BG'
            elif d['h'] <= 4 and d['w'] >= 100:
                prefix = 'HL'
            elif d['w'] <= 4 and d['h'] >= 10:
                prefix = 'VL'
            else:
                prefix = 'C'
            counters[prefix] = counters.get(prefix, 0) + 1
            name = f'!!{prefix}_P{page}_{counters[prefix]}'
        old = d['cnv'].group(0)
        new = old.replace(f'name="{d["cnv"].group(2)}"', f'name="{name}"', 1)
        out = out[:m.start()] + m.group(0).replace(old, new, 1) + out[m.end():]
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('out')
    ap.add_argument('--width', type=float, default=None)
    ap.add_argument('--height', type=float, default=None)
    a = ap.parse_args()
    assert os.path.abspath(a.src) != os.path.abspath(a.out), 'src and out must differ'

    z = zipfile.ZipFile(a.src)
    sw, sh = a.width or 0, a.height or 0
    if not sw or not sh:
        sw, sh = detect_size(z)
    entries = [(i, z.read(i.filename)) for i in z.infolist()]
    z.close()

    named = 0
    with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info, data in entries:
            m = re.fullmatch(r'ppt/slides/slide(\d+)\.xml', info.filename)
            if m:
                xml = data.decode('utf-8', 'ignore')
                before = len(re.findall(r'name="!!', xml))
                xml = name_shapes(xml, int(m.group(1)), sw, sh)
                named += len(re.findall(r'name="!!', xml)) - before
                data = xml.encode('utf-8')
            zo.writestr(info, data)
    print(f'slide size {sw:.0f}x{sh:.0f} pt; assigned {named} names -> {a.out}')


if __name__ == '__main__':
    main()
