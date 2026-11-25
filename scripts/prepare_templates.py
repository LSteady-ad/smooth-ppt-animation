# -*- coding: utf-8 -*-
"""Build the --tpl deck required by build_morph_anim.py.

The build script extracts real PowerPoint animation XML for presetIDs
10/42/55/2/53 from a template deck. Many decks only contain fade effects
even when the presetID says Zoom/Float/Grow; blindly cloning a fade block
and changing presetID makes short rating words look like they just pop in.

This tool harvests effects, classifies them by their REAL visual behavior,
and synthesizes a genuine scale animation when a real Zoom/Grow block is
missing. All generated blocks share one valid PowerPoint skeleton.

Usage:
  python prepare_templates.py --src <baseline.pptx> --out <tpl.pptx> \
      [--donors deckA.pptx deckB.pptx ...]
"""
import zipfile, re, os, argparse

PIDS = (10, 42, 55, 2, 53)
PREF = re.compile(r'<p:par>(<p:cTn[^>]*presetID="(\d+)".*?</p:cTn>)</p:par>', re.S)


def visual_kind(block):
    if '<p:animScale' in block:
        return 'scale'
    if '<p:animMotion' in block:
        return 'motion'
    if 'filter="fade"' in block:
        return 'fade'
    return 'none'


def harvest(path):
    found = {}
    z = zipfile.ZipFile(path)
    for n in z.namelist():
        if not re.fullmatch(r'ppt/slides/slide\d+\.xml', n):
            continue
        s = z.read(n).decode('utf-8', 'ignore')
        for m in PREF.finditer(s):
            pid = int(m.group(2))
            block = m.group(1)
            rank = {'scale': 3, 'motion': 2, 'fade': 1, 'none': 0}[visual_kind(block)]
            old = found.get(pid)
            if old is None or (rank, len(block)) > (
                    {'scale': 3, 'motion': 2, 'fade': 1, 'none': 0}[visual_kind(old)], len(old)):
                found[pid] = block
    z.close()
    return found


def scale_effect(cid, from_pct, to_pct):
    return (
        f'<p:animScale zoomContents="1"><p:cBhvr><p:cTn id="{cid}" dur="500" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl><p:spTgt spid="__SPID__"/></p:tgtEl></p:cBhvr>'
        f'<p:from x="{from_pct}" y="{from_pct}"/><p:to x="{to_pct}" y="{to_pct}"/>'
        '</p:animScale>'
    )


def motion_effect(cid, y_delta):
    return (
        f'<p:animMotion origin="layout" path="M 0 0 L 0 {y_delta} E" pathEditMode="relative" rAng="0">'
        f'<p:cBhvr><p:cTn id="{cid}" dur="500" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl><p:spTgt spid="__SPID__"/></p:tgtEl>'
        '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
        '</p:cBhvr></p:animMotion>'
    )


def make_block(pid, skeleton, source_block):
    block = re.sub(r'presetID="\d+"', f'presetID="{pid}"', skeleton, count=1)
    m = re.search(r'(<p:childTnLst>)(.*?)(</p:childTnLst>)', block, re.S)
    if not m:
        raise SystemExit('animation skeleton has no <p:childTnLst>')
    inner = m.group(2)
    sm = re.search(r'<p:set>.*?</p:set>', inner, re.S)
    if not sm:
        raise SystemExit('animation skeleton has no visibility <p:set>')
    set_block = sm.group(0)

    kind = visual_kind(source_block)
    if pid == 10:
        em = re.search(r'<p:animEffect[^>]*>.*?</p:animEffect>', source_block, re.S)
        if not em:
            raise SystemExit('no real fade effect found for presetID 10')
        eff = em.group(0)
    elif pid == 53:
        if kind == 'scale':
            em = re.search(r'<p:animScale.*?</p:animScale>', source_block, re.S)
            eff = em.group(0) if em else scale_effect(900053, '70000', '100000')
        else:
            eff = scale_effect(900053, '70000', '100000')
    elif pid == 55:
        eff = scale_effect(900055, '0', '100000')
    elif pid == 42:
        eff = motion_effect(900042, '0.10') + scale_effect(900142, '90000', '100000')
    else:  # pid == 2
        eff = motion_effect(900002, '-0.35')
    eff = re.sub(r'spid="\d+"', 'spid="__SPID__"', eff)
    return block[:m.start(2)] + set_block + eff + block[m.end(2):]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', required=True, help='read-only baseline deck')
    ap.add_argument('--out', required=True, help='template deck to create')
    ap.add_argument('--donors', nargs='*', default=[])
    a = ap.parse_args()
    assert os.path.abspath(a.src) != os.path.abspath(a.out), 'src and out must differ'

    pool = {}
    for f in [a.src] + a.donors:
        if os.path.exists(f):
            for pid, block in harvest(f).items():
                old = pool.get(pid)
                if old is None or (visual_kind(block), len(block)) > (visual_kind(old), len(old)):
                    pool[pid] = block
    if not pool:
        raise SystemExit('no animation presetID templates found in source or donors')

    fade_blocks = [(v, visual_kind(v)) for v in pool.values() if visual_kind(v) == 'fade']
    if not fade_blocks:
        raise SystemExit('no real fade skeleton found; cannot build templates')
    skeleton = max(fade_blocks, key=lambda x: len(x[0]))[0]
    skeleton = re.sub(r'presetID="\d+"', 'presetID="10"', skeleton, count=1)

    tpls = {}
    for pid in PIDS:
        source_block = pool.get(pid, skeleton)
        tpls[pid] = make_block(pid, skeleton, source_block)
        print(f'presetID {pid}: {visual_kind(source_block)} -> {visual_kind(tpls[pid])}')

    blocks = ''.join(f'<p:par>{tpls[p]}</p:par>' for p in PIDS)
    z = zipfile.ZipFile(a.src)
    with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == 'ppt/slides/slide1.xml':
                s = data.decode('utf-8', 'ignore')
                s = s.replace('</p:sld>', blocks + '</p:sld>', 1)
                data = s.encode('utf-8')
            zo.writestr(info, data)
    z.close()
    print('template deck:', a.out)


if __name__ == '__main__':
    main()
