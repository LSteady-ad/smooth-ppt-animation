# -*- coding: utf-8 -*-
"""
Build a Morph-animated PowerPoint deck from an existing .pptx.

Pipeline: strip old animation/off-screen clones -> keep on-screen (user-edited)
content -> apply naming tiers, cross-slide fly-in/out, selective in-page
animations, z-order alignment, and layer fixes -> write a NEW output file.

Usage:
  python build_morph_anim.py --src <input.pptx> --out <output.pptx> \
      [--ref <clean_named_deck.pptx>] [--tpl <deck_with_animation_templates.pptx>] \
      [--width 1152] [--height 648]

  --ref   : a pre-animation deck with the !! naming scheme; used by the strip
            pass to distinguish original off-screen shapes from generated clones.
            If omitted, the strip pass keeps ALL off-screen shapes.
  --tpl   : any deck containing presetID 10/42/55/2/53 animation templates.
            If omitted, templates are extracted from --src.

The source is NEVER modified. SRC and OUT must be different absolute paths.
See SKILL.md and references/ for the full methodology.
"""
import zipfile, re, os, shutil, itertools, collections, argparse, tempfile

def parse_args():
    ap = argparse.ArgumentParser(description='Add Morph animation to a PPTX deck')
    ap.add_argument('--src', required=True, help='read-only source deck')
    ap.add_argument('--out', required=True, help='new output deck (must differ from --src)')
    ap.add_argument('--ref', default=None, help='clean named deck for the strip pass')
    ap.add_argument('--tpl', default=None, help='deck containing animation templates')
    ap.add_argument('--width', type=float, default=0.0, help='slide width in points (auto-detected if omitted)')
    ap.add_argument('--height', type=float, default=0.0, help='slide height in points (auto-detected if omitted)')
    ap.add_argument('--slides', type=int, default=None, help='slide count (default: auto-detect)')
    ap.add_argument('--work', default=None, help='temp unpack dir (default: auto)')
    return ap.parse_args()

ARGS = parse_args()
SRC = ARGS.src
OUT = ARGS.out
REF = ARGS.ref
TPL_SRC = ARGS.tpl or SRC
SW, SH = ARGS.width, ARGS.height
if not SW or not SH:
    _z = zipfile.ZipFile(SRC)
    _s = _z.read('ppt/presentation.xml').decode('utf-8', 'ignore')
    _m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', _s)
    _z.close()
    SW = float(_m.group(1)) / 12700 if _m else 960.0
    SH = float(_m.group(2)) / 12700 if _m else 540.0
print(f'slide size: {SW:.0f}x{SH:.0f} pt')
WORK = ARGS.work or os.path.join(tempfile.gettempdir(), 'smooth_ppt_animation_work')
assert os.path.abspath(SRC) != os.path.abspath(OUT), 'source and output must differ'

def slide_count():
    if ARGS.slides: return ARGS.slides
    z = zipfile.ZipFile(SRC)
    n = len([x for x in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', x)])
    z.close()
    return n
N = slide_count()

# 章节 -> 页内动画 presetID
# 01(淡入10) 02(上浮42) 03(展开55) 04(飞入2)
CHAPTER_ANIM = {'01': 10, '02': 42, '03': 55, '04': 2}
DEFAULT_ANIM = 10

# 8 方向环（顺时针），R=右 L=左 U=上 D=下，两字母=斜角
RING = ['R','RD','D','LD','L','LU','U','RU']
# 每页(A=1..14)的主飞入方向：逐页变化，避免总是右边/上面
PAGE_MAIN = {
    1:'R', 2:'U', 3:'L', 4:'D', 5:'RU', 6:'LD', 7:'LU', 8:'RD',
    9:'R', 10:'L', 11:'U', 12:'D', 13:'RU', 14:'LD',
}

def off_pos(ds, bs, gap):
    """按方向把形状放到屏外 gap 距离处"""
    bx,by,bw,bh = bs['x'],bs['y'],bs['w'],bs['h']
    if   ds=='R':  return (SW+gap, by)
    elif ds=='L':  return (-bw-gap, by)
    elif ds=='U':  return (bx, -bh-gap)
    elif ds=='D':  return (bx, SH+gap)
    elif ds=='RU': return (SW+gap, -bh-gap)
    elif ds=='RD': return (SW+gap, SH+gap)
    elif ds=='LU': return (-bw-gap, -bh-gap)
    else:          return (-bw-gap, SH+gap)

# 动画模板
def extract_tpls():
    z=zipfile.ZipFile(TPL_SRC)
    tpls={}
    for pid in [10,42,55,2,53]:
        for pn in range(1,N+1):
            s=z.read(f'ppt/slides/slide{pn}.xml').decode('utf-8','ignore')
            m=re.search(r'<p:par>(<p:cTn[^>]*presetID="'+str(pid)+r'".*?</p:cTn>)</p:par>',s,re.S)
            if m and len(m.group(1))>500:
                t=m.group(1)
                t=re.sub(r'spid="\d+"','spid="__SPID__"',t)
                t=re.sub(r' id="\d+"',' id="__ID__"',t)
                tpls[pid]=t; break
    z.close()
    return tpls

TPLS = extract_tpls()
print("模板:", {k:len(v) for k,v in TPLS.items()})
_missing = [p for p in (10, 53) if p not in TPLS]
if _missing:
    raise SystemExit('missing animation template presetID(s): ' + ', '.join(map(str, _missing))
                     + ' — run scripts/prepare_templates.py first')

# 解包
if os.path.exists(WORK): shutil.rmtree(WORK)
os.makedirs(WORK)
z=zipfile.ZipFile(SRC,'r')
for n in z.namelist():
    if n.endswith('/'):
        continue
    p=os.path.join(WORK,n.replace('/',os.sep)); os.makedirs(os.path.dirname(p),exist_ok=True)
    open(p,'wb').write(z.read(n))
z.close()

# ===== 阶段-1：从用户最新版剥离旧动画与旧屏外副本，保留用户文字/字体修改 =====
# 判定规则：屏内元素全部保留；屏外元素只有在干净参考稿中同名存在时才保留
# （如 !!TITLE_Pn、!!NAV_ITEM_n），其余都是之前脚本生成的屏外副本，删掉重建。
ref_names = {}
if REF:
    zref = zipfile.ZipFile(REF)
    for pn in range(1,N+1):
        ss = zref.read(f'ppt/slides/slide{pn}.xml').decode('utf-8','ignore')
        ref_names[pn] = set(re.findall(r'<p:cNvPr id="\d+" name="([^"]*)"', ss))
    zref.close()

removed_shapes = 0
for pn in range(1,N+1):
    fp = os.path.join(WORK, f'ppt/slides/slide{pn}.xml')
    ss = open(fp, encoding='utf-8').read()
    ss = re.sub(r'<p:timing>.*?</p:timing>', '', ss, flags=re.S)
    def clean_sh(m):
        global removed_shapes
        b = m.group(0)
        c = re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"', b)
        o = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', b)
        e = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', b)
        if not (c and o and e):
            return b
        x = int(o.group(1))/12700; y = int(o.group(2))/12700
        w = int(e.group(1))/12700; h = int(e.group(2))/12700
        on = (-10 <= x+w/2 <= SW+10 and -10 <= y+h/2 <= SH+10)
        # Only strip off-screen clones when a --ref deck was supplied.
        # Without --ref, keep all shapes (nothing is known to be generated).
        if ref_names and (not on) and c.group(2) not in ref_names.get(pn, set()):
            removed_shapes += 1
            return ''
        return b
    ss = re.sub(r'<(p:sp|p:pic)\b.*?</\1>', clean_sh, ss, flags=re.S)
    open(fp, 'w', encoding='utf-8').write(ss)
print('剥离旧屏外副本:', removed_shapes)

def parse_shapes(s):
    out=[]
    for m in re.finditer(r'<(p:sp|p:pic)\b.*?</\1>', s, re.S):
        b=m.group(0)
        cnv=re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"',b)
        if not cnv: continue
        off=re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"',b)
        ext=re.search(r'<a:ext cx="(\d+)" cy="(\d+)"',b)
        x=int(off.group(1))/12700 if off else 0
        y=int(off.group(2))/12700 if off else 0
        w=int(ext.group(1))/12700 if ext else 0
        h=int(ext.group(2))/12700 if ext else 0
        txt=''.join(re.findall(r'<a:t>([^<]*)</a:t>',b)).strip()
        fss=[int(t)/100 for t in re.findall(r'<a:rPr[^>]*sz="(\d+)"',b)]
        out.append({'id':cnv.group(1),'name':cnv.group(2),'x':x,'y':y,'w':w,'h':h,
                    'txt':txt,'fs':max(fss) if fss else 0,'block':b,'tag':m.group(1),
                    'pic':m.group(1)=='p:pic'})
    return out

def is_frame(s):
    n=s['name']
    return n.startswith(('!!BG_','!!NAV_','!!FOOTER_','!!HL_','!!VL_'))

def is_title(s):
    return s['name'].startswith('!!TITLE_')

def is_content(s):
    n=s['name']
    if is_frame(s) or is_title(s): return False
    if s['x']>SW+50: return False  # 已在屏外
    return True

def is_big(s):
    """主要/大元素：适合跨页进出场"""
    if is_title(s): return True
    if s['name'].startswith('!!IMG_'): return True
    # 大形状：卡片容器/大图形
    if s['w']>=150 and s['h']>=60: return True
    # 核心文字（中标题字号>=15）
    if s['fs']>=15 and s['txt']: return True
    return False

def is_necessary(s, allshapes=None):
    """反向识别：二级及以上"必要元素" -> 平滑；其余有文字者才是三级动画。
    必要 = 页面/卡片大标题、纯数字评分/序号、小图形配短标签(胶囊/◆/圆点/下方图形)、
           短小标题(<=8字且无句读)。
    注意：长描述句即使旁边有 7x7 装饰圆点，也仍是三级，不得被提升。
    """
    t=s['txt']
    if not t:
        return True
    L=len(t); fs=s['fs']
    short = (L<=8) and ('，' not in t) and ('。' not in t) and ('；' not in t)
    # 页面主标题 / 章节号
    if s['name'].startswith('!!TITLE_'):
        return True
    # 大字号标题(>=18)
    if fs>=18:
        return True
    # 纯数字(评分78/82/85、序号01/02): 与同行文字一起平滑，避免被拆散
    if t.isdigit():
        return True
    # 表格短评字(高/中/低/关键)：只认"无文字矩形 p:sp 单元格"，图片横幅不算表格
    if L<=2 and allshapes:
        for o in allshapes:
            if o['name']==s['name'] or o['txt'] or o['tag']!='p:sp':
                continue
            if o['w']>=150 and 40<=o['h']<=130:
                if o['x']-2<=s['x']<=o['x']+o['w'] and o['y']-2<=s['y']<=o['y']+o['h']:
                    return False
    # 同类卡片行内的描述句：即使只有7~8字且没标点，只要与同行长描述句字号一致，
    # 就按三级处理，避免同一张表里有的描述句平滑、有的动画
    if short and allshapes:
        peers=[o for o in allshapes if o['tag']=='p:sp' and not o['txt']
               and 300<=o['w']<=700 and 40<=o['h']<=120]
        bysize=collections.defaultdict(list)
        for o in peers: bysize[(round(o['w']),round(o['h']))].append(o)
        scx=s['x']+s['w']/2; scy=s['y']+s['h']/2
        for key,boxlist in bysize.items():
            if len(boxlist)<3: continue
            if not any(o['x']-2<=scx<=o['x']+o['w']+2 and
                       o['y']-2<=scy<=o['y']+o['h']+2 for o in boxlist):
                continue
            longfs=[]
            for o in boxlist:
                for q in allshapes:
                    if not q['txt'] or q['name']==s['name'] or len(q['txt'])<=8:
                        continue
                    qx=q['x']+q['w']/2; qy=q['y']+q['h']/2
                    if o['x']-2<=qx<=o['x']+o['w']+2 and o['y']-2<=qy<=o['y']+o['h']+2:
                        longfs.append(q['fs'])
            if len(longfs)>=2:
                med=sorted(longfs)[len(longfs)//2]
                if abs(fs-med)<1.0:
                    return False
            break
    # 小图形文字规则只对"短标签"生效，避免装饰圆点把长描述句提升为二级
    if short and allshapes:
        scx=s['x']+s['w']/2; scy=s['y']+s['h']/2
        for o in allshapes:
            if o['name']==s['name']:
                continue
            # 图形 = 图片 或 无文字的小形状(胶囊/菱形/圆点)
            is_graphic = o['pic'] or (not o['txt'] and o['w']<=140 and o['h']<=90)
            if not is_graphic or o['w']>140 or o['h']>90:
                continue
            ocx=o['x']+o['w']/2; ocy=o['y']+o['h']/2
            inside=(o['x']-3<=scx<=o['x']+o['w']+3 and o['y']-3<=scy<=o['y']+o['h']+3)
            left=(0<=s['x']-(o['x']+o['w'])<25 and abs(ocy-scy)<18)
            right=(0<=o['x']-(s['x']+s['w'])<25 and abs(ocy-scy)<18)
            below=(0<=o['y']-(s['y']+s['h'])<22 and abs(ocx-scx)<max(s['w'],o['w'])/2+12)
            above=(0<=s['y']-(o['y']+o['h'])<22 and abs(ocx-scx)<max(s['w'],o['w'])/2+12)
            if inside or left or right or below or above:
                return True
    # 短小标题 -> 二级及以上
    if short:
        return True
    return False

def is_tertiary(s, allshapes=None):
    """三级 = 有文字 且 非必要元素(反向取补)。即小字号的正文描述句。"""
    if not s['txt']:
        return False
    return not is_necessary(s, allshapes)

def set_offscreen(block, x_pt, y_pt=None):
    """把形状块移到屏外指定位置"""
    def off_repl(m):
        nx=int(round(x_pt*12700))
        ny=int(round(y_pt*12700)) if y_pt is not None else int(m.group(2))
        return f'<a:off x="{nx}" y="{ny}"'
    return re.sub(r'<a:off x="(-?\d+)" y="(-?\d+)"',off_repl,block,count=1)

def renumber(block, new_id):
    return re.sub(r'<p:cNvPr id="\d+"',f'<p:cNvPr id="{new_id}"',block,count=1)

# ===== 跨页复制图片的关系修复 =====
# 问题：p:pic 的 r:embed 是"页本地"关系ID。直接把 A 页的 pic 块复制到 B 页后，
# 同一个 rId 在 B 页指向的是另一张图 -> Morph 时形状对不上、强制变形（大图尤其明显）。
# 修复：按源页 rels 解析出真实 media 目标，在目标页复用已有关系或新建关系，改写副本的 r:embed。
RELS={}
for _pn in range(1,N+1):
    _fp=os.path.join(WORK,f'ppt/slides/_rels/slide{_pn}.xml.rels')
    _s=open(_fp,encoding='utf-8').read()
    _info={}; _map={}
    for _m in re.finditer(r'<Relationship\s+([^>]*?)/?>',_s):
        _a=dict(re.findall(r'(\w+)="([^"]*)"',_m.group(1)))
        if 'Id' not in _a: continue
        _info[_a['Id']]=(_a.get('Type',''),_a.get('Target',''),_a.get('TargetMode',''))
        _map[_a['Id']]=_a.get('Target','')
    RELS[_pn]={'s':_s,'map':_map,'info':_info,
               'used':set(re.findall(r'Id="(rId\d+)"',_s))}

REL_ATTRS='embed|link|id|dm|lo|qs|cs'

def fix_picture_rels(block, src_page, dst_page):
    """Remap every relationship reference in a copied shape block.

    Images use r:embed, but media/audio shapes also use r:link and p14:media
    uses r:embed on a *media* relationship. Copying only r:embed leaves stale
    r:link/r:id values pointing at unrelated parts on the destination slide,
    which makes PowerPoint report the file needs repair. New relationships
    must preserve the source Type and TargetMode, not assume 'image'.
    """
    if not re.search(r'r:(?:'+REL_ATTRS+r')="rId\d+"',block):
        return block
    srcinfo=RELS[src_page]['info']; dst=RELS[dst_page]
    rev=collections.defaultdict(list)
    for _r,(_t,_g,_m) in dst['info'].items():
        rev[(_t,_g,_m)].append(_r)
    remap={}
    def repl(m):
        attr,rid=m.group(1),m.group(2)
        if not rid: return m.group(0)
        if rid in remap: return f'r:{attr}="{remap[rid]}"'
        info=srcinfo.get(rid)
        if not info: return m.group(0)
        if rev.get(info):
            newrid=rev[info][0]
        else:
            n=1
            while f'rId{n}' in dst['used']: n+=1
            newrid=f'rId{n}'
            dst['used'].add(newrid)
            dst['info'][newrid]=info
            typ,tgt,mode=info
            mode_attr=f' TargetMode="{mode}"' if mode else ''
            rel=f'<Relationship Id="{newrid}" Type="{typ}" Target="{tgt}"{mode_attr}/>'
            dst['s']=dst['s'].replace('</Relationships>',rel+'</Relationships>')
            rev[info].append(newrid)
        remap[rid]=newrid
        return f'r:{attr}="{newrid}"'
    return re.sub(r'r:('+REL_ATTRS+r')="(rId\d*)"',repl,block)

def flush_rels():
    for _pn in range(1,N+1):
        _fp=os.path.join(WORK,f'ppt/slides/_rels/slide{_pn}.xml.rels')
        open(_fp,'w',encoding='utf-8').write(RELS[_pn]['s'])

# 读取所有页原始
slides_raw={i:open(os.path.join(WORK,f'ppt/slides/slide{i}.xml'),encoding='utf-8').read() for i in range(1,N+1)}

# ===== 阶段0：面积分级配对 =====
# 占页面 >=1/5 的大色块保留原名，即使两页样子/位置不同也允许 Morph 平滑
# 小于 1/5 的顶栏/横线/小装饰改成页唯一名，避免同名强制配对横飞
AREA=SW*SH
renamed=0; kept=0
for pi in range(1,N+1):
    s=slides_raw[pi]
    large=set()
    for m in re.finditer(r'<(p:sp|p:pic)\b.*?</\1>',s,re.S):
        b=m.group(0)
        c=re.search(r'name="(!!NAV_ITEM_\d+)"',b)
        o=re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"',b)
        e=re.search(r'<a:ext cx="(\d+)" cy="(\d+)"',b)
        if c and o and e:
            w=int(e.group(1))/12700; h=int(e.group(2))/12700
            if w*h/AREA>=0.2: large.add(c.group(1)); kept+=1
    def nav_repl(m):
        n=m.group(1)
        if f'!!NAV_ITEM_{n}' in large: return m.group(0)
        return f'name="!!NAV_P{pi}_ITEM_{n}"'
    s=re.sub(r'name="!!NAV_ITEM_(\d+)"',nav_repl,s)
    s=re.sub(r'name="!!HL_(\d+)"',
             lambda m:f'name="!!HL_P{pi}_{m.group(1)}"', s)
    slides_raw[pi]=s
all_shapes={i:parse_shapes(slides_raw[i]) for i in range(1,N+1)}
print('大面积保留同名 %d 个，其余装饰页唯一化'%kept)

# 全局最大 id
maxid=max(int(x) for s in slides_raw.values() for x in re.findall(r'<p:cNvPr id="(\d+)"',s))

# 每页的大色块(面积20%~85%，排除全屏背景)：用于把其上的文字锚定到色块一起平滑
def big_blocks(shapes):
    return [s for s in shapes
            if 0.2<=s['w']*s['h']/(SW*SH)<0.85
            and not s['name'].startswith('!!BG_')]
BIGB={i:big_blocks(all_shapes[i]) for i in range(1,N+1)}
# 记录 (A页, 元素名) -> A页色块名：这些字在A页贴着色块，层级必须在色块之上
ANCHORED={}

# ===== 阶段1：跨页进出场（修改 shape blocks，注入屏外副本）=====
# 策略：
# 对每对相邻页 A(=n) -> B(=n+1):
#   (a) B页的大元素 -> 复制到A页屏外(右侧或上下)，同名 -> 从屏外飞入
#   (b) A页的大元素 -> 复制到B页屏外(左侧)，同名 -> 飞出
# 注意：为避免同一页出现两个同名元素导致配对混乱，A页原有同名元素需改名(加 _OUT 后缀不参与配对)
# 但更简单：只做"飞入"，退出只针对标题(已有机制)，避免复杂冲突

added_in=0
for n in range(1,N):
    A=n; B=n+1
    Bshapes=all_shapes[B]
    # 全量内容元素做平滑副本，但三级大纲文字除外(它们用页内动画进入)
    bigs=[s for s in Bshapes if is_content(s) and not is_title(s) and not is_tertiary(s,Bshapes)]
    bigs.sort(key=lambda a:(a['y'],a['x']))

    # 同行紧邻元素归组(如序号块01+文字"现状判断")，整组统一方向，避免被x分界拆散
    group_of={}
    parent={s['name']:s['name'] for s in bigs}
    def find(x):
        while parent[x]!=x: x=parent[x]
        return x
    for i in range(len(bigs)):
        for j in range(i+1,len(bigs)):
            a,b=bigs[i],bigs[j]
            if abs(a['y']-b['y'])<28:  # 同一行
                gap=max(a['x']-(b['x']+b['w']), b['x']-(a['x']+a['w']))
                if -10<=gap<100:       # 水平紧邻
                    ra,rb=find(a['name']),find(b['name'])
                    if ra!=rb: parent[ra]=rb
    grp_center={}
    for s in bigs:
        r=find(s['name']); group_of[s['name']]=r
        grp_center.setdefault(r,[]).append(s['x']+s['w']/2)

    seen={s['name'] for s in all_shapes[A]}
    inserts=[]
    for idx,bs in enumerate(bigs):
        if bs['name'] in seen:
            continue
        seen.add(bs['name'])
        # 仅记录"字在色块上"的关系用于层级修正；位置一律放屏外（用户指定，避免封面重叠）
        for blk in BIGB[B]:
            if (blk['x']-2<=bs['x']<=blk['x']+blk['w'] and
                blk['y']-2<=bs['y']<=blk['y']+blk['h']):
                ablk=next((s for s in all_shapes[A] if s['name']==blk['name']),None)
                if ablk:
                    ANCHORED[(A,bs['name'])]=ablk['name']
                break
        # 按整组中心决定方向(序号块+文字同向)；行距分档造速度差
        cxm=sum(grp_center[group_of[bs['name']]])/len(grp_center[group_of[bs['name']]])
        row=int(bs['y']//110)
        if cxm < SW/2:
            ox=-bs['w']-80-((row+int(A))%3)*90; oy=bs['y']
        else:
            ox=SW+80+((row+int(A))%3)*90; oy=bs['y']
        clone=bs['block']
        maxid+=1
        clone=renumber(clone,maxid)
        clone=fix_picture_rels(clone,B,A)
        clone=set_offscreen(clone,ox,oy)
        inserts.append(clone)
        added_in+=1
    if inserts:
        slides_raw[A]=slides_raw[A].replace('</p:spTree>',''.join(inserts)+'</p:spTree>',1)

print("飞入副本添加:",added_in)

# ===== 阶段2：退出效果（主要元素飞出到下一页屏外）=====
# A页主标题已在B页有屏外副本。这里给大图片也做飞出。
added_out=0
for n in range(1,N):
    A=n; B=n+1
    Ashapes=all_shapes[A]
    # 飞出元素：大图片 / 大图形容器 / 大色块(含NAV_ITEM) / 大字号文字
    # 大面积色块也必须在下一页屏外放同名副本，切走时平滑出去
    def need_fly(s):
        # 常规大元素
        big = (s['w']*s['h']/(SW*SH)>=0.15
               and (s['name'].startswith('!!IMG_')
                    or s['name'].startswith('!!NAV_')
                    or (s['w']>=200 and s['h']>=100)
                    or (s['fs']>=18 and s['txt'])))
        # 封面小装饰椭圆：面积虽小，但若在下一页无同名，
        # Morph 可能把它与下一页的大元素误配，过渡中间帧出现圆角变形。
        # 给它们在下一页放同名屏外副本，平滑飞走，避免误配。
        cover_ellipse = (A==1 and s['name'].startswith('!!NAV_P1_ITEM_')
                         and re.search(r'<a:prstGeom prst="ellipse"',s['block']))
        return big or cover_ellipse
    outs=[s for s in Ashapes if not is_title(s) and s['x']<=SW
          and not s['name'].startswith('!!BG_') and need_fly(s)]
    outs.sort(key=lambda a:-(a['w']*a['h']))
    seen={s['name'] for s in all_shapes[B]}
    inserts=[]
    for idx,im in enumerate(outs):
        if im['name'] in seen:
            continue
        seen.add(im['name'])
        # 飞出方向取本页主方向的反向，每页不同；距离逐档递增形成速度差
        out_dir=RING[(RING.index(PAGE_MAIN[(A-1) % len(PAGE_MAIN) + 1])+4)%8]
        ox,oy=off_pos(out_dir,im,80+idx*95)
        clone=im['block']
        maxid+=1
        clone=renumber(clone,maxid)
        clone=fix_picture_rels(clone,A,B)
        clone=set_offscreen(clone,ox,oy)
        inserts.append(clone)
        added_out+=1
    if inserts:
        slides_raw[B]=slides_raw[B].replace('</p:spTree>',''.join(inserts)+'</p:spTree>',1)
print("飞出副本添加:",added_out)

# ===== 阶段3：页内动画（按章节差异化，卡片整体同步，只给未做跨页的元素）=====
def get_chapter(pi):
    # Try to read a chapter marker (01/02/03/04) from the page title text.
    # Fall back to dividing the body slides into 4 equal chapters.
    s = slides_raw[pi]
    m = re.search(r'<a:t>(0[1-9])</a:t>', s)
    if m:
        return m.group(1)
    body_start, body_end = 3, max(3, N - 1)
    span = (body_end - body_start + 1) / 4.0
    if span <= 0:
        return None
    idx = int((pi - body_start) / span)
    return f'{min(4, idx + 1):02d}' if 0 <= idx < 4 else None

def build_effect(tpl, spid, idgen, ntype, dur, delay=0):
    e=tpl.replace('__SPID__',str(spid))
    e=re.sub(r' id="__ID__"',lambda x:f' id="{next(idgen)}"',e)
    e=re.sub(r'nodeType="(?:afterEffect|clickEffect|withEffect)"',f'nodeType="{ntype}"',e)
    # Normalize every real animation duration (keep dur="1" for the visibility set).
    # Donor decks often contain 327/350 ms effects; without this they look like a pop-in.
    e=re.sub(r'dur="(\d+)"',lambda m:(f'dur="{dur}"' if int(m.group(1))>1 else m.group(0)),e)
    # 只改外层 cTn 的起始延迟（模板中第一个 cond），实现卡片间轻微错峰
    if delay:
        e=e.replace('<p:cond delay="0"/>',f'<p:cond delay="{delay}"/>',1)
    return f'<p:par>{e}</p:par>'

MORPH_XML = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
    '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="2000">'
    '<p159:morph option="byObject"/>'
    '</p:transition></mc:Choice>'
    '<mc:Fallback><p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="slow">'
    '<p:fade/></p:transition></mc:Fallback>'
    '</mc:AlternateContent>'
)

def ensure_morph(slide_xml):
    # Use PowerPoint's native MCE shape: p159 is declared on mc:Choice, p/p14 are
    # declared locally on p:transition, and no mc:Ignorable is added to the root.
    # This avoids repair prompts in strict Office/WPS parsers.
    if '<p159:morph' in slide_xml or '<mc:AlternateContent' in slide_xml:
        return slide_xml
    m = re.search(r'</p:clrMapOvr>', slide_xml)
    if m:
        return slide_xml[:m.end()] + MORPH_XML + slide_xml[m.end():]
    m = re.search(r'</p:cSld>', slide_xml)
    if m:
        return slide_xml[:m.end()] + MORPH_XML + slide_xml[m.end():]
    return slide_xml

idgen=itertools.count(50000)
anim_stat=collections.Counter()
for pi in range(1,N+1):
    s=slides_raw[pi]
    if pi==1:
        # 第一页：取消 Morph 切换，屏内全部元素(含背景/导航/装饰)都加动画
        s=re.sub(r'<mc:AlternateContent>.*?</mc:AlternateContent>','',s,count=1,flags=re.S)
    else:
        # 新源稿可能完全没有平滑切换；这里补齐默认 2000ms Morph，供屏外副本跨页运动使用
        s=ensure_morph(s)
    slides_raw[pi]=s
    shapes=parse_shapes(s)
    if pi==1:
        content=[sh for sh in shapes
                 if 0<=sh['x']+sh['w']/2<=SW and 0<=sh['y']+sh['h']/2<=SH
                 and (pi,sh['name']) not in ANCHORED]
    else:
        # P2+ 只给三级大纲文字加页内动画(其余内容靠平滑进入)
        # 必须排除阶段2放进来的"飞出副本"：它们虽然 x 在屏内，但 y 已在屏外
        def on_screen(sh):
            return (-10<=sh['x']+sh['w']/2<=SW+10 and
                    -10<=sh['y']+sh['h']/2<=SH+10)
        content=[sh for sh in shapes if is_content(sh) and on_screen(sh)
                 and is_tertiary(sh,shapes)]
    if not content:
        s=re.sub(r'<p:timing>.*?</p:timing>','',s,flags=re.S)
        slides_raw[pi]=s
        continue
    # 分组：描述长句按x列聚类(每卡片一列、逐卡进入)；短评字按行分组
    def is_short_desc(sh):
        return len(sh['txt'])<=2 and sh['fs']<15
    descs=[sh for sh in content if not is_short_desc(sh)]
    shorts=[sh for sh in content if is_short_desc(sh)]
    groups=[]
    # 描述句按列心x聚类(列间距约120pt)，每列一组
    desc_sorted=sorted(descs,key=lambda a:a['x'])
    cols=[]
    for sh in desc_sorted:
        cx=sh['x']+sh['w']/2
        if cols and abs(cx-cols[-1][0])<=90:
            cols[-1][1].append(sh)
        else:
            cols.append([cx,[sh]])
    for cx,members in cols:
        groups.append(sorted(members,key=lambda a:(a['y'],a['x'])))
    # 短评字按行分组
    sband=collections.defaultdict(list)
    for sh in shorts: sband[round(sh['y']/95)].append(sh)
    for b in sorted(sband):
        groups.append(sorted(sband[b],key=lambda a:(a['x'],a['y'])))
    while len(groups)>4:   # 合并相邻组，整页最多4组
        groups.sort(key=lambda g:min(a['y']*0+a['x'] for a in g))
        # 合并元素最少的两组，尽量保留列结构
        groups.sort(key=lambda g:len(g))
        a=groups.pop(0); b=groups.pop(0)
        groups.append(sorted(a+b,key=lambda s:(s['y'],s['x'])))
    groups.sort(key=lambda g:min(a['x'] for a in g))
    chap=get_chapter(pi)
    cid=next(idgen); wid=next(idgen)
    effs=[]
    dur=400
    # 5种节奏模式按页轮换(非规律序列)，避免章节式机械交替
    RHYTHM=['uniform','seq','reverse','wave','relay','seq','seq','reverse',
            'relay','wave','seq','uniform','reverse','relay','uniform']
    rhythm=RHYTHM[(pi-1)%len(RHYTHM)]
    ngrp=len(groups)
    base=230   # 组间基础错峰
    # 用户要求：第一个动的动画不能延迟，否则平滑结束后还要等一会才出字。
    # 只保留组间错峰 base(=230ms) 与组内接力，不再加全局基准延迟。
    base_delay = 0
    for order,g in enumerate(groups):
        if rhythm=='uniform':
            gdelay=0
        elif rhythm=='seq':
            gdelay=order*base
        elif rhythm=='reverse':
            gdelay=(ngrp-1-order)*base
        elif rhythm=='wave':
            gdelay=(order%2)*base
        else:  # relay: 组间错峰 + 组内接力
            gdelay=order*base
        for si,sh in enumerate(g):
            ntype='afterEffect' if si==0 else 'withEffect'
            # 逐元素按角色选动画：短评字(高/中/低/序号)用缩放逐个进；描述句用淡入逐个进
            short = len(sh['txt'])<=2 and sh['fs']<15
            pid = 53 if short else 10
            tpl = TPLS[pid]
            idelay = (si*110) if short else (si*90)   # 都逐个，评字间距略大
            effs.append(build_effect(tpl,sh['id'],idgen,ntype,dur,
                                     base_delay+gdelay+idelay))
            anim_stat[pid]+=1
    timing=(
        '<p:timing><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        f'<p:par><p:cTn id="{cid}" fill="hold" nodeType="clickPar">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst><p:par><p:cTn id="{wid}" fill="hold" nodeType="withGroup">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst>{"".join(effs)}</p:childTnLst>'
        '</p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn></p:seq>'
        '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')
    s=re.sub(r'<p:timing>.*?</p:timing>','',s,flags=re.S)
    s=s.replace('</p:sld>',timing+'</p:sld>')
    slides_raw[pi]=s
    print(f'P{pi}: 章节{chap or "-"} {len(groups)}组 节奏{rhythm} 动画{len(effs)}')

# ===== 阶段4：对齐跨页同名元素的 z-order，消除 Morph 层级闪烁 =====
def parse_sptree(s):
    st=s.find('<p:spTree>'); en=s.find('</p:spTree>')
    tree=s[st+len('<p:spTree>'):en]
    blocks=[]
    for m in re.finditer(r'<(p:sp|p:pic)\b.*?</\1>',tree,re.S):
        c=re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"',m.group(0))
        if c: blocks.append((c.group(2),m.group(0)))
    first=re.search(r'<(p:sp|p:pic)\b',tree)
    prefix=tree[:first.start()] if first else tree
    last=None
    for m in re.finditer(r'<(p:sp|p:pic)\b.*?</\1>',tree,re.S): last=m
    suffix=tree[last.end():] if last else ''
    return st,en,prefix,suffix,blocks

def is_off_block(b):
    o=re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"',b)
    e=re.search(r'<a:ext cx="(\d+)" cy="(\d+)"',b)
    if not(o and e): return True
    x=int(o.group(1))/12700; y=int(o.group(2))/12700
    w=int(e.group(1))/12700; h=int(e.group(2))/12700
    return not (0<=x+w/2<=SW and 0<=y+h/2<=SH)

TREE={i:parse_sptree(slides_raw[i]) for i in range(1,N+1)}
ORDER={i:[n for n,_ in TREE[i][4]] for i in range(1,N+1)}
OFF={i:{n for n,b in TREE[i][4] if is_off_block(b)} for i in range(1,N+1)}

# 只按"屏内元素"计算层级百分位；屏外副本插到目标百分位对应的屏内元素之前。
# 旧算法把副本和屏内元素放在同一 0~1 轴排序，副本互相挤占，导致相邻页层级差仍然过大。
ON={i:[n for n in ORDER[i] if n not in OFF[i]] for i in range(1,N+1)}
fixed=0; moved=0
for pi in range(1,N+1):
    st,en,prefix,suffix,blocks=TREE[pi]
    cur=ORDER[pi]; offs=[n for n in cur if n in OFF[pi]]
    if not offs: continue
    targets={}
    for n in offs:
        for q in (pi-1,pi+1):
            if q<1 or q>N: continue
            if n in ON[q]:
                targets[n]=ON[q].index(n)/max(len(ON[q])-1,1)
                break
            elif n in ORDER[q]:
                # 相邻页同名元素也在屏外(如大红块飞出副本链)：继承其全局层级，
                # 避免第一个副本在前面、下一个在最后一层导致切换闪烁
                targets[n]=ORDER[q].index(n)/max(len(ORDER[q])-1,1)
                break
    on=ON[pi]
    slots=[[] for _ in range(len(on)+1)]
    for n in offs:
        if n in targets:
            j=int(round(targets[n]*max(len(on)-1,0)))
            slots[j].append(n)
        else:
            slots[-1].append(n)
    new_order=[]
    for j in range(len(on)):
        new_order.extend(slots[j]); new_order.append(on[j])
    new_order.extend(slots[-1])
    if new_order!=cur:
        bmap={n:b for n,b in blocks}
        s=slides_raw[pi]
        slides_raw[pi]=(s[:st+len('<p:spTree>')]+prefix+
                        ''.join(bmap[n] for n in new_order)+suffix+s[en:])
        fixed+=1; moved+=len(targets)
print('z-order 对齐: %d 页, %d 个屏外副本'%(fixed,moved))

# ===== 阶段5：锚定字层级必须在对应色块之上（防止切换时白字被色块盖住/层级反转）=====
# 重新读取对齐后的顺序
TREE2={i:parse_sptree(slides_raw[i]) for i in range(1,N+1)}
order_fixed=0
for pi in range(1,N+1):
    items=[(p,n) for (p,n) in ANCHORED if p==pi]
    if not items: continue
    st,en,prefix,suffix,blocks=TREE2[pi]
    cur=[n for n,_ in blocks]
    bmap={n:b for n,b in blocks}
    new_order=list(cur)
    for (_,txtname),blkname in [(k,v) for k,v in ANCHORED.items() if k[0]==pi]:
        if txtname not in new_order or blkname not in new_order: continue
        # 移除字，插入到色块之后(更上层)
        new_order.remove(txtname)
        pos=new_order.index(blkname)+1
        new_order.insert(pos,txtname)
        order_fixed+=1
    if new_order!=cur:
        s=slides_raw[pi]
        slides_raw[pi]=(s[:st+len('<p:spTree>')]+prefix+
                        ''.join(bmap[n] for n in new_order)+suffix+s[en:])
print('锚定字层级修正: %d 处'%order_fixed)

# 写回
for pi in range(1,N+1):
    open(os.path.join(WORK,f'ppt/slides/slide{pi}.xml'),'w',encoding='utf-8').write(slides_raw[pi])

flush_rels()
# 打包
zo=zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED)
for dp,dn,fn in os.walk(WORK):
    for f in fn:
        full=os.path.join(dp,f); rel=os.path.relpath(full,WORK).replace(os.sep,'/')
        zo.write(full,rel)
zo.close()
print("页内动画统计:",dict(anim_stat))
print("输出:",OUT)
