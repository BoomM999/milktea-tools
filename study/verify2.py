import zipfile, re, os
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def col_to_idx(col):
    n=0
    for ch in col: n=n*26+(ord(ch)-ord('A')+1)
    return n-1
def load_xlsx(path):
    z=zipfile.ZipFile(path); shared=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(f'{NS}si'):
            shared.append(''.join(t.text or '' for t in si.iter(f'{NS}t')))
    wb=ET.fromstring(z.read('xl/workbook.xml')); sheets=[]
    for s in wb.find(f'{NS}sheets'):
        sheets.append((s.get('name'), s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')))
    rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid={r.get('Id'):r.get('Target') for r in rels}
    out={}
    for name,ridx in sheets:
        tg=rid[ridx]
        if not tg.startswith('xl/'): tg='xl/'+tg
        root=ET.fromstring(z.read(tg)); rows=[]
        for row in root.iter(f'{NS}row'):
            cells={}; mc=-1
            for c in row.findall(f'{NS}c'):
                m=re.match(r'([A-Z]+)(\d+)',c.get('r')); ci=col_to_idx(m.group(1))
                t=c.get('t'); v=c.find(f'{NS}v'); isn=c.find(f'{NS}is')
                if t=='s' and v is not None: val=shared[int(v.text)]
                elif t=='inlineStr' and isn is not None: val=''.join(tt.text or '' for tt in isn.iter(f'{NS}t'))
                elif v is not None: val=v.text
                else: val=''
                cells[ci]=val
                if ci>mc: mc=ci
            rows.append([cells.get(i,'') for i in range(mc+1)])
        out[name]=rows
    return out

base=os.path.dirname(os.path.abspath(__file__))
ing=load_xlsx(os.path.join(base,"标签简化_推荐糖配料表.xlsx"))
mp=load_xlsx(os.path.join(base,"制作密码导出.xlsx"))["制作密码组合"]

HEADER_TOKENS={'前场配料','杯型','制作（正常糖）','制作(正常糖）','制作(正常糖)','推荐糖','推荐糖度','制作(正常糖）'}

def is_product_header(r):
    if len(r)==0: return False
    c0=(r[0] or '').strip()
    if c0=='' : return False
    if c0 in HEADER_TOKENS: return False
    # product header: has col0 product name, and (col1 杯型 OR col2 制作 OR it's a category like 冰淇淋&茶?)
    # A category section header (e.g. 冰淇淋&茶) usually has empty col1,col2 and is in a row right before 杯型 row
    return True

# Build product blocks: a product header row (col0 nonempty, not header token) starts a block.
# Section/category headers (冰淇淋&茶, 招牌椰椰, 现磨咖啡...) also have col0 nonempty but col1,col2 empty.
# We treat a row as a true PRODUCT if col2 (制作) is nonempty OR col1 (杯型) nonempty.
def build_blocks(sheet_rows):
    blocks=[]
    i=0
    cur_section=None
    while i<len(sheet_rows):
        r=sheet_rows[i]
        c0=(r[0] if len(r)>0 else '').strip()
        c1=(r[1] if len(r)>1 else '').strip()
        c2=(r[2] if len(r)>2 else '').strip()
        if c0!='' and c0 not in HEADER_TOKENS and c1=='' and c2=='':
            # section/category header
            cur_section=c0
            i+=1
            continue
        if c0!='' and c0 not in HEADER_TOKENS and (c1!='' or c2!=''):
            # product header
            name=c0
            specs=[(c1,c2)]
            rec = r[-1].strip() if r else ''
            # gather subsequent rows where col0 empty and col1 nonempty (杯型 variants) OR col0 empty and col1 empty but col2 nonempty (continuation of same spec)
            j=i+1
            while j<len(sheet_rows):
                rr=sheet_rows[j]
                rc0=(rr[0] if len(rr)>0 else '').strip()
                rc1=(rr[1] if len(rr)>1 else '').strip()
                rc2=(rr[2] if len(rr)>2 else '').strip()
                if rc0!='': break
                if rc1!='' or rc2!='':
                    specs.append((rc1,rc2))
                j+=1
            blocks.append({'section':cur_section,'name':name,'specs':specs,'rec':rec,'row':i})
            i=j
        else:
            i+=1
    return blocks

sheet=ing["国内产品"]
blocks=build_blocks(sheet)
print(f"国内产品 product blocks: {len(blocks)}")
for b in blocks[:8]:
    print(f"  [{b['section']}] {b['name']} rec={b['rec']!r} specs={b['specs']}")

# parse T annotation
def parse_t(text):
    res={}
    for tag in ('T','M'):
        for m in re.finditer(r'[\(（][^\(\)（）]*'+tag+r'[^\(\)（）]*[\)）]', text):
            grp=m.group(0)
            full=re.search(r'[满全](\d+)',grp); half=re.search(r'半(\d+)',grp); third=re.search(r'三(\d+)',grp)
            # also handle pure "T10" without 满
            if full or half or third:
                res[tag]=(full.group(1) if full else None, half.group(1) if half else None, third.group(1) if third else None)
            else:
                # e.g. (T10、M15) -> capture first number as 满
                nums=re.findall(r'(\d+)',grp)
                if nums:
                    res[tag]=(nums[0], nums[1] if len(nums)>1 else None, nums[2] if len(nums)>2 else None)
    return res

# mp lookup
mp_rows=[r for r in mp[1:] if len(r)>=8]
def sugar_of(c):
    for tok in c.split(','):
        t=tok.strip()
        if t in ('标准糖','七分糖','五分糖','三分糖','不额外加糖','无糖'): return t
    return None
mp_lookup={}
for r in mp_rows:
    name=r[2]; spec=r[3]; combo=r[5]; pwd=r[7]
    s=sugar_of(combo)
    mp_lookup.setdefault((name,spec),{})[s]=pwd

def get_t(pwd):
    if pwd is None: return 'NA'
    m=re.search(r'#T(\d+)#',pwd)
    if m: return m.group(1)
    m=re.search(r'\|T(\d+)(?:\||$)',pwd)
    if m: return m.group(1)
    return 'NONE'

print("\n\n===== VERIFY T MAPPING =====")
checked=0; mism=0
for b in blocks:
    nm=b['name']
    for spec,make in b['specs']:
        if not make: continue
        mp_size=None
        for kw in ('大杯','中杯','桶','个'):
            if kw in spec: mp_size=kw; break
        if not mp_size: continue
        key=(nm,mp_size)
        if key not in mp_lookup: continue
        tmap=parse_t(make).get('T')
        if not tmap: continue
        full,half,third=tmap
        sf=get_t(mp_lookup[key].get('标准糖'))
        sh=get_t(mp_lookup[key].get('七分糖'))
        sf2=get_t(mp_lookup[key].get('五分糖'))
        st=get_t(mp_lookup[key].get('三分糖'))
        sn=get_t(mp_lookup[key].get('不额外加糖'))
        checked+=1
        ok = (full==sf or sf in ('NA','NONE')) and (half in (sh,sf2) or sh in ('NA','NONE')) and (third==st or st in ('NA','NONE'))
        if not ok:
            mism+=1
            print(f"  [MISMATCH] {nm}/{mp_size}: 配料(满{full},半{half},三{third}) 密码(标{sf},七{sh},五{sf2},三{st},无{sn})")
print(f"\nChecked {checked} blocks, mismatches={mism}")
