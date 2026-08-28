import zipfile, re, os
from xml.etree import ElementTree as ET
from collections import defaultdict

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
mp_rows=[r for r in mp[1:] if len(r)>=8]
def sugar_of(c):
    for tok in c.split(','):
        t=tok.strip()
        if t in ('标准糖','七分糖','五分糖','三分糖','不额外加糖','无糖'): return t
    return None
mp_lookup={}
for r in mp_rows:
    mp_lookup.setdefault((r[2],r[3]),{})[sugar_of(r[5])]=r[7]

def get_t(pwd, sym='T'):
    if not pwd: return 'NA'
    # T appears as #Txx# or |Txx| or after | like |T65  -> capture digits following sym
    # handle both T and M
    pat = r'[#|]'+sym+r'(\d+)'
    m=re.search(pat,pwd)
    if m: return m.group(1)
    # also bare Txx not preceded
    m=re.search(r'(?:^|[#|])'+sym+r'(\d+)', pwd)
    if m: return m.group(1)
    return 'NONE'

# For every product+size in mp that has a T annotation in ingredient, build mapping sugar->T and sugar->M
# parse ingredient T/M annotation
def parse_tm(text):
    res={}
    for tag in ('T','M'):
        for m in re.finditer(r'[\(（][^\(\)（）]*'+tag+r'[^\(\)（）]*[\)）]', text):
            grp=m.group(0)
            full=re.search(r'[满全](\d+)',grp); half=re.search(r'半(\d+)',grp); third=re.search(r'三(\d+)',grp)
            if full or half or third:
                res[tag]=(full.group(1) if full else None, half.group(1) if half else None, third.group(1) if third else None)
            else:
                nums=re.findall(r'(\d+)',grp)
                if nums: res[tag]=(nums[0], nums[1] if len(nums)>1 else None, nums[2] if len(nums)>2 else None)
    return res

# Collect all (name,spec) with T/M annotation from mp that also exist in ingredient
# We need ingredient make text. Build a lookup name+size -> list of (spec_text, make_text, rec)
ING={}
HEADER_TOKENS={'前场配料','杯型','制作（正常糖）','制作(正常糖）','制作(正常糖)','推荐糖','推荐糖度'}
def build(sheet_rows):
    blocks=[]; i=0; cur=None
    while i<len(sheet_rows):
        r=sheet_rows[i]
        c0=(r[0] if len(r)>0 else '').strip(); c1=(r[1] if len(r)>1 else '').strip(); c2=(r[2] if len(r)>2 else '').strip()
        if c0 and c0 not in HEADER_TOKENS and c1=='' and c2=='':
            cur=c0; i+=1; continue
        if c0 and c0 not in HEADER_TOKENS and (c1 or c2):
            specs=[(c1,c2)]; rec=r[-1].strip() if r else ''
            j=i+1
            while j<len(sheet_rows):
                rr=sheet_rows[j]; rc0=(rr[0] if len(rr)>0 else '').strip(); rc1=(rr[1] if len(rr)>1 else '').strip(); rc2=(rr[2] if len(rr)>2 else '').strip()
                if rc0: break
                if rc1 or rc2: specs.append((rc1,rc2))
                j+=1
            blocks.append((c0,specs,rec)); i=j
        else: i+=1
    return blocks

for sheet in ing.values():
    for nm,specs,rec in build(sheet):
        for spec,mk in specs:
            if not mk: continue
            sz=None
            for kw in ('大杯','中杯','桶','个'):
                if kw in spec: sz=kw; break
            if not sz: continue
            tm=parse_tm(mk)
            if 'T' not in tm and 'M' not in tm: continue
            key=(nm,sz)
            if key not in mp_lookup: continue
            ING.setdefault(key,[]).append((tm,rec))

# Now aggregate sugar->Tvalue and sugar->Mvalue across dataset, to confirm mapping
sugar_T=defaultdict(lambda: defaultdict(int))  # sugar -> Tvalue -> count
sugar_M=defaultdict(lambda: defaultdict(int))
# Also track normalization: does 标准糖 always use 满, 七分糖/五分糖 use 半, 三分糖 use 三?
print("===== GLOBAL SUGAR -> T VALUE MAPPING (ingredient 满/半/三 vs password) =====")
rows_report=[]
for key, lst in ING.items():
    nm,sz=key
    # take first occurrence tm (usually one per size)
    tm,rec=lst[0]
    full,half,third = (tm.get('T') or (None,None,None))
    # also M
    mfull,mhalf,mthird = (tm.get('M') or (None,None,None))
    # get password T/M per sugar
    for sugar in ('标准糖','七分糖','五分糖','三分糖','不额外加糖'):
        pwd=mp_lookup[key].get(sugar)
        if not pwd: continue
        pvT=get_t(pwd,'T'); pvM=get_t(pwd,'M')
        # expected
        expT = {'标准糖':full,'七分糖':half,'五分糖':half,'三分糖':third,'不额外加糖':None}[sugar]
        expM = {'标准糖':mfull,'七分糖':mhalf,'五分糖':mhalf,'三分糖':mthird,'不额外加糖':None}[sugar]
        if sugar in ('七分糖','五分糖'):
            # accept either 半
            ok = (pvT==half) or (pvT==full and half is None) or pvT=='NONE' and sugar=='不额外加糖'
        # record expected vs actual for T
        if expT is not None:
            sugar_T[sugar][pvT]+=1
        else:
            sugar_T[sugar][pvT]+=1
        if expM is not None:
            sugar_M[sugar][pvM]+=1
        else:
            sugar_M[sugar][pvM]+=1

for sugar in ('标准糖','七分糖','五分糖','三分糖','不额外加糖'):
    tvals=sorted(sugar_T[sugar].items(), key=lambda x:-x[1])
    mvals=sorted(sugar_M[sugar].items(), key=lambda x:-x[1])
    print(f"  {sugar:6}: T={tvals}  M={mvals}")
