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

# Parse ingredient table: collect product -> list of (杯型spec, 制作text, 推荐糖度)
# structure: columns: 0=category/section, 1=杯型, 2=制作(正常糖), ... last col=推荐糖度
# A product row has col0 = product name. A 杯型 row has col0 empty and col1 nonempty.
# Let's extract from 国内产品 + 华南 (most detailed), and verify against 制作密码导出.

def norm(s):
    return re.sub(r'\s+','', s).replace('（','(').replace('）',')').replace('﹕',':').replace('Ｔ','T').replace('：',':')

def parse_t_annotation(text):
    # find all (T:... ) or (M:...) or (满T..) groups in the 制作 string
    # pattern like (T：满40、半30、三20）
    res={}
    # T group
    for m in re.finditer(r'[\(（][^\(\)（）]*[TtＴ][^\(\)（）]*[\)）]', text):
        grp=m.group(0)
        # extract 满/半/三 values
        full=re.search(r'[满全](\d+)',grp); half=re.search(r'半(\d+)',grp); third=re.search(r'三(\d+)',grp)
        # for pure T without 满: handle like T10,T15 etc
        if full or half or third:
            res['T']=(full.group(1) if full else None, half.group(1) if half else None, third.group(1) if third else None)
    # M group
    for m in re.finditer(r'[\(（][^\(\)（）]*[MmＭ][^\(\)（）]*[\)）]', text):
        grp=m.group(0)
        full=re.search(r'[满全](\d+)',grp); half=re.search(r'半(\d+)',grp); third=re.search(r'三(\d+)',grp)
        if full or half or third:
            res['M']=(full.group(1) if full else None, half.group(1) if half else None, third.group(1) if third else None)
    return res

# Find a product in ingredient sheet by name, return its 制作 text(s) and 推荐糖度
def find_product(sheet_rows, name):
    name=norm(name)
    found=[]
    i=0
    while i < len(sheet_rows):
        r=sheet_rows[i]
        r0=norm(r[0]) if len(r)>0 else ''
        if r0==name and (len(r)>2 and r[2].strip()!='' or (len(r)>1 and r[1].strip()!='')):
            # product header row. collect following 杯型 rows until next product header (col0 nonempty)
            block=[]
            j=i
            # the header itself may have 杯型 in col1 and 制作 in col2
            cur_spec = r[1].strip() if len(r)>1 else ''
            cur_make = r[2].strip() if len(r)>2 else ''
            cur_rec = r[-1].strip() if len(r)>0 else ''
            # gather subsequent rows where col0 is empty (continuation / 杯型 variants)
            k=i+1
            specs=[(cur_spec,cur_make)]
            while k<len(sheet_rows):
                rr=sheet_rows[k]
                rr0=norm(rr[0]) if len(rr)>0 else ''
                if rr0!='' :
                    break
                # col1 nonempty => another 杯型 variant
                if len(rr)>1 and rr[1].strip()!='':
                    specs.append((rr[1].strip(), rr[2].strip() if len(rr)>2 else ''))
                k+=1
            rec = sheet_rows[i][-1].strip() if sheet_rows[i] else ''
            found.append((specs, rec, i))
            i=k
        else:
            i+=1
    return found

# Build make-password lookup by (name, spec, sugar)
mp_rows=[r for r in mp[1:] if len(r)>=8]
def sugar_of(combo):
    for tok in combo.split(','):
        t=tok.strip()
        if t in ('标准糖','七分糖','五分糖','三分糖','不额外加糖','无糖'): return t
    return None
def ice_of(combo):
    for tok in combo.split(','):
        t=tok.strip()
        if t in ('正常冰','少冰','去冰','常温','热'): return t
    return None
# determine spec mapping: 大杯/中杯/桶 etc
mp_lookup={}
for r in mp_rows:
    name=r[2]; spec=r[3]; combo=r[5]; pwd=r[7]
    s=sugar_of(combo); ic=ice_of(combo)
    mp_lookup.setdefault((name,spec),{})[(s,ic)]=pwd

# Now verify: for each product present in BOTH ingredient sheet (国内产品) and mp,
# check that 制作(正常糖) T-annotation 满 = 标准糖 T value, 半 = 七分/五分, 三 = 三分
sheet = ing["国内产品"]
# gather all product names in sheet
names=[]
for r in sheet:
    if len(r)>0 and r[0].strip()!='' and len(r)<=3:
        names.append(r[0].strip())
print("PRODUCTS IN 国内产品:", len(names))
print(names)
print("\n\n===== VERIFY T MAPPING against 制作密码导出 =====")
checked=0
for nm in names:
    fps=find_product(sheet, nm)
    for specs, rec, idx in fps:
        for spec, make in specs:
            if not make: continue
            # determine size keyword
            size=None
            for kw in ('大杯','中杯','桶','圣代杯','霸气华夫脆筒','咖啡杯','16A','22A','700','500','1L'):
                if kw in spec:
                    size=kw; break
            # try to match in mp by name + a size token; mp spec is like 大杯/中杯/桶
            mp_size=None
            for kw in ('大杯','中杯','桶','个'):
                if kw in spec: mp_size=kw; break
            if not mp_size: continue
            key=(nm, mp_size)
            if key not in mp_lookup: continue
            tmap=parse_t_annotation(make).get('T')
            if not tmap:
                # no T annotation (e.g., 冰淇淋) -> skip
                continue
            full,half,third=tmap
            # get passwords
            def gett(sugar):
                for (s,ic),pwd in mp_lookup[key].items():
                    if s==sugar:
                        m=re.search(r'#T(\d+)#',pwd) or re.search(r'\|T(\d+)(\||$)',pwd)
                        return m.group(1) if m else 'NONE'
                return 'NA'
            sf=gett('标准糖'); sh=gett('七分糖'); sf2=gett('五分糖'); st=gett('三分糖'); sn=gett('不额外加糖')
            checked+=1
            ok_full = (full==sf) or sf=='NA'
            ok_half = (half==sh) or (half==sf2) or sh=='NA'
            ok_third= (third==st) or st=='NA'
            flag = 'OK' if (ok_full and ok_half and ok_third) else '*** MISMATCH ***'
            print(f"[{flag}] {nm}/{mp_size}: 配料T(满{full},半{half},三{third}) | 密码T(标准{sf},七分{sh},五分{sf2},三分{st},无糖{sn})")
print(f"\nChecked {checked} product-size blocks with T annotations.")
