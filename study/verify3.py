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

def build_blocks(sheet_rows):
    blocks=[]; i=0; cur=None
    while i<len(sheet_rows):
        r=sheet_rows[i]
        c0=(r[0] if len(r)>0 else '').strip()
        c1=(r[1] if len(r)>1 else '').strip()
        c2=(r[2] if len(r)>2 else '').strip()
        if c0!='' and c0 not in HEADER_TOKENS and c1=='' and c2=='':
            cur=c0; i+=1; continue
        if c0!='' and c0 not in HEADER_TOKENS and (c1!='' or c2!=''):
            specs=[(c1,c2)]; rec=r[-1].strip() if r else ''
            j=i+1
            while j<len(sheet_rows):
                rr=sheet_rows[j]
                rc0=(rr[0] if len(rr)>0 else '').strip()
                rc1=(rr[1] if len(rr)>1 else '').strip()
                rc2=(rr[2] if len(rr)>2 else '').strip()
                if rc0!='': break
                if rc1!='' or rc2!='': specs.append((rc1,rc2))
                j+=1
            blocks.append({'section':cur,'name':c0,'specs':specs,'rec':rec,'row':i})
            i=j
        else: i+=1
    return blocks

# mp lookup by name+spec -> sugar -> pwd
mp_rows=[r for r in mp[1:] if len(r)>=8]
def sugar_of(c):
    for tok in c.split(','):
        t=tok.strip()
        if t in ('标准糖','七分糖','五分糖','三分糖','不额外加糖','无糖'): return t
    return None
mp_lookup={}
for r in mp_rows:
    mp_lookup.setdefault((r[2],r[3]),{})[sugar_of(r[5])]=r[7]

sheet=ing["国内产品"]
blocks=build_blocks(sheet)

def get_size(spec):
    for kw in ('大杯','中杯','桶','个'):
        if kw in spec: return kw
    return None

print("===== SIDE-BY-SIDE: 配料表 vs 制作密码 (国内产品) =====")
shown=0
for b in blocks:
    nm=b['name']
    for spec,make in b['specs']:
        if not make: continue
        sz=get_size(spec)
        if not sz: continue
        key=(nm,sz)
        if key not in mp_lookup: continue
        # use 标准糖 password as representative
        pwd=mp_lookup[key].get('标准糖')
        if not pwd: continue
        # split pwd: prefix before first | , then body
        if '|' in pwd:
            cup_part, body = pwd.split('|',1)
        else:
            cup_part, body = '', pwd
        # body uses # separators. Convert # to 、 for comparison with 制作
        body_norm = body.replace('#','、')
        # normalize spaces
        def n(s): return re.sub(r'\s+','',s)
        print(f"\n■ {nm} / {sz}  [rec={b['rec']}]")
        print(f"   杯型(配料): {spec!r}")
        print(f"   杯型(密码): {cup_part!r}")
        print(f"   制作(配料): {make!r}")
        print(f"   制作(密码·标): {body_norm!r}")
        shown+=1
        if shown>=10: break
    if shown>=10: break
print(f"\nShown {shown}")
