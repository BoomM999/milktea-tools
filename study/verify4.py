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
mp_rows=[r for r in mp[1:] if len(r)>=8]
def sugar_of(c):
    for tok in c.split(','):
        t=tok.strip()
        if t in ('标准糖','七分糖','五分糖','三分糖','不额外加糖','无糖'): return t
    return None
mp_lookup={}
for r in mp_rows:
    mp_lookup.setdefault((r[2],r[3]),{})[sugar_of(r[5])]=r[7]

# Focus: products with multi-segment 制作 (杯中/冰沙/雪克杯 etc) to see | mapping
print("===== | SEPARATOR ORIGIN (multi-segment 制作) =====")
samples = [
 ("波波多肉葡萄","大杯"),  # 杯中/冰沙
 ("葡萄吨吨桶","桶"),        # 水果桶/雪克杯
 ("牛油果酸奶昔","中杯"),    # 杯中/冰沙
 ("手打柠檬红茶","大杯"),    # 雪克杯
 ("西瓜小宝杯","圣代杯"),    # 冰沙/杯顶
]
for nm,sz in samples:
    pwd=mp_lookup.get((nm,sz),{}).get('标准糖')
    print(f"\n■ {nm}/{sz}")
    # find 制作 in ingredient - need manual from earlier dump; instead print the mp pwd structure
    print(f"   密码: {pwd!r}")
    # count pipes
    if pwd:
        segs=pwd.split('|')
        print(f"   段数={len(segs)}: {segs}")
