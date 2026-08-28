import zipfile, re, os
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def col_to_idx(col):
    n = 0
    for ch in col:
        n = n*26 + (ord(ch)-ord('A')+1)
    return n-1

def load_xlsx(path):
    z = zipfile.ZipFile(path)
    # shared strings
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(f'{NS}si'):
            # concatenate all text in this si
            txt = ''.join(t.text or '' for t in si.iter(f'{NS}t'))
            shared.append(txt)
    # workbook sheet names + rels
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = []
    for s in wb.find(f'{NS}sheets'):
        sheets.append((s.get('name'), s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')))
    # rels map
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid_to_target = {}
    for r in rels:
        rid_to_target[r.get('Id')] = r.get('Target')
    result = {}
    for name, rid in sheets:
        target = rid_to_target.get(rid)
        if not target:
            continue
        if not target.startswith('xl/'):
            target = 'xl/' + target
        root = ET.fromstring(z.read(target))
        rows_out = []
        for row in root.iter(f'{NS}row'):
            cells = {}
            maxc = -1
            for c in row.findall(f'{NS}c'):
                ref = c.get('r')
                m = re.match(r'([A-Z]+)(\d+)', ref)
                ci = col_to_idx(m.group(1))
                t = c.get('t')
                v = c.find(f'{NS}v')
                isnode = c.find(f'{NS}is')
                if t == 's' and v is not None:
                    val = shared[int(v.text)]
                elif t == 'inlineStr' and isnode is not None:
                    val = ''.join(tt.text or '' for tt in isnode.iter(f'{NS}t'))
                elif v is not None:
                    val = v.text
                else:
                    val = ''
                cells[ci] = val
                if ci > maxc:
                    maxc = ci
            rowlist = [cells.get(i, '') for i in range(maxc+1)]
            rows_out.append(rowlist)
        result[name] = rows_out
    return result

def dump(path, label, max_rows=80):
    print("="*90)
    print(f"FILE: {label}  ({os.path.basename(path)})")
    print("="*90)
    data = load_xlsx(path)
    for sheet, rows in data.items():
        print(f"\n=== SHEET: {sheet}  rows={len(rows)} ===")
        for i, r in enumerate(rows):
            if i >= max_rows:
                print(f"... truncated, total {len(rows)} rows")
                break
            print(f"[{i}] " + " | ".join("" if c is None else str(c) for c in r))

base = os.path.dirname(os.path.abspath(__file__))
dump(os.path.join(base, "标签简化_推荐糖配料表.xlsx"), "标签简化_推荐糖配料表")
dump(os.path.join(base, "制作密码导出.xlsx"), "制作密码导出")
