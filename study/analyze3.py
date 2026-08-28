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
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(f'{NS}si'):
            shared.append(''.join(t.text or '' for t in si.iter(f'{NS}t')))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = []
    for s in wb.find(f'{NS}sheets'):
        sheets.append((s.get('name'), s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid_to_target = {r.get('Id'): r.get('Target') for r in rels}
    result = {}
    for name, rid in sheets:
        target = rid_to_target.get(rid)
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
            rows_out.append([cells.get(i, '') for i in range(maxc+1)])
        result[name] = rows_out
    return result

base = os.path.dirname(os.path.abspath(__file__))
mp = load_xlsx(os.path.join(base, "制作密码导出.xlsx"))["制作密码组合"]

# Build structure: group by 商品名称 + 规格名称
from collections import defaultdict
groups = defaultdict(list)
for r in mp[1:]:
    if len(r) < 8:
        continue
    cat, spu, name, spec, sku, combo, comboid, pwd = r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7]
    groups[(name, spec)].append((cat, combo, comboid, pwd))

# Print a few food-tea products in full to see sugar mapping
targets = [("百香冰茶","大杯"), ("多肉杨枝甘露","大杯"), ("莓莓水果奶茶","大杯"), ("华夫珍珠奶","大杯"), ("手打柠檬红茶","大杯")]
print("############ DETAILED MAKE-PASSWORD GROUPS ############")
for key in targets:
    if key in groups:
        print(f"\n##### {key[0]} / {key[1]}  ({len(groups[key])} rows) #####")
        for g in groups[key]:
            print(f"  combo={g[1]!r:30} comboid={g[2]!r:22} pwd={g[3]!r}")

# Sugar level -> T value mapping analysis for 百香冰茶 (大杯)
print("\n\n############ SUGAR -> T MAPPING (百香冰茶 大杯) ############")
for g in groups[("百香冰茶","大杯")]:
    combo, pwd = g[1], g[3]
    # extract T token
    tm = re.search(r'#T(\d+)#', pwd)
    tval = tm.group(1) if tm else "NONE"
    print(f"  {combo:18} -> T={tval}")

# Scan all unique sugar tokens and unique T-presence patterns
print("\n\n############ ALL SUGAR TOKENS (col 做法组合) ############")
sugar_tokens = defaultdict(int)
for r in mp[1:]:
    if len(r) < 6: continue
    combo = r[5]
    for tok in combo.split(','):
        sugar_tokens[tok.strip()] += 1
for k,v in sorted(sugar_tokens.items(), key=lambda x:-x[1]):
    print(f"  {k}: {v}")

# For combos containing sugar, see T value distribution
print("\n\n############ SUGAR LEVEL -> T VALUE (where T present) ############")
# We'll map per product base name (strip sugar/ice) to list of (sugar, T)
prod_sugar_t = defaultdict(list)
for r in mp[1:]:
    if len(r) < 8: continue
    name, spec, combo, pwd = r[2], r[3], r[5], r[7]
    sugar = [t.strip() for t in combo.split(',') if '糖' in t]
    if not sugar: 
        continue
    sugar = sugar[0]
    tm = re.search(r'#T(\d+)#', pwd)
    tval = tm.group(1) if tm else None
    prod_sugar_t[(name, spec)].append((sugar, tval))

# aggregate: for each sugar level, what T values appear
sugar_tvals = defaultdict(set)
sugar_count = defaultdict(int)
for key, lst in prod_sugar_t.items():
    for sugar, tval in lst:
        sugar_tvals[sugar].add(tval)
        sugar_count[sugar]+=1
for k in sorted(sugar_tvals):
    print(f"  {k:8} count={sugar_count[k]:5} Tvalues={sorted(sugar_tvals[k], key=lambda x:(x is None, x))}")
