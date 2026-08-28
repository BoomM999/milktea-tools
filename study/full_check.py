import sys, os, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
import gen_password as G

base = os.path.dirname(os.path.abspath(__file__))
ING = os.path.join(base, "标签简化_推荐糖配料表.xlsx")
MP = os.path.join(base, "制作密码导出.xlsx")

# load real export
mp_data = G.load_xlsx(MP)["制作密码组合"]
def sugar_of(c):
    for tok in c.split(','):
        t=tok.strip()
        if t in ('标准糖','七分糖','五分糖','三分糖','不额外加糖','无糖'): return t
    return None
mp_lookup={}
for r in mp_data[1:]:
    if len(r)<8: continue
    mp_lookup.setdefault((r[2],r[3]),{})[sugar_of(r[5])]=r[7]

# generate from ingredient
items = G.collect_all(ING)
gen_lookup={}
for it in items:
    for sugar in G.SUGARS:
        pwd = G.gen_password_for(it['name'], it['spec'], it['make'], sugar)
        gen_lookup.setdefault((it['name'],it['size']),{})[sugar]=pwd

# compare: for each (name,size) present in BOTH, compare each sugar's password (structural, ignore 门店措辞归一差异)
# We compare the T/M numeric values and segment structure, normalize 物料名称差异 by comparing only the T/M tokens and # count
def norm_pwd(p):
    if not p: return ''
    # extract cup prefix
    if '|' in p:
        cup, body = p.split('|',1)
    else:
        cup, body = '', p
    # normalize 物料名称: keep only structure markers and T/M numbers
    # replace any non-(#|T|M|digit) runs with nothing to focus on sugar mapping
    tm = re.findall(r'[TM]\d+', body)
    segs = body.split('|')
    nsegs = len(segs)
    return (cup, nsegs, tm)

total=0; exact=0; sugar_ok=0; mism=[]
for key, sugs in gen_lookup.items():
    if key not in mp_lookup: continue
    for sugar in sugs:
        real = mp_lookup[key].get(sugar)
        gen = sugs[sugar]
        if real is None: continue
        total+=1
        # compare T/M values
        rt = re.findall(r'[TM]\d+', real.split('|',1)[1] if '|' in real else real)
        gt = re.findall(r'[TM]\d+', gen.split('|',1)[1] if '|' in gen else gen)
        if sorted(rt)==sorted(gt):
            sugar_ok+=1
        else:
            mism.append((key,sugar,'TM',rt,gt))
        # exact structural (cup+nsegs+tm)
        if norm_pwd(real)==norm_pwd(gen):
            exact+=1

print(f"对照总数: {total}")
print(f"糖度T/M数值完全吻合: {sugar_ok}/{total} ({100*sugar_ok/total:.1f}%)")
print(f"结构(cup+段数+TM)完全吻合: {exact}/{total} ({100*exact/total:.1f}%)")
print(f"\n糖度不吻合样例(前15):")
for m in mism[:15]:
    print(f"  {m[0]} / {m[1]}: 导出TM={m[3]} 生成TM={m[4]}")
