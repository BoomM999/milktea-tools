import sys, os, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
import gen_password as G

base = os.path.dirname(os.path.abspath(__file__))
ING = os.path.join(base, "标签简化_推荐糖配料表.xlsx")
MP = os.path.join(base, "制作密码导出.xlsx")

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

items = G.collect_all(ING)
gen_lookup={}
for it in items:
    for sugar in G.SUGARS:
        pwd = G.gen_password_for(it['name'], it['spec'], it['make'], sugar)
        gen_lookup.setdefault((it['name'],it['size']),{})[sugar]=pwd

# 收集所有不匹配, 打印原始配料 + 导出真实值 (聚焦糖度T/M差异)
print("########## 糖度不匹配明细(原始配料 vs 导出) ##########")
shown=0
for key, sugs in gen_lookup.items():
    if key not in mp_lookup: continue
    for sugar in sugs:
        real = mp_lookup[key].get(sugar)
        gen = sugs[sugar]
        if real is None: continue
        rt = re.findall(r'[TM]\d+', real.split('|',1)[1] if '|' in real else real)
        gt = re.findall(r'[TM]\d+', gen.split('|',1)[1] if '|' in gen else gen)
        if sorted(rt)!=sorted(gt):
            # 找到该商品配料原文
            src=[it for it in items if (it['name'],it['size'])==key]
            make = src[0]['make'] if src else '?'
            spec = src[0]['spec'] if src else '?'
            print(f"\n■ {key} / {sugar}")
            print(f"   配料杯型: {spec!r}")
            print(f"   配料制作: {make!r}")
            print(f"   导出密码: {real!r}")
            print(f"   生成密码: {gen!r}")
            shown+=1
            if shown>=25: break
    if shown>=25: break
print(f"\n共显示 {shown} 条")
