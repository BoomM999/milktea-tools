#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_password.py — 标签简化+推荐糖配料表 -> 制作密码

用法:
  python gen_password.py <配料表.xlsx>                 # 打印全部商品解析
  python gen_password.py <配料表.xlsx> "商品名"         # 打印指定商品各糖度密码
  python gen_password.py <配料表.xlsx> "商品名" --sheet 国内产品
  python gen_password.py <配料表.xlsx> --export out.xlsx  # 展开全部(商品x规格x温度x糖度)导出

规则依据: 已与"制作密码导出_*.xlsx"样本逐条核对(见 SKILL.md)。
"""
import sys, os, re, zipfile
from xml.etree import ElementTree as ET
from collections import defaultdict

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# ---------- xlsx 底层读取(绕过 openpyxl 3.1.5 / py3.14 样式解析 bug) ----------
def col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

def load_xlsx(path):
    z = zipfile.ZipFile(path)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(f'{NS}si'):
            shared.append(''.join(t.text or '' for t in si.iter(f'{NS}t')))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = []
    for s in wb.find(f'{NS}sheets'):
        sheets.append((s.get('name'),
                       s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid = {r.get('Id'): r.get('Target') for r in rels}
    out = {}
    for name, ridx in sheets:
        tg = rid[ridx]
        if not tg.startswith('xl/'):
            tg = 'xl/' + tg
        root = ET.fromstring(z.read(tg))
        rows = []
        for row in root.iter(f'{NS}row'):
            cells = {}
            mc = -1
            for c in row.findall(f'{NS}c'):
                m = re.match(r'([A-Z]+)(\d+)', c.get('r'))
                ci = col_to_idx(m.group(1))
                t = c.get('t')
                v = c.find(f'{NS}v')
                isn = c.find(f'{NS}is')
                if t == 's' and v is not None:
                    val = shared[int(v.text)]
                elif t == 'inlineStr' and isn is not None:
                    val = ''.join(tt.text or '' for tt in isn.iter(f'{NS}t'))
                elif v is not None:
                    val = v.text
                else:
                    val = ''
                cells[ci] = val
                if ci > mc:
                    mc = ci
            rows.append([cells.get(i, '') for i in range(mc + 1)])
        out[name] = rows
    return out

# ---------- 文本规整 ----------
def half(s):
    return (s.replace('（', '(').replace('）', ')')
             .replace('：', ':').replace('、', '#')
             .replace('﹕', ':').replace('Ｔ', 'T').replace('Ｍ', 'M')
             .replace('，', ','))

# ---------- 糖度占位符解析 ----------
def parse_tm_group(grp):
    """从 (T:满40、半30、三20) 或 (满T35M15、半T25M10、三T15M5) 提取 T 与 M 各自的 (满,半,三)
    关键: T 与 M 在同组内满/半/三档数值不同, 必须分别提取。
    例 (满T35M15、半T25M10、三T15M5):
        T -> (满35, 半25, 三15)
        M -> (满15, 半10, 三5)
    例 (T:满40、半30、三20):
        T -> (满40, 半30, 三20), 无 M
    """
    g = re.sub(r'[、，,:：]', '', grp)  # 去掉分隔符与冒号, 紧贴便于提取
    has_T = bool(re.search(r'[TtＴ]', g))
    has_M = bool(re.search(r'[MmＭ]', g))

    def grab(level_re, sym):
        # level_re 如 r'满' ; 在 g 中找 level 后紧跟的 [T]?数字 [M]?数字
        # 兼容 满40 / 满T40 / 满T40M15
        m = re.search(level_re + r'[TtＴ]?(\d+)[MmＭ]?(\d+)?', g)
        if not m:
            return None
        a = m.group(1)
        b = m.group(2) if m.group(2) else None
        if sym == 'T':
            return a  # T 取第一个数字
        else:
            return b if b else a  # M 取第二个数字(若有)否则同第一个

    def extract_for(sym):
        full = grab(r'[满全]', sym)
        halfv = grab(r'半', sym)
        third = grab(r'三', sym)
        return (full, halfv, third)

    out = {}
    if has_T:
        out['T'] = extract_for('T')
    if has_M:
        out['M'] = extract_for('M')
    if not has_T and not has_M:
        # 无符号: 纯 (满40、半30、三20) 当作 T
        full = re.search(r'[满全](\d+)', g)
        halfv = re.search(r'半(\d+)', g)
        third = re.search(r'三(\d+)', g)
        if full or halfv or third:
            out['T'] = (full.group(1) if full else None,
                        halfv.group(1) if halfv else None,
                        third.group(1) if third else None)
    return out

def extract_tm(text):
    """返回 {'T': (满,半,三) or None, 'M': (满,半,三) or None}"""
    out = {'T': None, 'M': None}
    for m in re.finditer(r'[\(（][^\(\)（）]*[\)）]', text):
        g = parse_tm_group(m.group(0))
        if 'T' in g:
            out['T'] = g['T']
        if 'M' in g:
            out['M'] = g['M']
    return out

def replace_tm_in_segment(seg, sugar):
    """
    把一段(已 # 分隔、半角括号)中的 T/M 占位符按糖度替换。
    保留符号: 标准糖满 -> #T满# ; 七分/五分半 -> #T半# ; 三分三 -> #T三#
    不额外加糖/无糖 -> 删除整项(连同符号)
    支持 T/M 同组: (满T35M15、半T25M10、三T15M5) -> 标准糖: T35 M15
    """
    def val_for(sym_base, sugar):
        full, halfv, third = sym_base
        if sugar in ('标准糖',):
            return full
        elif sugar in ('七分糖', '五分糖'):
            return halfv if halfv is not None else full
        elif sugar == '三分糖':
            return third if third is not None else halfv
        else:
            return None  # 不额外加糖 -> 删除

    def repl(m):
        grp = m.group(0)
        tm = parse_tm_group(grp)
        parts = []
        if 'T' in tm:
            v = val_for(tm['T'], sugar)
            if v is not None:
                parts.append(f"T{v}")
        if 'M' in tm:
            v = val_for(tm['M'], sugar)
            if v is not None:
                parts.append(f"M{v}")
        return '#'.join(parts)  # 同组内 T/M 用 # 连接; 若全删则返回 ''
    seg = re.sub(r'[\(（][^\(\)（）]*[\)）]', repl, seg)

    # 处理已拆成 #T满# / #M满# 形式(清理不额外加糖残留)
    if sugar in ('不额外加糖', '无糖'):
        seg = re.sub(r'#?T\d+#?', '', seg)
        seg = re.sub(r'#?M\d+#?', '', seg)
        seg = re.sub(r'#{2,}', '#', seg).strip('#')
    return seg

# ---------- 杯型解析 ----------
def parse_cup_specs(spec_text):
    """'大杯\n冰/少冰：700磨+杯盖\n常温：22A+杯盖' -> [('冰/少冰','700磨+杯盖'),('常温','22A+杯盖')]"""
    out = []
    for line in spec_text.split('\n'):
        line = line.strip()
        if not line:
            continue
        if '：' in line:
            temp, cup = line.split('：', 1)
        elif ':' in line:
            temp, cup = line.split(':', 1)
        else:
            temp, cup = '', line
        out.append((temp.strip(), cup.strip()))
    return out

# 杯型名归一(配料表写法 -> 制作密码前缀)
CUP_NORM = [
    ('700磨砂注塑杯', '700磨+杯盖'),
    ('700磨', '700磨+杯盖'),
    ('700塑', '700塑'),
    ('咖啡专用杯', '咖啡杯'),
    ('咖啡杯', '咖啡杯'),
    ('22A纸杯', '22A+杯盖'),
    ('22A荔枝杯', '22A+杯盖'),
    ('22A', '22A+杯盖'),
    ('16A纸杯', '16A+杯盖'),
    ('16A', '16A+杯盖'),
    ('1L水果桶', '1L水果桶'),
    ('水果桶', '1L水果桶'),
    ('圣代杯', '圣代杯'),
    ('霸气华夫脆筒', '霸气华夫脆筒'),
]

def norm_cup(cup):
    for key, val in CUP_NORM:
        if key in cup:
            return val
    return cup

def pick_cup(specs, temp_pref='冰'):
    """选杯型: 优先含 temp_pref(冰) 的档;否则取第一个非空杯型。
    杯型文本里的尺寸词(大杯/中杯/桶)不是杯型名, 需剔除后取真实杯型描述。"""
    if not specs:
        return ''
    SIZE_WORDS = ('大杯', '中杯', '桶', '个')
    def real_cup(cup):
        # 去掉开头的尺寸词
        c = cup
        for sw in SIZE_WORDS:
            if c.startswith(sw):
                c = c[len(sw):].strip()
                break
        return c
    # 优先匹配含"冰"的温度描述
    for temp, cup in specs:
        if '冰' in temp and cup:
            return norm_cup(real_cup(cup))
    # 否则取第一个有真实杯型的
    for temp, cup in specs:
        if cup:
            rc = real_cup(cup)
            if rc:
                return norm_cup(rc)
    return ''

# ---------- 制作文本 -> 密码段 ----------
def make_to_segments(make_text, sugar=None):
    """'杯中：葡萄颗粒2（捣碎）、茶冻1、晶球1\n冰沙：葡萄汁60ml、绿120、冰270'
       -> ['杯中#葡萄颗粒2(捣碎)#茶冻1#晶球1', '冰沙#葡萄汁60ml#绿120#冰270']
    顺序: 先去段标签 -> 替换 T/M 占位符(此时 、 仍保留, 糖度组完整) -> 最后 、-># 与括号半角化"""
    segs = []
    for line in make_text.split('\n'):
        line = line.strip()
        if not line:
            continue
        # 去掉段标签冒号: 杯中：/ 冰沙：/ 雪克杯：/ 雪克：/ 水果桶：/ 杯顶：
        line = re.sub(r'^(杯中|冰沙|雪克杯|雪克|水果桶|杯顶)[：:]', '', line)
        # 若 sugar 给定, 先替换 T/M 占位符(、 尚在, 组完整)
        if sugar is not None:
            line = replace_tm_in_segment(line, sugar)
        # 最后统一转换: 、-># , 全角括号->半角, 全角冒号->半角
        line = (line.replace('（', '(').replace('）', ')')
                     .replace('：', ':').replace('、', '#')
                     .replace('﹕', ':').replace('Ｔ', 'T').replace('Ｍ', 'M')
                     .replace('，', ','))
        segs.append(line)
    return segs

# ---------- 商品块提取 ----------
HEADER_TOKENS = {'前场配料', '杯型', '制作（正常糖）', '制作(正常糖）', '制作(正常糖)',
                 '推荐糖', '推荐糖度', '制作(正常糖）'}

def build_blocks(sheet_rows):
    """每行(每个温度档)作为独立商品-规格单元。
    同一商品的多行 = 不同温度档各自一套配方(冰/少冰/常温/热), 不合并。
    仅当某行杯型为空时, 向右继承最近非空杯型。
    """
    blocks = []
    i = 0
    cur = None
    last_cup = ''
    while i < len(sheet_rows):
        r = sheet_rows[i]
        c0 = (r[0] if len(r) > 0 else '').strip()
        c1 = (r[1] if len(r) > 1 else '').strip()
        c2 = (r[2] if len(r) > 2 else '').strip()
        if c0 and c0 not in HEADER_TOKENS and c1 == '' and c2 == '':
            cur = c0
            last_cup = ''
            i += 1
            continue
        if c0 and c0 not in HEADER_TOKENS and (c1 or c2):
            # 该行自身即一个温度档配方
            sp = c1 if c1 else last_cup
            if c1:
                last_cup = c1
            mk = c2
            rec = r[-1].strip() if r else ''
            # 如果杯型为空且制作也为空, 跳过(纯继承行)
            if not sp and not mk:
                i += 1
                continue
            blocks.append({'section': cur, 'name': c0, 'spec': sp,
                            'make': mk, 'rec': rec,
                            'temp': _temp_of(sp)})
            i += 1
        else:
            i += 1
    return blocks

def _temp_of(spec):
    """从杯型文本推断温度档(用于多温度展开时选择杯型)"""
    for t in ('热', '常温', '少冰', '去冰', '冰'):
        if t in spec:
            return t
    return '冰'

def size_of(spec):
    for kw in ('大杯', '中杯', '桶', '个'):
        if kw in spec:
            return kw
    return '大杯'

SUGARS = ['标准糖', '七分糖', '五分糖', '三分糖', '不额外加糖']

def gen_password_for(name, spec, make, sugar):
    cup = pick_cup(parse_cup_specs(spec))
    segs = make_to_segments(make, sugar)
    # 每段内替换 T/M (make_to_segments 已按 sugar 处理; 此处保险再跑一次清理残留)
    new_segs = [replace_tm_in_segment(s, sugar) for s in segs]
    new_segs = [s for s in new_segs if s]
    body = '|'.join(new_segs)
    if cup:
        return f"{cup}|{body}" if body else cup
    return body

# ---------- 主流程 ----------
def collect_all(xlsx_path, sheet_filter=None):
    data = load_xlsx(xlsx_path)
    result = []
    for sheet, rows in data.items():
        if sheet_filter and sheet != sheet_filter:
            continue
        for b in build_blocks(rows):
            if not b.get('make'):
                continue
            spec = b['spec']; make = b['make']
            sz = size_of(spec)
            result.append({
                'sheet': sheet,
                'section': b['section'],
                'name': b['name'],
                'size': sz,
                'spec': spec,
                'make': make,
                'temp': b.get('temp', '冰'),
                'rec': b['rec'],
            })
    return result

def parse_combo(combo):
    """从 '去冰,七分糖' 或 '七分糖/去冰' 解析出 (温度, 糖度)。
    温度集合: 去冰/少冰/常温/正常冰/热 ; 糖度集合: SUGARS。"""
    parts = [t.strip() for t in re.split(r'[,，/]', combo) if t.strip()]
    temp = None; sugar = None
    for t in parts:
        if t in ('去冰', '少冰', '常温', '正常冰', '热'):
            temp = t
        elif t in SUGARS:
            sugar = t
    return temp, sugar

def get_ingredient(items, name):
    """从 collect_all 结果取某商品第一条(全温度通用)做法。"""
    for it in items:
        if it['name'] == name and it.get('make'):
            return it
    return None

def main():
    args = sys.argv[1:]
    if not args:
        print("用法: python gen_password.py <配料表.xlsx> [商品名] [--sheet 名] [--export out.csv] [--from-export 导出表.xlsx] [--fill-list 清单.txt]")
        sys.exit(1)
    xlsx = args[0]
    name_filter = None
    sheet_filter = None
    export = None
    from_export = None
    fill_list = None
    i = 1
    while i < len(args):
        a = args[i]
        if a == '--sheet':
            sheet_filter = args[i + 1]; i += 2
        elif a == '--export':
            export = args[i + 1]; i += 2
        elif a == '--from-export':
            from_export = args[i + 1]; i += 2
        elif a == '--fill-list':
            fill_list = args[i + 1]; i += 2
        else:
            name_filter = a; i += 1
    items = collect_all(xlsx, sheet_filter)
    if name_filter:
        items = [it for it in items if name_filter in it['name']]
    if not items:
        print("未找到匹配商品。")
        sys.exit(0)

    # 模式1: --fill-list 补用户贴的带 - 清单(斜杠/逗号标签均可)
    if fill_list:
        if not name_filter:
            print("--fill-list 需配合商品名参数。")
            sys.exit(1)
        ing = get_ingredient(items, name_filter)
        if not ing:
            print(f"未找到商品 {name_filter} 的配料。"); sys.exit(1)
        with open(fill_list, 'r', encoding='utf-8') as f:
            lines = [l.rstrip('\n') for l in f]
        out = []
        i = 0
        while i < len(lines):
            ln = lines[i]
            if not ln.strip():
                i += 1; continue
            temp, sugar = parse_combo(ln)
            if temp or sugar:
                # 标签行
                out.append(ln)
                nxt = lines[i + 1] if i + 1 < len(lines) else ''
                if nxt.strip() and '|' in nxt:
                    out.append(nxt)      # 已填密码, 保留
                    i += 2
                else:
                    pwd = gen_password_for(ing['name'], ing['spec'], ing['make'], sugar or '标准糖')
                    out.append(pwd)      # 补密码
                    i += 2 if nxt.strip() else 1
            else:
                # 非标签行(如独立密码或 - ), 原样保留并跳过
                out.append(ln)
                i += 1
        print('\n'.join(out))
        return

    # 模式2: --from-export 按导出表已有组合逐行生成(带标签)
    if from_export:
        if not name_filter:
            print("--from-export 需配合商品名参数。")
            sys.exit(1)
        ing = get_ingredient(items, name_filter)
        if not ing:
            print(f"未找到商品 {name_filter} 的配料。"); sys.exit(1)
        mp = load_xlsx(from_export)['制作密码组合']
        combos = [r[5] for r in mp[1:] if len(r) >= 8 and r[2] == name_filter]
        out = []
        for c in combos:
            _, sugar = parse_combo(c)
            pwd = gen_password_for(ing['name'], ing['spec'], ing['make'], sugar or '标准糖')
            out.append(f"{c}|{pwd}")
        print('\n'.join(out))
        return

    if export:
        # 展开 商品 x 规格 x 糖度 -> 行
        import csv
        with open(export, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['商品分类', '商品名称', '规格名称', '做法组合', '制作密码'])
            for it in items:
                for sugar in SUGARS:
                    pwd = gen_password_for(it['name'], it['spec'], it['make'], sugar)
                    w.writerow([it['section'] or '', it['name'], it['size'], sugar, pwd])
        print(f"已导出: {export} ({len(items)*len(SUGARS)} 行)")
        return
    # 默认: 打印
    for it in items:
        print(f"\n=== {it['name']} / {it['size']}  [{it['sheet']}] 推荐糖={it['rec']} ===")
        print(f"  杯型(配料): {it['spec']!r}")
        print(f"  制作(配料): {it['make']!r}")
        for sugar in SUGARS:
            pwd = gen_password_for(it['name'], it['spec'], it['make'], sugar)
            print(f"  [{sugar}] {pwd}")

if __name__ == '__main__':
    main()
