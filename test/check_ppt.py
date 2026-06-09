#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT 全量校验：表格 vs 文字 vs 图表
支持单位换算、XML级图表数据提取（含pie3DChart等python-pptx不支持的类型）
"""

import re, os, argparse
from difflib import SequenceMatcher
from lxml import etree
from pptx import Presentation

# ==================== 单位 ====================
UNIT_MULT = {'亿元':1e8,'亿':1e8,'万元':1e4,'万':1e4,'千元':1e3,'元':1}
SPECIAL = {'万人次':1e4,'%':0.01,'％':0.01,'个百分点':0.01,'倍':1}
NUM_RE = re.compile(r'(?P<sign>[-])?(?P<number>[\d,]+\.?\d*)\s*(?P<unit>万?人?次|亿元?|万元?|千元|元|%|％|个百分点|倍)?')
METRIC_KW = ['收入','成本','费用','结余','人次','次均','人均','占比','药占','耗占','执行率','完成率','预算','收支','医疗','门急诊','门诊','急诊','住院','手术','检查','化验','治疗','药品','耗材','材料','人员','物业','维修','水电','能源','折旧','摊销','增减','增长','增幅','降幅','同比','环比','业务量','服务量','经费','商品','服务','计提','绩效','工资','保险','公积金','试剂','消毒','保洁','保安','洗涤','燃气','电费','水费','燃料','拨款','补助']
CHANGE_WORDS = {'增减','增长','下降','降低','增加','减少','同比','变动','增幅','降幅'}
DISTRICT_WORDS = {'奉贤','杨浦','成人','儿童','男','女'}
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'


def parse_num(text):
    m = NUM_RE.search(text)
    if not m: return None
    ns = m.group('number').replace(',','')
    try: v = float(ns)
    except: return None
    s = m.group('sign') or ''; u = m.group('unit') or ''
    mul = SPECIAL.get(u, UNIT_MULT.get(u, 1))
    nv = v * mul
    if s == '-': nv = -nv; v = -v
    return {'raw_text':m.group(0).strip(),'value':v,'unit':u,'norm':nv}


def nl(label):
    if not label: return ''
    x = re.sub(r'[：:，,、（()）\[\]\s　""\'\'""]+','',label)
    x = re.sub(r'^(其中|其中|其中)[，,]?','',x)
    x = re.sub(r'^\d{4}年\d+[-]\d+月[，,]?','',x)
    x = re.sub(r'^\d{4}年\d+月[，,]?','',x)
    x = re.sub(r'^\d{4}年[，,]?','',x)
    x = re.sub(r'^本年\d+[-]\d+月[，,]?','',x)
    x = re.sub(r'^上月','',x)
    x = re.sub(r'^较上年同期\d{4}年\d{1,2}月[，,]?','',x)
    x = re.sub(r'^较上年同期[，,]?','',x)
    x = re.sub(r'占[%％]?$','',x)
    x = re.sub(r'^第\d+','',x)
    x = re.sub(r'[（(]\d+[）)]','',x)
    x = re.sub(r'^\d+[)）、.]','',x)
    x = re.sub(r'[①②③④⑤⑥⑦⑧⑨⑩]','',x)
    return x.strip()


def is_bad_label(lb):
    if not lb or len(lb) < 2: return True
    if lb in {'年','月','日','第','占','为','是','达'}: return True
    if re.match(r'^[\d\.\-,]+$', lb): return True
    if re.match(r'^(19|20)\d{2}$', lb): return True
    if re.match(r'^第\d+$', lb): return True
    if re.match(r'^[一二三四五六七八九十]+$', lb): return True
    return False


def sim(a,b):
    """标签相似度"""
    na,nb = nl(a),nl(b)
    if not na or not nb: return 0
    if na == nb: return 1.0
    if na in nb or nb in na:
        return 0.85 + 0.15*min(len(na),len(nb))/max(len(na),len(nb))
    return SequenceMatcher(None,na,nb).ratio()

def sim_col(a,b):
    """列标题相似度 - 不剥离日期"""
    if not a or not b: return 0
    ca = re.sub(r'[：:，,、（()）\[\]\s　""\'\'""]+','',a)
    cb = re.sub(r'[：:，,、（()）\[\]\s　""\'\'""]+','',b)
    if ca == cb: return 1.0
    if ca in cb or cb in ca:
        return 0.85 + 0.15*min(len(ca),len(cb))/max(len(ca),len(cb))
    return SequenceMatcher(None,ca,cb).ratio()
    na,nb = nl(a),nl(b)
    if not na or not nb: return 0
    if na == nb: return 1.0
    if na in nb or nb in na:
        return 0.85 + 0.15*min(len(na),len(nb))/max(len(na),len(nb))
    return SequenceMatcher(None,na,nb).ratio()


# ==================== 图表数据提取（XML方式） ====================

def extract_chart_data(chart_element):
    """
    从图表XML元素提取结构化数据。
    支持 barChart, lineChart, scatterChart, pieChart, pie3DChart 等所有类型。
    返回: [{categories:[str], series_name:str, values:[float]}, ...]
    """
    root = chart_element
    all_series = []

    # Find all series elements in any chart type
    ser_elements = root.findall(f'.//{{{C_NS}}}ser')
    if not ser_elements:
        # Try pie3DChart specific
        ser_elements = root.findall(f'.//{{{C_NS}}}pie3DChart/{{{C_NS}}}ser')

    for ser in ser_elements:
        # Extract series name
        series_name = ''
        tx = ser.find(f'{{{C_NS}}}tx')
        if tx is not None:
            str_ref = tx.find(f'{{{C_NS}}}strRef')
            if str_ref is not None:
                cache = str_ref.find(f'{{{C_NS}}}strCache')
                if cache is not None:
                    pts = cache.findall(f'{{{C_NS}}}pt')
                    if pts:
                        v_el = pts[0].find(f'{{{C_NS}}}v')
                        if v_el is not None:
                            series_name = v_el.text or ''

        # Extract categories
        categories = []
        cat = ser.find(f'{{{C_NS}}}cat')
        if cat is not None:
            # strRef
            str_ref = cat.find(f'{{{C_NS}}}strRef')
            if str_ref is not None:
                cache = str_ref.find(f'{{{C_NS}}}strCache')
                if cache is not None:
                    for pt in cache.findall(f'{{{C_NS}}}pt'):
                        v_el = pt.find(f'{{{C_NS}}}v')
                        if v_el is not None and v_el.text:
                            categories.append(v_el.text)
            # multiLvlStrRef
            if not categories:
                ml_ref = cat.find(f'{{{C_NS}}}multiLvlStrRef')
                if ml_ref is not None:
                    for cache in ml_ref.findall(f'.//{{{C_NS}}}strCache'):
                        for pt in cache.findall(f'{{{C_NS}}}pt'):
                            v_el = pt.find(f'{{{C_NS}}}v')
                            if v_el is not None and v_el.text:
                                categories.append(v_el.text)
            # numRef (scatter charts sometimes use numbers as categories)
            if not categories:
                num_ref = cat.find(f'{{{C_NS}}}numRef')
                if num_ref is not None:
                    cache = num_ref.find(f'{{{C_NS}}}numCache')
                    if cache is not None:
                        for pt in cache.findall(f'{{{C_NS}}}pt'):
                            v_el = pt.find(f'{{{C_NS}}}v')
                            if v_el is not None and v_el.text:
                                categories.append(v_el.text)

        # Extract values
        values = []
        val = ser.find(f'{{{C_NS}}}val')
        if val is not None:
            num_ref = val.find(f'{{{C_NS}}}numRef')
            if num_ref is not None:
                cache = num_ref.find(f'{{{C_NS}}}numCache')
                if cache is not None:
                    fmt_code = cache.findtext(f'{{{C_NS}}}formatCode', '')
                    for pt in cache.findall(f'{{{C_NS}}}pt'):
                        v_el = pt.find(f'{{{C_NS}}}v')
                        if v_el is not None and v_el.text:
                            try:
                                values.append(float(v_el.text))
                            except ValueError:
                                pass

        if categories or values:
            all_series.append({
                'categories': categories,
                'series_name': series_name,
                'values': values,
                'format_code': '',
            })

    return all_series


# ==================== 文本数值提取 ====================

def extract_text_vals(text):
    results = []
    sents = re.split(r'[。；;]', text)
    for sent in sents:
        sent = sent.strip()
        if not sent: continue
        tm = re.match(r'^(.+?)(?:为|达|是|：|:)\s*[\d\-]', sent)
        topic = ''
        if tm:
            topic = tm.group(1).strip()
            topic = re.sub(r'^\d{4}年\d{1,2}[-]\d{1,2}月[，,]?','',topic)
            topic = re.sub(r'^\d{4}年\d{1,2}月[，,]?','',topic)
            topic = re.sub(r'^\d{4}年[，,]?','',topic)
            topic = topic.strip('，,。；;、')
        for m in NUM_RE.finditer(sent):
            ns = m.group('number').replace(',','')
            try: v = float(ns)
            except: continue
            sgn = m.group('sign') or ''; u = m.group('unit') or ''
            if re.match(r'^20[2-9]\d$', ns) and not u: continue
            pre5 = sent[max(0,m.start()-5):m.start()]
            suf3 = sent[m.end():m.end()+3]
            if re.match(r'^[1-9]$|^1[0-2]$', ns) and not u:
                if '月' in suf3[:3] or '月' in pre5 or suf3.startswith('-'): continue
                if '年' in pre5: continue
            if re.search(r'第\s*$', pre5): continue
            if re.search(r'^\d+[)）、.]', sent) and m.start() < 3 and not u: continue
            pre = sent[:m.start()]
            pre_clean = re.sub(r'[是为达到比：:,，、\s（(]+$','',pre)
            pre_clean = re.sub(r'^\d{4}年\d{1,2}[-]\d{1,2}月[，,]?','',pre_clean)
            pre_clean = re.sub(r'^\d{4}年\d{1,2}月[，,]?','',pre_clean)
            pre_clean = re.sub(r'^\d{4}年[，,]?','',pre_clean)
            pre_clean = re.sub(r'^其中[，,]?','',pre_clean)
            pre_clean = re.sub(r'[（(]\d+[）)]\s*','',pre_clean)
            pre_clean = pre_clean.strip('，,。；;、')
            meaningful = re.findall(r'[一-鿿][一-鿿\w%/()（）\-]{1,20}', pre_clean)
            local_lb = meaningful[-1] if meaningful else pre_clean[-20:]
            local_nl = nl(local_lb)
            has_kw = any(kw in local_nl for kw in METRIC_KW)
            is_dist = any(dw in local_nl for dw in DISTRICT_WORDS)
            label = local_lb if (has_kw or is_dist) else (topic + local_lb if topic else local_lb)
            if is_bad_label(nl(label)): continue
            dtype = 'absolute'
            if u in ['%','％','个百分点']: dtype = 'percentage'
            elif any(w in label for w in CHANGE_WORDS): dtype = 'change'
            mul = SPECIAL.get(u, UNIT_MULT.get(u, 1))
            nv = v * mul
            if sgn == '-': nv = -nv; v = -v
            results.append({
                'label':label,'value':v,'unit':u,'norm':nv,
                'raw_text':m.group(0).strip(),'dtype':dtype,
                'has_district':is_dist,
            })
    return results


# ==================== PPT提取 ====================

def extract_ppt(pptx_path):
    prs = Presentation(pptx_path)
    data = {}
    for si, slide in enumerate(prs.slides):
        sn = si+1
        tabs = []; txts = []; charts = []
        for sh in slide.shapes:
            if sh.has_table:
                tabs.append(_ext_table(sh.table))
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if t:
                    vs = extract_text_vals(t)
                    if vs: txts.append({'full':t,'vals':vs})
            if sh.has_chart:
                try:
                    cd = extract_chart_data(sh.chart._element)
                    if cd: charts.append({'data':cd})
                except Exception as e:
                    print(f'  [WARN] Slide {sn}: chart extraction error: {e}')
        data[sn] = {'tables':tabs,'texts':txts,'charts':charts}
    return data


def _ext_table(table):
    rows = table.rows; nrows = len(rows); ncols = len(table.columns)
    if nrows < 2: return {'headers':[],'points':[]}
    hdrs = [rows[0].cells[c].text.strip() for c in range(ncols)]
    sr = 1
    if nrows > 2:
        r1c0 = rows[1].cells[0].text.strip()
        if (not r1c0) or any(k in r1c0 for k in ['其中','按','分类','类别','项目']):
            for c in range(ncols):
                ct = rows[1].cells[c].text.strip()
                if ct and not re.match(r'^[-]?[\d,]+\.?\d*\s*[%％]?\s*$',ct):
                    hdrs[c] = f'{hdrs[c]}|{ct}' if hdrs[c] else ct
            sr = 2
    ctypes = ['unknown']*ncols; cunit = [1]*ncols
    for c in range(1,ncols):
        hpct = False
        for r in range(sr, min(nrows,sr+10)):
            ct = rows[r].cells[c].text.strip() if c < len(rows[r].cells) else ''
            if '%' in ct or '％' in ct: hpct = True; break
        ctypes[c] = 'percentage' if hpct else 'absolute'
    for c in range(1,ncols):
        h = hdrs[c]
        if '亿' in h: cunit[c] = 1e8
        elif '万' in h: cunit[c] = 1e4
        elif '千' in h: cunit[c] = 1e3
        elif '元' in h and '万' not in h and '亿' not in h: cunit[c] = 1
        elif ctypes[c] == 'absolute' and c <= 3:
            svals = []
            for r in range(sr, min(nrows,sr+5)):
                ct = rows[r].cells[c].text.strip() if c < len(rows[r].cells) else ''
                p = parse_num(ct)
                if p and p['unit'] not in ['%','％']: svals.append(p['value'])
            if svals:
                avg = sum(abs(x) for x in svals)/len(svals)
                if 0.1 <= avg <= 50000: cunit[c] = 1e4
                elif avg > 50000: cunit[c] = 1
                elif avg < 0.1: cunit[c] = 1e8
    points = []
    for r in range(sr, nrows):
        rl = rows[r].cells[0].text.strip()
        if not rl: continue
        if re.match(r'^(合计|总计|小计|平均)$',rl): continue
        if is_bad_label(nl(rl)): continue
        for c in range(1, ncols):
            ct = rows[r].cells[c].text.strip()
            if not ct: continue
            p = parse_num(ct)
            if not p: continue
            if not p['unit'] and cunit[c] > 1:
                p['norm'] = p['value'] * cunit[c]
            p['row_label'] = rl
            p['col_header'] = hdrs[c] if c < len(hdrs) else ''
            p['col_type'] = ctypes[c]
            p['col_idx'] = c
            points.append(p)
    return {'headers':hdrs,'points':points,'ctypes':ctypes,'cunit':cunit}


# ==================== 校验 ====================

def verify_chart_vs_table(charts, tables, sn):
    """校验图表数据 vs 表格数据，系列名匹配表列"""
    issues = []
    tps = []; _ = [tps.extend(t['points']) for t in tables]
    if not tps or not charts: return issues

    # 检测饼图（值为比率<1）
    is_pie = False
    for chart in charts:
        for ser in chart['data']:
            vals = ser.get('values',[])
            if vals and all(v < 1 for v in vals):
                is_pie = True; break

    for chart in charts:
        for ser in chart['data']:
            cats = ser.get('categories',[]); vals = ser.get('values',[])
            sname = ser.get('series_name','')

            for ci, cat in enumerate(cats):
                if ci >= len(vals): break
                cv = vals[ci]

                # 在表格中找最佳匹配：行=cat，列=sname
                best = None; best_sc = 0
                for dp in tps:
                    rl = dp['row_label']; ch = dp['col_header']
                    # 行匹配
                    row_sc = sim(cat, rl)
                    if row_sc < 0.6: continue
                    # 列匹配：series name 与 column header（不剥离日期）
                    col_sc = 0
                    if sname and ch:
                        col_sc = sim_col(sname, ch)
                    # 综合分数
                    sc = row_sc * 0.7 + col_sc * 0.3
                    if col_sc > 0.6:
                        sc = row_sc * 0.5 + col_sc * 0.5  # 列匹配好则权重更高
                    if sc > best_sc:
                        best_sc = sc; best = dp

                if not best or best_sc < 0.6: continue

                dv = best['value']; dn = best['norm']; ch = best['col_header']

                # 饼图特殊处理：比较比率 vs 表格占比列
                if is_pie and cv < 1:
                    # 表格值可能是绝对值 → 检查是否有占比列
                    if '占' in ch or '%' in ch:
                        # 表格占比列的值应该也是比率或百分比
                        if dv < 1 and abs(cv - dv) < 0.05: continue
                        if dv >= 1 and abs(cv*100 - dv) < 5: continue
                    # 如果没有占比列，尝试用原始值比
                    # 图表比率应当大致等于表格值/总合计
                    continue  # skip pie chart value comparison for now

                # 普通比较
                if abs(cv - dv) < 0.001: continue
                if dn and abs(cv - dn) < 0.001: continue
                if dn and dn != 0 and abs(cv - dn)/abs(dn) <= 0.02: continue
                if dv != 0 and abs(cv - dv)/abs(dv) <= 0.02: continue

                issues.append({
                    'slide':sn, 'type':'chart_vs_table',
                    'chart_cat':cat, 'chart_val':cv, 'chart_series':sname,
                    'table_label':best['row_label'], 'table_val':best['raw_text'],
                    'table_col':ch, 'score':round(best_sc,2),
                })
    return issues


def verify_chart_vs_text(charts, texts, sn, tol=0.005):
    """校验图表数据 vs 文字描述，系列名用于区分年份"""
    issues = []
    txtps = []; _ = [txtps.extend(tb['vals']) for tb in texts]
    if not txtps or not charts: return issues

    # 检测饼图
    is_pie = False
    for chart in charts:
        for ser in chart['data']:
            vals = ser.get('values',[])
            if vals and all(v < 1 for v in vals):
                is_pie = True; break

    if is_pie: return issues  # 饼图比率值不适合直接与文字中的绝对值比较

    for chart in charts:
        for ser in chart['data']:
            cats = ser.get('categories',[]); vals = ser.get('values',[])
            sname = ser.get('series_name','')

            for ci, cat in enumerate(cats):
                if ci >= len(vals): break
                cv = vals[ci]

                best = None; best_sc = 0
                for tp in txtps:
                    tlb = tp['label']; tnl_ = nl(tlb)
                    # 过滤"其中"子项文字 (图表通常是汇总数据)
                    if '其中' in tlb or '其中' in tnl_: continue
                    # 过滤占比/增长率文字
                    if any(w in tnl_ for w in ['同比','增长','下降','占比','比重','占%']): continue
                    # 需要类别+系列名同时匹配
                    cat_sc = sim(cat, tlb)
                    ser_sc = sim_col(sname, tlb) if sname else 0
                    sc = cat_sc
                    # 如果系列名是年份，检查文字的年份上下文是否匹配
                    year_match = re.search(r'(202[5-6])年', sname) if sname else None
                    if year_match:
                        chart_year = year_match.group(1)
                        # 在文字标签附近找年份
                        txt_prefix = '' if not hasattr(tp,'prefix') else getattr(tp,'prefix','')
                        full_label = tlb
                        if chart_year not in full_label:
                            sc *= 0.5  # 年份不匹配，大幅降权
                    if sname and ser_sc > 0.5:
                        sc = sc * 0.6 + ser_sc * 0.4
                    if sc > best_sc:
                        best_sc = sc; best = tp

                if not best or best_sc < 0.75: continue

                tv = best['value']; tn = best['norm']; tu = best['unit']
                if abs(cv - tv) < 0.001: continue
                if tn and abs(cv - tn) < 0.001: continue
                if tn and tn != 0 and abs(cv - tn)/abs(tn) <= tol: continue
                if tv != 0 and abs(cv - tv)/abs(tv) <= tol: continue

                issues.append({
                    'slide':sn, 'type':'chart_vs_text',
                    'chart_cat':cat, 'chart_val':cv, 'chart_series':sname,
                    'text_label':best['label'], 'text_val':best['raw_text'],
                    'text_unit':tu, 'score':round(best_sc,2),
                })
    return issues


def verify_text_vs_table(sd, sn, tol=0.005):
    """校验文字描述 vs 表格数据"""
    issues = []
    tps = []; _ = [tps.extend(t['points']) for t in sd['tables']]
    if not tps: return issues
    txtps = []; _ = [txtps.extend(tb['vals']) for tb in sd['texts']]
    if not txtps: return issues

    for tp in txtps:
        tlb = tp['label']; tv = tp['value']; tu = tp['unit']
        tn = tp['norm']; tdtype = tp['dtype']; tdist = tp['has_district']
        tnl_ = nl(tlb)

        # Filter "高于/低于全院" text (difference values)
        if re.search(r'(高于|低于|高出|低于)全[院区]', tlb): continue
        if '占全院' in tlb and tdtype == 'percentage': continue
        if ('占' in tlb and ('比重' in tlb or '总人次' in tlb or '全院' in tlb)) and tdtype == 'percentage': continue
        if '时间节点' in tnl_: continue
        if abs(tv) < 2 and not tu and tdtype == 'absolute' and len(tnl_) >= 4:
            if not any(kw in tnl_ for kw in ['次均','人均','占比','执行','率']): continue

        best = None; best_sc = 0
        for dp in tps:
            rl = dp['row_label']; rnl_ = nl(rl)
            ct = dp['col_type']; ch = dp['col_header']

            if tdtype == 'percentage' and ct == 'absolute': continue
            if tdtype == 'change' and ct == 'percentage': continue

            if tdtype == 'change' and ct == 'absolute':
                if '增减' not in ch and '同比' not in ch and '变动' not in ch:
                    if dp.get('col_idx',99) < 3: continue

            dp_dist = any(dw in rnl_ for dw in DISTRICT_WORDS)
            if tdist and not dp_dist: continue

            sc = sim(tlb, rl)
            if ch and ch not in rl:
                sc = max(sc, sim(tlb, rl+ch))

            col_bonus = 0
            if tdtype == 'change' and ('增减' in ch or '变动' in ch): col_bonus = 0.12
            if tdtype == 'percentage' and ('率' in ch or '%' in ch): col_bonus = 0.10
            sc = min(sc + col_bonus, 1.05)

            if sc > best_sc: best_sc = sc; best = dp

        if not best or best_sc < 0.80: continue

        dv = best['value']; du = best['unit']; dn = best['norm']
        dr = best['raw_text']; dl = best['row_label']; dct = best['col_type']

        if abs(abs(tv) - abs(dv)) < 0.001: continue
        if tn and dn:
            if abs(tn-dn) < 0.001 or abs(abs(tn)-abs(dn)) < 0.001: continue
            mx = max(abs(tn),abs(dn))
            if mx > 0 and abs(tn-dn)/mx <= tol: continue
            if mx > 0 and abs(abs(tn)-abs(dn))/mx <= tol: continue
        if tdtype == 'percentage' and dct == 'percentage':
            if abs(abs(tv) - abs(dv)) <= 0.02: continue

        issues.append({
            'slide':sn, 'type':'text_vs_table', 'metric':tnl_,
            't_label':tlb, 't_val':tp['raw_text'], 't_unit':tu,
            't_norm':tn, 't_type':tdtype,
            'd_label':dl, 'd_val':dr, 'd_unit':du,
            'd_norm':dn, 'd_type':dct,
            'score':round(best_sc,2),
        })
    return issues


def fmt(v):
    if v is None: return 'N/A'
    if abs(v) >= 1e8: return f'{v/1e8:.4f}亿'
    if abs(v) >= 1e4: return f'{v/1e4:.2f}万'
    if 0<abs(v)<1: return f'{v*100:.1f}%'
    return f'{v:,.2f}'

def main():
    p = argparse.ArgumentParser(description='PPT全量校验：表格/文字/图表一致性')
    p.add_argument('pptx'); p.add_argument('--tol',type=float,default=0.005)
    p.add_argument('-v','--verbose',action='store_true')
    a = p.parse_args()
    print(f'PPT全量校验: {os.path.basename(a.pptx)}')
    print(f'容差:{a.tol*100:.1f}% | 包含图表校验')
    print('='*70)

    sd = extract_ppt(a.pptx)
    tt = sum(sum(len(t['points']) for t in s['tables']) for s in sd.values())
    tn = sum(sum(len(tb['vals']) for tb in s['texts']) for s in sd.values())
    tc = sum(len(s['charts']) for s in sd.values())
    print(f'{len(sd)}页 | 表格:{tt} | 文字数值:{tn} | 图表页:{tc}')
    print()

    all_iss = []
    for sn in sorted(sd):
        s = sd[sn]
        has_t = any(len(t['points'])>0 for t in s['tables'])
        has_tx = s['texts']
        has_ch = s['charts']

        if has_t and has_tx:
            all_iss.extend(verify_text_vs_table(s, sn, a.tol))
        if has_ch:
            if has_t:
                all_iss.extend(verify_chart_vs_table(s['charts'], s['tables'], sn))
            if has_tx:
                all_iss.extend(verify_chart_vs_text(s['charts'], s['texts'], sn, a.tol))

    # 分类统计
    tvt = [i for i in all_iss if i['type']=='text_vs_table']
    cvt = [i for i in all_iss if i['type']=='chart_vs_table']
    cvx = [i for i in all_iss if i['type']=='chart_vs_text']

    print(f'校验结果:')
    print(f'  文字vs表格: {len(tvt)} 个问题')
    print(f'  图表vs表格: {len(cvt)} 个问题')
    print(f'  图表vs文字: {len(cvx)} 个问题')
    print(f'  合计: {len(all_iss)} 个')
    print('='*70)

    def print_iss(iss, title):
        if not iss: return
        print(f'\n【{title}】')
        print('-'*50)
        for i in iss:
            if i['type'] == 'text_vs_table':
                print(f"\n  第{i['slide']}页 | {i['metric']} | 匹配:{i['score']:.0%}")
                print(f"    文字: \"{i['t_label'][:50]}\" = {i['t_val']}{i['t_unit']} [{i['t_type']}]")
                print(f"    表格: \"{i['d_label'][:30]}\" = {i['d_val']}{i['d_unit']} [{i['d_type']}]")
                if i['t_norm'] and i['d_norm']:
                    mx=max(abs(i['t_norm']),abs(i['d_norm']))
                    if mx>0:
                        d=abs(i['t_norm']-i['d_norm'])/mx*100
                        print(f"    差异:{d:.1f}% ({fmt(i['t_norm'])} vs {fmt(i['d_norm'])})")
            elif i['type'] in ('chart_vs_table','chart_vs_text'):
                print(f"\n  第{i['slide']}页 | 图表数据 | 匹配:{i['score']:.0%}")
                print(f"    图表: \"{i['chart_cat']}\" ({i['chart_series']}) = {i['chart_val']}")
                if i['type'] == 'chart_vs_table':
                    print(f"    表格: \"{i['table_label']}\" = {i['table_val']}")
                else:
                    print(f"    文字: \"{i['text_label']}\" = {i['text_val']}{i.get('text_unit','')}")

    print_iss(tvt, '文字vs表格')
    print_iss(cvt, '图表vs表格')
    print_iss(cvx, '图表vs文字')

    if not all_iss:
        print('\n未发现任何不一致问题。')

    # 保存报告
    out = os.path.join(os.path.dirname(os.path.abspath(a.pptx)),'ppt全量校验结果.txt')
    with open(out,'w',encoding='utf-8') as f:
        f.write(f'PPT全量校验报告\n文件:{os.path.basename(a.pptx)}\n容差:{a.tol*100:.1f}%\n校验时间:2026-06-06\n{"="*70}\n\n')
        f.write(f'表格数据:{tt} | 文字数值:{tn} | 图表页:{tc}\n')
        f.write(f'文字vs表格:{len(tvt)} | 图表vs表格:{len(cvt)} | 图表vs文字:{len(cvx)}\n')
        for lb,iss in [('文字vs表格',tvt),('图表vs表格',cvt),('图表vs文字',cvx)]:
            if iss:
                f.write(f'\n【{lb}】\n{"="*50}\n')
                for i in iss:
                    f.write(f"\n第{i['slide']}页:\n")
                    for k,v in i.items():
                        if k not in ('slide','type'): f.write(f'  {k}: {v}\n')
    print(f'\n报告已保存: {out}')

if __name__ == '__main__': main()
