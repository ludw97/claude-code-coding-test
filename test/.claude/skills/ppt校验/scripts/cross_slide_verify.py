#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT 跨页数据一致性校验（最终版）
关键改进：使用表格第一列标题作为指标主题上下文
"""

import re, os, argparse, sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ppt_verify import extract_ppt, parse_num, nl, fmt

# ==================== 表格上下文感知 ====================

def get_table_topic(table):
    """获取表格的主题（第一列标题通常指示度量内容）"""
    headers = table.get('headers', [])
    if headers and len(headers) >= 2:
        h1 = headers[1]  # 第二列标题（第一列是行标签列）
        if h1:
            # 去除日期
            h1_clean = re.sub(r'\d{4}年\d{1,2}[-]\d{1,2}月', '', h1)
            h1_clean = re.sub(r'\d{4}年\d{1,2}月', '', h1_clean)
            h1_clean = re.sub(r'\d{4}年', '', h1_clean)
            h1_clean = h1_clean.strip()
            if h1_clean:
                return h1_clean
    return ''


def make_context_key(table_topic, row_label, col_header):
    """生成包含表格主题上下文的键"""
    parts = []

    # 表格主题
    if table_topic:
        topic_nl = nl(table_topic)
        if topic_nl and len(topic_nl) >= 2:
            parts.append(topic_nl)

    # 行标签（最小清洗）
    if row_label:
        rl = re.sub(r'[\s　]+', '', row_label)
        rl = re.sub(r'[：:，,、。；;]+$', '', rl)
        rl = rl.strip()
        if rl and len(rl) >= 2:
            parts.append(rl)

    # 日期
    if col_header:
        dates = []
        for m in re.finditer(r'(202[5-6])年(\d{1,2})月', col_header):
            dates.append(f'{m.group(1)}年{m.group(2)}月')
        for m in re.finditer(r'(202[5-6])年(\d{1,2})[-](\d{1,2})月', col_header):
            dates.append(f'{m.group(1)}年{m.group(2)}-{m.group(3)}月')
        if dates:
            parts.append(dates[0])

    return '|'.join(parts)


def main():
    p = argparse.ArgumentParser(description='PPT跨页一致性（表格上下文感知版）')
    p.add_argument('pptx')
    p.add_argument('--tol', type=float, default=0.01)
    a = p.parse_args()

    print(f'PPT跨页一致性校验（上下文感知版）')
    print(f'文件: {os.path.basename(a.pptx)}')
    print(f'差异阈值: {a.tol*100:.1f}%')
    print('='*70)

    slide_data = extract_ppt(a.pptx)
    all_metrics = []

    # 收集：只从表格收集（表格有明确的主题上下文）
    for sn in sorted(slide_data):
        sd = slide_data[sn]
        for table in sd['tables']:
            topic = get_table_topic(table)
            for dp in table['points']:
                rl = dp.get('row_label', '')
                ch = dp.get('col_header', '')
                ckey = make_context_key(topic, rl, ch)
                if not ckey:
                    continue

                all_metrics.append({
                    'slide': sn, 'context_key': ckey,
                    'label': f'[{topic}] {rl} ({ch})',
                    'norm': dp['norm'], 'unit': dp.get('unit', ''),
                    'raw_text': dp['raw_text'], 'dtype': dp.get('col_type', 'absolute'),
                })

    print(f'收集表格指标（含上下文）: {len(all_metrics)}')

    # 按上下文键分组
    groups = defaultdict(list)
    for m in all_metrics:
        groups[m['context_key']].append(m)

    # 过滤出跨页的组
    multi_slide = {}
    for k, v in groups.items():
        slides = {m['slide'] for m in v}
        if len(slides) >= 2:
            multi_slide[k] = v

    print(f'跨页出现的上下文指标组: {len(multi_slide)}')

    errors = []
    warnings = []

    for ckey, metrics in multi_slide.items():
        by_slide = defaultdict(list)
        for m in metrics:
            by_slide[m['slide']].append(m)

        slides = sorted(by_slide.keys())

        for i in range(len(slides)):
            for j in range(i+1, len(slides)):
                s1, s2 = slides[i], slides[j]
                for m1 in by_slide[s1]:
                    for m2 in by_slide[s2]:
                        n1, n2 = m1['norm'], m2['norm']
                        if n1 is None or n2 is None:
                            continue
                        if m1['dtype'] != m2['dtype']:
                            continue
                        if n1 == 0 and n2 == 0:
                            continue
                        if n1 == 0 or n2 == 0:
                            continue
                        ratio = max(abs(n1), abs(n2)) / min(abs(n1), abs(n2))
                        if ratio > 10:
                            continue

                        mx = max(abs(n1), abs(n2))
                        diff_pct = abs(n1 - n2) / mx
                        if diff_pct <= 0.002:
                            continue

                        u1, u2 = m1.get('unit', ''), m2.get('unit', '')
                        unit_diff = bool(u1 and u2 and u1 != u2)

                        item = {
                            'context_key': ckey,
                            'slide1': s1, 'slide2': s2,
                            'label1': m1['label'], 'label2': m2['label'],
                            'norm1': n1, 'norm2': n2,
                            'unit1': u1, 'unit2': u2,
                            'raw1': m1['raw_text'], 'raw2': m2['raw_text'],
                            'diff_pct': diff_pct, 'unit_diff': unit_diff,
                        }

                        if diff_pct <= 0.01:
                            warnings.append(('尾差/进位退位', item))
                        elif unit_diff and diff_pct <= 0.02:
                            warnings.append(('单位差异', item))
                        elif diff_pct <= 0.015:
                            warnings.append(('可能尾差', item))
                        else:
                            errors.append(item)

    # 去重
    seen = set()
    dedup_errors = []
    for e in errors:
        sig = (min(e['slide1'], e['slide2']), max(e['slide1'], e['slide2']),
               e['context_key'], round(e['diff_pct'], 4))
        if sig not in seen:
            seen.add(sig)
            dedup_errors.append(e)

    dedup_warnings = []
    seen_w = set()
    for wt, w in warnings:
        sig = (min(w['slide1'], w['slide2']), max(w['slide1'], w['slide2']),
               w['context_key'], round(w['diff_pct'], 4))
        if sig not in seen_w:
            seen_w.add(sig)
            dedup_warnings.append((wt, w))

    # 只保留真正有问题的
    real_errors = [e for e in dedup_errors if e['diff_pct'] > 0.01]
    real_warnings = [(wt, w) for wt, w in dedup_warnings if w['diff_pct'] > 0.002]

    print(f'\n校验结果:')
    print(f'  错误（差异>1%）: {len(real_errors)} 个')
    print(f'  提示（尾差/单位差异）: {len(real_warnings)} 个')
    print('='*70)

    if real_errors:
        print(f'\n{"="*60}')
        print(f'【跨页数据不一致 - 错误】')
        print(f'{"="*60}')

        # 按 context_key 分组输出
        by_key = defaultdict(list)
        for e in real_errors:
            by_key[e['context_key']].append(e)

        idx = 0
        for ckey in sorted(by_key):
            items = by_key[ckey]
            print(f'\n  --- {ckey} ({len(items)}处不一致) ---')
            for e in items[:3]:  # 每组最多显示3个
                idx += 1
                print(f'  #{idx} 差异:{e["diff_pct"]*100:.1f}%')
                print(f'    第{e["slide1"]}页: {e["label1"][:80]}')
                print(f'      值: {e["raw1"]}{e["unit1"]} ({fmt(e["norm1"])})')
                print(f'    第{e["slide2"]}页: {e["label2"][:80]}')
                print(f'      值: {e["raw2"]}{e["unit2"]} ({fmt(e["norm2"])})')
            if len(items) > 3:
                print(f'    ... 还有 {len(items)-3} 处')

    if real_warnings:
        print(f'\n{"="*60}')
        print(f'【跨页数据差异 - 提示（尾差/进位/单位差异，不作为错误）】')
        print(f'{"="*60}')
        for idx, (wtype, w) in enumerate(real_warnings[:30], 1):
            print(f'\n  #{idx} [{wtype}] | {w["context_key"]}')
            print(f'    第{w["slide1"]}页: {w["label1"][:70]}')
            print(f'      值: {w["raw1"]}{w["unit1"]} ({fmt(w["norm1"])})')
            print(f'    第{w["slide2"]}页: {w["label2"][:70]}')
            print(f'      值: {w["raw2"]}{w["unit2"]} ({fmt(w["norm2"])})')
            print(f'    差异: {w["diff_pct"]*100:.2f}%')

    if not real_errors and not real_warnings:
        print('\n未发现跨页数据不一致问题。')

    # 保存报告
    out = os.path.join(os.path.dirname(os.path.abspath(a.pptx)), '跨页一致性校验结果.txt')
    with open(out, 'w', encoding='utf-8') as f:
        f.write(f'PPT跨页一致性校验报告（上下文感知版）\n')
        f.write(f'文件: {os.path.basename(a.pptx)}\n阈值: {a.tol*100:.1f}%\n校验时间: 2026-07-05\n')
        f.write(f'{"="*70}\n\n')
        f.write(f'收集表格指标（含上下文）: {len(all_metrics)}\n')
        f.write(f'跨页出现的上下文指标组: {len(multi_slide)}\n')
        f.write(f'错误（差异>1%）: {len(real_errors)} 个\n')
        f.write(f'提示（尾差/单位差异）: {len(real_warnings)} 个\n')

        if real_errors:
            f.write(f'\n{"="*60}\n【跨页数据不一致 - 错误（按指标分组）】\n{"="*60}\n')
            by_key = defaultdict(list)
            for e in real_errors:
                by_key[e['context_key']].append(e)
            for ckey in sorted(by_key):
                f.write(f'\n--- {ckey} ---\n')
                for e in by_key[ckey]:
                    f.write(f'  第{e["slide1"]}页: {e["label1"]}\n')
                    f.write(f'    值: {e["raw1"]}{e["unit1"]} ({fmt(e["norm1"])})\n')
                    f.write(f'  第{e["slide2"]}页: {e["label2"]}\n')
                    f.write(f'    值: {e["raw2"]}{e["unit2"]} ({fmt(e["norm2"])})\n')
                    f.write(f'  差异: {e["diff_pct"]*100:.2f}%\n')

        if real_warnings:
            f.write(f'\n{"="*60}\n【跨页数据差异 - 提示】\n{"="*60}\n')
            for idx, (wtype, w) in enumerate(real_warnings, 1):
                f.write(f'\n#{idx} [{wtype}] | {w["context_key"]}\n')
                f.write(f'  第{w["slide1"]}页: {w["label1"]}\n')
                f.write(f'    值: {w["raw1"]}{w["unit1"]} ({fmt(w["norm1"])})\n')
                f.write(f'  第{w["slide2"]}页: {w["label2"]}\n')
                f.write(f'    值: {w["raw2"]}{w["unit2"]} ({fmt(w["norm2"])})\n')
                f.write(f'  差异: {w["diff_pct"]*100:.2f}%\n')

    print(f'\n报告已保存: {out}')


if __name__ == '__main__':
    main()
