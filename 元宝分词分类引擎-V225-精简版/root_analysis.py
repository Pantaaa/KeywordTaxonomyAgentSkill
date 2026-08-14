# -*- coding: utf-8 -*-
"""通用-其他 专项词根分析（V225 迭代用）
提取高频词根 + 典型样本 + 可规则化评估
"""
import sys, io, os, glob, re, csv
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass
import openpyxl
from collections import Counter

OUT_DIR = r'D:\DSH-分词工作区\输出'

def iter_oo():
    files = sorted(glob.glob(os.path.join(OUT_DIR, '*.xlsx')))
    for fp in files:
        wb = openpyxl.load_workbook(fp, read_only=True)
        for ws in wb.worksheets:
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    continue
                if not row or len(row) < 3 or not row[0] or not row[2]:
                    continue
                label = str(row[0])
                if label.startswith('通用-其他-其他-'):
                    yield str(row[2]).strip()
        wb.close()

def main():
    kws = list(iter_oo())
    print(f'通用-其他 关键词: {len(kws):,} 条')
    uniq = list(dict.fromkeys(kws))
    print(f'去重后: {len(uniq):,} 条')

    # 2-4字词根统计（每词取最长候选一次）
    root_counter = Counter()
    for kw in uniq:
        for n in (4, 3, 2):
            for i in range(len(kw) - n + 1):
                sub = kw[i:i+n]
                if all('\u4e00' <= ch <= '\u9fff' for ch in sub):
                    root_counter[sub] += 1
                    break

    print(f'\n=== 高频词根 Top 80（含典型样本）===')
    rows = []
    for sub, cnt in root_counter.most_common(2000):
        if cnt < 500:
            continue
        typical = next((k for k in uniq if sub in k), '')
        rows.append((sub, cnt, typical))
        if len(rows) >= 80:
            break

    print('| 词根 | 次数 | 典型样本 |')
    print('|------|------|----------|')
    for sub, cnt, typ in rows:
        print(f'| {sub} | {cnt:,} | {typ[:35]} |')

    # 保存到文件供下一步用
    with open(os.path.join(OUT_DIR, '_词根分析.txt'), 'w', encoding='utf-8') as f:
        f.write('| 词根 | 次数 | 典型样本 |\n')
        f.write('|------|------|----------|\n')
        for sub, cnt, typ in rows:
            f.write(f'| {sub} | {cnt:,} | {typ[:35]} |\n')
    print(f'\n已保存: 输出/_词根分析.txt')

if __name__ == '__main__':
    main()
