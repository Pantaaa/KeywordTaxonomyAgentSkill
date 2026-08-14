#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""元宝 V225 输出质量分析工具（增强版：全标签占比）

输出内容（用户要求）：
  1. 所有标签的占比（L1 完整表 + L2 全量表）
  2. 未明确分类的"其他"占比（重点分析与优化建议）
  3. 三维度报告（准确/速度/合规）+ 抽检 + 词根分析

用法:
  python analyze_output.py              # 全量分析输出目录
  python analyze_output.py --sample 50  # 指定抽检条数
  python analyze_output.py --top 50     # L2 表条数（默认 50）
输出: 输出/_质量报告_YYYY-MM-DD.md
"""
import sys, io, os, glob, re, random, time, argparse
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass
from collections import Counter
import openpyxl

OUT_DIR = r'D:\DSH-分词工作区\输出'
SKILL_DIR = r'D:\DSH-分词工作区\skills\元宝分词分类引擎-V225'
DATE = time.strftime('%Y-%m-%d')


def iter_results(files):
    """流式遍历输出文件: yield (一级标签, 二级标签, 关键词)"""
    for f in files:
        wb = openpyxl.load_workbook(f, read_only=True)
        for ws in wb.worksheets:
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    continue
                if row and len(row) >= 3 and row[2]:
                    yield str(row[0]), str(row[1]), str(row[2])
        wb.close()


def l1_of(label):
    if not label:
        return '未知'
    return '行业-AI' if label.startswith('行业-AI') else label.split('-')[0]


def l2_of(label):
    """从一级标签提取 L2（通用-L1-L2-DATE / 行业-AI-L2-DATE / 品牌-名-DATE / 竞品-名-DATE）"""
    if not label:
        return '?'
    parts = label.split('-')
    if len(parts) >= 4:  # 通用-xxx-yyy-date
        return parts[2]
    if len(parts) == 3:  # 行业-AI-yyy-date
        return parts[1]
    return parts[1] if len(parts) >= 2 else '?'


def analyze(sample_n=50, top_l2=50, seed=20260814):
    files = sorted(glob.glob(os.path.join(OUT_DIR, '*.xlsx')))
    if not files:
        print('❌ 输出目录无 xlsx')
        return
    print(f'读取 {len(files)} 个输出文件...')
    t0 = time.time()

    l1_cnt = Counter()
    l2_cnt = Counter()          # 全量 L2（含品牌/竞品名）
    l1l2_cnt = Counter()        # 完整 L1-L2 路径
    l1_l2_map = {}              # L1 -> Counter(L2)
    oo_keywords = []            # 通用-其他 关键词样本
    all_rows = []               # 抽检样本池
    total = filtered = 0

    for full_label, l2, kw in iter_results(files):
        total += 1
        l1 = l1_of(full_label)
        l2c = re.sub(r'-\d+$', '', l2)
        l1_cnt[l1] += 1
        l2_cnt[l2c] += 1
        l1l2_cnt[f'{l1}-{l2c}'] += 1
        l1_l2_map.setdefault(l1, Counter())[l2c] += 1
        if l1 == '通用' and l2c == '其他':
            oo_keywords.append(kw)
        if total % 5000 == 1:
            all_rows.append((full_label, l2, kw))

    dt = time.time() - t0
    speed = total / dt if dt else 0

    lines = []
    lines.append(f'# 元宝分词分类引擎 V225 — 质量报告（{DATE}）')
    lines.append('')
    lines.append('## 处理概况')
    lines.append(f'- 文件: {len(files)} 个 | 有效: {total:,} 条 | 过滤: {filtered:,}')
    lines.append(f'- 耗时: {dt:.0f}s（{speed:,.0f} q/s，含流式读取）')
    lines.append('')

    # ===== 全标签占比 =====
    lines.append('## 一、全标签占比')
    lines.append('')
    lines.append('### L1 一级标签占比（全部）')
    lines.append('')
    lines.append('| 一级标签 | 数量 | 占比 | 说明 |')
    lines.append('|----------|------|------|------|')
    for k, v in l1_cnt.most_common():
        pct = v / total * 100
        note = '🔴 含大量"其他"，需优化' if k == '通用' else ''
        lines.append(f'| {k} | {v:,} | {pct:.2f}% | {note} |')
    lines.append('')
    lines.append('### L1-L2 全路径占比（Top 50）')
    lines.append('')
    lines.append('| 路径 | 数量 | 占比 |')
    lines.append('|------|------|------|')
    for k, v in l1l2_cnt.most_common(top_l2):
        lines.append(f'| {k} | {v:,} | {v/total*100:.2f}% |')
    lines.append('')

    # ===== "其他"占比（重点） =====
    lines.append('## 二、"其他"标签占比（重点优化对象）')
    lines.append('')
    other_paths = {k: v for k, v in l1l2_cnt.items() if k.endswith('-其他')}
    other_total = sum(other_paths.values())
    lines.append(f'- **全部"其他"合计: {other_total:,} 条（{other_total/total*100:.2f}%）**')
    lines.append('')
    if other_paths:
        lines.append('| "其他"路径 | 数量 | 占全部比例 |')
        lines.append('|-----------|------|-----------|')
        for k, v in sorted(other_paths.items(), key=lambda x: -x[1]):
            lines.append(f'| {k} | {v:,} | {v/total*100:.2f}% |')
        lines.append('')
    oo = other_paths.get('通用-其他', 0)
    lines.append(f'### 通用-其他 深度分析')
    lines.append(f'- 通用-其他: {oo:,} 条（{oo/total*100:.2f}%）')
    lines.append('')

    # ===== 维度3 合规 =====
    lines.append('## 三、维度3 - 规则执行度')
    bad_fmt = sum(1 for l1, l2, kw in all_rows if not re.match(r'^(品牌|竞品|行业-AI|通用)-', str(l1)))
    lines.append(f'- 标签格式合规（抽检池）: {len(all_rows)-bad_fmt}/{len(all_rows)} ✅')
    lines.append(f'- 单元拆分: 每 5000 条/单元（引擎内置）')
    lines.append(f'- 违禁词/纯数字过滤: 引擎三层防御，本批过滤 {filtered:,} 条')
    lines.append('')

    # ===== 抽检 =====
    lines.append('## 四、抽检结果（人工复核样本）')
    random.seed(seed)
    sample = random.sample(all_rows, min(sample_n, len(all_rows)))
    lines.append(f'- 抽检数量: {len(sample)} 条（每 5000 条取 1 条构建样本池后随机）')
    lines.append('')
    lines.append('| # | 关键词 | 一级标签 | 二级标签 | 备注 |')
    lines.append('|---|--------|----------|----------|------|')
    for i, (l1, l2, kw) in enumerate(sample, 1):
        l2c = re.sub(r'-\d+$', '', l2)
        lines.append(f'| {i} | {kw} | {l1} | {l2c} | 待人工复核 |')
    lines.append('')

    # ===== 通用-其他词根分析 =====
    lines.append('## 五、通用-其他 可扩展词根分析')
    lines.append('')
    if oo_keywords:
        lines.append(f'- 通用-其他关键词: {len(oo_keywords):,} 条')
        lines.append('- 词根提炼规则: 2-4 字中文子串，出现 ≥100 次（迭代规范）')
        lines.append('')
        root_counter = Counter()
        for kw in oo_keywords:
            for n in (4, 3, 2):
                for i in range(len(kw) - n + 1):
                    sub = kw[i:i+n]
                    if all('\u4e00' <= ch <= '\u9fff' for ch in sub):
                        root_counter[sub] += 1
                        break
        lines.append('### 高频词根 Top 40（≥100次）')
        lines.append('')
        lines.append('| 词根 | 次数 | 典型样本 | 建议归类 |')
        lines.append('|------|------|----------|----------|')
        shown = 0
        for sub, cnt in root_counter.most_common(500):
            if cnt < 100:
                continue
            typical = next((k for k in oo_keywords if sub in k), '')
            lines.append(f'| {sub} | {cnt:,} | {typical[:30]} | 待定 |')
            shown += 1
            if shown >= 40:
                break
        lines.append('')
        # 可规则化评估
        sys.path.insert(0, SKILL_DIR)
        try:
            import auto_learn as al
            classified = sum(1 for k in oo_keywords if al.suggest_label(k) != '其他')
            lines.append(f'### 可规则化评估')
            lines.append(f'- 建议规则命中: {classified}/{len(oo_keywords):,} = {classified/len(oo_keywords)*100:.1f}%')
            lines.append('- （建议规则命中 → 可加入引擎规则，供下一次迭代决策）')
        except Exception as e:
            lines.append(f'- 可规则化评估失败: {e}')
    else:
        lines.append('- 通用-其他: 0 条 🎉')

    # 输出到 stdout（用户要求：占比报告在反馈中体现，不单独输出文件）
    report = '\n'.join(lines)
    print('<<<REPORT_BEGIN>>>')
    print(report)
    print('<<<REPORT_END>>>')
    print('=== 汇总 ===')
    print(f'有效: {total:,} | L1 标签数: {len(l1_cnt)} | L2 标签数: {len(l2_cnt)}')
    print(f'通用-其他: {oo:,}（{oo/total*100:.2f}%） | 全部其他: {other_total:,}（{other_total/total*100:.2f}%）')
    print(f'抽检: {len(sample)} 条 | 高频词根: {shown} 个')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--sample', type=int, default=50)
    ap.add_argument('--top', type=int, default=50)
    args = ap.parse_args()
    analyze(sample_n=args.sample, top_l2=args.top)
