#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用-其他 → BGE 语义细分整合流程（V225 工作区版）

步骤：
  1. 从输出目录提取所有"通用-其他-其他"关键词 → 中间 CSV
  2. 调用 BGE 语义细分（阈值 0.65）→ 得到细分 CSV + 剩余
  3. 输出统计（细分/剩余 + 各标签分布）

用法:
  python bge_rerun.py [--threshold 0.65] [--sample 50000] [--dry-run]
  --dry-run 只统计不跑模型；--sample 限制样本量（验证用）
"""
import sys, io, os, glob, csv, time, argparse
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass
import openpyxl

OUT_DIR = r'D:\DSH-分词工作区\输出'
TMP_OO = os.path.join(OUT_DIR, '_通用其他_候选.csv')
TMP_OUT = os.path.join(OUT_DIR, '_bge_细分')
os.makedirs(TMP_OUT, exist_ok=True)


def find_bge_model():
    """自动探测 BGE 模型路径（精简版 skill 不含模型，从外部引用）"""
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bge_model'),
        r'D:\元宝分词分类引擎-V225\bge_model',
        os.environ.get('BGE_MODEL_DIR', ''),
    ]
    for p in candidates:
        if p and os.path.exists(os.path.join(p, 'model.safetensors')):
            return p
    raise FileNotFoundError(
        '未找到 BGE 模型（model.safetensors）。请设置环境变量 BGE_MODEL_DIR 指向模型目录，'
        '或使用 D:\\元宝分词分类引擎-V225\\bge_model')


def extract_oo(sample=None):
    """从输出 xlsx 提取通用-其他关键词 → CSV"""
    files = sorted(glob.glob(os.path.join(OUT_DIR, '*.xlsx')))
    n = 0
    with open(TMP_OO, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['关键词'])
        for fp in files:
            wb = openpyxl.load_workbook(fp, read_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    if not row or len(row) < 3 or not row[0]:
                        continue
                    label = str(row[0])   # 如 通用-其他-其他-26/08/14
                    l2 = str(row[1]) if row[1] else ''
                    # 通用-其他: 一级标签 = 通用-其他-其他-日期 且 L2 以"其他"开头
                    if label.startswith('通用-其他-其他-') and l2.startswith('其他'):
                        w.writerow([str(row[2]).strip()])
                        n += 1
                        if sample and n >= sample:
                            wb.close()
                            print(f'提取通用-其他: {n} 条（sample 上限）')
                            return n
            wb.close()
    print(f'提取通用-其他: {n} 条')
    return n


def bge_classify(threshold=0.65):
    """BGE 语义细分"""
    import numpy as np
    import torch
    torch.set_num_threads(8)
    from sentence_transformers import SentenceTransformer

    # 复用 bge_chunk 的种子句与标签映射
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import bge_chunk as bc

    kws = []
    with open(TMP_OO, encoding='utf-8-sig') as f:
        r = csv.reader(f)
        next(r, None)
        for row in r:
            if row and row[0].strip():
                kws.append(row[0].strip())
    uniq = list(dict.fromkeys(kws))
    print(f'加载 {len(kws)} 条，去重后 {len(uniq)} 条')

    print('加载 BGE 模型...')
    t0 = time.time()
    model_path = find_bge_model()
    print(f'模型路径: {model_path}')
    model = SentenceTransformer(model_path)
    label_names = list(bc.SEED_SENTENCES.keys())
    centers = {}
    for label, sents in bc.SEED_SENTENCES.items():
        vecs = model.encode(sents, normalize_embeddings=True)
        centers[label] = np.mean(vecs, axis=0)
    center_matrix = np.vstack([centers[l] for l in label_names])
    center_matrix = center_matrix / (np.linalg.norm(center_matrix, axis=1, keepdims=True) + 1e-9)

    print(f'BGE 编码 {len(uniq)} 条...')
    vecs = model.encode(uniq, normalize_embeddings=True, batch_size=1024)
    scores = vecs @ center_matrix.T
    print(f'编码耗时 {time.time()-t0:.0f}s')

    by_label = {}
    remain = []
    for i, kw in enumerate(uniq):
        best_idx = int(np.argmax(scores[i]))
        best = float(scores[i][best_idx])
        if best >= threshold:
            label = bc.LABEL_MAP[label_names[best_idx]]
            by_label.setdefault(label, []).append(kw)
        else:
            remain.append(kw)

    # 写结果
    for label, kws_list in by_label.items():
        safe = label.replace('通用-', '').replace('行业-AI', 'AI').replace('-', '_')
        outpath = os.path.join(TMP_OUT, f'{safe}.csv')
        with open(outpath, 'w', encoding='utf-8', newline='') as f:
            w = csv.writer(f)
            w.writerow(['标签', '关键词'])
            for kw in kws_list:
                w.writerow([label, kw])
    with open(os.path.join(TMP_OUT, 'remaining.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f)
        w.writerow(['关键词'])
        for kw in remain:
            w.writerow([kw])

    total = len(uniq)
    classified = total - len(remain)
    print(f'\n=== BGE 语义细分结果（阈值 {threshold}）===')
    print(f'候选: {total:,} | 已细分: {classified:,}（{classified/total*100:.1f}%） | 保持其他: {len(remain):,}（{len(remain)/total*100:.1f}%）')
    print('\n细分标签分布:')
    for label, kws_list in sorted(by_label.items(), key=lambda x: -len(x[1])):
        print(f'  {label:<24} {len(kws_list):>7,}')
    return {'total': total, 'classified': classified, 'remain': len(remain), 'by_label': {k: len(v) for k, v in by_label.items()}}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--threshold', type=float, default=0.65)
    ap.add_argument('--sample', type=int, default=None)
    ap.add_argument('--dry-run', action='store_true', help='只统计候选，不跑模型')
    args = ap.parse_args()

    n = extract_oo(sample=args.sample)
    if args.dry_run:
        print(f'候选通用-其他: {n:,} 条（dry-run，不跑 BGE）')
    else:
        bge_classify(threshold=args.threshold)
