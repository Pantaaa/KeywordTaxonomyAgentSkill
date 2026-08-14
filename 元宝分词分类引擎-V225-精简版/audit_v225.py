"""V225：人工抽查 30 条（第十轮）"""
import sys, io, os, random
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

in_dir = r'D:\分词工具处理2.5\输出\20260807元宝06合并分类'
random.seed(6543)

def collect_oo(max_n=80000):
    samples = []
    for fn in [f'其他_其他_26-08-14_part{i}.xlsx' for i in range(1, 3)]:
        p = os.path.join(in_dir, fn)
        if not os.path.exists(p):
            continue
        wb = openpyxl.load_workbook(p, read_only=True)
        ws = wb.active
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            if row and len(row) >= 3 and row[2] and row[2] != '关键词':
                samples.append(str(row[2]).strip())
                if len(samples) >= max_n:
                    break
        wb.close()
        if len(samples) >= max_n:
            break
    return samples

print('===== 通用-其他 人工抽查 30 条（第十轮） =====')
oo = collect_oo()
for kw in random.sample(oo, 30):
    print(f'    {kw}')
