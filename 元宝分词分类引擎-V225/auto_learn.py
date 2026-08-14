"""元宝引擎 自动规则学习工具（强化学习能力，V199）
从"其他"标签输出自动提炼高频词根 + 候选归类建议，加速迭代
用法:
  python auto_learn.py --input "其他.xlsx" [--top 50] [--min-cnt 100] [--col 3]
输出: 高频词根 + 典型样本 + 候选归类建议（供人工审查/自动写入引擎）
"""
import sys, io, os, re, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from collections import Counter
import openpyxl

def is_chinese(s):
    return all('\u4e00' <= ch <= '\u9fff' for ch in s)

def load_keywords(fpath, col=3):
    kws = []
    wb = openpyxl.load_workbook(fpath, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            if not row or len(row) < col or not row[col-1]:
                continue
            kw = str(row[col-1]).strip()
            if kw and kw != '关键词':
                kws.append(kw)
    wb.close()
    return kws

# 候选归类（词根命中 → 建议标签）
SUGGEST_RULES = [
    # (正则, 标签)
    (r'(教程|入门|学习|课程|培训)', '教育-课程培训'),
    (r'(价格|多少钱|报价|购买|哪里买|优惠|批发|零售)', '商品'),
    (r'(软件|下载|安装|在线|工具|编辑器|生成器)', '工具'),
    (r'(作文|范文|文案|报告|总结|方案|策划)', '文本创作'),
    (r'(旅游|景点|酒店|机票|攻略|出行)', '出行'),
    (r'(医院|疾病|治疗|药品|健康|体检)', '医疗健康'),
    (r'(足球|篮球|健身|运动|赛事|比赛)', '体育运动'),
    (r'(股票|基金|理财|保险|贷款|银行)', '金融'),
    (r'(房产|楼盘|房价|租房|买房|装修)', '房产'),
    (r'(招聘|求职|简历|面试|岗位)', '社交-招聘求职'),
    (r'(律师|法律|合同|诉讼|法规)', '政务-法律'),
    (r'(婚礼|婚纱|婚庆|结婚|喜糖)', '商品-婚庆用品'),
    (r'(美食|菜谱|餐厅|做饭|烹饪)', '美食餐饮'),
    (r'(化妆|护肤|口红|面膜|美容)', '商品-美妆护肤'),
    (r'(衣服|穿搭|服饰|时尚|鞋子)', '商品-服饰鞋包'),
    (r'(家具|家居|家电|收纳|清洁)', '商品-家居用品'),
    (r'(汽车|车型|二手车|驾驶|油耗)', '商品-汽车用品'),
    (r'(AI|人工智能|机器人|大模型|智能)', '行业-AI'),
]

def suggest_label(kw):
    for pat, label in SUGGEST_RULES:
        if re.search(pat, kw, re.I):
            return label
    return '其他'

def main():
    ap = argparse.ArgumentParser(description='自动规则学习')
    ap.add_argument('--input', required=True, help='"其他"标签输出 xlsx')
    ap.add_argument('--top', type=int, default=50)
    ap.add_argument('--min-cnt', type=int, default=100)
    ap.add_argument('--col', type=int, default=3)
    args = ap.parse_args()

    kws = load_keywords(args.input, args.col)
    print(f'加载关键词: {len(kws)} 条')

    # 高频词根
    root_counter = Counter()
    for kw in kws:
        for n in (3, 2):
            for i in range(len(kw) - n + 1):
                sub = kw[i:i+n]
                if is_chinese(sub):
                    root_counter[sub] += 1

    print(f'\n=== 高频词根 + 候选归类建议（Top {args.top}）===')
    shown = 0
    for sub, cnt in root_counter.most_common(500):
        if cnt < args.min_cnt:
            continue
        typical = [kw for kw in kws if sub in kw][:3]
        # 建议归类（取典型样本的多数建议）
        labels = Counter(suggest_label(k) for k in [kw for kw in kws if sub in kw][:10])
        best_label = labels.most_common(1)[0][0] if labels else '其他'
        print(f'\n  [{cnt}] {sub} → 建议: {best_label}')
        for t in typical:
            print(f'      {t}')
        shown += 1
        if shown >= args.top:
            break

    # 可归类的统计（建议非"其他"的比例）
    classified = 0
    for kw in kws:
        if suggest_label(kw) != '其他':
            classified += 1
    print(f'\n=== 可规则化评估 ===')
    print(f'样本中建议可归类: {classified}/{len(kws)} = {classified/len(kws)*100:.1f}%')
    print('（建议规则命中 → 可加入引擎 _detect_l1 / 二级字典）')

if __name__ == '__main__':
    main()
