"""V187b BGE 后处理（分块版）：每 part 一个批次，独立输出 CSV
用法: python bge_chunk.py <part编号> <阈值>
"""
import sys, io, os, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import numpy as np
from collections import Counter
import csv, openpyxl

os.environ['HF_ENDPOINT'] = 'https://hf-mirror.com'
# V191: 效率优化 - 用满 CPU 多线程
import torch
torch.set_num_threads(8)
from sentence_transformers import SentenceTransformer

SEED_SENTENCES = {
    '工具': ['软件工具下载', '在线编辑器', '文件格式转换工具', '网盘云存储', '制作生成工具',
            '办公软件', '在线工具网站', '免费工具软件', '工具app'],
    '文本创作': ['写文案', '工作总结报告', '演讲稿致辞', '范文模板', '文章写作',
            '广告语标语', '通知公告', '祝福贺词', '策划方案'],
    '教育': ['课程学习', '教育培训', '考试作业', '教学课件', '学校招生',
            '考研升学', '在线课程', '知识学习', '学习方法'],
    '图片创作': ['图片处理', '海报设计', '视频剪辑', '摄影修图', '绘画素材',
            '壁纸头像', '动画制作', '平面设计', '图片编辑'],
    '行业AI': ['人工智能', '大模型', '机器人', '智能客服', '数字人',
            'AI工具', '机器学习', '智能应用'],
    '出行': ['旅游攻略', '景点门票', '酒店预订', '机票火车票', '城市旅行',
            '自驾游', '民宿', '旅行路线'],
    '金融': ['理财投资', '保险', '基金股票', '银行贷款', '财务税务',
            '证券', '经济金融'],
    '医疗健康': ['医院看病', '疾病治疗', '药品', '体检', '养生保健',
            '医生咨询', '健康知识'],
    '娱乐': ['游戏', '电影', '音乐', '综艺', '明星', '动漫', '小说', '直播'],
    '体育运动': ['足球', '篮球', '健身', '跑步', '比赛', '赛事', '运动训练',
            '体育新闻', '运动员'],
    '房产': ['楼盘', '房价', '租房', '买房', '装修', '物业', '二手房', '户型'],
    '招聘求职': ['招聘', '求职', '简历', '面试', '岗位', '薪资', '入职', '人才招聘'],
    '法律': ['法律咨询', '律师', '合同', '诉讼', '维权', '法规'],
    '美食': ['菜谱', '食谱', '餐厅', '做饭', '烹饪', '美食推荐'],
    '美妆': ['化妆', '护肤', '口红', '面膜', '美容', '彩妆'],
    '服装': ['衣服', '穿搭', '服饰', '时尚', '鞋子', '包包'],
    '家居': ['家居', '家具', '装修', '收纳', '清洁', '家居用品'],
    '汽车': ['汽车', '车型', '买车', '二手车', '汽车保养', '汽车配件'],
    '婚庆': ['婚礼', '婚纱', '婚庆', '结婚', '婚宴', '订婚'],
}
LABEL_MAP = {
    '工具': '通用-工具-其他', '文本创作': '通用-文本创作-其他', '教育': '通用-教育-其他',
    '图片创作': '通用-图片创作-其他', '行业AI': '行业-AI-AI应用', '出行': '通用-出行-其他',
    '金融': '通用-金融-其他', '医疗健康': '通用-医疗健康-其他', '娱乐': '通用-娱乐-其他',
    '体育运动': '通用-体育运动-其他', '房产': '通用-房产-其他',
    # V188: L1 不新增，法律→政务-L2，婚庆→商品-L2
    # V198: 招聘求职 → 社交-L2（用户指定 2026-08-08：招聘求职应归社交标签）
    '招聘求职': '通用-社交-招聘求职', '法律': '通用-政务-法律', '婚庆': '通用-商品-婚庆用品',
    '美食': '通用-美食餐饮-其他', '美妆': '通用-商品-美妆护肤', '服装': '通用-商品-服饰鞋包',
    '家居': '通用-商品-家居用品', '汽车': '通用-商品-汽车用品',
}

def main():
    part = int(sys.argv[1])
    threshold = float(sys.argv[2]) if len(sys.argv) > 2 else 0.65
    # V197: 参数化输入（默认 20260803 全量，可指定 20260807 元宝06）
    in_dir = sys.argv[3] if len(sys.argv) > 3 else r'D:\分词工具处理2.5\输出\20260803后合并分类'
    prefix = sys.argv[4] if len(sys.argv) > 4 else '通用_其他_其他_part'
    out_dir = sys.argv[5] if len(sys.argv) > 5 else os.path.join(in_dir, '语义细分')
    os.makedirs(out_dir, exist_ok=True)

    fpath = os.path.join(in_dir, f'{prefix}{part}.xlsx')
    print(f'[part{part}] 读取 {fpath}')
    t0 = time.time()
    kws = []
    wb = openpyxl.load_workbook(fpath, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            if not row or len(row) < 3 or not row[2]:
                continue
            kws.append(str(row[2]).strip())
    wb.close()
    print(f'[part{part}] 读取 {len(kws)} 条, {time.time()-t0:.0f}s')

    uniq = list(dict.fromkeys(kws))
    print(f'[part{part}] 去重后 {len(uniq)} 条')

    print(f'[part{part}] 加载 BGE...')
    model = SentenceTransformer(r'D:\元宝分词分类引擎-V225\bge_model')
    label_names = list(SEED_SENTENCES.keys())
    centers = {}
    for label, sents in SEED_SENTENCES.items():
        vecs = model.encode(sents, normalize_embeddings=True)
        centers[label] = np.mean(vecs, axis=0)
    center_matrix = np.vstack([centers[l] for l in label_names])
    center_matrix = center_matrix / (np.linalg.norm(center_matrix, axis=1, keepdims=True) + 1e-9)

    print(f'[part{part}] BGE 编码 {len(uniq)} 条...')
    t0 = time.time()
    vecs = model.encode(uniq, normalize_embeddings=True, batch_size=1024)
    scores = vecs @ center_matrix.T
    print(f'[part{part}] 编码耗时 {time.time()-t0:.0f}s')

    # 分类 + 按标签收集
    by_label = {}
    remain = []
    for i, kw in enumerate(uniq):
        best_idx = int(np.argmax(scores[i]))
        best = float(scores[i][best_idx])
        if best >= threshold:
            label = LABEL_MAP[label_names[best_idx]]
            by_label.setdefault(label, []).append(kw)
        else:
            remain.append(kw)

    # 写出本 part 结果
    for label, kws_list in by_label.items():
        safe = label.replace('通用-', '').replace('行业-AI', 'AI').replace('-', '_')
        outpath = os.path.join(out_dir, f'part{part}_{safe}.csv')
        with open(outpath, 'w', encoding='utf-8', newline='') as f:
            w = csv.writer(f)
            for kw in kws_list:
                w.writerow([label, kw])
    # 剩余（保持其他）
    with open(os.path.join(out_dir, f'part{part}_remaining.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f)
        for kw in remain:
            w.writerow([kw])

    cnt = {k: len(v) for k, v in by_label.items()}
    print(f'[part{part}] 已细分 {len(uniq)-len(remain)} ({sum(cnt.values())}), 保持其他 {len(remain)}')
    print(f'[part{part}] 标签分布: {cnt}')
    print(f'[part{part}] 完成, 总耗时 {time.time()-t0:.0f}s')

if __name__ == '__main__':
    main()
