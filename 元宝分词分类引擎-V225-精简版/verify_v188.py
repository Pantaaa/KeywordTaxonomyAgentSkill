"""V188 验证：图片创作优先 + 视频创作二级 + L1 不新增"""
import sys, io, importlib.util
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

spec = importlib.util.spec_from_file_location("engine", r"D:\元宝分词分类引擎-V225\classify_engine_v225.py")
mod = importlib.util.module_from_spec(spec)
sys.modules['engine'] = mod
spec.loader.exec_module(mod)

print('=== V188 验证 ===')
tests = [
    # 图片/视频 → 图片创作（用户指定）
    ('在线视频剪辑', '图片创作'),
    ('海报制作', '图片创作'),
    ('制图软件哪个好用', '图片创作'),
    ('在线图片处理软件工具', '图片创作'),
    ('照片文字自动生成视频', '图片创作'),
    ('视频剪辑软件', '图片创作'),
    ('图片设计', '图片创作'),
    ('修图工具', '图片创作'),
    # 视频创作二级
    ('在线视频剪辑', '视频创作'),
    ('视频制作', '视频创作'),
    ('照片生成视频', '视频创作'),
    # 图片设计/处理二级
    ('海报设计', '图片创作'),  # 含"海报"→图片处理二级，L1 图片创作 ✅
    ('平面设计', '图片设计'),
    # 工具保持（非图片/视频）
    ('翻译软件', '工具'),
    ('数据分析', '工具'),
    ('代码生成器', '工具'),
    ('配音 在线', '工具'),
    ('公众号排版', '工具'),
    # 回归
    ('智能马桶', '商品'),
    ('nba赛程', '体育运动'),
    ('腾讯元宝', '品牌'),
    ('高考作文ai写作范文', '行业-AI'),
    ('成都', '出行'),
    ('135', None),
    ('ppt模板下载', None),
    ('色情直播', None),
]

passed = failed = 0
for kw, exp in tests:
    r = mod.classify_keyword(kw)
    if exp is None:
        if r is None:
            passed += 1
            print(f'  ✅ {kw:22s} → 被过滤')
        else:
            failed += 1
            print(f'  ❌ {kw:22s} → 应过滤, 得到 {r["一级标签"]}')
        continue
    l1 = r['一级标签'] if r else 'FILTERED'
    l2 = r['二级标签'] if r else ''
    ok = (r is not None and exp in l1)
    if ok:
        passed += 1
        print(f'  ✅ {kw:22s} → {l1} | {l2}')
    else:
        failed += 1
        print(f'  ❌ {kw:22s} → {l1} | {l2} （期望一级含: {exp}）')

print(f'\n通过: {passed}/{passed+failed} | 失败: {failed}')
