"""技能版本升级工具（单版本替换模式）- 元宝分词分类引擎迭代自动更新

⚠️ 重要：必须从外部目录运行（如 D:\），避免 Windows 占用当前技能目录无法删除：

  cd /d D:\\
  python "D:\\元宝分词分类引擎-V225\\upgrade_skill.py" --old V169 --new V170 --ws

功能（避免堆积历史版本，升级后自动替换）：
  1. 复制 D:\元宝分词分类引擎-{OLD} → D:\元宝分词分类引擎-{NEW}
  2. 重命名 classify_engine_v{old}.py → classify_engine_v{new}.py
  3. 更新所有 .py/.md 文件中的技能版本标识（保留 V167/V168 等历史修复引用）
  4. 同步更新工作区技能 skills/元宝分词分类引擎-{NEW}/
  5. 自动记录版本变更到 VERSIONING.md
  6. 【替换】升级成功后删除旧版本目录（D盘 + 工作区），始终只保留 1 个版本
"""
import sys, io, os, shutil, argparse, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

D_ROOT = 'D:\\'
WS_SKILLS = r'D:\DSH-分词工作区\skills'

def parse_args():
    p = argparse.ArgumentParser(description='元宝分词分类引擎技能版本升级（单版本替换模式）')
    p.add_argument('--old', required=True, help='旧版本号，如 V169')
    p.add_argument('--new', required=True, help='新版本号，如 V170')
    p.add_argument('--ws', action='store_true', help='同时更新工作区技能目录')
    p.add_argument('--keep-old', action='store_true', help='保留旧版本目录（默认删除）')
    return p.parse_args()

def update_version_marks(content, old_ver, new_ver):
    old_num = old_ver.lower().lstrip('v')
    new_num = new_ver.lower().lstrip('v')
    content = content.replace(f'元宝分词分类引擎-{old_ver}', f'元宝分词分类引擎-{new_ver}')
    content = content.replace(f'classify_engine_v{old_num}', f'classify_engine_v{new_num}')
    content = content.replace(f'元宝分词分类引擎 {old_ver} 独立版', f'元宝分词分类引擎 {new_ver} 独立版')
    content = content.replace(f'# 元宝分词分类引擎 {old_ver} — 独立版', f'# 元宝分词分类引擎 {new_ver} — 独立版')
    content = content.replace(f'# 元宝分词分类引擎 {old_ver} — 规则与引擎完全集成', f'# 元宝分词分类引擎 {new_ver} — 规则与引擎完全集成')
    content = content.replace(f'name: 元宝分词分类引擎-{old_ver}', f'name: 元宝分词分类引擎-{new_ver}')
    content = content.replace(f'name: 元宝分词分类引擎-{old_ver}-独立版', f'name: 元宝分词分类引擎-{new_ver}-独立版')
    content = content.replace(f'**当前版本**：**{old_ver}**', f'**当前版本**：**{new_ver}**')
    content = content.replace(f'实现 {old_ver} 规则', f'实现 {new_ver} 规则')
    return content

def main():
    args = parse_args()
    old_ver, new_ver = args.old, args.new
    src = os.path.join(D_ROOT, f'元宝分词分类引擎-{old_ver}')
    dst = os.path.join(D_ROOT, f'元宝分词分类引擎-{new_ver}')
    if not os.path.exists(src):
        print(f'❌ 源目录不存在: {src}')
        return

    # 1. 复制
    print(f'1. 复制 {src} → {dst}')
    if os.path.exists(dst):
        shutil.rmtree(dst)
    shutil.copytree(src, dst, ignore=shutil.ignore_patterns('__pycache__', '*.txt'))
    print('   完成')

    # 2. 重命名引擎
    old_num, new_num = old_ver.lower().lstrip('v'), new_ver.lower().lstrip('v')
    old_engine = os.path.join(dst, f'classify_engine_v{old_num}.py')
    new_engine = os.path.join(dst, f'classify_engine_v{new_num}.py')
    if os.path.exists(old_engine):
        os.rename(old_engine, new_engine)
        print(f'2. 引擎重命名: classify_engine_v{old_num}.py → classify_engine_v{new_num}.py')

    # 3. 更新版本标识
    print(f'3. 更新版本标识 {old_ver} → {new_ver}')
    for fname in os.listdir(dst):
        if not fname.endswith(('.py', '.md')):
            continue
        fpath = os.path.join(dst, fname)
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception:
            continue
        new_content = update_version_marks(content, old_ver, new_ver)
        if new_content != content:
            with open(fpath, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print(f'   更新: {fname}')

    # 4. 同步工作区
    if args.ws:
        ws_src = os.path.join(WS_SKILLS, f'元宝分词分类引擎-{old_ver}')
        ws_dst = os.path.join(WS_SKILLS, f'元宝分词分类引擎-{new_ver}')
        if os.path.exists(ws_src):
            print(f'4. 同步工作区: {ws_src} → {ws_dst}')
            if os.path.exists(ws_dst):
                shutil.rmtree(ws_dst)
            shutil.copytree(ws_src, ws_dst, ignore=shutil.ignore_patterns('__pycache__'))
            # 工作区引擎文件重命名（与 D 盘步骤 2 对齐，V225 修复：此前工作区引擎名未同步）
            ws_old_engine = os.path.join(ws_dst, f'classify_engine_v{old_num}.py')
            ws_new_engine = os.path.join(ws_dst, f'classify_engine_v{new_num}.py')
            if os.path.exists(ws_old_engine) and not os.path.exists(ws_new_engine):
                os.rename(ws_old_engine, ws_new_engine)
                print(f'   工作区引擎重命名: classify_engine_v{old_num}.py → classify_engine_v{new_num}.py')
            for fname in os.listdir(ws_dst):
                if not fname.endswith(('.py', '.md')):
                    continue
                fpath = os.path.join(ws_dst, fname)
                try:
                    with open(fpath, 'r', encoding='utf-8') as f:
                        content = f.read()
                except Exception:
                    continue
                new_content = update_version_marks(content, old_ver, new_ver)
                if new_content != content:
                    with open(fpath, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                    print(f'   更新: {fname}')
        else:
            print(f'4. ⚠️ 工作区旧技能目录不存在: {ws_src}（跳过）')

    # 5. 记录版本
    versioning = os.path.join(dst, 'VERSIONING.md')
    today = datetime.datetime.now().strftime('%Y-%m-%d')
    with open(versioning, 'a', encoding='utf-8') as f:
        f.write(f'- {new_ver}: {today} 由 upgrade_skill.py 从 {old_ver} 升级（替换模式）\n')
    print(f'5. 版本记录已追加: VERSIONING.md')

    # 6. 删除旧版本（单版本替换）
    if not args.keep_old:
        print(f'6. 删除旧版本（单版本替换）...')
        # 确保 cwd 不在旧目录（外部运行时 cwd 已是 D:\ 或其它）
        try:
            os.chdir(dst)
        except Exception:
            pass
        # 删除 D 盘旧版本（先删除内部文件，再删目录）
        if os.path.exists(src):
            shutil.rmtree(src, ignore_errors=True)
            if not os.path.exists(src):
                print(f'   已删除 D盘: {src}')
            else:
                print(f'   ⚠️ D盘旧目录删除失败（可能被占用），请手动删除: {src}')
        if args.ws:
            ws_src = os.path.join(WS_SKILLS, f'元宝分词分类引擎-{old_ver}')
            if os.path.exists(ws_src):
                shutil.rmtree(ws_src, ignore_errors=True)
                if not os.path.exists(ws_src):
                    print(f'   已删除 工作区: {ws_src}')
                else:
                    print(f'   ⚠️ 工作区旧目录删除失败，请手动删除: {ws_src}')

    print(f'\n✅ 升级完成: {old_ver} → {new_ver}（{"替换模式" if not args.keep_old else "保留旧版"}）')
    print(f'当前版本目录: {dst}')

if __name__ == '__main__':
    main()
