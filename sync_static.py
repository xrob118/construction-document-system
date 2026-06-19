import os
import shutil

project_root = os.path.dirname(os.path.abspath(__file__))
dist_dir = os.path.join(project_root, 'frontend', 'dist')
static_dir = os.path.join(project_root, 'backend', 'static')

# 删除旧的 static 目录
if os.path.exists(static_dir):
    shutil.rmtree(static_dir)
    print('已删除旧 static 目录')

# 复制新的构建产物
shutil.copytree(dist_dir, static_dir)
print(f'新构建产物已复制到 {static_dir}')

# 列出新文件
for f in sorted(os.listdir(os.path.join(static_dir, 'assets'))):
    if 'ProjectForm' in f or 'ProjectList' in f:
        print('  -', f)
