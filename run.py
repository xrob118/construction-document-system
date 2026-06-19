"""
施工资料管理系统 - 启动脚本
支持开发模式和生产模式
"""
import sys
import os
import webbrowser
import argparse
import subprocess

def check_backend_deps():
    """检查后端依赖是否安装"""
    try:
        import fastapi
        import uvicorn
        import sqlalchemy
        import docx
        return True
    except ImportError:
        return False

def install_backend_deps():
    """安装后端依赖"""
    req_path = os.path.join(os.path.dirname(__file__), 'backend', 'requirements.txt')
    if os.path.exists(req_path):
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', req_path])
    else:
        print("未找到 requirements.txt，请手动安装依赖")

def check_frontend_build():
    """检查前端是否已构建"""
    static_dir = os.path.join(os.path.dirname(__file__), 'backend', 'static')
    return os.path.exists(static_dir) and os.path.exists(os.path.join(static_dir, 'index.html'))

def build_frontend():
    """构建前端"""
    frontend_dir = os.path.join(os.path.dirname(__file__), 'frontend')
    if not os.path.exists(os.path.join(frontend_dir, 'node_modules')):
        print("正在安装前端依赖...")
        subprocess.check_call(['npm', 'install'], cwd=frontend_dir, shell=True)

    print("正在构建前端...")
    subprocess.check_call(['npm', 'run', 'build'], cwd=frontend_dir, shell=True)

    # 将构建产物复制到后端static目录
    import shutil
    dist_dir = os.path.join(frontend_dir, 'dist')
    static_dir = os.path.join(os.path.dirname(__file__), 'backend', 'static')

    if os.path.exists(static_dir):
        shutil.rmtree(static_dir)
    shutil.copytree(dist_dir, static_dir)
    print(f"前端构建产物已复制到 {static_dir}")

def start_backend(host='127.0.0.1', port=8080, dev=False):
    """启动后端服务"""
    backend_dir = os.path.join(os.path.dirname(__file__), 'backend')
    sys.path.insert(0, backend_dir)

    os.environ['DEV_MODE'] = '1' if dev else '0'

    import uvicorn
    uvicorn.run(
        "main:app",
        host=host,
        port=port,
        reload=dev,
        log_level="info"
    )

def main():
    parser = argparse.ArgumentParser(description='施工资料管理系统')
    parser.add_argument('--dev', action='store_true', help='开发模式（前后端分离）')
    parser.add_argument('--host', default='127.0.0.1', help='服务地址（默认127.0.0.1，联网模式使用0.0.0.0）')
    parser.add_argument('--port', type=int, default=8080, help='服务端口（默认8080）')
    parser.add_argument('--no-browser', action='store_true', help='不自动打开浏览器')
    parser.add_argument('--build-frontend', action='store_true', help='仅构建前端')
    parser.add_argument('--install-deps', action='store_true', help='仅安装依赖')
    args = parser.parse_args()

    if args.install_deps:
        install_backend_deps()
        return

    if args.build_frontend:
        build_frontend()
        return

    # 检查后端依赖
    if not check_backend_deps():
        print("检测到缺少后端依赖，正在安装...")
        install_backend_deps()

    if not args.dev:
        # 生产模式需要前端构建产物
        if not check_frontend_build():
            print("检测到前端未构建，正在构建...")
            build_frontend()

    # 启动后端
    if not args.no_browser and not args.dev:
        import threading
        def open_browser():
            import time
            time.sleep(2)
            webbrowser.open(f'http://localhost:{args.port}')
        threading.Thread(target=open_browser, daemon=True).start()

    print(f"\n{'='*50}")
    print(f"  施工资料管理系统")
    print(f"  模式: {'开发模式' if args.dev else '生产模式'}")
    print(f"  地址: http://{args.host}:{args.port}")
    if args.host == '0.0.0.0':
        print(f"  联网访问: http://<本机IP>:{args.port}")
    print(f"{'='*50}\n")

    start_backend(host=args.host, port=args.port, dev=args.dev)

if __name__ == '__main__':
    main()
