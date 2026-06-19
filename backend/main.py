"""FastAPI 应用入口"""

import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from fastapi import Request

from database import create_tables, init_seed_data, migrate_processes_code, migrate_processes_duration_days, migrate_projects_new_columns, migrate_processes_new_columns, migrate_schedule_tasks_new_columns, migrate_generated_docs_new_columns, migrate_workers_role_cert_to_json
from routers import project, construction_design, survey, tech_briefing, safety_briefing, schedule, templates

# 获取项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 创建 FastAPI 应用
app = FastAPI(title="施工资料管理系统", description="施工资料管理系统后端API")

# 配置 CORS 中间件，允许所有来源
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---- 应用启动事件 ----
@app.on_event("startup")
def startup_event():
    """应用启动时初始化数据库和种子数据"""
    create_tables()
    migrate_projects_new_columns()
    migrate_processes_code()
    migrate_processes_duration_days()
    migrate_processes_new_columns()
    migrate_schedule_tasks_new_columns()
    migrate_generated_docs_new_columns()
    migrate_workers_role_cert_to_json()
    init_seed_data()


# ---- 注册路由 ----
app.include_router(project.router)
app.include_router(construction_design.router)
app.include_router(survey.router)
app.include_router(tech_briefing.router)
app.include_router(safety_briefing.router)
app.include_router(schedule.router)
app.include_router(templates.router)


# ---- 静态文件与前端托管 ----

# 挂载 output 目录为静态文件目录，用于下载生成的文档
output_dir = os.path.join(BASE_DIR, "output")
os.makedirs(output_dir, exist_ok=True)
app.mount("/output", StaticFiles(directory=output_dir), name="output")

# 前端静态文件托管（生产模式）
# 优先检查 static 目录（重启脚本复制目标），其次检查 dist 目录
frontend_dir = None
for candidate in ["static", "dist"]:
    path = os.path.join(BASE_DIR, candidate)
    if os.path.isdir(path) and os.path.exists(os.path.join(path, "index.html")):
        frontend_dir = path
        break

if frontend_dir:
    # 挂载 assets 子目录
    assets_dir = os.path.join(frontend_dir, "assets")
    if os.path.isdir(assets_dir):
        app.mount("/assets", StaticFiles(directory=assets_dir), name="frontend-assets")

    # SPA 回退：所有非API、非静态文件的GET请求返回 index.html
    @app.get("/{full_path:path}")
    async def serve_spa(full_path: str, request: Request):
        # 如果请求的是具体文件（有扩展名），尝试直接返回
        if "." in full_path.split("/")[-1]:
            file_path = os.path.join(frontend_dir, full_path)
            if os.path.isfile(file_path):
                return FileResponse(file_path)

        # 否则返回 index.html（SPA 路由回退）
        return FileResponse(os.path.join(frontend_dir, "index.html"))
