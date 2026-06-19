# 施工资料智能生成系统

> 面向电力施工单位的「文档智能体」—— 选择工程、点击按钮，5 秒内输出 5 份标准 Word 档案 + 1 张进度横道图。
>
> 🏆 参赛项目：[TRAE AI 创造力大赛 · 学习工作赛道](https://www.trae.cn/ai-creativity)

![Made with TRAE](https://img.shields.io/badge/Made%20with-TRAE-orange)
![Vue 3](https://img.shields.io/badge/Vue-3-42b883)
![FastAPI](https://img.shields.io/badge/FastAPI-009688)
![License](https://img.shields.io/badge/License-MIT-blue)

---

## 🎯 一句话

把资料员从「4 小时敲字」里解放出来——选工程，点生成，5 秒下载标准 Word 文档。

---

## 🚀 1 分钟 Demo（GitHub Codespaces）

> 最快方式：点下面的按钮，云端自动启动完整 Demo。

[![Open in GitHub Codespaces](https://github.com/codespaces/badge.svg)](https://codespaces.new/xrob118/construction-document-system?quickstart=1)

启动后访问 **8080 端口**（Codespaces 会自动转发）即可看到完整 UI。

---

## 🖥️ 本地 5 分钟启动

### 前置环境
- Python 3.10+
- Node.js 18+（仅首次构建需要）

### 步骤

```bash
# 1. 克隆仓库
git clone https://github.com/xrob118/construction-document-system.git
cd construction-document-system

# 2. 安装 Python 依赖
cd backend
pip install -r requirements.txt

# 3. （可选）重置数据库为演示数据
python seed.py

# 4. 启动后端（默认端口 8080）
python run.py

# 5. 浏览器访问
# http://127.0.0.1:8080
```

> 💡 仓库内已包含构建好的前端（`backend/static/`），无需 npm install 即可运行。
> 如需修改前端代码：`cd frontend && npm install && npm run build`，再把 `dist/` 同步到 `backend/static/`。

---

## ✨ 五大模块 · 一键生成

每个模块都遵循统一的交互：**「选择工程 → 点击生成 → 预览 → 下载」**。人员、工艺、日期等数据由后端从工程关联自动拉取——前端没有任何多余选项。

| # | 模块 | 输出 | 占位符 |
|---|---|---|---|
| M-01 | 施工组织设计 | DOCX | 17 |
| M-02 | 项目勘察单 | DOCX | 11 |
| M-03 | 技术交底记录 | DOCX | 6 |
| M-04 | 安全交底记录 | DOCX | 6 |
| M-05 | 施工进度横道图 | XLSX | — |

> 详见 [项目结构](#-项目结构) 与 [`backend/services/data_service.py`](backend/services/data_service.py)。

---

## 📊 看得见的效率

| 指标 | 优化前 | 优化后 |
|---|---|---|
| 单份文档耗时 | 4 小时 | **5 秒** |
| 文档退回率 | 68% | **0%**（自动校验） |
| 每项目年节约人工 | — | **≈ 200 小时** |

---

## 🏗️ 项目结构

```
construction-document-system/
├── backend/                       # FastAPI 后端
│   ├── main.py                    # 应用入口
│   ├── run.py                     # 启动脚本
│   ├── database.py                # SQLite 连接
│   ├── models.py                  # ORM 模型
│   ├── seed.py                    # 演示数据初始化
│   ├── requirements.txt           # Python 依赖
│   ├── data/
│   │   └── construction.db        # SQLite 数据库（含演示数据）
│   ├── routers/                   # API 路由
│   │   ├── construction_design.py # 施工组织设计
│   │   ├── survey.py              # 项目勘察单
│   │   ├── tech_briefing.py       # 技术交底
│   │   ├── safety_briefing.py     # 安全交底
│   │   ├── schedule.py            # 施工进度横道图
│   │   ├── project.py             # 工程基础数据
│   │   └── templates.py           # 模板管理
│   ├── services/                  # 业务服务
│   │   ├── data_service.py        # 数据组合（5 模块共享）
│   │   ├── doc_generator.py       # Word/XLSX 生成器
│   │   └── ...
│   ├── templates/                 # Word/XLSX 模板（含 *_v2.docx）
│   └── static/                    # 前端构建产物（自动服务）
├── frontend/                      # Vue 3 前端
│   ├── src/
│   │   ├── views/                 # 页面（5 模块 + 数据基础）
│   │   ├── api/                   # axios 接口封装
│   │   ├── router/
│   │   └── ...
│   └── dist/                      # 构建产物（同步到 backend/static）
├── competition/                   # 比赛材料
│   └── product.html               # 创意产物 HTML 附件
├── .devcontainer/                 # GitHub Codespaces 配置
├── .gitignore
├── LICENSE
└── README.md
```

---

## 🛠️ 技术栈

| 层级 | 选型 | 理由 |
|---|---|---|
| 前端 | Vue 3 + Vite + Element Plus | 上手快，组件全 |
| 后端 | FastAPI + SQLAlchemy | Python 生态，文档自动生成 |
| 数据库 | SQLite | 零部署，单文件 |
| 文档生成 | python-docx + openpyxl | 业界标准库 |
| 文档预览 | mammoth | DOCX → HTML 实时预览 |

---

## 🔌 核心 API

| 端点 | 方法 | 用途 |
|---|---|---|
| `/api/projects` | GET | 工程列表 |
| `/api/construction-design/generate` | POST | 生成施工组织设计 |
| `/api/survey/generate` | POST | 生成项目勘察单 |
| `/api/tech-briefing/generate` | POST | 生成技术交底 |
| `/api/safety-briefing/generate` | POST | 生成安全交底 |
| `/api/schedule-tasks/import-template` | POST | 从模板导入进度任务 |
| `/api/schedule-tasks/generate` | POST | 生成横道图 xlsx |
| `/api/projects/preview/{doc_id}` | GET | 文档预览 HTML |

> 列表型接口统一返回 `{ items, total, page, page_size }`。

---

## 🧪 演示数据

`seed.py` 会写入一份虚构的演示工程（"某市科技园区 10kV 配电工程"）+ 30 名虚构人员 + 12 条典型工艺 + 20 个进度任务。**所有数据均为虚构**，用于：

- 本地快速体验
- GitHub Codespaces 首次启动
- TRAE 比赛公开 Demo

如需使用自己的真实数据，只需不运行 `seed.py`，在前端界面手动录入即可。

---

## 🤝 参与贡献

欢迎 PR、Issue 与 Fork。本项目采用 MIT 协议，可自由用于商业或学习用途。

---

## 📜 协议

[MIT](LICENSE) © 2026 Construction Document System Contributors

---

## 🏆 关于比赛

本项目为 **TRAE AI 创造力大赛初赛** 参赛作品（学习工作赛道）。  
使用 [TRAE IDE](https://www.trae.cn/) + Auto 模型开发。  
创意产物 HTML 见 [`competition/product.html`](competition/product.html)。