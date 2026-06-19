"""数据库连接与初始化"""

import os
import json
from sqlalchemy import create_engine, inspect, text
from sqlalchemy.orm import sessionmaker

from models import Base, Project, Process, Worker, ProjectMember, DocTemplate

# 数据库文件路径
DB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DB_DIR, "construction.db")

# 确保数据目录存在
os.makedirs(DB_DIR, exist_ok=True)

# SQLAlchemy 数据库引擎
DATABASE_URL = f"sqlite:///{DB_PATH}"
engine = create_engine(DATABASE_URL, echo=False, connect_args={"check_same_thread": False})

# 会话工厂
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


def get_db():
    """获取数据库会话的依赖注入函数"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def create_tables():
    """创建所有数据表（包括新表 project_approvals, project_processes）"""
    Base.metadata.create_all(bind=engine)


def _category_num(category):
    """工艺分类编号: 土建=1, 电气=2, 其他=0"""
    if not category:
        return 0
    return {"土建": 1, "电气": 2}.get(category, 0)


def _sub_category_num(sub_category):
    """工序等级编号: 一般=0, 特殊=2, 其他=0"""
    if not sub_category:
        return 0
    if "特殊" in sub_category:
        return 2
    return 0


def migrate_processes_code():
    """为 processes 表添加 code 列并为已存在数据自动填充编号"""
    db = SessionLocal()
    try:
        inspector = inspect(engine)
        if 'processes' not in inspector.get_table_names():
            return

        columns = [c['name'] for c in inspector.get_columns('processes')]

        if 'code' not in columns:
            with engine.connect() as conn:
                conn.execute(text("ALTER TABLE processes ADD COLUMN code VARCHAR(20)"))
                conn.commit()

        # 找出未编号的工艺，按 id 顺序填充
        processes = db.query(Process).filter((Process.code == None) | (Process.code == '')).order_by(Process.id).all()
        if not processes:
            return

        counters = {}
        for p in processes:
            key = (p.category or '', p.sub_category or '')
            if key not in counters:
                counters[key] = 0
            cat_num = _category_num(p.category)
            sub_num = _sub_category_num(p.sub_category)
            letter = chr(ord('A') + counters[key])
            p.code = f"{cat_num}-{sub_num}-{letter}"
            counters[key] += 1
        db.commit()
        print(f"已为 {len(processes)} 条工艺自动填充编号")
    except Exception as e:
        db.rollback()
        print(f"工艺编号迁移失败: {e}")
    finally:
        db.close()


def migrate_processes_duration_days():
    """为 processes 表添加 duration_days 列（如不存在）"""
    try:
        inspector = inspect(engine)
        if 'processes' not in inspector.get_table_names():
            return
        columns = [c['name'] for c in inspector.get_columns('processes')]
        if 'duration_days' in columns:
            return
        with engine.connect() as conn:
            conn.execute(text("ALTER TABLE processes ADD COLUMN duration_days INTEGER"))
            conn.commit()
        print("已为 processes 表添加 duration_days 列")
    except Exception as e:
        print(f"添加 duration_days 列失败: {e}")


def migrate_workers_role_cert_to_json():
    """为 workers 表添加 role2/role3/certification2/certification3 列（如不存在）"""
    try:
        inspector = inspect(engine)
        if 'workers' not in inspector.get_table_names():
            return
        columns = [c['name'] for c in inspector.get_columns('workers')]
        new_cols = {
            'role2': 'ALTER TABLE workers ADD COLUMN role2 VARCHAR(100)',
            'role3': 'ALTER TABLE workers ADD COLUMN role3 VARCHAR(100)',
            'certification2': 'ALTER TABLE workers ADD COLUMN certification2 VARCHAR(200)',
            'certification3': 'ALTER TABLE workers ADD COLUMN certification3 VARCHAR(200)',
        }
        for col_name, sql in new_cols.items():
            if col_name not in columns:
                with engine.connect() as conn:
                    conn.execute(text(sql))
                    conn.commit()
                print(f"已为 workers 表添加 {col_name} 列")
    except Exception as e:
        print(f"workers 表添加列失败: {e}")


def migrate_projects_new_columns():
    """为 projects 表添加新字段（如不存在），兼容已有数据"""
    try:
        inspector = inspect(engine)
        if 'projects' not in inspector.get_table_names():
            return

        existing_columns = [c['name'] for c in inspector.get_columns('projects')]

        # 定义 projects 表所有可能的列及其 SQL 类型（与 models.py 保持一致）
        new_columns = {
            "project_type": "VARCHAR(200)",
            "location": "VARCHAR(300)",
            "voltage_level": "VARCHAR(50)",
            "line_name": "VARCHAR(200)",
            "company_name": "VARCHAR(200)",
            "work_task": "VARCHAR(500)",
            "subcontractor": "VARCHAR(200)",
            "subcontractor_civil": "VARCHAR(200)",
            "subcontractor_electric": "VARCHAR(200)",
            "start_date": "VARCHAR(20)",
            "end_date": "VARCHAR(20)",
            "description": "TEXT",
            "survey_unit": "VARCHAR(200)",
            "survey_department": "VARCHAR(200)",
            "survey_number": "VARCHAR(50)",
            "survey_leader": "VARCHAR(50)",
            "survey_members": "VARCHAR(500)",
            "power_off_range": "TEXT",
            "live_parts": "TEXT",
            "danger_points": "TEXT",
            "briefing_host": "VARCHAR(50)",
            "equipment_list": "TEXT",
            "quality_control": "TEXT",
            "schedule_note": "VARCHAR(500)",
        }

        added = []
        with engine.connect() as conn:
            for col_name, col_type in new_columns.items():
                if col_name not in existing_columns:
                    conn.execute(text(f"ALTER TABLE projects ADD COLUMN {col_name} {col_type}"))
                    added.append(col_name)
            conn.commit()

        if added:
            print(f"已为 projects 表添加新列: {', '.join(added)}")
    except Exception as e:
        print(f"projects 表迁移失败: {e}")


def migrate_processes_new_columns():
    """为 processes 表添加 sort_order 和 depends_on 列（如不存在）"""
    try:
        inspector = inspect(engine)
        if 'processes' not in inspector.get_table_names():
            return

        existing_columns = [c['name'] for c in inspector.get_columns('processes')]

        new_columns = {
            "sort_order": "INTEGER DEFAULT 0",
            "depends_on": "VARCHAR(200)",
        }

        added = []
        with engine.connect() as conn:
            for col_name, col_type in new_columns.items():
                if col_name not in existing_columns:
                    conn.execute(text(f"ALTER TABLE processes ADD COLUMN {col_name} {col_type}"))
                    added.append(col_name)
            conn.commit()

        if added:
            print(f"已为 processes 表添加新列: {', '.join(added)}")
    except Exception as e:
        print(f"processes 表新列迁移失败: {e}")


def migrate_schedule_tasks_new_columns():
    """为 schedule_tasks 表添加 process_id 列（如不存在）"""
    try:
        inspector = inspect(engine)
        if 'schedule_tasks' not in inspector.get_table_names():
            return

        existing_columns = [c['name'] for c in inspector.get_columns('schedule_tasks')]

        new_columns = {
            "process_id": "INTEGER",
        }

        added = []
        with engine.connect() as conn:
            for col_name, col_type in new_columns.items():
                if col_name not in existing_columns:
                    conn.execute(text(f"ALTER TABLE schedule_tasks ADD COLUMN {col_name} {col_type}"))
                    added.append(col_name)
            conn.commit()

        if added:
            print(f"已为 schedule_tasks 表添加新列: {', '.join(added)}")
    except Exception as e:
        print(f"schedule_tasks 表新列迁移失败: {e}")


def migrate_generated_docs_new_columns():
    """为 generated_docs 表添加 pdf_path 列（如不存在）"""
    try:
        inspector = inspect(engine)
        if 'generated_docs' not in inspector.get_table_names():
            return

        existing_columns = [c['name'] for c in inspector.get_columns('generated_docs')]

        new_columns = {
            "pdf_path": "VARCHAR(500)",
        }

        added = []
        with engine.connect() as conn:
            for col_name, col_type in new_columns.items():
                if col_name not in existing_columns:
                    conn.execute(text(f"ALTER TABLE generated_docs ADD COLUMN {col_name} {col_type}"))
                    added.append(col_name)
            conn.commit()

        if added:
            print(f"已为 generated_docs 表添加新列: {', '.join(added)}")
    except Exception as e:
        print(f"generated_docs 表新列迁移失败: {e}")


def init_seed_data():
    """插入种子数据（仅在表为空时插入）"""
    db = SessionLocal()
    try:
        # 检查是否已有数据，避免重复插入
        if db.query(Process).first() is not None:
            return

        # ---- 插入施工工艺种子数据 ----
        # 注意：depends_on 需要在插入后用实际 ID 回填
        processes_data = [
            {
                "name": "土方开挖",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["施工准备", "测量放线", "表层清理", "机械开挖", "人工修整", "验槽"], ensure_ascii=False),
                "standards": "基底标高允许偏差0~-50mm；基底平整度≤20mm；边坡坡度符合设计要求",
                "equipment": "0.3~0.6m³反铲挖掘机1台",
                "hazards": "基坑坍塌、管线损坏、机械伤害",
                "safety_measures": "1、开挖前核对地下管线位置。2、作业区设置围挡和警示牌。3、分层开挖，严禁超挖。",
                "description": "土方开挖工程是指将土方挖除至设计标高的施工过程。",
                "duration_days": 3,
                "sort_order": 10,
                "depends_on": None,
            },
            {
                "name": "钢筋绑扎",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["钢筋加工", "钢筋配料", "钢筋连接", "钢筋绑扎", "隐蔽验收"], ensure_ascii=False),
                "standards": "钢筋规格符合设计要求；绑扎间距偏差±10mm；保护层厚度偏差±3mm",
                "equipment": "钢筋弯曲机1台、钢筋切断机1台",
                "hazards": "钢筋划伤、机械伤害、高处坠落",
                "safety_measures": "1、佩戴防割手套。2、机械操作持证上岗。3、绑扎平台稳固可靠。",
                "description": "钢筋绑扎工程是指按照设计图纸要求，将钢筋加工成型后进行绑扎安装的施工过程。",
                "duration_days": 5,
                "sort_order": 20,
                "depends_on": "土方开挖",  # 占位符，插入后用实际 ID 替换
            },
            {
                "name": "模板安装",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["模板设计", "模板制作", "模板安装", "支撑加固", "预检验收"], ensure_ascii=False),
                "standards": "模板平整度≤5mm；模板垂直度≤6mm；接缝严密不漏浆",
                "equipment": "木工圆锯1台、冲击钻2把",
                "hazards": "模板倒塌、高处坠落、机械伤害",
                "safety_measures": "1、支撑体系经计算复核。2、高处作业系安全带。3、拆除时由上而下逐层进行。",
                "description": "模板安装工程是指按照设计要求安装模板及支撑系统的施工过程。",
                "duration_days": 4,
                "sort_order": 30,
                "depends_on": "钢筋绑扎",
            },
            {
                "name": "混凝土浇筑",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["浇筑准备", "模板检查", "混凝土拌合", "混凝土运输", "混凝土浇筑", "振捣密实", "养护"], ensure_ascii=False),
                "standards": "混凝土强度满足设计要求；浇筑连续进行无冷缝；振捣密实无蜂窝麻面",
                "equipment": "汽车泵或地泵1套、振捣器2台",
                "hazards": "模板胀模、泵管甩动、机械伤害",
                "safety_measures": "1、浇筑前检查模板和支撑。2、泵管接头牢固。3、振捣器绝缘良好。",
                "description": "混凝土浇筑工程是指将拌合好的混凝土浇注入模并振捣密实的施工过程。",
                "duration_days": 3,
                "sort_order": 40,
                "depends_on": "模板安装",
            },
            {
                "name": "砌体工程",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["材料准备", "放线定位", "砌筑", "勾缝", "养护"], ensure_ascii=False),
                "standards": "砂浆饱满度≥80%；垂直度偏差≤5mm；平整度偏差≤8mm",
                "equipment": "砂浆搅拌机1台",
                "hazards": "墙体倒塌、高处坠落、砂浆飞溅",
                "safety_measures": "1、砌筑高度每日不超过1.5m。2、脚手架搭设合格。3、佩戴防护眼镜。",
                "description": "砌体工程是指使用砖、石、砌块等材料进行砌筑的施工过程。",
                "duration_days": 6,
                "sort_order": 50,
                "depends_on": "混凝土浇筑",
            },
            {
                "name": "防水工程",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["基层处理", "防水层施工", "保护层施工", "闭水试验"], ensure_ascii=False),
                "standards": "防水层无渗漏；搭接宽度≥100mm；闭水试验24h无渗漏",
                "equipment": "热熔枪1把、滚刷2把",
                "hazards": "烫伤、有毒气体、滑跌",
                "safety_measures": "1、热熔施工佩戴隔热手套。2、作业区域通风良好。3、防水层未固化前禁止踩踏。",
                "description": "防水工程是指为防止水渗入建筑物而进行防水层施工的过程。",
                "duration_days": 4,
                "sort_order": 60,
                "depends_on": "砌体工程",
            },
            {
                "name": "脚手架搭设",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["方案编制", "基础处理", "立杆搭设", "横杆安装", "剪刀撑设置", "验收"], ensure_ascii=False),
                "standards": "立杆间距偏差±20mm；步距偏差±20mm；连墙件设置符合规范",
                "equipment": "钢管脚手架1套、扳手4把",
                "hazards": "高处坠落、架体倒塌、物体打击",
                "safety_measures": "1、持证架子工操作。2、连墙件随搭随设。3、作业面满铺脚手板并设挡脚板。",
                "description": "脚手架搭设工程是指为施工作业提供操作平台和安全防护而搭设脚手架的过程。",
                "duration_days": 2,
                "sort_order": 15,
                "depends_on": None,
            },
            {
                "name": "抹灰工程",
                "project_type": "配电房",
                "category": "土建",
                "sub_category": "一般",
                "flow_steps": json.dumps(["基层处理", "贴灰饼", "抹底层灰", "抹中层灰", "抹面层灰", "养护"], ensure_ascii=False),
                "standards": "表面平整度≤4mm；阴阳角方正≤4mm；无空鼓裂缝",
                "equipment": "砂浆搅拌机1台、抹子4把",
                "hazards": "高处坠落、砂浆入眼、空鼓脱落",
                "safety_measures": "1、操作平台稳固。2、佩戴防护眼镜。3、基层清理干净，分层抹灰。",
                "description": "抹灰工程是指用砂浆涂抹在建筑物表面进行找平和装饰的施工过程。",
                "duration_days": 5,
                "sort_order": 70,
                "depends_on": "防水工程",
            },
        ]

        # 先插入所有工艺（depends_on 暂用名称占位）
        created_processes = {}
        for item in processes_data:
            depends_on_name = item.pop("depends_on", None)
            process = Process(**item)
            db.add(process)
            db.flush()  # flush 以获取自增 ID
            created_processes[process.name] = {
                "obj": process,
                "depends_on_name": depends_on_name,
            }

        # 回填 depends_on：将名称替换为实际 ID
        for name, info in created_processes.items():
            dep_name = info["depends_on_name"]
            if dep_name and dep_name in created_processes:
                dep_id = created_processes[dep_name]["obj"].id
                info["obj"].depends_on = str(dep_id)

        # ---- 插入施工人员种子数据 ----
        workers_data = [
            {"name": "张建国", "role": "项目经理", "team": "项目管理部", "certification": "一级建造师"},
            {"name": "李明辉", "role": "技术负责人", "team": "技术部", "certification": "高级工程师"},
            {"name": "王安全", "role": "安全员", "team": "安全管理部", "certification": "安全员C证"},
            {"name": "赵质控", "role": "质量员", "team": "质量管理部", "certification": "质量员证书"},
            {"name": "刘施工", "role": "施工员", "team": "施工一部", "certification": "施工员证书"},
            {"name": "陈测量", "role": "测量员", "team": "技术部", "certification": "测量员证书"},
            {"name": "周材料", "role": "材料员", "team": "物资部", "certification": "材料员证书"},
            {"name": "吴资料", "role": "资料员", "team": "综合部", "certification": "资料员证书"},
            {"name": "孙钢筋", "role": "钢筋工长", "team": "钢筋班组", "certification": "钢筋工技能证"},
            {"name": "马模板", "role": "木工工长", "team": "模板班组", "certification": "木工技能证"},
            {"name": "黄混凝土", "role": "混凝土工长", "team": "混凝土班组", "certification": "混凝土工技能证"},
            {"name": "杨架子", "role": "架子工长", "team": "架子班组", "certification": "架子工操作证"},
        ]

        for item in workers_data:
            worker = Worker(**item)
            db.add(worker)

        db.commit()
        print("种子数据初始化完成")
    except Exception as e:
        db.rollback()
        print(f"种子数据初始化失败: {e}")
    finally:
        db.close()
