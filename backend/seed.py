"""
seed.py — 公开版 Demo 数据初始化

作用：清空数据库并写入虚构的工程、人员、工艺、设备等演示数据。
所有数据均为虚构，用于：
- TRAE AI 创造力大赛公开 Demo
- GitHub 公开仓库随仓库分发的体验数据
- 不含任何真实工程、人名、单位、地名

用法：
    python seed.py
"""

import os
import sys
from datetime import date, timedelta

# 让脚本能 import backend 包
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from database import engine, SessionLocal, Base
from models import (
    Project, Worker, ProjectMember, Process, ProjectProcess,
    ScheduleTask,
)


def reset_database():
    """删除并重建所有表"""
    print(">> 重建数据库表结构 ...")
    Base.metadata.drop_all(engine)
    Base.metadata.create_all(engine)


def seed_workers(db):
    """30 名虚构人员，覆盖项目经理/技术/安全/材料/施工员/工作负责人等角色"""
    workers_data = [
        ("张明", "项目经理", "施工班组成员", "质检员", "工程业务部", "B证"),
        ("李华", "项目经理", "电气试验员", None, "工程业务部", "B证"),
        ("王芳", "项目经理", "施工班组成员", None, "工程业务部", "B证"),
        ("赵刚", "技术负责人", "施工班组成员", None, "工程业务部", "电力工程技术副高级"),
        ("陈静", "技术负责人", "质检员", None, "工程业务部", "电力工程技术中级"),
        ("刘伟", "安全员", "施工班组成员", None, "工程业务部", "安全员C证"),
        ("黄敏", "安全员", "施工班组成员", "资料员", "工程业务部", "安全员C证"),
        ("周强", "材料员", "施工班组成员", None, "工程业务部", "高压电工证"),
        ("吴磊", "施工员", "施工班组成员", None, "工程业务部", "施工员证"),
        ("徐勇", "工作负责人", "施工班组成员", None, "工程业务部", "安全准入证书"),
        ("孙浩", "工作负责人", "施工班组成员", None, "工程业务部", "安全准入证书"),
        ("马超", "工作负责人", "施工班组成员", None, "工程业务部", "安全准入证书"),
        ("朱琳", "施工班组成员", None, None, "工程业务部", "高压电工证"),
        ("胡杰", "施工班组成员", None, None, "工程业务部", "高压电工证"),
        ("林峰", "施工班组成员", "工地电工", None, "工程业务部", "高压电工证"),
        ("何涛", "施工班组成员", None, None, "工程业务部", "高压电工证"),
        ("高翔", "施工班组成员", None, None, "工程业务部", "高压电工证"),
        ("罗斌", "电气试验员", "安全员", None, "工程业务部", "电气试验证"),
        ("宋晨", "电气试验员", None, None, "工程业务部", "电气试验证"),
        ("韩雪", "施工班组成员", "资料员", None, "工程业务部", "高压电工证"),
        ("冯雷", "项目工程师", "施工班组成员", None, "工程业务部", "电力工程技术副高级"),
        ("邓超", "项目工程师", "施工班组成员", "施工员", "工程业务部", "送电线路 技师"),
        ("彭丽", "项目工程师", "质检员", "资料员", "工程业务部", "电力电缆 技师"),
        ("曾鹏", "设计人员", "项目经理", None, "工程业务部", "设计专业证书"),
        ("蒋雯", "设计人员", "安全员", None, "工程业务部", "设计专业证书"),
        ("蔡旭", "企业管理人员", None, None, "工程业务部", "A证"),
        ("余晖", "施工班组成员", None, None, "工程业务部", "高压电工证"),
        ("潘涛", "施工班组成员", None, None, "工程业务部", "高压电工证"),
        ("袁野", "电力电缆工", None, None, "工程业务部", "特种高压电力电缆工证"),
        ("方明", "施工班组成员", "安全员", None, "工程业务部", "高压电工证"),
    ]
    workers = []
    for name, role, role2, role3, team, cert in workers_data:
        w = Worker(
            name=name, role=role, role2=role2, role3=role3,
            team=team, certification=cert,
        )
        db.add(w)
        workers.append(w)
    db.flush()
    print(f"   写入 {len(workers)} 名人员")
    return workers


def seed_processes(db):
    """12 条典型配电/电缆施工工艺"""
    processes_data = [
        ("10kV 电缆敷设", "高压电缆敷设；穿管保护；防火封堵",
         "电缆挤压、机械损伤", "做好防护，使用合格工器具"),
        ("10kV 电缆终端制作", "剥切、压接、密封、接地",
         "触电、烫伤", "佩戴绝缘手套，使用专用工具"),
        ("10kV 电缆中间接头制作", "接头安装、密封、接地",
         "触电、火灾", "现场配置灭火器，专人监护"),
        ("10kV 变压器安装", "就位、接线、试验",
         "挤压、坠落", "使用吊装带，专人指挥"),
        ("10kV 开关柜安装", "就位、母线连接、二次接线",
         "挤压、触电", "断电作业，挂接地线"),
        ("10kV 接地装置施工", "接地极埋设、扁钢敷设、焊接",
         "触电、灼伤", "使用绝缘工具，佩戴护目镜"),
        ("高压试验", "绝缘电阻、直流耐压、串联谐振",
         "触电、电击", "设置围栏，专人监护"),
        ("继电保护调试", "保护定值校验、二次回路检查",
         "触电、误操作", "断开跳闸回路，确认状态"),
        ("10kV 架空线路架设", "电杆组立、横担安装、导线架设",
         "高处坠落、触电", "使用安全带，作业前验电"),
        ("箱式变电站安装", "基础检查、箱体就位、接地",
         "挤压、坠落", "使用吊装带，专人指挥"),
        ("低压配电柜安装", "就位、母线连接、二次接线",
         "挤压、触电", "断电作业，挂接地线"),
        ("电力电缆故障测寻", "故障定位、路径探测、识别",
         "触电、机械伤害", "断电后作业，使用绝缘工具"),
    ]
    procs = []
    for name, description, hazards, safety in processes_data:
        p = Process(name=name, description=description, hazards=hazards, safety_measures=safety)
        db.add(p)
        procs.append(p)
    db.flush()
    print(f"   写入 {len(procs)} 条工艺")
    return procs


def seed_demo_project(db, workers, processes):
    """1 个虚构示范工程"""
    today = date.today()
    p = Project(
        project_code="DEMO-2026-001",
        project_name="某市科技园区 10kV 配电工程（示范工程）",
        project_type="配电工程",
        voltage_level="10kV",
        subcontractor="示范电力工程有限公司",
        company_name="示范电力工程有限公司",
        survey_unit="示范电力工程有限公司",
        survey_department="工程项目部",
        survey_number="DEMO-2026-001-KC",
        survey_leader="张明",
        survey_members="张明、李华、王芳、赵刚、刘伟",
        line_name="10kV 示范 I 回",
        work_task="新建 10kV 配电室 1 座，敷设电缆 2.5 公里",
        power_off_range="10kV 示范线全线",
        live_parts="10kV 母线及以下带电部位",
        danger_points="高压触电、高处坠落、电缆中间接头故障",
        briefing_host="赵刚",
        location="某市科技园区",
        start_date=str(today),
        end_date=str(today + timedelta(days=60)),
        description="本工程为示范工程，用于 TRAE AI 创造力大赛公开 Demo，所有数据均为虚构。",
    )
    db.add(p)
    db.flush()

    # 工程成员（10 人，覆盖关键角色）
    member_assignments = [
        ("张明", "项目经理"),
        ("赵刚", "技术负责人"),
        ("陈静", "技术员"),
        ("刘伟", "安全员"),
        ("黄敏", "安全员"),
        ("周强", "材料员"),
        ("吴磊", "施工员"),
        ("徐勇", "工作负责人"),
        ("孙浩", "工作负责人"),
        ("马超", "工作负责人"),
    ]
    member_ids = []
    for name, role in member_assignments:
        m_type = "manager" if role in ("项目经理", "技术负责人", "技术员", "安全员", "材料员", "施工员") else "worker"
        m = ProjectMember(project_id=p.id, name=name, member_type=m_type, role=role)
        db.add(m)
        member_ids.append(m.id)
    db.flush()

    # 关联工艺（全部 12 项）
    for proc in processes:
        pp = ProjectProcess(project_id=p.id, process_id=proc.id)
        db.add(pp)
    db.flush()

    print(f"   写入 1 个示范工程（{p.project_code}）")
    print(f"   关联工程成员 {len(member_assignments)} 人")
    print(f"   关联工艺 {len(processes)} 项")
    return p


def seed_schedule_tasks(db, project_id):
    """横道图示范任务（20 个任务，覆盖前期→施工→试验→验收）"""
    today = date.today()
    tasks = [
        ("施工准备", 0, 5, 100, "前期"),
        ("施工组织设计编制", 0, 3, 100, "前期"),
        ("现场勘察", 2, 3, 100, "前期"),
        ("材料进场", 4, 2, 100, "前期"),
        ("基础施工", 6, 8, 0, "土建"),
        ("设备就位", 14, 6, 0, "电气"),
        ("电缆敷设", 18, 10, 0, "电气"),
        ("电缆终端制作", 26, 4, 0, "电气"),
        ("电缆中间接头制作", 28, 5, 0, "电气"),
        ("变压器安装", 30, 4, 0, "电气"),
        ("开关柜安装", 32, 5, 0, "电气"),
        ("接地装置施工", 36, 4, 0, "电气"),
        ("二次接线", 38, 6, 0, "电气"),
        ("高压试验", 44, 4, 0, "试验"),
        ("继电保护调试", 46, 4, 0, "试验"),
        ("系统调试", 50, 3, 0, "试验"),
        ("试运行", 52, 3, 0, "验收"),
        ("竣工验收", 54, 3, 0, "验收"),
        ("资料归档", 56, 3, 0, "验收"),
        ("工程交付", 59, 1, 0, "验收"),
    ]
    for i, (name, start_offset, duration, progress, group) in enumerate(tasks):
        s = today + timedelta(days=start_offset)
        e = s + timedelta(days=duration)
        t = ScheduleTask(
            project_id=project_id,
            task_name=name,
            start_date=str(s),
            end_date=str(e),
            progress=progress,
            sort_order=i,
        )
        db.add(t)
    db.flush()
    print(f"   写入 {len(tasks)} 个进度任务")


def main():
    reset_database()
    db = SessionLocal()
    try:
        print(">> 写入演示数据 ...")
        workers = seed_workers(db)
        processes = seed_processes(db)
        project = seed_demo_project(db, workers, processes)
        seed_schedule_tasks(db, project.id)
        db.commit()
        print(">> 完成！数据库已就绪，所有数据均为虚构。")
    finally:
        db.close()


if __name__ == "__main__":
    main()