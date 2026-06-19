"""SQLAlchemy 数据模型定义"""

from datetime import datetime
from sqlalchemy import Column, Integer, String, Text, DateTime, ForeignKey, JSON, Float
from sqlalchemy.orm import relationship, backref
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()


class Project(Base):
    """工程信息模型"""
    __tablename__ = "projects"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    # === 基本信息 ===
    project_code = Column(String(50), unique=True, nullable=False, comment="工程编号")
    project_name = Column(String(200), nullable=False, comment="工程名称")
    project_type = Column(String(200), nullable=True, comment="工程类别(逗号分隔多选)")
    location = Column(String(300), nullable=True, comment="工程地点")
    voltage_level = Column(String(50), nullable=True, comment="电压等级")
    line_name = Column(String(200), nullable=True, comment="线路名称/设备双重名称")
    company_name = Column(String(200), nullable=True, comment="编制单位")
    work_task = Column(String(500), nullable=True, comment="工作任务")

    # === 分包与工期 ===
    subcontractor = Column(String(200), nullable=True, comment="分包单位(兼容旧数据)")
    subcontractor_civil = Column(String(200), nullable=True, comment="土建分包单位")
    subcontractor_electric = Column(String(200), nullable=True, comment="电气分包单位")
    start_date = Column(String(20), nullable=True, comment="开工日期")
    end_date = Column(String(20), nullable=True, comment="竣工日期")
    description = Column(Text, nullable=True, comment="工程概况")

    # === 勘察信息 ===
    survey_unit = Column(String(200), nullable=True, comment="勘察单位")
    survey_department = Column(String(200), nullable=True, comment="勘察部门")
    survey_number = Column(String(50), nullable=True, comment="勘察编号")
    survey_leader = Column(String(50), nullable=True, comment="勘察负责人")
    survey_members = Column(String(500), nullable=True, comment="勘察人员(逗号分隔)")
    power_off_range = Column(Text, nullable=True, comment="停电范围")
    live_parts = Column(Text, nullable=True, comment="保留的带电部位")
    danger_points = Column(Text, nullable=True, comment="作业现场危险点")

    # === 交底信息 ===
    briefing_host = Column(String(50), nullable=True, comment="交底主持人")

    # === 施工机具清单(JSON数组) ===
    # 格式: [{"name":"吊车","code":"001","unit":"台","quantity":1}, ...]
    equipment_list = Column(Text, nullable=True, comment="施工机具清单(JSON)")

    # === 质量控制点(JSON数组) ===
    # 格式: [{"project":"隐蔽工程","sub_project":"平面布置","basis":"设计图纸","method":"跟踪检查","responsible":"施工负责人","record":"评级记录"}, ...]
    quality_control = Column(Text, nullable=True, comment="质量控制点(JSON)")

    # === 工期说明 ===
    schedule_note = Column(String(500), nullable=True, comment="工期说明")

    # === 时间戳 ===
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment="更新时间")

    # 关联
    documents = relationship("GeneratedDoc", back_populates="project", cascade="all, delete-orphan")
    approvals = relationship("ProjectApproval", back_populates="project", cascade="all, delete-orphan")
    process_links = relationship("ProjectProcess", back_populates="project", cascade="all, delete-orphan")


class Process(Base):
    """施工工艺模型"""
    __tablename__ = "processes"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    code = Column(String(20), nullable=True, unique=True, index=True, comment="工艺编号(如 1-0-A, 1-2-B)")
    name = Column(String(200), nullable=False, comment="工艺名称")
    project_type = Column(String(100), nullable=True, comment="工程部位: 配电房/电缆通道排管/电缆井")
    category = Column(String(100), nullable=True, comment="工艺分类: 土建/电气")
    sub_category = Column(String(100), nullable=True, comment="工序等级: 一般/特殊（需专项施工方案）")
    flow_steps = Column(Text, nullable=True, comment="施工工序(JSON字符串)")
    duration_days = Column(Integer, nullable=True, comment="施工所需天数")
    # 排序权重：数值越小越先施工（同类别内排序）
    sort_order = Column(Integer, default=0, comment="施工排序(数值越小越先施工)")
    # 工序依赖：必须在此工序之前完成的工艺ID列表(逗号分隔)
    depends_on = Column(String(200), nullable=True, comment="前置依赖工艺ID(逗号分隔)")
    standards = Column(Text, nullable=True, comment="施工标准")
    equipment = Column(Text, nullable=True, comment="所需施工设备名称、型号及数量")
    hazards = Column(Text, nullable=True, comment="可能的危险源识别")
    safety_measures = Column(Text, nullable=True, comment="应采取的安全措施")
    description = Column(Text, nullable=True, comment="工艺说明")


class Worker(Base):
    """施工人员模型"""
    __tablename__ = "workers"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    name = Column(String(100), nullable=False, comment="姓名")
    role = Column(String(100), nullable=True, comment="角色/职务")
    role2 = Column(String(100), nullable=True, comment="角色/职务2")
    role3 = Column(String(100), nullable=True, comment="角色/职务3")
    team = Column(String(200), nullable=True, comment="所属班组")
    certification = Column(String(200), nullable=True, comment="资质证书")
    certification2 = Column(String(200), nullable=True, comment="资质证书2")
    certification3 = Column(String(200), nullable=True, comment="资质证书3")


class ScheduleTask(Base):
    """施工进度任务模型"""
    __tablename__ = "schedule_tasks"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False, comment="关联工程ID")
    task_name = Column(String(200), nullable=False, comment="任务名称")
    process_id = Column(Integer, nullable=True, comment="关联施工工艺ID")
    parent_id = Column(Integer, nullable=True, comment="父任务ID（用于分组）")
    start_date = Column(String(20), nullable=False, comment="计划开始日期")
    end_date = Column(String(20), nullable=False, comment="计划结束日期")
    actual_start = Column(String(20), nullable=True, comment="实际开始日期")
    actual_end = Column(String(20), nullable=True, comment="实际结束日期")
    progress = Column(Integer, default=0, comment="完成进度(0-100)")
    responsible = Column(String(100), nullable=True, comment="负责人")
    sort_order = Column(Integer, default=0, comment="排序序号")
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment="更新时间")

    # 关联工程
    project = relationship("Project", backref="schedule_tasks")


class ProjectMember(Base):
    """项目人员模型（管理人员和施工人员）"""
    __tablename__ = "project_members"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False, comment="关联工程ID")
    name = Column(String(100), nullable=False, comment="姓名")
    member_type = Column(String(50), nullable=False, comment="人员类型: manager/worker")
    role = Column(String(100), nullable=True, comment="职务/工种")

    # 关联工程
    project = relationship("Project", backref=backref("member_list", cascade="all, delete-orphan"))


class ProjectApproval(Base):
    """项目审批信息模型（审批表保持空白供手签）"""
    __tablename__ = "project_approvals"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False, comment="关联工程ID")
    role = Column(String(50), nullable=False, comment="角色: 监理/审核/批准/编制")
    organization = Column(String(200), nullable=True, comment="单位")
    sort_order = Column(Integer, default=0, comment="排序序号")

    # 关联工程
    project = relationship("Project", back_populates="approvals")


class ProjectProcess(Base):
    """项目关联施工工艺模型"""
    __tablename__ = "project_processes"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False, comment="关联工程ID")
    process_id = Column(Integer, ForeignKey("processes.id"), nullable=False, comment="关联工艺ID")

    # 关联
    project = relationship("Project", back_populates="process_links")
    process = relationship("Process")


class DocTemplate(Base):
    """文档模板模型"""
    __tablename__ = "doc_templates"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    doc_type = Column(String(50), nullable=False, unique=True, comment="文档类型: construction_design/survey/tech_briefing/safety_briefing/gantt_chart")
    file_name = Column(String(300), nullable=True, comment="原始文件名")
    file_path = Column(String(500), nullable=True, comment="文件存储路径")
    uploaded_at = Column(DateTime, default=datetime.now, comment="上传时间")


class GeneratedDoc(Base):
    """生成文档记录模型"""
    __tablename__ = "generated_docs"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False, comment="关联工程ID")
    doc_type = Column(String(50), nullable=False, comment="文档类型")
    file_path = Column(String(500), nullable=True, comment="文件存储路径")
    pdf_path = Column(String(500), nullable=True, comment="PDF预览文件路径")
    created_at = Column(DateTime, default=datetime.now, comment="生成时间")

    # 关联工程
    project = relationship("Project", back_populates="documents")
