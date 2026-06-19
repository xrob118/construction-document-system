"""数据组合规则模块

负责从数据库中查询数据并组合成文档模板所需的占位符字典。
"""

import json
import os
import re
import sys
from datetime import datetime, timedelta

# 将父目录加入路径，以便导入 models 和 database
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models import Project, Process, Worker, GeneratedDoc, ProjectProcess, ProjectApproval, ProjectMember
from database import SessionLocal


# 数量/单位拆分：把 "1套" / "2台" / "1.5吨" / "10" / "10 个" 拆成 ("1", "套") / ("10", "")
_QTY_RE = re.compile(r'^\s*(\d+(?:\.\d+)?)\s*(.*?)\s*$')


def _split_quantity_and_unit(qty_str):
    """从 "1套" / "2台" / "1.5吨" 等中拆分数量和单位

    Returns:
        (数量字符串, 单位字符串) 元组。如果只有数字则单位为空字符串。
    """
    if not qty_str:
        return "", ""
    m = _QTY_RE.match(str(qty_str))
    if not m:
        return str(qty_str), ""
    num = m.group(1)
    unit = m.group(2).strip()
    return num, unit


# 5.2 施工标准：把 "1、xxx\n2、yyy\n3、zzz" 这种以"数字+、"开头的行拆成多段
# 用于解决单个工艺的施工标准挤成 6-10 行超长段，Word 无法在中间分页的问题
_SUBSTEP_RE = re.compile(r'(?:^|\n)\s*(\d+)\s*[、.．]\s*')


def _split_standards_into_substeps(text):
    """把施工标准文本按编号子步骤拆成多段

    输入示例：
        "1、混凝土强度等级、抗渗等级、厚度、保护层、垫层/底板标高和坡度必须符合设计。
         2、浇筑前钢筋...
         3、..."

    返回：
        ["1、混凝土强度等级...", "2、浇筑前钢筋...", "3、..."]
    """
    if not text:
        return []
    matches = list(_SUBSTEP_RE.finditer(text))
    if len(matches) < 2:
        return [text.strip()] if text.strip() else []
    parts = []
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        seg = text[start:end].strip()
        if seg:
            parts.append(seg)
    return parts


class DataService:
    """数据服务类，负责组合各类文档所需的数据"""

    @staticmethod
    def _get_worker_display_role(w, worker_roles=None):
        """获取工人在文档中的显示角色

        如果 worker_roles 中指定了该工人的角色，则使用指定角色；
        否则使用该工人的第一个角色（role）。
        """
        if worker_roles and w.id in worker_roles:
            return worker_roles[w.id]
        return w.role or "施工人员"

    @staticmethod
    def _get_worker_cert_for_role(w, role):
        """根据角色获取对应的证书（角色1↔证书1，角色2↔证书2，角色3↔证书3）"""
        if role == w.role:
            return w.certification or ""
        elif role == w.role2:
            return w.certification2 or ""
        elif role == w.role3:
            return w.certification3 or ""
        return ""

    @staticmethod
    def _get_worker_all_roles(w):
        """获取工人的所有角色列表"""
        return [r for r in [w.role, w.role2, w.role3] if r]

    def _auto_fetch_workers(self, db, project_id, worker_ids):
        """按"规则三"自动拉取工程相关人员

        优先级：
        1. 显式传入的 worker_ids
        2. ProjectMember（工程关联成员）
        3. Worker 全局库兜底（按 role 关键字匹配）
        """
        if worker_ids:
            return db.query(Worker).filter(Worker.id.in_(worker_ids)).all()

        members = db.query(ProjectMember).filter(
            ProjectMember.project_id == project_id,
        ).all()
        if members:
            class _MemberProxy:
                def __init__(self, m):
                    self.id = m.id
                    self.name = m.name
                    self.role = m.role
                    self.role2 = None
                    self.role3 = None
                    self.team = None
                    self.certification = None
                    self.certification2 = None
                    self.certification3 = None
            return [_MemberProxy(m) for m in members]

        # ProjectMember 为空时从 Worker 全局库兜底
        all_workers = db.query(Worker).all()
        result = []
        for w in all_workers:
            roles_text = ' '.join(filter(None, [w.role, w.role2, w.role3]))
            class _WorkerProxy:
                def __init__(self, src, merged_role):
                    self.id = src.id
                    self.name = src.name
                    self.role = merged_role
                    self.role2 = None
                    self.role3 = None
                    self.team = src.team
                    self.certification = src.certification
                    self.certification2 = src.certification2
                    self.certification3 = src.certification3
            if roles_text and any(k in roles_text for k in ['项目经理', '技术', '安全', '材料', '工作负责人', '施工员', '质检', '试验', '班组']):
                result.append(_WorkerProxy(w, roles_text))
        return result

    def get_construction_design_data(self, project_id, worker_ids, process_ids, worker_roles=None):
        """组合施工组织设计数据

        Args:
            project_id: 工程ID
            worker_ids: 人员ID列表
            process_ids: 工艺ID列表

        Returns:
            包含模板中所有占位符对应值的字典
        """
        db = SessionLocal()
        try:
            # 查询工程信息
            project = db.query(Project).filter(Project.id == project_id).first()
            if not project:
                raise ValueError(f"未找到ID为 {project_id} 的工程")

            # 查询人员信息（自动拉取）
            workers = self._auto_fetch_workers(db, project_id, worker_ids)

            # 查询工艺信息 - 优先从 ProjectProcess 关联读取
            if process_ids:
                processes = db.query(Process).filter(Process.id.in_(process_ids)).all()
            else:
                # 从 ProjectProcess 关联读取
                process_links = db.query(ProjectProcess).filter(
                    ProjectProcess.project_id == project_id
                ).all()
                process_ids = [pl.process_id for pl in process_links]
                processes = db.query(Process).filter(Process.id.in_(process_ids)).all() if process_ids else []

            # 组合人员信息字符串
            workers_info = []
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                cert = self._get_worker_cert_for_role(w, display_role)
                parts = [w.name, display_role]
                if w.team:
                    parts.append(w.team)
                if cert:
                    parts.append(cert)
                workers_info.append(f"{'，'.join(parts)}")

            # 组合工艺信息
            processes_info = []
            # 五、技术措施 拆为 5.1 施工工序 / 5.2 施工标准（每段独立成页）
            tech_procedures_parts = []  # 5.1 施工工序
            tech_standards_parts = []   # 5.2 施工标准（不含"依据"行）
            tech_basis_parts = []       # 施工依据（各工艺的"依据"行，单独成页，放在工器具前一页）
            # 六、安全措施 = 6.1 危险源识别 + 6.2 安全措施（同级别连续，不分页）
            safety_hazards_parts = []   # 6.1 危险源识别
            safety_measures_only_parts = []  # 6.2 安全措施
            emergency_measures_parts = []

            for idx, p in enumerate(processes, 1):
                flow_steps = json.loads(p.flow_steps) if p.flow_steps else []
                process_text = f"{idx}、【{p.name}】\n"
                if p.project_type:
                    process_text += f"工程部位：{p.project_type}\n"
                if p.category:
                    process_text += f"工艺分类：{p.category}\n"
                if p.sub_category:
                    process_text += f"工序等级：{p.sub_category}\n"
                if flow_steps:
                    process_text += f"施工工序：{' → '.join(flow_steps)}\n"
                if p.duration_days is not None:
                    process_text += f"施工天数：{p.duration_days} 天\n"
                if p.equipment:
                    process_text += f"施工设备：{p.equipment}\n"
                if p.hazards:
                    process_text += f"危险源识别：{p.hazards}\n"
                if p.safety_measures:
                    process_text += f"安全措施：{p.safety_measures}\n"
                processes_info.append(process_text)

                # 5.1 施工工序：flow_steps
                if flow_steps:
                    tech_procedures_parts.append(f"{idx}、【{p.name}】\n施工工序：{' → '.join(flow_steps)}")

                # 5.2 施工标准：解析 standards 字段，"依据"行单独提取
                if p.standards:
                    std_text = p.standards.strip()
                    basis_match = re.match(r'^(依据[：:]\s*.*?)\n', std_text, re.DOTALL)
                    if basis_match:
                        basis_content = basis_match.group(1).strip()
                        basis_value = basis_content[len('依据'):].lstrip('：:').strip()
                        rest_text = std_text[basis_match.end():].strip()
                        # "依据"行放入 tech_basis_parts
                        tech_basis_parts.append(f"{idx}、【{p.name}】\n依据：{basis_value}")
                        # 步骤 1-8 放入 tech_standards_parts（按编号子步骤拆段，
                        # 避免单个工艺挤成 6-10 行的超长段，Word 没法在中间分页，
                        # 进而导致 5.2 跨 6 页且每页只有 2-3 段、底部大量空白）
                        if rest_text:
                            sub_steps = _split_standards_into_substeps(rest_text)
                            header = f"{idx}、【{p.name}】"
                            if sub_steps and len(sub_steps) > 1:
                                tech_standards_parts.append(header + "\n" + sub_steps[0])
                                for step in sub_steps[1:]:
                                    tech_standards_parts.append(step)
                            else:
                                tech_standards_parts.append(f"{idx}、【{p.name}】\n{rest_text}")
                    else:
                        sub_steps = _split_standards_into_substeps(std_text)
                        if sub_steps and len(sub_steps) > 1:
                            tech_standards_parts.append(f"{idx}、【{p.name}】\n{sub_steps[0]}")
                            for step in sub_steps[1:]:
                                tech_standards_parts.append(step)
                        else:
                            tech_standards_parts.append(f"{idx}、【{p.name}】\n{std_text}")

                # 6.1 危险源识别：hazards
                if p.hazards:
                    safety_hazards_parts.append(f"{idx}、【{p.name}】\n危险源识别：{p.hazards}")
                # 6.2 安全措施：safety_measures
                if p.safety_measures:
                    safety_measures_only_parts.append(f"{idx}、【{p.name}】\n安全措施：{p.safety_measures}")

                # 组合应急措施：各工艺的 hazards + safety_measures
                emergency_parts = []
                if p.hazards:
                    emergency_parts.append(f"危险源：{p.hazards}")
                if p.safety_measures:
                    emergency_parts.append(f"应对措施：{p.safety_measures}")
                if emergency_parts:
                    emergency_measures_parts.append(f"{idx}、【{p.name}】\n" + "\n".join(emergency_parts))

            # 构建项目经理和技术负责人
            project_manager = ""
            tech_leader = ""
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                if display_role == "项目经理":
                    project_manager = w.name
                elif display_role == "技术负责人":
                    tech_leader = w.name

            # 解析施工机具清单 JSON
            # 数量字段格式: "1套" / "2台" / "1.5吨" 等，单位嵌在数字后面
            # 没有"unit"字段时，把数字和单位拆开：数量列只放数字，单位列放"套/台/..."
            equipment_table = []
            if project.equipment_list:
                try:
                    equip_data = json.loads(project.equipment_list)
                    if isinstance(equip_data, list):
                        for idx, item in enumerate(equip_data, 1):
                            if isinstance(item, dict):
                                qty_str = str(item.get("quantity", "")).strip()
                                num, unit = _split_quantity_and_unit(qty_str)
                                # 若条目自带 unit 字段，优先使用
                                explicit_unit = item.get("unit")
                                if explicit_unit:
                                    unit = str(explicit_unit).strip() or unit
                                # 名称里若已经包含单位（如"汽泵/地泵或料斗"），不再合并
                                equipment_table.append([
                                    str(idx),
                                    item.get("name", ""),
                                    item.get("code", ""),
                                    unit or "台",
                                    num or qty_str,
                                ])
                except (json.JSONDecodeError, TypeError):
                    pass

            # 解析质量控制点 JSON
            quality_table = []
            if project.quality_control:
                try:
                    qc_data = json.loads(project.quality_control)
                    if isinstance(qc_data, list):
                        for item in qc_data:
                            if isinstance(item, dict):
                                quality_table.append([
                                    item.get("project", ""),
                                    item.get("sub_project", ""),
                                    item.get("basis", ""),
                                    item.get("method", ""),
                                    item.get("responsible", ""),
                                    item.get("record", ""),
                                ])
                except (json.JSONDecodeError, TypeError):
                    pass

            # 读取审批信息（保持空白供手签）
            approvals = db.query(ProjectApproval).filter(
                ProjectApproval.project_id == project_id
            ).order_by(ProjectApproval.sort_order).all()

            # 审批机构：取审批表第一个非空的 organization，否则用编制单位
            approval_org = ""
            for ap in approvals:
                if ap.organization:
                    approval_org = ap.organization
                    break
            if not approval_org:
                approval_org = project.company_name or "监理单位"

            # 组合数据字典
            # 计算日期：开工日期前2天
            date_str = ""
            if project.start_date:
                try:
                    start_dt = datetime.strptime(project.start_date, "%Y-%m-%d")
                    date_str = (start_dt - timedelta(days=2)).strftime("%Y年%-m月")
                except ValueError:
                    date_str = datetime.now().strftime("%Y年%m月")
            else:
                date_str = datetime.now().strftime("%Y年%m月")

            # 工程概况：留空（用户要求删除第3页描述文字）
            project_overview = ""

            # 施工作业计划：开竣工时间
            schedule_info = f"施工计划：开工时间{project.start_date or '待定'}，竣工时间{project.end_date or '待定'}。"

            # 作业主要内容：work_task 字段
            work_content = project.work_task or ""

            # 组织措施人员
            # 角色匹配采用"宽松匹配"：数据库中常用"工作负责人"/"试验工"等角色名，
            # 而模板占位符是"技术员"/"材料员"——这里做语义映射
            org_leader = ""
            org_tech = ""
            org_safety = ""
            org_material = ""
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                role = display_role or ""
                if role == "项目经理" or "项目经理" in role:
                    if not org_leader:
                        org_leader = w.name
                # 技术员: 含"技术"/"工作负责人"/"试验工"/"质检员"/"施工员"
                if ("技术" in role or "工作负责人" in role
                        or "试验" in role or "质检" in role or "施工员" in role):
                    if not org_tech:
                        org_tech = w.name
                # 安全员: 含"安全"
                if "安全" in role:
                    if not org_safety:
                        org_safety = w.name
                # 材料员: 含"材料"/"施工班组"/"班组"
                if "材料" in role or "班组" in role:
                    if not org_material:
                        org_material = w.name

            # 应急组长 = 项目经理
            emergency_leader = org_leader or project_manager

            # 施工人员准备：施工人员姓名和岗位
            worker_prepare_lines = []
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                cert = self._get_worker_cert_for_role(w, display_role)
                line = f"{w.name}，职务：{display_role}"
                if cert:
                    line += f"，持证：{cert}"
                worker_prepare_lines.append(line)
            worker_prepare = "\n".join(worker_prepare_lines)

            # 现场施工：工艺流程 + 工艺标准组合
            construction_parts = []
            for idx, p in enumerate(processes, 1):
                flow_steps = json.loads(p.flow_steps) if p.flow_steps else []
                parts = [f"{idx}、【{p.name}】"]
                if flow_steps:
                    parts.append(f"施工工序：{' → '.join(flow_steps)}")
                if p.standards:
                    parts.append(f"施工标准：{p.standards}")
                if p.equipment:
                    parts.append(f"施工设备：{p.equipment}")
                if p.description:
                    parts.append(f"说明：{p.description}")
                construction_parts.append("\n".join(parts))
            construction_content = "\n\n".join(construction_parts) if construction_parts else ""

            # 安全措施：每个工艺的危险源识别 + 安全措施
            safety_parts = []
            for idx, p in enumerate(processes, 1):
                parts = [f"{idx}、【{p.name}】"]
                if p.hazards:
                    parts.append(f"危险源识别：{p.hazards}")
                if p.safety_measures:
                    parts.append(f"安全措施：{p.safety_measures}")
                if len(parts) > 1:
                    safety_parts.append("\n".join(parts))
            safety_content = "\n\n".join(safety_parts) if safety_parts else ""

            # 施工工艺标准及验收：各工艺标准（按编号子步骤拆段，
            # 避免单个工艺挤成 6-10 行的超长段，Word 没法在中间分页）
            standard_parts = []
            for idx, p in enumerate(processes, 1):
                if p.standards:
                    std_text = p.standards.strip()
                    sub_steps = _split_standards_into_substeps(std_text)
                    header = f"{idx}、【{p.name}】"
                    if len(sub_steps) > 1:
                        standard_parts.append(header + "\n" + sub_steps[0])
                        for step in sub_steps[1:]:
                            standard_parts.append(step)
                    else:
                        standard_parts.append(f"{header}\n{std_text}")
            process_standards = "\n\n".join(standard_parts) if standard_parts else ""

            # 五、技术措施 → 5.1 施工工序 + 5.2 施工标准（每段独立成页）
            # 六、危险源识别（独立章节）
            # 七、安全措施（独立章节，与六连续不分页）
            # 八、文明施工（单独占页，暂无内容）
            # 九、应急处置措施
            # 施工依据 → 单独成页，放在工器具前一页
            # 标记 "__PB__" 在 doc_generator 解析为分页符（<w:br w:type="page"/>）
            tech_procedures_text = "5.1 施工工序\n\n" + "\n\n".join(tech_procedures_parts) if tech_procedures_parts else ""
            tech_standards_text = "5.2 施工标准\n\n" + "\n\n".join(tech_standards_parts) if tech_standards_parts else ""
            tech_basis_text = "5.3 施工依据\n\n" + "\n\n".join(tech_basis_parts) if tech_basis_parts else ""
            # 六、危险源识别（独立章节）
            safety_hazards_text = "六、危险源识别\n\n" + "\n\n".join(safety_hazards_parts) if safety_hazards_parts else ""
            # 七、安全措施（独立章节，与六连续不分页）
            safety_measures_text = "七、安全措施\n\n" + "\n\n".join(safety_measures_only_parts) if safety_measures_only_parts else ""
            # 八、文明施工（单独占页，暂无内容）
            civilized_construction_text = "八、文明施工\n（暂无内容）"
            # 拼装：5.1 + 5.2 + 施工依据 之间用空行连接（不再强制分页），让其自然排版
            tech_parts = []
            if tech_procedures_text:
                tech_parts.append(tech_procedures_text)
            if tech_standards_text:
                tech_parts.append(tech_standards_text)
            if tech_basis_text:
                tech_parts.append(tech_basis_text)
            tech_measures_parts_joined = ("\n\n".join(tech_parts)).strip("\n")
            # 安全措施：六 + 七 连续（不分页）
            # 减少不必要的 __PB__，让内容自然排版，避免大片空白
            safety_parts = []
            if safety_hazards_text:
                safety_parts.append(safety_hazards_text)
            if safety_measures_text:
                # 七、安全措施不再强制拆分，让 Word 自然分页
                safety_parts.append(safety_measures_text)
            # 六、七之间不再加分页符，连续排版
            safety_measures_parts_joined = ("\n\n".join(safety_parts)).strip("\n")
            # 六、危险源识别前不再强制分页，让内容自然排版，避免大片空白
            # 八、文明施工（暂无内容）— 拼装时把它挪到九、应急处置措施的开头
            # 紧跟在 18 段安全措施之后会单独占两行（19 页），合并到应急章节更紧凑
            # 拼装方式调整，不动原始数据

            # 九、应急处置措施：不再强制分页，让 Word 自然排版
            # 去掉 __PB__，让九紧跟在安全措施后面，避免大片空白
            if emergency_measures_parts:
                emergency_measures_text = "\n\n".join(emergency_measures_parts)
            else:
                emergency_measures_text = ""
            # 八、文明施工（暂无内容）— 拼到九、应急处置措施的末尾（同段内，不分页）
            # 让"八"作为收尾段紧跟九最后一节，节省单占 2 行的空白
            # 用 "\n" 拼接（行内换行），让 doc_generator 把"八"和九最后一段合并在同一段
            # 拼装方式调整，不动原始数据
            if civilized_construction_text:
                em = emergency_measures_text
                if em.endswith("__PB__"):
                    em = em[:-len("__PB__")].rstrip("\n")
                # 用 "\n" 接续：em 末尾段的最后一个换行后跟 civilized_construction_text
                emergency_measures_text = (em + "\n" + civilized_construction_text) if em else civilized_construction_text

            data = {
                "title": "施工组织设计",
                "project_name": f"{project.project_code} {project.project_name}" if project.project_code else (project.project_name or ""),
                "project_code": project.project_code or "",
                "approval_org": approval_org,
                "company_name": project.company_name or "",
                "subcontractor": project.subcontractor or "",
                "project_type": project.project_type or "",
                "location": project.location or "",
                "start_date": project.start_date or "",
                "end_date": project.end_date or "",
                # description 置空（用户要求删除第3页描述文字）
                "description": "",
                "date": date_str,
                "project_manager": project_manager,
                "tech_leader": tech_leader,
                "workers": "\n".join(workers_info),
                "processes": "\n\n".join(processes_info),
                # 新增占位符
                "project_overview": project_overview,
                "schedule_info": schedule_info,
                "work_content": work_content,
                "org_leader": org_leader,
                "org_tech": org_tech,
                "org_safety": org_safety,
                "org_material": org_material,
                "emergency_leader": emergency_leader,
                "worker_prepare": worker_prepare,
                "construction_content": construction_content,
                "safety_measures_content": safety_content,
                "process_standards": process_standards,
                # 兼容旧字段：tech_measures / safety_measures 已按 5.1+5.2 / 6.1+6.2 拼装，含 __PB__ 分页符标记
                "tech_measures": tech_measures_parts_joined,
                "safety_measures": safety_measures_parts_joined,
                "emergency_measures": emergency_measures_text,
                "equipment_table": equipment_table,
                "quality_table": quality_table,
                "content": project_overview,
            }
            return data
        finally:
            db.close()

    def get_survey_data(self, project_id, worker_ids, content_items, worker_roles=None):
        """组合项目勘察单数据

        Args:
            project_id: 工程ID
            worker_ids: 人员ID列表
            content_items: 勘察内容项列表

        Returns:
            包含模板中所有占位符对应值的字典
        """
        db = SessionLocal()
        try:
            project = db.query(Project).filter(Project.id == project_id).first()
            if not project:
                raise ValueError(f"未找到ID为 {project_id} 的工程")

            # 规则三：自动拉取人员（与施工组织设计保持一致）
            workers = self._auto_fetch_workers(db, project_id, worker_ids)

            # 拉取项目关联工艺（用于自动生成安全措施/勘察内容）
            process_links = db.query(ProjectProcess).filter(
                ProjectProcess.project_id == project_id
            ).all()
            process_ids = [pl.process_id for pl in process_links]
            processes = db.query(Process).filter(Process.id.in_(process_ids)).all() if process_ids else []

            # 自动生成勘察内容：每条工艺的危险源 + 安全措施
            if not content_items and processes:
                content_items = []
                for p in processes:
                    if p.hazards:
                        content_items.append(f"【{p.name}】危险源：{p.hazards}")
                    if p.safety_measures:
                        content_items.append(f"【{p.name}】安全措施：{p.safety_measures}")

            # 自动生成安全措施（模板 {{safety_measures}} 占位符）
            # 只含"危险源识别 + 安全措施"两项，不含施工工序
            safety_measures_parts = []
            for p in processes:
                parts = [f"【{p.name}】"]
                if p.hazards:
                    parts.append(f"危险源识别：{p.hazards}")
                if p.safety_measures:
                    parts.append(f"安全措施：{p.safety_measures}")
                if len(parts) > 1:
                    safety_measures_parts.append("\n".join(parts))
            safety_measures_text = "\n\n".join(safety_measures_parts) if safety_measures_parts else ""

            # 组合人员信息
            workers_info = []
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                cert = self._get_worker_cert_for_role(w, display_role)
                if cert:
                    workers_info.append(f"{w.name}（{display_role}，{cert}）")
                else:
                    workers_info.append(f"{w.name}（{display_role}）")

            # 组合勘察内容
            survey_content = []
            if content_items:
                for idx, item in enumerate(content_items, 1):
                    survey_content.append(f"{idx}. {item}")

            # 查找勘察负责人与人员（项目字段优先，否则从 ProjectMember 兜底）
            # 注意：模板上"勘察负责人"/"勘察人员"/"记录人"/"勘察日期"均要求手写留空，
            # 这里统一置为空字符串，让模板里的下划线占位区显示
            survey_leader = ""
            survey_members = ""

            # 查找项目经理
            project_manager = ""
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                if display_role == "项目经理":
                    project_manager = w.name

            # 勘察编号兜底（项目没填则用项目编号+日期生成）
            survey_number = project.survey_number or (
                f"{project.project_code}-KC-{datetime.now().strftime('%Y%m%d')}" if project.project_code else ""
            )

            # 勘察日期：手写留空（模板上有下划线占位区）
            survey_date_str = ""

            data = {
                "title": "项目勘察单",
                "project_name": project.project_name or "",
                "project_code": project.project_code or "",
                "survey_unit": project.survey_unit or "",
                "survey_department": project.survey_department or "",
                "survey_number": survey_number,
                "survey_leader": survey_leader,
                "survey_members": survey_members,
                "line_name": project.line_name or "",
                "work_task": project.work_task or "",
                "power_off_range": project.power_off_range or "",
                "live_parts": project.live_parts or "",
                "danger_points": project.danger_points or "",
                "safety_measures": safety_measures_text,
                "survey_date": survey_date_str,
                "subcontractor": project.subcontractor or "",
                "project_type": project.project_type or "",
                "location": project.location or "",
                "start_date": project.start_date or "",
                "end_date": project.end_date or "",
                "description": project.description or "",
                "date": survey_date_str,
                "project_manager": project_manager,
                "workers": "\n".join(workers_info),
                "content": "\n".join(survey_content) if survey_content else "无勘察内容",
            }
            return data
        finally:
            db.close()

    def get_tech_briefing_data(self, project_id, worker_ids, content_items, worker_roles=None):
        """组合技术交底数据

        Args:
            project_id: 工程ID
            worker_ids: 人员ID列表
            content_items: 交底内容项列表

        Returns:
            包含模板中所有占位符对应值的字典
        """
        db = SessionLocal()
        try:
            project = db.query(Project).filter(Project.id == project_id).first()
            if not project:
                raise ValueError(f"未找到ID为 {project_id} 的工程")

            # 规则三：自动拉取人员
            workers = self._auto_fetch_workers(db, project_id, worker_ids)

            # 组合人员信息
            workers_info = []
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                cert = self._get_worker_cert_for_role(w, display_role)
                if cert:
                    workers_info.append(f"{w.name}（{display_role}，{cert}）")
                else:
                    workers_info.append(f"{w.name}（{display_role}）")

            # 从 ProjectProcess 关联读取工艺
            process_links = db.query(ProjectProcess).filter(
                ProjectProcess.project_id == project_id
            ).all()
            process_ids = [pl.process_id for pl in process_links]
            processes = db.query(Process).filter(Process.id.in_(process_ids)).all() if process_ids else []

            # 组合交底作业项目：工艺名称列表
            work_items_list = [p.name for p in processes]
            work_items = "、".join(work_items_list) if work_items_list else ""

            # 组合技术交底内容：各工艺的 standards + flow_steps
            tech_content_parts = []
            for p in processes:
                parts = [f"【{p.name}】"]
                flow_steps = json.loads(p.flow_steps) if p.flow_steps else []
                if flow_steps:
                    parts.append(f"施工工序：{' → '.join(flow_steps)}")
                if p.standards:
                    parts.append(f"施工标准：{p.standards}")
                if p.equipment:
                    parts.append(f"施工设备：{p.equipment}")
                if p.duration_days is not None:
                    parts.append(f"施工天数：{p.duration_days} 天")
                if len(parts) > 1:
                    tech_content_parts.append("\n".join(parts))

            tech_briefing_content = "\n\n".join(tech_content_parts) if tech_content_parts else ""

            # 组合交底内容（兼容旧接口）
            tech_content = []
            if content_items:
                for idx, item in enumerate(content_items, 1):
                    tech_content.append(f"{idx}. {item}")

            # 查找技术负责人
            tech_leader = ""
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                if display_role == "技术负责人":
                    tech_leader = w.name

            data = {
                "title": "技术交底记录",
                "project_name": project.project_name or "",
                "project_code": project.project_code or "",
                "company_name": project.company_name or "",
                "subcontractor": project.subcontractor or "",
                "project_type": project.project_type or "",
                "location": project.location or "",
                "start_date": project.start_date or "",
                "end_date": project.end_date or "",
                "description": project.description or "",
                "date": datetime.now().strftime("%Y年%m月%d日"),
                "briefing_date": datetime.now().strftime("%Y年%m月%d日"),
                "briefing_host": project.briefing_host or tech_leader or "",
                "tech_leader": tech_leader,
                "workers": "\n".join(workers_info),
                "work_items": work_items,
                "tech_briefing_content": tech_briefing_content,
                "content": "\n".join(tech_content) if tech_content else "无交底内容",
            }
            return data
        finally:
            db.close()

    def get_safety_briefing_data(self, project_id, worker_ids, content_items, worker_roles=None):
        """组合安全交底数据

        Args:
            project_id: 工程ID
            worker_ids: 人员ID列表
            content_items: 交底内容项列表

        Returns:
            包含模板中所有占位符对应值的字典
        """
        db = SessionLocal()
        try:
            project = db.query(Project).filter(Project.id == project_id).first()
            if not project:
                raise ValueError(f"未找到ID为 {project_id} 的工程")

            # 规则三：自动拉取人员
            workers = self._auto_fetch_workers(db, project_id, worker_ids)

            # 组合人员信息
            workers_info = []
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                cert = self._get_worker_cert_for_role(w, display_role)
                if cert:
                    workers_info.append(f"{w.name}（{display_role}，{cert}）")
                else:
                    workers_info.append(f"{w.name}（{display_role}）")

            # 从 ProjectProcess 关联读取工艺
            process_links = db.query(ProjectProcess).filter(
                ProjectProcess.project_id == project_id
            ).all()
            process_ids = [pl.process_id for pl in process_links]
            processes = db.query(Process).filter(Process.id.in_(process_ids)).all() if process_ids else []

            # 交底工作内容：填写工程概况
            work_content = project.description or ""

            # 危险源归到一个标题下展示
            hazard_parts = []
            for p in processes:
                if p.hazards:
                    hazard_parts.append(f"{p.name}：{p.hazards}")
            hazards_text = "\n".join(hazard_parts) if hazard_parts else ""

            # 安全措施归到一个标题下
            measure_parts = []
            for p in processes:
                if p.safety_measures:
                    measure_parts.append(f"{p.name}：{p.safety_measures}")
            measures_text = "\n".join(measure_parts) if measure_parts else ""

            # 组合主要交底内容：危险源 + 安全措施
            briefing_parts = []
            if hazards_text:
                briefing_parts.append("危险源识别：\n" + hazards_text)
            if measures_text:
                briefing_parts.append("安全措施：\n" + measures_text)
            safety_briefing_content = "\n\n".join(briefing_parts)

            # 组合安全交底内容（兼容旧接口）
            safety_content = []
            if content_items:
                for idx, item in enumerate(content_items, 1):
                    safety_content.append(f"{idx}. {item}")

            # 查找安全员
            safety_officer = ""
            for w in workers:
                display_role = self._get_worker_display_role(w, worker_roles)
                if display_role == "安全员":
                    safety_officer = w.name

            data = {
                "title": "施工安全交底",
                "project_name": project.project_name or "",
                "project_code": project.project_code or "",
                "company_name": project.company_name or "",
                "subcontractor": project.subcontractor or "",
                "project_type": project.project_type or "",
                "location": project.location or "",
                "start_date": project.start_date or "",
                "end_date": project.end_date or "",
                "description": project.description or "",
                "date": "",
                "briefing_date": "",
                "briefing_host": project.briefing_host or safety_officer or "",
                "safety_officer": safety_officer,
                "workers": "\n".join(workers_info),
                "work_items": work_content,
                "safety_briefing_content": safety_briefing_content,
                "hazards": hazards_text,
                "safety_measures": measures_text,
                "content": "\n".join(safety_content) if safety_content else "无交底内容",
            }
            return data
        finally:
            db.close()
