<template>
  <div class="project-form">
    <!-- 页面标题区 -->
    <div class="page-header">
      <div class="header-left">
        <el-icon size="28" color="#409EFF"><DocumentAdd /></el-icon>
        <div>
          <h2 class="header-title">新建工程</h2>
          <p class="header-desc">请填写工程信息，带 * 为必填项</p>
        </div>
      </div>
      <div class="header-right">
        <el-button @click="handleSaveDraft" :loading="savingDraft">
          <el-icon><FolderChecked /></el-icon>保存草稿
        </el-button>
        <el-button v-if="hasDraft" type="warning" plain @click="handleLoadDraft">
          <el-icon><RefreshLeft /></el-icon>恢复草稿
        </el-button>
      </div>
    </div>

    <el-form ref="formRef" :model="form" :rules="rules" label-width="120px" class="form-body">

      <!-- 1. 基本信息 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><EditPen /></el-icon>
            <span>基本信息</span>
          </div>
        </template>
        <el-row :gutter="24">
          <el-col :span="12">
            <el-form-item label="项目编号" prop="project_code">
              <el-input v-model="form.project_code" placeholder="请输入项目编号" clearable>
                <template #prefix><el-icon><Key /></el-icon></template>
              </el-input>
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="项目名称" prop="project_name">
              <el-input v-model="form.project_name" placeholder="请输入项目名称" clearable>
                <template #prefix><el-icon><EditPen /></el-icon></template>
              </el-input>
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="24">
          <el-col :span="12">
            <el-form-item label="工程类别" prop="project_type">
              <el-checkbox-group v-model="form.project_type">
                <el-checkbox
                  v-for="opt in projectTypeOptions"
                  :key="opt"
                  :value="opt"
                  border
                >
                  {{ opt }}
                </el-checkbox>
              </el-checkbox-group>
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="工程地点" prop="location">
              <el-input v-model="form.location" placeholder="请输入工程地点" clearable>
                <template #prefix><el-icon><Location /></el-icon></template>
              </el-input>
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="24">
          <el-col :span="12">
            <el-form-item label="电压等级">
              <el-input v-model="form.voltage_level" placeholder="请输入电压等级" clearable />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="线路名称">
              <el-input v-model="form.line_name" placeholder="请输入线路名称" clearable />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="24">
          <el-col :span="12">
            <el-form-item label="编制单位">
              <el-input v-model="form.company_name" placeholder="请输入编制单位" clearable>
                <template #prefix><el-icon><OfficeBuilding /></el-icon></template>
              </el-input>
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="工作任务">
              <el-input v-model="form.work_task" placeholder="请输入工作任务" clearable />
            </el-form-item>
          </el-col>
        </el-row>
      </el-card>

      <!-- 2. 分包与工期 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><Calendar /></el-icon>
            <span>分包与工期</span>
          </div>
        </template>
        <el-row :gutter="24">
          <el-col :span="12">
            <el-form-item label="土建分包单位" prop="subcontractor_civil">
              <el-input v-model="form.subcontractor_civil" placeholder="请输入土建分包单位" clearable>
                <template #prefix><el-icon><OfficeBuilding /></el-icon></template>
              </el-input>
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="电气分包单位" prop="subcontractor_electric">
              <el-input v-model="form.subcontractor_electric" placeholder="请输入电气分包单位" clearable>
                <template #prefix><el-icon><OfficeBuilding /></el-icon></template>
              </el-input>
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="24">
          <el-col :span="12">
            <el-form-item label="开工日期" prop="start_date">
              <el-date-picker
                v-model="form.start_date"
                type="date"
                placeholder="请选择开工日期"
                value-format="YYYY-MM-DD"
                style="width: 100%"
              />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="竣工日期" prop="end_date">
              <el-date-picker
                v-model="form.end_date"
                type="date"
                placeholder="请选择竣工日期"
                value-format="YYYY-MM-DD"
                style="width: 100%"
                :disabled-date="(date) => form.start_date && date < new Date(form.start_date)"
              />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item v-if="form.start_date && form.end_date" label="工期">
          <el-tag type="info" size="large">共 {{ dayCount }} 天</el-tag>
        </el-form-item>
      </el-card>

      <!-- 3. 工程概况 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><Document /></el-icon>
            <span>工程概况</span>
          </div>
        </template>
        <el-form-item label="工程概况" prop="description">
          <el-input
            v-model="form.description"
            type="textarea"
            :rows="6"
            placeholder="请输入工程概况描述，包括工程规模、结构形式、主要施工内容等"
            show-word-limit
            maxlength="2000"
          />
        </el-form-item>
      </el-card>

      <!-- 4. 人员信息 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><User /></el-icon>
            <span>人员信息</span>
          </div>
        </template>

        <div class="member-section">
          <div class="member-section-title">
            <el-icon><UserFilled /></el-icon> 项目管理人员
            <span class="member-section-tip">多角色人员可在不同分组中分别勾选，每次勾选对应一个角色</span>
          </div>
          <div class="member-role-groups">
            <div v-for="(group, role) in managerRoleGroups" :key="role" class="member-role-group">
              <div class="member-role-label">{{ role }}</div>
              <div class="member-role-items">
                <el-checkbox
                  v-for="slot in group"
                  :key="slot.key"
                  :model-value="selectedSlots.includes(slot.key)"
                  @change="toggleSlot(slot)"
                  border
                  class="member-checkbox"
                >
                  {{ slot.worker_name }}
                  <el-tag v-if="slot.cert" size="small" type="warning" style="margin-left:4px">{{ slot.cert }}</el-tag>
                </el-checkbox>
              </div>
            </div>
          </div>
        </div>

        <el-divider />

        <div class="member-section">
          <div class="member-section-title">
            <el-icon><Avatar /></el-icon> 施工人员
            <span class="member-section-tip">勾选需要的施工班组人员</span>
          </div>
          <div class="member-role-groups">
            <div v-for="(group, role) in workerRoleGroups" :key="role" class="member-role-group">
              <div class="member-role-label">{{ role }}</div>
              <div class="member-role-items">
                <el-checkbox
                  v-for="slot in group"
                  :key="slot.key"
                  :model-value="selectedSlots.includes(slot.key)"
                  @change="toggleSlot(slot)"
                  border
                  class="member-checkbox"
                >
                  {{ slot.worker_name }}
                  <el-tag v-if="slot.cert" size="small" type="warning" style="margin-left:4px">{{ slot.cert }}</el-tag>
                </el-checkbox>
              </div>
            </div>
          </div>
        </div>

        <el-divider />

      </el-card>

      <!-- 5. 勘察信息 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><Search /></el-icon>
            <span>勘察信息</span>
          </div>
        </template>
        <el-row :gutter="24">
          <el-col :span="12">
            <el-form-item label="勘察单位">
              <el-input v-model="form.survey_unit" placeholder="请输入勘察单位" clearable />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="勘察部门">
              <el-input v-model="form.survey_department" placeholder="请输入勘察部门" clearable />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="停电范围">
          <el-input
            v-model="form.power_off_range"
            type="textarea"
            :rows="3"
            placeholder="请输入停电范围"
          />
        </el-form-item>
        <el-form-item label="保留带电部位">
          <el-input
            v-model="form.live_parts"
            type="textarea"
            :rows="3"
            placeholder="请输入保留带电部位"
          />
        </el-form-item>
        <el-form-item label="作业现场危险点">
          <el-input
            v-model="form.danger_points"
            type="textarea"
            :rows="3"
            placeholder="请输入作业现场危险点"
          />
        </el-form-item>
        <el-form-item label="装设位置">
          <el-input
            v-model="form.safety_measures"
            type="textarea"
            :rows="4"
            placeholder="请输入接地线、绝缘隔板、遮栏、围栏、标示牌等装设位置"
          />
        </el-form-item>
      </el-card>

      <!-- 6. 施工机具清单 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><SetUp /></el-icon>
            <span>施工机具清单</span>
            <el-tag size="small" type="info" style="margin-left: 8px">根据所选工艺自动生成</el-tag>
          </div>
        </template>
        <el-table :data="form.equipment_list" border size="small" class="dynamic-table" empty-text="请先选择施工工艺，机具将自动生成">
          <el-table-column label="序号" width="60" align="center">
            <template #default="{ $index }">{{ $index + 1 }}</template>
          </el-table-column>
          <el-table-column label="来源工艺" width="140">
            <template #default="{ row }">
              <el-tag size="small" :type="row.source_category === '电气' ? 'success' : 'warning'">{{ row.source_process }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="器具名称" min-width="160">
            <template #default="{ row }">{{ row.name }}</template>
          </el-table-column>
          <el-table-column label="规格/描述" min-width="200">
            <template #default="{ row }">{{ row.description || '-' }}</template>
          </el-table-column>
          <el-table-column label="数量" width="120">
            <template #default="{ row }">{{ row.quantity || '-' }}</template>
          </el-table-column>
          <el-table-column label="操作" width="70" align="center">
            <template #default="{ $index }">
              <el-button type="danger" :icon="Delete" circle size="small" @click="form.equipment_list.splice($index, 1)" />
            </template>
          </el-table-column>
        </el-table>
      </el-card>

      <!-- 7. 危险源识别 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><Warning /></el-icon>
            <span>危险源识别</span>
            <el-tag size="small" type="info" style="margin-left: 8px">根据所选工艺自动生成</el-tag>
          </div>
        </template>
        <el-table :data="form.hazard_list" border size="small" class="dynamic-table" empty-text="请先选择施工工艺，危险源将自动生成">
          <el-table-column label="序号" width="60" align="center">
            <template #default="{ $index }">{{ $index + 1 }}</template>
          </el-table-column>
          <el-table-column label="来源工艺" width="140">
            <template #default="{ row }">
              <el-tag size="small" :type="row.category === '电气' ? 'success' : 'warning'">{{ row.process_name }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="危险源" min-width="200">
            <template #default="{ row }">{{ row.hazard }}</template>
          </el-table-column>
          <el-table-column label="安全措施" min-width="300">
            <template #default="{ row }">
              <div style="white-space: pre-line">{{ row.safety_measures }}</div>
            </template>
          </el-table-column>
        </el-table>
      </el-card>

      <!-- 8. 施工工艺选择 -->
      <el-card shadow="hover" class="section-card">
        <template #header>
          <div class="section-header">
            <el-icon><Operation /></el-icon>
            <span>施工工艺选择</span>
          </div>
        </template>
        <div class="process-header">
          <span class="process-hint">选择本工程涉及的施工工艺</span>
          <el-input
            v-model="processSearch"
            placeholder="搜索工艺名称"
            clearable
            style="width: 240px"
          >
            <template #prefix><el-icon><Search /></el-icon></template>
          </el-input>
        </div>

        <!-- 土建工艺 -->
        <div v-if="civilProcesses.length > 0" class="process-group">
          <div class="process-group-title">
            <el-tag type="warning" effect="dark" size="small">土建</el-tag>
            <span>土建施工工艺</span>
            <span class="process-group-count">已选 {{ selectedCivilCount }} 项</span>
          </div>
          <el-table :data="civilProcesses" border size="small" class="process-table" @row-click="(row) => toggleProcess(row.id)">
            <el-table-column width="50" align="center">
              <template #default="{ row }">
                <el-checkbox :model-value="selectedProcesses.includes(row.id)" @change="toggleProcess(row.id)" @click.stop />
              </template>
            </el-table-column>
            <el-table-column label="编号" prop="code" width="100" align="center" />
            <el-table-column label="工艺名称" prop="name" min-width="200" />
            <el-table-column label="施工流程" min-width="300">
              <template #default="{ row }">{{ formatFlowSteps(row.flow_steps) }}</template>
            </el-table-column>
            <el-table-column label="工期(天)" prop="duration_days" width="90" align="center" />
          </el-table>
        </div>

        <!-- 电气工艺 -->
        <div v-if="electricProcesses.length > 0" class="process-group">
          <div class="process-group-title">
            <el-tag type="success" effect="dark" size="small">电气</el-tag>
            <span>电气施工工艺</span>
            <span class="process-group-count">已选 {{ selectedElectricCount }} 项</span>
          </div>
          <el-table :data="electricProcesses" border size="small" class="process-table" @row-click="(row) => toggleProcess(row.id)">
            <el-table-column width="50" align="center">
              <template #default="{ row }">
                <el-checkbox :model-value="selectedProcesses.includes(row.id)" @change="toggleProcess(row.id)" @click.stop />
              </template>
            </el-table-column>
            <el-table-column label="编号" prop="code" width="100" align="center" />
            <el-table-column label="工艺名称" prop="name" min-width="200" />
            <el-table-column label="施工流程" min-width="300">
              <template #default="{ row }">{{ formatFlowSteps(row.flow_steps) }}</template>
            </el-table-column>
            <el-table-column label="工期(天)" prop="duration_days" width="90" align="center" />
          </el-table>
        </div>

        <el-empty v-if="civilProcesses.length === 0 && electricProcesses.length === 0" description="暂无施工工艺数据，可在首页导入" :image-size="60" />
        <div class="process-summary">
          已选择 <el-tag type="primary" size="small">{{ selectedProcesses.length }}</el-tag> 项施工工艺
        </div>
      </el-card>

      <!-- 提交按钮 -->
      <div class="form-actions">
        <el-button @click="handleReset">重置</el-button>
        <el-button type="primary" size="large" :loading="submitting" @click="handleSubmit">
          <el-icon><CircleCheck /></el-icon>提交工程
        </el-button>
      </div>
    </el-form>
  </div>
</template>

<script setup>
import { ref, reactive, computed, watch, onMounted, onUnmounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Delete, Plus, Warning } from '@element-plus/icons-vue'
import { createProject, getProcesses, getProjectTypes, getWorkers } from '../api'

const router = useRouter()

const submitting = ref(false)
const savingDraft = ref(false)
const formRef = ref(null)

// 表单数据
const form = reactive({
  project_code: '',
  project_name: '',
  project_type: [],
  location: '',
  voltage_level: '',
  line_name: '',
  company_name: '',
  work_task: '',
  subcontractor: '',
  subcontractor_civil: '',
  subcontractor_electric: '',
  start_date: '',
  end_date: '',
  description: '',
  survey_leader: '',
  survey_members: '',
  briefing_host: '',
  survey_unit: '',
  survey_department: '',
  survey_number: '',
  power_off_range: '',
  live_parts: '',
  danger_points: '',
  safety_measures: '',
  equipment_list: [],
  hazard_list: [],
  process_ids: []
})

// 项目管理人员列表
const managers = reactive([])
// 施工人员列表
const workers = reactive([])

// 人员数据库列表
const workerList = ref([])
const selectedSlots = ref([])

// 判断角色是否为管理类
const isManagerRole = (role) => {
  if (!role) return false
  const workerKeywords = ['施工班组', '继保', '电试', '电工', '电缆工', '试验工', '试验员']
  return !workerKeywords.some(kw => role.includes(kw))
}

// 获取工人的所有角色-证书配对槽位
const getWorkerSlots = (w) => {
  const slots = []
  const pairs = [
    { role: w.role, cert: w.certification, index: 1 },
    { role: w.role2, cert: w.certification2, index: 2 },
    { role: w.role3, cert: w.certification3, index: 3 },
  ]
  for (const p of pairs) {
    if (p.role) {
      slots.push({
        key: `${w.id}_${p.index}`,
        role: p.role,
        cert: p.cert || '',
        worker_id: w.id,
        worker_name: w.name,
        role_index: p.index,
        is_manager: isManagerRole(p.role),
      })
    }
  }
  return slots
}

// 所有人员的所有槽位
const allSlots = computed(() => {
  const result = []
  for (const w of workerList.value) {
    result.push(...getWorkerSlots(w))
  }
  return result
})

// 管理人员按角色分组
const managerRoleGroups = computed(() => {
  const groups = {}
  allSlots.value.filter(s => s.is_manager).forEach(slot => {
    const role = slot.role
    if (!groups[role]) groups[role] = []
    groups[role].push(slot)
  })
  return groups
})

// 施工人员按角色分组
const workerRoleGroups = computed(() => {
  const groups = {}
  allSlots.value.filter(s => !s.is_manager).forEach(slot => {
    const role = slot.role
    if (!groups[role]) groups[role] = []
    groups[role].push(slot)
  })
  return groups
})

// 切换槽位选择
const toggleSlot = (slot) => {
  const idx = selectedSlots.value.indexOf(slot.key)
  if (idx >= 0) {
    selectedSlots.value.splice(idx, 1)
  } else {
    selectedSlots.value.push(slot.key)
  }
  // 重建 managers 和 workers 列表
  rebuildMembers()
}

// 从选中的 slot 重建 managers/workers
const rebuildMembers = () => {
  managers.splice(0, managers.length)
  workers.splice(0, workers.length)
  for (const slotKey of selectedSlots.value) {
    const slot = allSlots.value.find(s => s.key === slotKey)
    if (!slot) continue
    const member = {
      name: slot.worker_name,
      role: slot.role,
      cert: slot.cert,
      member_type: slot.is_manager ? 'manager' : 'worker',
    }
    if (slot.is_manager) {
      managers.push(member)
    } else {
      workers.push(member)
    }
  }
}

const loadWorkerList = async () => {
  try {
    const res = await getWorkers()
    workerList.value = res.items || res || []
  } catch (error) {
    console.error('加载人员列表失败：', error)
  }
}

// 工程部位类别选项
const projectTypeOptions = ref([])

// 施工工艺
const processList = ref([])
const selectedProcesses = ref([])
const processSearch = ref('')

// 草稿标记
const hasDraft = ref(false)

// 计算工期天数
const dayCount = computed(() => {
  if (form.start_date && form.end_date) {
    const start = new Date(form.start_date)
    const end = new Date(form.end_date)
    const diff = Math.ceil((end - start) / (1000 * 60 * 60 * 24))
    return diff > 0 ? diff : 0
  }
  return 0
})

// 按数据库 code 排序
const sortedProcessList = computed(() => {
  return [...processList.value].sort((a, b) => {
    if (a.code && b.code) return a.code.localeCompare(b.code)
    return (a.id || 0) - (b.id || 0)
  })
})

// 筛选后的施工工艺
const filteredProcesses = computed(() => {
  const list = sortedProcessList.value
  if (!processSearch.value) return list
  const kw = processSearch.value.toLowerCase()
  return list.filter(p =>
    p.name.toLowerCase().includes(kw) || (p.category || '').toLowerCase().includes(kw) || (p.code || '').toLowerCase().includes(kw)
  )
})

// 土建工艺
const civilProcesses = computed(() => {
  return filteredProcesses.value.filter(p => p.category === '土建')
})

// 电气工艺
const electricProcesses = computed(() => {
  return filteredProcesses.value.filter(p => p.category === '电气')
})

// 已选土建工艺数
const selectedCivilCount = computed(() => {
  return civilProcesses.value.filter(p => selectedProcesses.value.includes(p.id)).length
})

// 已选电气工艺数
const selectedElectricCount = computed(() => {
  return electricProcesses.value.filter(p => selectedProcesses.value.includes(p.id)).length
})

// 验证规则
const rules = {
  project_code: [{ required: true, message: '请输入项目编号', trigger: 'blur' }],
  project_name: [{ required: true, message: '请输入项目名称', trigger: 'blur' }],
  project_type: [{ required: true, type: 'array', min: 1, message: '请选择工程类别', trigger: 'change' }],
  location: [{ required: true, message: '请输入工程地点', trigger: 'blur' }],
  subcontractor_civil: [{ required: true, message: '请输入土建分包单位', trigger: 'blur' }],
  start_date: [{ required: true, message: '请选择开工日期', trigger: 'change' }],
  end_date: [{ required: true, message: '请选择竣工日期', trigger: 'change' }],
  description: [{ required: true, message: '请输入工程概况', trigger: 'blur' }]
}

// 格式化工艺流程
const formatFlowSteps = (flowSteps) => {
  if (!flowSteps) return ''
  try {
    const arr = typeof flowSteps === 'string' ? JSON.parse(flowSteps) : flowSteps
    return arr.join(' → ')
  } catch {
    return flowSteps
  }
}

// 切换工艺选择
const toggleProcess = (id) => {
  const idx = selectedProcesses.value.indexOf(id)
  if (idx >= 0) {
    selectedProcesses.value.splice(idx, 1)
  } else {
    selectedProcesses.value.push(id)
  }
}

// 当选择的工艺变化时，自动更新机具清单和危险源识别
watch(selectedProcesses, (newVal) => {
  // 更新机具清单
  const equipmentMap = new Map()
  newVal.forEach(pid => {
    const p = processList.value.find(x => x.id === pid)
    if (p && p.equipment) {
      const items = p.equipment.split(/[、,，]/).filter(Boolean)
      items.forEach(item => {
        const trimmed = item.trim()
        const qtyMatch = trimmed.match(/(\d+)\s*(台|把|套|只|个|根|块|张|米|m)/)
        const qty = qtyMatch ? qtyMatch[0] : ''
        const name = trimmed.replace(/\d+\s*(台|把|套|只|个|根|块|张|米|m)/g, '').trim() || trimmed
        if (!equipmentMap.has(name)) {
          equipmentMap.set(name, {
            name,
            description: trimmed,
            quantity: qty,
            source_process: p.name,
            source_category: p.category
          })
        }
      })
    }
  })
  form.equipment_list = Array.from(equipmentMap.values())

  // 更新危险源识别
  const hazards = []
  newVal.forEach(pid => {
    const p = processList.value.find(x => x.id === pid)
    if (p && p.hazards) {
      const hazardItems = p.hazards.split(/[、,，]/).filter(Boolean)
      hazardItems.forEach(h => {
        hazards.push({
          process_name: p.name,
          category: p.category,
          hazard: h.trim(),
          safety_measures: p.safety_measures || ''
        })
      })
    }
  })
  form.hazard_list = hazards
}, { deep: true })

// 保存草稿
const handleSaveDraft = async () => {
  savingDraft.value = true
  try {
    const draft = {
      form: { ...form },
      managers: managers.map(m => ({ ...m })),
      workers: workers.map(w => ({ ...w })),
      selectedProcesses: [...selectedProcesses.value],
      selectedSlots: [...selectedSlots.value],
      savedAt: new Date().toLocaleString()
    }
    localStorage.setItem('project_draft', JSON.stringify(draft))
    hasDraft.value = true
    ElMessage.success('草稿已保存')
  } catch (error) {
    ElMessage.error('保存草稿失败')
  } finally {
    savingDraft.value = false
  }
}

// 恢复草稿
const handleLoadDraft = async () => {
  try {
    await ElMessageBox.confirm('恢复草稿将覆盖当前填写的内容，是否继续？', '提示', {
      confirmButtonText: '确定恢复',
      cancelButtonText: '取消',
      type: 'warning'
    })
    const draftStr = localStorage.getItem('project_draft')
    if (draftStr) {
      const draft = JSON.parse(draftStr)
      Object.assign(form, draft.form)
      // project_type 在草稿中可能是字符串或数组，统一转为数组
      if (typeof form.project_type === 'string') {
        form.project_type = form.project_type ? form.project_type.split(',').filter(Boolean) : []
      }
      managers.splice(0, managers.length, ...draft.managers)
      workers.splice(0, workers.length, ...draft.workers)
      selectedProcesses.value = draft.selectedProcesses || []
      selectedSlots.value = draft.selectedSlots || []
      ElMessage.success(`已恢复草稿（保存于 ${draft.savedAt}）`)
    }
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('恢复草稿失败')
    }
  }
}

// 重置表单
const handleReset = () => {
  ElMessageBox.confirm('确定要重置所有已填写的内容吗？', '提示', {
    confirmButtonText: '确定',
    cancelButtonText: '取消',
    type: 'warning'
  }).then(() => {
    formRef.value?.resetFields()
    managers.splice(0, managers.length)
    workers.splice(0, workers.length)
    selectedProcesses.value = []
    selectedSlots.value = []
    form.equipment_list.splice(0, form.equipment_list.length)
    form.hazard_list.splice(0, form.hazard_list.length)
  }).catch(() => {})
}

// 提交表单
const handleSubmit = async () => {
  if (!formRef.value) return
  try {
    await formRef.value.validate()
  } catch {
    ElMessage.warning('请完善必填项')
    return
  }
  submitting.value = true
  try {
    const members = [...managers, ...workers].filter(m => m.name.trim())
    const data = {
      ...form,
      project_type: form.project_type.join(','),
      members: members.map(m => ({ name: m.name, member_type: m.member_type, role: m.role, cert: m.cert })),
      process_ids: selectedProcesses.value,
      equipment_list: JSON.stringify(form.equipment_list),
      quality_control: JSON.stringify(form.hazard_list),
      approvals: []
    }
    await createProject(data)
    localStorage.removeItem('project_draft')
    hasDraft.value = false
    ElMessage.success('工程创建成功！')
    router.push('/project/list')
  } catch (error) {
    ElMessage.error('创建失败，请重试')
    console.error('创建工程失败：', error)
  } finally {
    submitting.value = false
  }
}

// 加载工程部位类别列表
const loadProjectTypes = async () => {
  try {
    const res = await getProjectTypes()
    projectTypeOptions.value = res.items || []
  } catch (error) {
    console.error('加载工程类别失败：', error)
  }
}

// 加载施工工艺列表
const loadProcessList = async () => {
  try {
    const res = await getProcesses()
    processList.value = res.items || []
  } catch (error) {
    console.error('加载施工工艺失败：', error)
  }
}

// 检查是否有草稿
const checkDraft = () => {
  hasDraft.value = !!localStorage.getItem('project_draft')
}

// 页面离开前提示
const handleBeforeUnload = (e) => {
  if (form.project_code || form.project_name) {
    e.preventDefault()
    e.returnValue = ''
  }
}

onMounted(() => {
  checkDraft()
  loadProjectTypes()
  loadProcessList()
  loadWorkerList()
  window.addEventListener('beforeunload', handleBeforeUnload)
})

onUnmounted(() => {
  window.removeEventListener('beforeunload', handleBeforeUnload)
})
</script>

<style scoped>
.project-form {
  max-width: 960px;
}

.page-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 20px;
  padding: 16px 20px;
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 1px 4px rgba(0, 21, 41, 0.08);
}

.header-left {
  display: flex;
  align-items: center;
  gap: 12px;
}

.header-title {
  font-size: 20px;
  font-weight: 600;
  color: #303133;
  margin: 0;
}

.header-desc {
  font-size: 13px;
  color: #909399;
  margin: 4px 0 0;
}

.header-right {
  display: flex;
  gap: 8px;
}

/* 分区卡片 */
.section-card {
  margin-bottom: 16px;
  border-radius: 8px;
}

.section-card :deep(.el-card__header) {
  padding: 12px 20px;
  background: #f5f7fa;
}

.section-header {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 15px;
  font-weight: 600;
  color: #303133;
}

.section-card :deep(.el-card__body) {
  padding: 20px 24px;
}

/* 人员列表 */
.member-section {
  margin-bottom: 8px;
}

.member-section-title {
  font-size: 14px;
  font-weight: 600;
  color: #606266;
  margin-bottom: 12px;
  display: flex;
  align-items: center;
  gap: 4px;
}

.selected-tags {
  margin-top: 8px;
  display: flex;
  flex-wrap: wrap;
}

/* 人员分组框体 */
.member-section-tip {
  font-size: 12px;
  color: #909399;
  font-weight: 400;
  margin-left: 8px;
}

.member-role-groups {
  display: flex;
  flex-direction: column;
  gap: 12px;
  padding: 8px 0;
}

.member-role-group {
  display: flex;
  align-items: flex-start;
  gap: 12px;
  padding: 10px 12px;
  background: #fafbfc;
  border: 1px solid #ebeef5;
  border-radius: 6px;
}

.member-role-label {
  min-width: 96px;
  padding: 4px 10px;
  font-size: 13px;
  font-weight: 600;
  color: #409eff;
  background: #ecf5ff;
  border-radius: 4px;
  text-align: center;
  white-space: nowrap;
}

.member-role-items {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  flex: 1;
}

.member-checkbox {
  margin-right: 0 !important;
}

.member-list {
  padding: 0 0 8px;
}

.member-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 10px;
}

.member-index {
  min-width: 28px;
  text-align: center;
  font-weight: 600;
}

/* 动态表格 */
.dynamic-table {
  width: 100%;
}

.dynamic-table :deep(.el-input__wrapper) {
  box-shadow: none;
}

.dynamic-table :deep(.el-input__wrapper:hover),
.dynamic-table :deep(.el-input__wrapper:focus-within) {
  box-shadow: 0 0 0 1px var(--el-input-hover-border-color) inset;
}

.table-add-btn {
  margin-top: 12px;
  text-align: center;
}

/* 施工工艺选择 */
.process-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
}

.process-hint {
  font-size: 14px;
  color: #606266;
}

.process-group {
  margin-bottom: 16px;
}

.process-group-title {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 10px;
  font-size: 14px;
  font-weight: 600;
  color: #303133;
}

.process-group-count {
  font-size: 12px;
  color: #909399;
  font-weight: 400;
}

.process-table {
  width: 100%;
}

.process-table :deep(.el-table__row) {
  cursor: pointer;
}

.process-table :deep(.el-table__row:hover) {
  background: #ecf5ff;
}

.process-summary {
  margin-top: 12px;
  padding: 8px 14px;
  background: #f5f7fa;
  border-radius: 6px;
  font-size: 14px;
  color: #606266;
  display: flex;
  align-items: center;
  gap: 4px;
}

/* 提交按钮 */
.form-actions {
  display: flex;
  justify-content: center;
  gap: 16px;
  padding: 24px 0 8px;
}
</style>
