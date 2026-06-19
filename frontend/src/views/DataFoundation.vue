<template>
  <div class="data-foundation">
    <!-- 文档模板管理 -->
    <el-card shadow="hover" class="section-card">
      <template #header>
        <span class="section-title">文档模板管理</span>
      </template>
      <el-row :gutter="16">
        <el-col :span="8" v-for="tpl in templates" :key="tpl.doc_type">
          <el-card shadow="hover" class="template-card">
            <div class="template-header">
              <span class="template-label">{{ tpl.label }}</span>
              <el-tag v-if="tpl.file_name" type="success" size="small">已上传</el-tag>
              <el-tag v-else type="info" size="small">未上传</el-tag>
            </div>
            <div class="template-info" v-if="tpl.file_name">
              <div class="info-row">
                <span class="info-label">文件名：</span>
                <span class="info-value">{{ tpl.file_name }}</span>
              </div>
              <div class="info-row" v-if="tpl.uploaded_at">
                <span class="info-label">上传时间：</span>
                <span class="info-value">{{ tpl.uploaded_at }}</span>
              </div>
            </div>
            <div class="template-info" v-else>
              <div class="no-template">暂未上传模板</div>
            </div>
            <div class="template-actions">
              <el-upload
                :show-file-list="false"
                :accept="getAccept(tpl.doc_type)"
                :before-upload="(file) => beforeUploadTemplate(file, tpl.doc_type)"
                :http-request="() => {}"
              >
                <el-button type="primary" size="small" :loading="uploading[tpl.doc_type]">
                  {{ tpl.file_name ? '重新上传' : '上传模板' }}
                </el-button>
              </el-upload>
              <el-button
                v-if="tpl.file_name"
                type="success"
                size="small"
                plain
                @click="handleDownloadTemplate(tpl)"
              >
                下载
              </el-button>
              <el-button
                v-if="tpl.file_name"
                type="danger"
                size="small"
                plain
                @click="handleDeleteTemplate(tpl)"
              >
                删除
              </el-button>
            </div>
          </el-card>
        </el-col>
      </el-row>
    </el-card>

    <!-- 施工工艺数据管理 -->
    <el-card shadow="hover" class="section-card">
      <template #header>
        <div class="section-header">
          <span class="section-title">施工工艺数据管理</span>
          <div class="section-header-right">
            <el-tag type="info" size="small">已入库 {{ processCount }} 条</el-tag>
          </div>
        </div>
      </template>

      <!-- 操作栏 -->
      <el-row :gutter="20" align="middle">
        <el-col :span="16">
          <div class="action-bar">
            <el-upload
              :show-file-list="false"
              accept=".xlsx,.xls,.csv"
              :before-upload="beforeUploadProcess"
              :http-request="() => {}"
              :disabled="importingProcess"
            >
              <el-button type="primary" :loading="importingProcess" size="small">
                <el-icon><Upload /></el-icon>
                {{ importingProcess ? '导入中...' : '导入施工工艺' }}
              </el-button>
            </el-upload>
            <el-button type="success" plain size="small" @click="handleExportProcesses">
              <el-icon><Download /></el-icon>
              导出施工工艺
            </el-button>
            <el-button type="danger" plain size="small" @click="handleDeleteAllProcesses">
              <el-icon><Delete /></el-icon>
              全部删除
            </el-button>
            <span class="import-hint">支持 .xlsx / .xls / .csv，导入后会弹出预览/覆盖确认</span>
          </div>
        </el-col>
      </el-row>

      <el-alert
        v-if="processImportResult"
        :title="processImportResult.message"
        :type="processImportResult.errors && processImportResult.errors.length > 0 ? 'warning' : 'success'"
        show-icon
        closable
        style="margin-top: 12px"
        @close="processImportResult = null"
      />

      <!-- 工艺表格 -->
      <el-table :data="processList" stripe style="width: 100%; margin-top: 16px" max-height="500">
        <el-table-column prop="code" label="编号" width="90" sortable />
        <el-table-column prop="name" label="工艺名称" min-width="150" sortable />
        <el-table-column prop="project_type" label="工程部位" width="130" sortable>
          <template #default="{ row }">{{ row.project_type || '-' }}</template>
        </el-table-column>
        <el-table-column prop="category" label="工艺分类" width="80" sortable />
        <el-table-column prop="sub_category" label="工序等级" width="80" sortable />
        <el-table-column prop="duration_days" label="施工天数" width="90" sortable align="center">
          <template #default="{ row }">{{ row.duration_days != null ? row.duration_days + ' 天' : '-' }}</template>
        </el-table-column>
        <el-table-column label="施工机具" show-overflow-tooltip width="160">
          <template #default="{ row }">{{ row.equipment || '-' }}</template>
        </el-table-column>
        <el-table-column label="危险源" show-overflow-tooltip width="160">
          <template #default="{ row }">{{ row.hazards || '-' }}</template>
        </el-table-column>
        <el-table-column label="操作" width="140" fixed="right">
          <template #default="{ row }">
            <el-button type="primary" link size="small" @click="handleEditProcess(row)">编辑</el-button>
            <el-button type="danger" link size="small" @click="handleDeleteProcess(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <!-- 编辑工艺对话框 -->
    <el-dialog
      v-model="editDialogVisible"
      :title="'编辑施工工艺：' + editForm.name"
      width="800px"
      :close-on-click-modal="false"
      destroy-on-close
    >
      <el-form :model="editForm" label-width="90px" label-position="top">
        <el-row :gutter="16">
          <el-col :span="3">
            <el-form-item label="编号">
              <el-input v-model="editForm.code" disabled placeholder="自动生成" />
            </el-form-item>
          </el-col>
          <el-col :span="9">
            <el-form-item label="工艺名称">
              <el-input v-model="editForm.name" />
            </el-form-item>
          </el-col>
          <el-col :span="6">
            <el-form-item label="工程部位">
              <el-input v-model="editForm.project_type" placeholder="配电房/电缆通道排管/电缆井" />
            </el-form-item>
          </el-col>
          <el-col :span="3">
            <el-form-item label="分类">
              <el-select v-model="editForm.category" placeholder="选择">
                <el-option label="土建" value="土建" />
                <el-option label="电气" value="电气" />
              </el-select>
            </el-form-item>
          </el-col>
          <el-col :span="3">
            <el-form-item label="等级">
              <el-select v-model="editForm.sub_category" placeholder="选择">
                <el-option label="一般" value="一般" />
                <el-option label="特殊" value="特殊（需专项施工方案）" />
              </el-select>
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="施工标准">
          <el-input v-model="editForm.standards" type="textarea" :rows="3" placeholder="施工标准、规范依据等" />
        </el-form-item>
        <el-form-item label="施工工序">
          <el-input v-model="editForm.flow_steps_display" type="textarea" :rows="3" placeholder="用 → 分隔各步骤，如：测量放线 → 机械开挖 → 验槽" />
          <span class="form-hint">各步骤之间用 → 分隔，保存后自动转为 JSON 存储</span>
        </el-form-item>
        <el-form-item label="施工所需天数">
          <el-input-number v-model="editForm.duration_days" :min="0" :step="1" placeholder="如：7" style="width: 200px" />
          <span class="form-hint" style="margin-left: 12px">单位：天，0 表示不填</span>
        </el-form-item>
        <el-form-item label="施工机具">
          <el-input v-model="editForm.equipment" type="textarea" :rows="2" placeholder="设备名称、型号及数量" />
        </el-form-item>
        <el-form-item label="危险源识别">
          <el-input v-model="editForm.hazards" type="textarea" :rows="2" placeholder="可能的危险源" />
        </el-form-item>
        <el-form-item label="安全措施">
          <el-input v-model="editForm.safety_measures" type="textarea" :rows="3" placeholder="应采取的安全措施" />
        </el-form-item>
        <el-form-item label="工艺说明">
          <el-input v-model="editForm.description" type="textarea" :rows="2" placeholder="补充说明" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="editDialogVisible = false">取消</el-button>
        <el-button type="primary" :loading="savingProcess" @click="handleSaveProcess">保存</el-button>
      </template>
    </el-dialog>

    <!-- 工艺导入预览/确认覆盖对话框 -->
    <el-dialog
      v-model="importPreviewVisible"
      title="工艺导入预览"
      width="900px"
      :close-on-click-modal="false"
      destroy-on-close
    >
      <div v-if="importPreviewData">
        <el-row :gutter="16" style="margin-bottom: 12px">
          <el-col :span="12">
            <el-statistic title="新增编号" :value="importPreviewData.new_count" :value-style="{ color: '#67c23a' }" />
          </el-col>
          <el-col :span="12">
            <el-statistic title="冲突编号（已存在）" :value="importPreviewData.overwrite_count" :value-style="{ color: '#e6a23c' }" />
          </el-col>
        </el-row>

        <el-alert
          v-if="importPreviewData.overwrite_count > 0"
          type="warning"
          :closable="false"
          show-icon
          style="margin-bottom: 12px"
        >
          <template #title>
            检测到 {{ importPreviewData.overwrite_count }} 个编号已存在，请勾选需要覆盖的项
          </template>
        </el-alert>

        <el-table
          v-if="importPreviewData.overwrite_items.length > 0"
          :data="importPreviewData.overwrite_items"
          border
          max-height="300"
          @selection-change="onOverwriteSelectionChange"
        >
          <el-table-column type="selection" width="55" />
          <el-table-column prop="code" label="编号" width="100" />
          <el-table-column prop="name" label="文件中的工艺名称" min-width="180" />
          <el-table-column prop="existing_name" label="已存在工艺名称" min-width="180" />
          <el-table-column prop="category" label="分类" width="80" />
          <el-table-column prop="sub_category" label="等级" width="160" />
        </el-table>

        <el-divider v-if="importPreviewData.new_items.length > 0">新增清单（将自动导入）</el-divider>
        <el-table
          v-if="importPreviewData.new_items.length > 0"
          :data="importPreviewData.new_items"
          border
          max-height="200"
        >
          <el-table-column prop="code" label="编号" width="100" />
          <el-table-column prop="name" label="工艺名称" min-width="180" />
          <el-table-column prop="category" label="分类" width="80" />
          <el-table-column prop="sub_category" label="等级" width="160" />
        </el-table>
      </div>

      <template #footer>
        <el-button @click="cancelImportPreview">取消</el-button>
        <el-button
          type="primary"
          :loading="confirmingImport"
          :disabled="importPreviewData && importPreviewData.new_count === 0 && selectedOverwriteCodes.length === 0"
          @click="handleConfirmImport"
        >
          确认导入（覆盖 {{ selectedOverwriteCodes.length }} 项，新增 {{ importPreviewData ? importPreviewData.new_count : 0 }} 项）
        </el-button>
      </template>
    </el-dialog>

    <!-- 施工人员数据管理 -->
    <el-card shadow="hover" class="section-card">
      <template #header>
        <div class="section-header">
          <span class="section-title">施工人员数据管理</span>
          <el-tag type="info" size="small">已入库 {{ workerCount }} 条</el-tag>
        </div>
      </template>
      <el-row :gutter="20" align="middle">
        <el-col :span="12">
          <div class="import-area">
            <el-upload
              :show-file-list="false"
              accept=".xlsx,.xls,.csv"
              :before-upload="beforeUploadWorker"
              :http-request="() => {}"
              :disabled="importingWorker"
            >
              <el-button type="primary" :loading="importingWorker">
                <el-icon><Upload /></el-icon>
                {{ importingWorker ? '导入中...' : '导入施工人员' }}
              </el-button>
            </el-upload>
            <el-button
              type="success"
              :loading="exportingWorker"
              @click="handleExportWorkers"
              style="margin-left: 8px"
            >
              <el-icon><Download /></el-icon>
              {{ exportingWorker ? '导出中...' : '导出施工人员' }}
            </el-button>
            <span class="import-hint">支持 .xlsx / .xls / .csv 格式</span>
            <div class="form-hint">
              <el-checkbox v-model="workerOverwriteMode">覆盖已存在的人员（按姓名匹配）</el-checkbox>
            </div>
          </div>
        </el-col>
        <el-col :span="12">
          <div class="import-help">
            <div class="help-title">文件格式要求：</div>
            <div class="help-text">第一行为表头，必须包含"姓名"列</div>
            <div class="help-text">可选列：角色/职务、角色/职务2、角色/职务3、所属班组、资质证书、资质证书2、资质证书3</div>
            <div class="help-text">每人最多3个角色和3本证书，勾选"覆盖"后按姓名匹配更新</div>
          </div>
        </el-col>
      </el-row>

      <el-alert
        v-if="workerImportResult"
        :title="workerImportResult.message"
        :type="workerImportResult.errors && workerImportResult.errors.length > 0 ? 'warning' : 'success'"
        :description="workerImportResult.errors && workerImportResult.errors.length > 0 ? workerImportResult.errors.join('；') : ''"
        show-icon
        closable
        style="margin-top: 12px"
        @close="workerImportResult = null"
      />

      <el-table :data="workerList" stripe style="width: 100%; margin-top: 16px" max-height="400" v-if="workerList.length > 0">
        <el-table-column prop="name" label="姓名" width="100" />
        <el-table-column label="角色1 / 证书1" min-width="160">
          <template #default="{ row }">
            <span v-if="row.role">{{ row.role }}<el-tag v-if="row.certification" size="small" type="warning" style="margin-left:4px">{{ row.certification }}</el-tag></span>
            <span v-else>-</span>
          </template>
        </el-table-column>
        <el-table-column label="角色2 / 证书2" min-width="160">
          <template #default="{ row }">
            <span v-if="row.role2">{{ row.role2 }}<el-tag v-if="row.certification2" size="small" type="warning" style="margin-left:4px">{{ row.certification2 }}</el-tag></span>
            <span v-else>-</span>
          </template>
        </el-table-column>
        <el-table-column label="角色3 / 证书3" min-width="160">
          <template #default="{ row }">
            <span v-if="row.role3">{{ row.role3 }}<el-tag v-if="row.certification3" size="small" type="warning" style="margin-left:4px">{{ row.certification3 }}</el-tag></span>
            <span v-else>-</span>
          </template>
        </el-table-column>
        <el-table-column prop="team" label="所属班组" width="120" />
        <el-table-column label="操作" width="120" fixed="right">
          <template #default="{ row }">
            <el-button type="primary" link size="small" @click="handleEditWorker(row)">编辑</el-button>
            <el-button type="danger" link size="small" @click="handleDeleteWorker(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <!-- 编辑施工人员对话框 -->
    <el-dialog
      v-model="workerEditVisible"
      :title="workerEditForm.id ? '编辑施工人员：' + workerEditForm.name : '新增施工人员'"
      width="600px"
      :close-on-click-modal="false"
      destroy-on-close
    >
      <el-form :model="workerEditForm" label-width="90px">
        <el-form-item label="姓名" required>
          <el-input v-model="workerEditForm.name" placeholder="请输入姓名" />
        </el-form-item>
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="角色/职务1">
              <el-input v-model="workerEditForm.role" placeholder="如：项目经理" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="资质证书1">
              <el-input v-model="workerEditForm.certification" placeholder="对应角色1的证书" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="角色/职务2">
              <el-input v-model="workerEditForm.role2" placeholder="如：技术负责人" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="资质证书2">
              <el-input v-model="workerEditForm.certification2" placeholder="对应角色2的证书" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="角色/职务3">
              <el-input v-model="workerEditForm.role3" placeholder="如：安全员" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="资质证书3">
              <el-input v-model="workerEditForm.certification3" placeholder="对应角色3的证书" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="所属班组">
          <el-input v-model="workerEditForm.team" placeholder="请输入所属班组" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="workerEditVisible = false">取消</el-button>
        <el-button type="primary" :loading="savingWorker" @click="handleSaveWorker">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import {
  getTemplates, uploadTemplate, deleteTemplate,
  getProcesses, updateProcess, importProcesses, exportProcesses,
  previewImportProcesses, confirmImportProcesses,
  getWorkers, importWorkers, exportWorkers
} from '../api'

// ---- 文档模板 ----
const templates = ref([])
const uploading = reactive({})

const getAccept = (docType) => {
  const acceptMap = {
    construction_design: '.doc,.docx',
    survey: '.doc,.docx',
    tech_briefing: '.doc,.docx',
    safety_briefing: '.doc,.docx',
    gantt_chart: '.xlsx,.xls'
  }
  return acceptMap[docType] || '.doc,.docx'
}

const loadTemplates = async () => {
  try {
    const res = await getTemplates()
    templates.value = res.items || []
  } catch (error) {
    console.error('加载模板列表失败：', error)
  }
}

const beforeUploadTemplate = async (file, docType) => {
  const acceptStr = getAccept(docType)
  const allowedExts = acceptStr.split(',')
  const fileExt = '.' + file.name.split('.').pop().toLowerCase()
  if (!allowedExts.includes(fileExt)) {
    ElMessage.error(`仅支持 ${acceptStr} 格式的模板文件`)
    return false
  }
  uploading[docType] = true
  try {
    await uploadTemplate(docType, file)
    ElMessage.success('模板上传成功')
    loadTemplates()
  } catch (error) {
    ElMessage.error('模板上传失败')
    console.error('上传模板失败：', error)
  } finally {
    uploading[docType] = false
  }
  return false
}

const handleDownloadTemplate = (tpl) => {
  window.open(`/api/templates/download/${tpl.doc_type}`, '_blank')
}

const handleDeleteTemplate = async (tpl) => {
  try {
    await ElMessageBox.confirm(`确定删除"${tpl.label}"模板吗？`, '提示', {
      confirmButtonText: '确定',
      cancelButtonText: '取消',
      type: 'warning',
    })
    await deleteTemplate(tpl.doc_type)
    ElMessage.success('模板删除成功')
    loadTemplates()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('模板删除失败')
      console.error('删除模板失败：', error)
    }
  }
}

// ---- 施工工艺 ----
const processList = ref([])
const processCount = ref(0)
const importingProcess = ref(false)
const processImportResult = ref(null)

// 编辑对话框
const editDialogVisible = ref(false)
const savingProcess = ref(false)
const editForm = reactive({
  id: null,
  code: '',
  name: '',
  project_type: '',
  category: '',
  sub_category: '',
  standards: '',
  flow_steps_display: '',
  duration_days: null,
  equipment: '',
  hazards: '',
  safety_measures: '',
  description: '',
})

// 工艺导入预览相关
const importPreviewVisible = ref(false)
const importPreviewData = ref(null)
const confirmingImport = ref(false)
const selectedOverwriteCodes = ref([])

const loadProcessList = async () => {
  try {
    const res = await getProcesses()
    processList.value = res.items || []
    processCount.value = res.total || processList.value.length
  } catch (error) {
    console.error('加载工艺列表失败：', error)
  }
}

const formatFlowStepsDisplay = (flowSteps) => {
  if (!flowSteps) return ''
  try {
    const arr = typeof flowSteps === 'string' ? JSON.parse(flowSteps) : flowSteps
    return arr.join(' → ')
  } catch {
    return flowSteps
  }
}

const handleEditProcess = (row) => {
  editForm.id = row.id
  editForm.code = row.code || ''
  editForm.name = row.name || ''
  editForm.project_type = row.project_type || ''
  editForm.category = row.category || ''
  editForm.sub_category = row.sub_category || ''
  editForm.standards = row.standards || ''
  editForm.flow_steps_display = formatFlowStepsDisplay(row.flow_steps)
  editForm.duration_days = row.duration_days != null ? row.duration_days : null
  editForm.equipment = row.equipment || ''
  editForm.hazards = row.hazards || ''
  editForm.safety_measures = row.safety_measures || ''
  editForm.description = row.description || ''
  editDialogVisible.value = true
}

const handleSaveProcess = async () => {
  if (!editForm.name.trim()) {
    ElMessage.warning('工艺名称不能为空')
    return
  }
  savingProcess.value = true
  try {
    // 将 flow_steps_display 转为 JSON 数组
    let flowSteps = null
    if (editForm.flow_steps_display.trim()) {
      const steps = editForm.flow_steps_display.split('→').map(s => s.trim()).filter(s => s)
      if (steps.length > 0) {
        flowSteps = JSON.stringify(steps)
      }
    }

    await updateProcess(editForm.id, {
      name: editForm.name.trim(),
      project_type: editForm.project_type.trim() || null,
      category: editForm.category || null,
      sub_category: editForm.sub_category || null,
      standards: editForm.standards.trim() || null,
      flow_steps: flowSteps,
      duration_days: editForm.duration_days,
      equipment: editForm.equipment.trim() || null,
      hazards: editForm.hazards.trim() || null,
      safety_measures: editForm.safety_measures.trim() || null,
      description: editForm.description.trim() || null,
    })
    ElMessage.success(`工艺"${editForm.name}"修改成功`)
    editDialogVisible.value = false
    loadProcessList()
  } catch (error) {
    ElMessage.error('保存失败')
    console.error('保存工艺失败：', error)
  } finally {
    savingProcess.value = false
  }
}

const beforeUploadProcess = async (file) => {
  const allowedExtensions = ['.xlsx', '.xls', '.csv']
  const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase()
  if (!allowedExtensions.includes(fileExt)) {
    ElMessage.error('仅支持 .xlsx, .xls, .csv 格式')
    return false
  }
  importingProcess.value = true
  processImportResult.value = null
  try {
    const res = await previewImportProcesses(file)
    importPreviewData.value = res
    importPreviewVisible.value = true
    // 默认勾选所有冲突项（用户可在弹窗中取消勾选）
    selectedOverwriteCodes.value = (res.overwrite_items || []).map(item => item.code)
  } catch (error) {
    ElMessage.error('文件解析失败')
    console.error('解析工艺文件失败：', error)
  } finally {
    importingProcess.value = false
  }
  return false
}

const onOverwriteSelectionChange = (selection) => {
  selectedOverwriteCodes.value = selection.map(item => item.code)
}

const handleConfirmImport = async () => {
  if (!importPreviewData.value) return
  confirmingImport.value = true
  try {
    const res = await confirmImportProcesses(
      importPreviewData.value.preview_id,
      selectedOverwriteCodes.value
    )
    importPreviewVisible.value = false
    processImportResult.value = res
    ElMessage.success(res.message || '导入完成')
    loadProcessList()
  } catch (error) {
    ElMessage.error('导入失败')
    console.error('确认工艺导入失败：', error)
  } finally {
    confirmingImport.value = false
  }
}

const cancelImportPreview = () => {
  importPreviewVisible.value = false
  importPreviewData.value = null
  selectedOverwriteCodes.value = []
}

const handleExportProcesses = async () => {
  try {
    const blob = await exportProcesses()
    const url = window.URL.createObjectURL(new Blob([blob]))
    const link = document.createElement('a')
    link.href = url
    const now = new Date()
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`
    link.setAttribute('download', `施工工艺_导出_${dateStr}.xlsx`)
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
    ElMessage.success('导出成功')
  } catch (error) {
    ElMessage.error('导出失败')
    console.error('导出工艺失败：', error)
  }
}

const handleDeleteAllProcesses = async () => {
  try {
    await ElMessageBox.confirm('确定删除全部施工工艺数据吗？此操作不可恢复！', '危险操作', {
      confirmButtonText: '确定删除',
      cancelButtonText: '取消',
      type: 'error',
      confirmButtonClass: 'el-button--danger',
    })
    const { default: request } = await import('../api')
    const res = await request.delete('/processes')
    ElMessage.success(res.data?.message || res.message || '删除成功')
    loadProcessList()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败')
      console.error('删除全部工艺失败：', error)
    }
  }
}

const handleDeleteProcess = async (row) => {
  try {
    await ElMessageBox.confirm(`确定删除工艺"${row.name}"吗？`, '提示', {
      confirmButtonText: '确定',
      cancelButtonText: '取消',
      type: 'warning',
    })
    const { default: request } = await import('../api')
    await request.delete(`/processes/${row.id}`)
    ElMessage.success('删除成功')
    loadProcessList()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败')
      console.error('删除工艺失败：', error)
    }
  }
}

// ---- 施工人员 ----
const workerList = ref([])
const workerCount = ref(0)
const importingWorker = ref(false)
const exportingWorker = ref(false)
const workerOverwriteMode = ref(false)
const workerImportResult = ref(null)

// 编辑施工人员
const workerEditVisible = ref(false)
const savingWorker = ref(false)
const workerEditForm = reactive({
  id: null,
  name: '',
  role: '',
  role2: '',
  role3: '',
  team: '',
  certification: '',
  certification2: '',
  certification3: '',
})

const loadWorkerList = async () => {
  try {
    const res = await getWorkers()
    workerList.value = res.items || []
    workerCount.value = workerList.value.length
  } catch (error) {
    console.error('加载人员列表失败：', error)
  }
}

const beforeUploadWorker = async (file) => {
  const allowedExtensions = ['.xlsx', '.xls', '.csv']
  const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase()
  if (!allowedExtensions.includes(fileExt)) {
    ElMessage.error('仅支持 .xlsx, .xls, .csv 格式')
    return false
  }
  importingWorker.value = true
  workerImportResult.value = null
  try {
    const res = await importWorkers(file, workerOverwriteMode.value)
    workerImportResult.value = res
    ElMessage.success(res.message || '导入完成')
    loadWorkerList()
  } catch (error) {
    ElMessage.error('导入失败')
    console.error('导入施工人员失败：', error)
  } finally {
    importingWorker.value = false
  }
  return false
}

const handleExportWorkers = async () => {
  if (exportingWorker.value) return
  exportingWorker.value = true
  try {
    const blob = await exportWorkers()
    const url = window.URL.createObjectURL(new Blob([blob]))
    const link = document.createElement('a')
    link.href = url
    const now = new Date()
    const yyyy = now.getFullYear()
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const dd = String(now.getDate()).padStart(2, '0')
    link.setAttribute('download', `施工人员_导出_${yyyy}${mm}${dd}.xlsx`)
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
    ElMessage.success('导出成功')
  } catch (error) {
    ElMessage.error('导出失败')
    console.error('导出施工人员失败：', error)
  } finally {
    exportingWorker.value = false
  }
}

const handleDeleteWorker = async (row) => {
  try {
    await ElMessageBox.confirm(`确定删除人员"${row.name}"吗？`, '提示', {
      confirmButtonText: '确定',
      cancelButtonText: '取消',
      type: 'warning',
    })
    const { default: request } = await import('../api')
    await request.delete(`/workers/${row.id}`)
    ElMessage.success('删除成功')
    loadWorkerList()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败')
      console.error('删除人员失败：', error)
    }
  }
}

const handleEditWorker = (row) => {
  workerEditForm.id = row.id
  workerEditForm.name = row.name || ''
  workerEditForm.role = row.role || ''
  workerEditForm.role2 = row.role2 || ''
  workerEditForm.role3 = row.role3 || ''
  workerEditForm.team = row.team || ''
  workerEditForm.certification = row.certification || ''
  workerEditForm.certification2 = row.certification2 || ''
  workerEditForm.certification3 = row.certification3 || ''
  workerEditVisible.value = true
}

const handleSaveWorker = async () => {
  if (!workerEditForm.name.trim()) {
    ElMessage.warning('姓名不能为空')
    return
  }
  savingWorker.value = true
  try {
    const { default: request } = await import('../api')
    await request.put(`/workers/${workerEditForm.id}`, {
      name: workerEditForm.name.trim(),
      role: workerEditForm.role.trim() || null,
      role2: workerEditForm.role2.trim() || null,
      role3: workerEditForm.role3.trim() || null,
      team: workerEditForm.team.trim() || null,
      certification: workerEditForm.certification.trim() || null,
      certification2: workerEditForm.certification2.trim() || null,
      certification3: workerEditForm.certification3.trim() || null,
    })
    ElMessage.success('保存成功')
    workerEditVisible.value = false
    loadWorkerList()
  } catch (error) {
    ElMessage.error('保存失败')
    console.error('保存人员失败：', error)
  } finally {
    savingWorker.value = false
  }
}

onMounted(() => {
  loadTemplates()
  loadProcessList()
  loadWorkerList()
})
</script>

<style scoped>
.data-foundation {
  max-width: 1200px;
}

.section-card {
  margin-bottom: 20px;
  border-radius: 8px;
}

.section-title {
  font-size: 16px;
  font-weight: 600;
  color: #303133;
}

.section-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.template-card {
  border-radius: 8px;
  margin-bottom: 12px;
}

.template-header {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 8px;
}

.template-label {
  font-size: 14px;
  font-weight: 600;
  color: #303133;
}

.template-info {
  margin-bottom: 8px;
}

.info-row {
  font-size: 12px;
  color: #606266;
  margin-bottom: 2px;
}

.info-label {
  color: #909399;
}

.info-value {
  color: #303133;
}

.no-template {
  font-size: 12px;
  color: #909399;
}

.template-actions {
  display: flex;
  gap: 8px;
}

.import-area {
  display: flex;
  align-items: center;
  gap: 12px;
}

.section-header-right {
  display: flex;
  align-items: center;
  gap: 8px;
}

.action-bar {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}

.form-hint {
  font-size: 11px;
  color: #909399;
  margin-top: -8px;
  display: block;
}

.import-hint {
  font-size: 12px;
  color: #909399;
}

.import-help {
  background: #f5f7fa;
  border-radius: 6px;
  padding: 12px 16px;
}

.help-title {
  font-size: 13px;
  font-weight: 600;
  color: #303133;
  margin-bottom: 6px;
}

.help-text {
  font-size: 12px;
  color: #606266;
  line-height: 1.8;
}

.expand-content {
  padding: 12px 20px;
}

.expand-row {
  font-size: 13px;
  color: #606266;
  line-height: 2;
}

.expand-label {
  color: #909399;
  font-weight: 600;
  margin-right: 4px;
}
</style>
