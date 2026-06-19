<template>
  <div class="construction-design">
    <el-card shadow="hover" class="form-card">
      <el-form label-width="120px">
        <!-- 选择工程 -->
        <el-form-item label="选择工程" required>
          <el-select
            v-model="selectedProjectId"
            placeholder="请选择工程"
            filterable
            style="width: 100%"
            @change="handleProjectChange"
          >
            <el-option
              v-for="p in projects"
              :key="p.id"
              :label="`${p.project_code} - ${p.project_name}`"
              :value="p.id"
            />
          </el-select>
        </el-form-item>
      </el-form>

      <!-- 生成按钮 -->
      <div class="generate-action">
        <el-button
          type="primary"
          size="large"
          :loading="generating"
          :disabled="!selectedProjectId"
          @click="handleGenerate"
        >
          生成施工组织设计
        </el-button>
      </div>

      <!-- 生成结果 -->
      <el-result
        v-if="generateResult"
        icon="success"
        title="施工组织设计生成成功！"
        sub-title="您可以预览或下载生成的文档"
      >
        <template #extra>
          <el-button type="primary" @click="openPreview(generateResult.id)">预览文档</el-button>
          <el-button @click="handleDownload">下载文档</el-button>
        </template>
      </el-result>
    </el-card>

    <!-- 生成历史 -->
    <el-card shadow="hover" class="history-card">
      <template #header>
        <div class="history-header">
          <span class="section-title">施工组织设计生成历史</span>
          <div class="history-actions">
            <span v-if="selectedHistoryIds.length > 0" class="selected-tip">
              已选 {{ selectedHistoryIds.length }} 条
            </span>
            <el-button
              v-if="selectedHistoryIds.length > 0"
              type="danger"
              size="small"
              @click="handleBatchDelete"
            >
              批量删除
            </el-button>
            <el-button @click="loadHistory" :loading="historyLoading" size="small">
              刷新
            </el-button>
          </div>
        </div>
      </template>
      <el-table
        :data="history"
        stripe
        style="width: 100%"
        v-loading="historyLoading"
        @selection-change="handleSelectionChange"
      >
        <el-table-column type="selection" width="50" />
        <el-table-column prop="project_name" label="工程名称" />
        <el-table-column prop="created_at" label="生成时间" width="180" />
        <el-table-column label="操作" width="240">
          <template #default="{ row }">
            <el-button type="primary" link size="small" @click="openPreview(row.id)">预览</el-button>
            <el-button type="primary" link size="small" @click="downloadHistory(row)">下载</el-button>
            <el-button type="danger" link size="small" @click="handleDelete(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <!-- 预览弹窗（iframe 预览后端 PDF） -->
    <el-dialog v-model="previewVisible" title="文档预览" width="900px" top="2vh" destroy-on-close>
      <iframe :src="previewUrl" style="width: 100%; height: 75vh; border: none;"></iframe>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import {
  getProjects,
  generateConstructionDesign,
  getConstructionDesignHistory,
  deleteConstructionDesign,
  batchDeleteConstructionDesign
} from '../api'

// 工程列表与选择
const projects = ref([])
const selectedProjectId = ref('')

// 生成状态
const generating = ref(false)
const generateResult = ref(null)

// 历史记录
const history = ref([])
const historyLoading = ref(false)
// 历史记录多选
const selectedHistoryIds = ref([])
const handleSelectionChange = (rows) => {
  selectedHistoryIds.value = rows.map(r => r.id)
}

// 单条删除
const handleDelete = async (row) => {
  try {
    await ElMessageBox.confirm(
      `确定要删除 "${row.project_name || '该工程'}" 的施工组织设计记录吗？\n（同时清理对应的 .docx 与 .pdf 文件）`,
      '删除确认',
      {
        confirmButtonText: '删除',
        cancelButtonText: '取消',
        type: 'warning',
        confirmButtonClass: 'el-button--danger',
      }
    )
  } catch {
    return
  }
  try {
    await deleteConstructionDesign(row.id)
    ElMessage.success('删除成功')
    loadHistory()
  } catch (error) {
    ElMessage.error('删除失败：' + (error.message || '未知错误'))
  }
}

// 批量删除
const handleBatchDelete = async () => {
  const ids = selectedHistoryIds.value
  if (!ids.length) return
  try {
    await ElMessageBox.confirm(
      `确定要删除已选中的 ${ids.length} 条记录吗？\n（同时清理对应的 .docx 与 .pdf 文件，且不可恢复）`,
      '批量删除确认',
      {
        confirmButtonText: '全部删除',
        cancelButtonText: '取消',
        type: 'warning',
        confirmButtonClass: 'el-button--danger',
      }
    )
  } catch {
    return
  }
  try {
    await batchDeleteConstructionDesign(ids)
    ElMessage.success(`已删除 ${ids.length} 条记录`)
    selectedHistoryIds.value = []
    loadHistory()
  } catch (error) {
    ElMessage.error('批量删除失败：' + (error.message || '未知错误'))
  }
}

// 选择工程后清空上一次结果
const handleProjectChange = () => {
  generateResult.value = null
}

// 生成施工组织设计（不传任何选择项，后端自动从工程关联数据组合）
const handleGenerate = async () => {
  generating.value = true
  try {
    const res = await generateConstructionDesign({
      project_id: selectedProjectId.value,
      worker_ids: [],
      process_ids: [],
      worker_roles: {},
    })
    generateResult.value = res
    ElMessage.success('施工组织设计生成成功！')
    loadHistory()
  } catch (error) {
    ElMessage.error('生成失败，请重试')
    console.error('生成施工组织设计失败：', error)
  } finally {
    generating.value = false
  }
}

// 下载文档
const handleDownload = () => {
  if (generateResult.value?.download_url) {
    window.open(generateResult.value.download_url, '_blank')
  } else {
    ElMessage.warning('暂无下载链接')
  }
}

// 下载历史文档
const downloadHistory = (row) => {
  if (row.download_url) {
    window.open(row.download_url, '_blank')
  } else {
    ElMessage.warning('暂无下载链接')
  }
}

// 预览弹窗状态
const previewUrl = ref('')
const previewVisible = ref(false)

// 打开预览（通过后端 PDF 预览 API）
const openPreview = (docId) => {
  previewUrl.value = `/api/projects/preview/${docId}?type=pdf`
  previewVisible.value = true
}

// 加载工程列表
const loadProjects = async () => {
  try {
    const res = await getProjects()
    projects.value = res.items || []
  } catch (error) {
    console.error('加载工程列表失败：', error)
  }
}

// 加载历史记录
const loadHistory = async () => {
  historyLoading.value = true
  try {
    const res = await getConstructionDesignHistory()
    const items = res.items || res || []
    history.value = items.map(item => ({
      ...item,
      project_name: item.project_name
        ? (item.project_code ? `${item.project_code} - ${item.project_name}` : item.project_name)
        : (() => {
            const proj = projects.value.find(p => p.id === item.project_id)
            if (!proj) return ''
            return proj.project_code ? `${proj.project_code} - ${proj.project_name}` : proj.project_name
          })()
    }))
  } catch (error) {
    console.error('加载历史记录失败：', error)
  } finally {
    historyLoading.value = false
  }
}

onMounted(() => {
  loadProjects()
  loadHistory()
})
</script>

<style scoped>
.construction-design {
  max-width: 1000px;
}

.form-card {
  margin-bottom: 20px;
  border-radius: 8px;
}

.generate-action {
  display: flex;
  justify-content: center;
  margin-top: 24px;
}

.history-card {
  border-radius: 8px;
}

.history-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.history-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}

.selected-tip {
  font-size: 12px;
  color: #909399;
}

.section-title {
  font-size: 16px;
  font-weight: 600;
  color: #303133;
}
</style>
