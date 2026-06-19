<template>
  <div class="gantt-chart">
    <!-- 表单区域 -->
    <el-card shadow="hover" class="form-card">
      <template #header>
        <span class="section-title">施工进度横道图</span>
      </template>

      <el-form label-width="120px">
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

      <el-divider />

      <div class="generate-action">
        <el-button
          type="primary"
          size="large"
          :loading="generating"
          :disabled="!selectedProjectId"
          @click="handleGenerate"
        >
          <el-icon><TrendCharts /></el-icon>
          生成施工进度横道图
        </el-button>
      </div>

      <el-result
          v-if="generateResult"
          icon="success"
          title="施工进度横道图生成成功！"
          sub-title="您可以预览或下载生成的文档"
        >
          <template #extra>
            <el-button type="primary" @click="handlePreview(generateResult.id)">预览文档</el-button>
            <el-button @click="handleDownload(generateResult.download_url)">下载文档</el-button>
          </template>
        </el-result>
    </el-card>

    <!-- 生成历史 -->
    <el-card shadow="hover" class="history-card">
      <template #header>
        <div class="history-header">
          <span class="section-title">施工进度横道图生成历史</span>
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
            <el-button type="primary" link size="small" @click="handlePreview(row.id)">预览</el-button>
            <el-button type="primary" link size="small" @click="handleDownload(row.download_url)">下载</el-button>
            <el-button type="danger" link size="small" @click="handleDelete(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <!-- Excel 预览弹窗 -->
    <el-dialog
      v-model="previewVisible"
      title="施工进度横道图预览"
      width="98%"
      top="1vh"
      destroy-on-close
      class="gantt-preview-dialog"
    >
      <div v-loading="previewLoading" class="preview-container">
        <div v-if="previewHtml" class="excel-preview" v-html="previewHtml"></div>
        <el-empty v-else-if="!previewLoading" description="暂无预览内容" />
      </div>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import * as XLSX from 'xlsx'
import {
  getProjects,
  importScheduleFromTemplate,
  generateScheduleChart,
  getScheduleChartHistory,
  deleteScheduleHistory,
  batchDeleteScheduleHistory
} from '../api'

const projects = ref([])
const selectedProjectId = ref('')
const generating = ref(false)
const generateResult = ref(null)

// 历史记录
const history = ref([])
const historyLoading = ref(false)
const selectedHistoryIds = ref([])
const handleSelectionChange = (rows) => {
  selectedHistoryIds.value = rows.map(r => r.id)
}

// 预览
const previewVisible = ref(false)
const previewLoading = ref(false)
const previewHtml = ref('')

// ===== 生成横道图 =====
const handleGenerate = async () => {
  generating.value = true
  try {
    await importScheduleFromTemplate(selectedProjectId.value)
    const res = await generateScheduleChart(selectedProjectId.value)
    generateResult.value = res
    ElMessage.success(res.message || '横道图生成成功！')
    await loadHistory()
  } catch (error) {
    ElMessage.error('生成失败，请重试')
    console.error('生成施工进度横道图失败：', error)
  } finally {
    generating.value = false
  }
}

// ===== 预览Excel =====
const handlePreview = async (docId) => {
  previewVisible.value = true
  previewLoading.value = true
  previewHtml.value = ''

  try {
    // 下载Excel文件的二进制数据
    const downloadUrl = `/api/schedule-tasks/download/${docId}`
    const response = await fetch(downloadUrl)
    if (!response.ok) throw new Error('下载失败')
    const arrayBuffer = await response.arrayBuffer()

    // 用SheetJS解析
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true })
    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]

    // 转为HTML，保留样式
    const html = XLSX.utils.sheet_to_html(worksheet, {
      id: 'gantt-table',
      editable: false,
    })

    // 后处理：给蓝色填充的单元格加样式
    previewHtml.value = enhanceHtmlStyles(html, worksheet)
  } catch (error) {
    console.error('预览失败：', error)
    ElMessage.error('预览失败：' + (error.message || '未知错误'))
    previewHtml.value = ''
  } finally {
    previewLoading.value = false
  }
}

// 增强HTML样式：根据单元格填充色添加背景色
const enhanceHtmlStyles = (html, worksheet) => {
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')

  // 收集有填充色的单元格
  const blueCells = new Set()
  const headerCells = new Set()
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C })
      const cell = worksheet[addr]
      if (!cell || !cell.s || !cell.s.fgColor) continue
      const rgb = cell.s.fgColor.rgb
      if (rgb) {
        const upper = rgb.toUpperCase()
        if (upper.includes('00B0F0')) {
          blueCells.add(addr)
        } else if (upper.includes('D9E1F2')) {
          headerCells.add(addr)
        }
      }
    }
  }

  // 给HTML中的td添加样式
  return html.replace(/<td([^>]*)>/g, (match, attrs) => {
    // 提取id属性中的单元格地址
    const idMatch = attrs.match(/id="[^"]*-(\w+)"/)
    if (!idMatch) return match
    const cellAddr = idMatch[1]
    let style = ''
    if (blueCells.has(cellAddr)) {
      style = 'background-color:#00B0F0;'
    } else if (headerCells.has(cellAddr)) {
      style = 'background-color:#D9E1F2;'
    }
    // 所有单元格加边框
    style += 'border:1px solid #999;padding:1px 4px;font-size:11px;font-family:宋体;white-space:nowrap;'
    if (style) {
      if (attrs.includes('style="')) {
        return match.replace('style="', `style="${style}`)
      } else {
        return `<td${attrs} style="${style}">`
      }
    }
    return match
  })
}

// 下载
const handleDownload = (downloadUrl) => {
  if (downloadUrl) {
    window.open(downloadUrl, '_blank')
  } else {
    ElMessage.warning('暂无下载链接，请先生成横道图')
  }
}

// 单条删除
const handleDelete = async (row) => {
  try {
    await ElMessageBox.confirm(
      `确定要删除 "${row.project_name || '该工程'}" 的施工进度横道图记录吗？`,
      '删除确认',
      {
        confirmButtonText: '删除',
        cancelButtonText: '取消',
        type: 'warning',
        confirmButtonClass: 'el-button--danger',
      }
    )
  } catch { return }
  try {
    await deleteScheduleHistory(row.id)
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
      `确定要删除已选中的 ${ids.length} 条记录吗？（不可恢复）`,
      '批量删除确认',
      {
        confirmButtonText: '全部删除',
        cancelButtonText: '取消',
        type: 'warning',
        confirmButtonClass: 'el-button--danger',
      }
    )
  } catch { return }
  try {
    await batchDeleteScheduleHistory(ids)
    ElMessage.success(`已删除 ${ids.length} 条记录`)
    selectedHistoryIds.value = []
    loadHistory()
  } catch (error) {
    ElMessage.error('批量删除失败：' + (error.message || '未知错误'))
  }
}

// ===== 数据加载 =====
const loadProjects = async () => {
  try {
    const res = await getProjects()
    projects.value = res.items || []
    await loadHistory()
  } catch (error) {
    console.error('加载工程列表失败：', error)
  }
}

const loadHistory = async () => {
  historyLoading.value = true
  try {
    const res = await getScheduleChartHistory()
    const items = res.items || []
    history.value = items.map(item => ({
      ...item,
      project_name: projects.value.find(p => p.id === item.project_id)?.project_name || ''
    }))
  } catch (error) {
    console.error('加载历史记录失败：', error)
    history.value = []
  } finally {
    historyLoading.value = false
  }
}

const handleProjectChange = () => {
  generateResult.value = null
}

onMounted(() => {
  loadProjects()
})
</script>

<style scoped>
.gantt-chart {
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

/* 预览容器 */
.preview-container {
  min-height: 400px;
  max-height: 85vh;
  overflow: auto;
}

/* Excel预览表格样式 */
.excel-preview {
  overflow: auto;
}

.excel-preview :deep(table) {
  border-collapse: collapse;
  font-family: "SimSun", "宋体", serif;
  font-size: 11px;
}

.excel-preview :deep(td),
.excel-preview :deep(th) {
  border: 1px solid #999;
  padding: 1px 4px;
  white-space: nowrap;
  vertical-align: middle;
}

.excel-preview :deep(tr:first-child td),
.excel-preview :deep(tr:first-child th) {
  background-color: #D9E1F2;
  font-weight: bold;
  font-size: 10px;
  text-align: center;
}

.excel-preview :deep(tr:nth-child(2) td),
.excel-preview :deep(tr:nth-child(2) th) {
  background-color: #D9E1F2;
  font-size: 9px;
  text-align: center;
}
</style>
