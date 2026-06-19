import axios from 'axios'

// 创建 axios 实例
const request = axios.create({
  baseURL: '/api',
  timeout: 30000
})

// 请求拦截器
request.interceptors.request.use(
  (config) => config,
  (error) => Promise.reject(error)
)

// 响应拦截器：统一提取 data
request.interceptors.response.use(
  (response) => response.data,
  (error) => {
    console.error('请求错误：', error)
    return Promise.reject(error)
  }
)

// ========== 项目相关接口 ==========
// 获取项目列表（返回 {items, total, page, page_size}）
export const getProjects = () => request.get('/projects')
// 获取单个项目详情
export const getProject = (id) => request.get(`/projects/${id}`)
// 创建项目
export const createProject = (data) => request.post('/projects', data)
// 更新项目
export const updateProject = (id, data) => request.put(`/projects/${id}`, data)
// 删除项目
export const deleteProject = (id) => request.delete(`/projects/${id}`)

// ========== 工程批量导入导出 ==========
// 导入工程Excel
export const importProjects = (file) => {
  const formData = new FormData()
  formData.append('file', file)
  return request.post('/projects/import', formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
  })
}

// 导出工程Excel
export const exportProjects = () => request.get('/projects/export', { responseType: 'blob' })

// 下载导入模板
export const downloadImportTemplate = () => request.get('/projects/import-template', { responseType: 'blob' })

// 自动排期
export const autoSchedule = (projectId) => request.post(`/projects/${projectId}/auto-schedule`)

// 文档预览（默认 docx，与下载的 Word 文档完全一致；type='pdf' 走 PDF 内联预览）
export const previewDocument = (docId, type = 'pdf') => request.get(`/projects/preview/${docId}`, {
  params: { type },
  responseType: 'blob',
})

// ========== 工程部位类别接口 ==========
// 获取工程部位类别列表（从施工工艺数据中提取）
export const getProjectTypes = () => request.get('/project-types')

// ========== 工艺相关接口 ==========
// 获取施工工艺列表（返回 {items: [...]}）
export const getProcesses = () => request.get('/processes')
// 修改单条施工工艺
export const updateProcess = (id, data) => request.put(`/processes/${id}`, data)
// 导出施工工艺为 Excel
export const exportProcesses = () => request.get('/processes/export', { responseType: 'blob' })
// 删除全部施工工艺
export const deleteAllProcesses = () => request.delete('/processes')
// 预览工艺导入：解析文件，返回新增/覆盖候选清单
export const previewImportProcesses = (file) => {
  const formData = new FormData()
  formData.append('file', file)
  return request.post('/processes/import/preview', formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
  })
}
// 确认工艺导入：根据 preview_id 和 overwrite_codes 执行导入
export const confirmImportProcesses = (previewId, overwriteCodes = []) => {
  return request.post('/processes/import/confirm', {
    preview_id: previewId,
    overwrite_codes: overwriteCodes
  })
}
// 批量导入施工工艺（兼容旧版，内部走 preview+confirm）
export const importProcesses = (file, overwrite = false) => {
  const formData = new FormData()
  formData.append('file', file)
  return request.post(`/processes/import?overwrite=${overwrite}`, formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
  })
}

// ========== 人员相关接口 ==========
// 获取施工人员列表（返回 {items: [...]}）
export const getWorkers = () => request.get('/workers')
// 导出施工人员为 Excel
export const exportWorkers = () => request.get('/workers/export', { responseType: 'blob' })
// 批量导入施工人员（支持 .xlsx, .xls, .csv，overwrite 参数控制覆盖模式）
export const importWorkers = (file, overwrite = false) => {
  const formData = new FormData()
  formData.append('file', file)
  return request.post(`/workers/import?overwrite=${overwrite}`, formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
  })
}

// ========== 施工组织设计相关接口 ==========
// 生成施工组织设计
export const generateConstructionDesign = (data) => request.post('/construction-design/generate', data)
// 获取施工组织设计生成历史
export const getConstructionDesignHistory = () => request.get('/construction-design/history')
// 删除施工组织设计历史记录
export const deleteConstructionDesign = (id) => request.delete(`/construction-design/${id}`)
// 批量删除施工组织设计历史记录（ids 通过 query 传，用 paramsSerializer 处理数组）
export const batchDeleteConstructionDesign = (ids) => request.delete('/construction-design/batch', { data: { ids } })

// ========== 勘察单相关接口 ==========
// 生成勘察单
export const generateSurvey = (data) => request.post('/survey/generate', data)
// 获取勘察单生成历史
export const getSurveyHistory = () => request.get('/survey/history')
// 删除勘察单历史记录
export const deleteSurvey = (id) => request.delete(`/survey/${id}`)
// 批量删除勘察单历史记录
export const batchDeleteSurvey = (ids) => request.delete('/survey/batch', { data: { ids } })

// ========== 技术交底相关接口 ==========
// 生成技术交底
export const generateTechBriefing = (data) => request.post('/tech-briefing/generate', data)
// 获取技术交底生成历史
export const getTechBriefingHistory = () => request.get('/tech-briefing/history')
// 删除技术交底历史记录
export const deleteTechBriefing = (id) => request.delete(`/tech-briefing/${id}`)
// 批量删除技术交底历史记录
export const batchDeleteTechBriefing = (ids) => request.delete('/tech-briefing/batch', { data: { ids } })

// ========== 安全交底相关接口 ==========
// 生成安全交底
export const generateSafetyBriefing = (data) => request.post('/safety-briefing/generate', data)
// 获取安全交底生成历史
export const getSafetyBriefingHistory = () => request.get('/safety-briefing/history')
// 删除安全交底历史记录
export const deleteSafetyBriefing = (id) => request.delete(`/safety-briefing/${id}`)
// 批量删除安全交底历史记录
export const batchDeleteSafetyBriefing = (ids) => request.delete('/safety-briefing/batch', { data: { ids } })

// ========== 施工进度横道图相关接口 ==========
// 获取施工进度任务列表
export const getScheduleTasks = (projectId) => request.get('/schedule-tasks', { params: { project_id: projectId } })
// 创建施工进度任务
export const createScheduleTask = (data) => request.post('/schedule-tasks', data)
// 批量创建施工进度任务
export const batchCreateScheduleTasks = (data) => request.post('/schedule-tasks/batch', data)
// 更新施工进度任务
export const updateScheduleTask = (id, data) => request.put(`/schedule-tasks/${id}`, data)
// 删除施工进度任务
export const deleteScheduleTask = (id) => request.delete(`/schedule-tasks/${id}`)
// 删除工程下所有进度任务
export const deleteScheduleTasksByProject = (projectId) => request.delete(`/schedule-tasks/by-project/${projectId}`)
// 从 Excel 模板导入施工进度
export const importScheduleFromTemplate = (projectId) => request.post(`/schedule-tasks/import-template?project_id=${projectId}`)
// 上传 Excel 文件导入施工进度
export const importScheduleFromExcel = (projectId, file) => {
  const formData = new FormData()
  formData.append('file', file)
  return request.post(`/schedule-tasks/import-excel?project_id=${projectId}`, formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
  })
}
// 生成施工进度横道图文档（xlsx）
export const generateScheduleChart = (projectId) => request.post('/schedule-tasks/generate', { project_id: projectId })
// 获取施工进度横道图生成历史
export const getScheduleChartHistory = (projectId) => request.get('/schedule-tasks/history', { params: { project_id: projectId } })
// 删除施工进度横道图历史记录
export const deleteScheduleHistory = (id) => request.delete(`/schedule-tasks/history/${id}`)
// 批量删除施工进度横道图历史记录
export const batchDeleteScheduleHistory = (ids) => request.delete('/schedule-tasks/history/batch', { data: { ids } })

// ========== 文档模板管理接口 ==========
// 上传文档模板（支持 .docx, .doc, .xlsx, .xls, .md, .pdf）
export const uploadTemplate = (docType, file) => {
  const formData = new FormData()
  formData.append('file', file)
  return request.post(`/templates/upload?doc_type=${docType}`, formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
  })
}
// 获取模板列表
export const getTemplates = () => request.get('/templates')
// 删除模板
export const deleteTemplate = (docType) => request.delete(`/templates/${docType}`)

export default request
