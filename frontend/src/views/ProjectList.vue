<template>
  <div class="project-list">
    <!-- 搜索栏 -->
    <el-card shadow="hover" class="search-card">
      <el-row :gutter="16" align="middle">
        <el-col :span="6">
          <el-input
            v-model="searchKeyword"
            placeholder="搜索项目编号/名称"
            clearable
            @clear="handleSearch"
            @keyup.enter="handleSearch"
          >
            <template #prefix>
              <el-icon><Search /></el-icon>
            </template>
          </el-input>
        </el-col>
        <el-col :span="5">
          <el-select v-model="filterType" placeholder="工程类别筛选" clearable @change="handleSearch">
            <el-option v-for="opt in projectTypeOptions" :key="opt" :label="opt" :value="opt" />
          </el-select>
        </el-col>
        <el-col :span="3">
          <el-button type="primary" @click="handleSearch">
            <el-icon><Search /></el-icon>搜索
          </el-button>
        </el-col>
        <el-col :span="10" style="text-align: right">
          <el-button type="success" @click="importDialogVisible = true">
            <el-icon><Upload /></el-icon>导入
          </el-button>
          <el-button type="warning" @click="handleExport">
            <el-icon><Download /></el-icon>导出
          </el-button>
          <el-button plain @click="handleDownloadTemplate">
            <el-icon><Document /></el-icon>下载模板
          </el-button>
        </el-col>
      </el-row>
    </el-card>

    <!-- 工程列表 -->
    <el-card shadow="hover" class="table-card">
      <el-table :data="pagedData" stripe style="width: 100%" v-loading="loading">
        <el-table-column prop="project_code" label="项目编号" width="160" />
        <el-table-column prop="project_name" label="项目名称" min-width="180" show-overflow-tooltip />
        <el-table-column label="分包单位" min-width="200">
          <template #default="{ row }">
            <div v-if="row.subcontractor_civil || row.subcontractor_electric">
              <div v-if="row.subcontractor_civil"><el-tag size="small" type="warning">土建</el-tag> {{ row.subcontractor_civil }}</div>
              <div v-if="row.subcontractor_electric"><el-tag size="small" type="success">电气</el-tag> {{ row.subcontractor_electric }}</div>
            </div>
            <span v-else>{{ row.subcontractor }}</span>
          </template>
        </el-table-column>
        <el-table-column label="工程类别" min-width="140">
          <template #default="{ row }">
            <template v-if="row.project_type">
              <el-tag v-for="t in row.project_type.split(',').filter(Boolean)" :key="t" size="small" style="margin: 2px">{{ t }}</el-tag>
            </template>
          </template>
        </el-table-column>
        <el-table-column prop="start_date" label="开工日期" width="120" />
        <el-table-column label="操作" width="260" fixed="right">
          <template #default="{ row }">
            <el-button type="primary" link size="small" @click="viewDetail(row)">查看详情</el-button>
            <el-button type="warning" link size="small" @click="editProject(row)">编辑</el-button>
            <el-button type="success" link size="small" @click="handleAutoSchedule(row)">排期</el-button>
            <el-button type="danger" link size="small" @click="handleDelete(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>

      <!-- 分页 -->
      <div class="pagination-wrap">
        <el-pagination
          v-model:current-page="currentPage"
          v-model:page-size="pageSize"
          :page-sizes="[10, 20, 50]"
          :total="filteredData.length"
          layout="total, sizes, prev, pager, next, jumper"
          background
        />
      </div>
    </el-card>

    <!-- 导入对话框 -->
    <el-dialog v-model="importDialogVisible" title="导入工程" width="600px" @close="resetImportState">
      <el-upload
        ref="importUploadRef"
        :auto-upload="false"
        :limit="1"
        accept=".xlsx,.xls,.csv"
        :on-change="handleImportFileChange"
        :on-exceed="() => ElMessage.warning('只能上传一个文件')"
        drag
      >
        <el-icon style="font-size: 48px; color: #c0c4cc"><Upload /></el-icon>
        <div style="margin-top: 8px">将文件拖到此处，或<em>点击上传</em></div>
        <template #tip>
          <div class="el-upload__tip">支持 .xlsx / .xls / .csv 格式</div>
        </template>
      </el-upload>

      <!-- 校验结果 -->
      <div v-if="importResult" style="margin-top: 16px">
        <el-alert
          :title="`校验完成：成功 ${importResult.success_count || 0} 条，失败 ${importResult.fail_count || 0} 条`"
          :type="importResult.fail_count > 0 ? 'warning' : 'success'"
          show-icon
          :closable="false"
        />
        <el-table v-if="importResult.errors && importResult.errors.length" :data="importResult.errors" size="small" border style="margin-top: 12px">
          <el-table-column prop="row" label="行号" width="80" />
          <el-table-column prop="field" label="字段" width="120" />
          <el-table-column prop="message" label="错误信息" />
        </el-table>
      </div>

      <template #footer>
        <el-button @click="importDialogVisible = false">取消</el-button>
        <el-button type="primary" :loading="importing" :disabled="!importFile" @click="handleImportPreview">上传校验</el-button>
        <el-button type="success" :loading="importing" :disabled="!importResult || importResult.success_count === 0" @click="handleConfirmImport">确认导入</el-button>
      </template>
    </el-dialog>

    <!-- 查看详情弹窗 -->
    <el-dialog v-model="detailVisible" title="工程详情" width="800px" top="3vh">
      <template v-if="currentRow">
        <el-descriptions :column="2" border>
          <el-descriptions-item label="项目编号">{{ currentRow.project_code }}</el-descriptions-item>
          <el-descriptions-item label="项目名称">{{ currentRow.project_name }}</el-descriptions-item>
          <el-descriptions-item label="工程类别">
            <el-tag v-for="t in (currentRow.project_type || '').split(',').filter(Boolean)" :key="t" size="small" style="margin: 2px">{{ t }}</el-tag>
          </el-descriptions-item>
          <el-descriptions-item label="工程地点">{{ currentRow.location }}</el-descriptions-item>
          <el-descriptions-item label="分包单位">{{ currentRow.subcontractor }}</el-descriptions-item>
          <el-descriptions-item label="电压等级">{{ currentRow.voltage_level }}</el-descriptions-item>
          <el-descriptions-item label="线路名称">{{ currentRow.line_name }}</el-descriptions-item>
          <el-descriptions-item label="编制单位">{{ currentRow.preparation_unit }}</el-descriptions-item>
          <el-descriptions-item label="开工日期">{{ currentRow.start_date }}</el-descriptions-item>
          <el-descriptions-item label="竣工日期">{{ currentRow.end_date }}</el-descriptions-item>
          <el-descriptions-item label="工作任务" :span="2">{{ currentRow.work_task }}</el-descriptions-item>
          <el-descriptions-item label="工程概况" :span="2">{{ currentRow.description }}</el-descriptions-item>
        </el-descriptions>

        <!-- 勘察信息 -->
        <el-collapse style="margin-top: 16px">
          <el-collapse-item title="勘察信息" name="survey">
            <el-descriptions :column="2" border size="small">
              <el-descriptions-item label="勘察负责人">{{ currentRow.survey_leader }}</el-descriptions-item>
              <el-descriptions-item label="勘察人员">{{ currentRow.survey_members }}</el-descriptions-item>
              <el-descriptions-item label="勘察日期">{{ currentRow.survey_date }}</el-descriptions-item>
              <el-descriptions-item label="勘察结果" :span="2">{{ currentRow.survey_result }}</el-descriptions-item>
            </el-descriptions>
          </el-collapse-item>

          <!-- 交底信息 -->
          <el-collapse-item title="交底信息" name="briefing">
            <el-descriptions :column="2" border size="small">
              <el-descriptions-item label="交底主持人">{{ currentRow.briefing_host }}</el-descriptions-item>
              <el-descriptions-item label="交底日期">{{ currentRow.briefing_date }}</el-descriptions-item>
              <el-descriptions-item label="交底内容" :span="2">{{ currentRow.briefing_content }}</el-descriptions-item>
            </el-descriptions>
          </el-collapse-item>
        </el-collapse>

        <!-- 项目管理人员 -->
        <h4 class="detail-member-title">项目管理人员</h4>
        <el-table :data="detailManagers" size="small" border v-if="detailManagers.length">
          <el-table-column prop="name" label="姓名" />
          <el-table-column prop="role" label="职务" />
        </el-table>
        <div v-else class="no-member-tip">暂无管理人员</div>

        <!-- 施工人员 -->
        <h4 class="detail-member-title">施工人员</h4>
        <el-table :data="detailWorkers" size="small" border v-if="detailWorkers.length">
          <el-table-column prop="name" label="姓名" />
          <el-table-column prop="role" label="工种" />
        </el-table>
        <div v-else class="no-member-tip">暂无施工人员</div>

        <!-- 施工机具清单 -->
        <h4 class="detail-member-title">施工机具清单</h4>
        <el-table :data="currentRow.equipments || []" size="small" border v-if="(currentRow.equipments || []).length">
          <el-table-column prop="name" label="机具名称" />
          <el-table-column prop="model" label="规格型号" />
          <el-table-column prop="quantity" label="数量" width="80" />
          <el-table-column prop="unit" label="单位" width="80" />
        </el-table>
        <div v-else class="no-member-tip">暂无机具信息</div>

        <!-- 质量控制点 -->
        <h4 class="detail-member-title">质量控制点</h4>
        <el-table :data="currentRow.quality_controls || []" size="small" border v-if="(currentRow.quality_controls || []).length">
          <el-table-column prop="control_point" label="控制点" />
          <el-table-column prop="inspection_method" label="检验方法" />
          <el-table-column prop="standard" label="质量标准" />
          <el-table-column prop="responsible" label="责任人" />
        </el-table>
        <div v-else class="no-member-tip">暂无质量控制点</div>

        <!-- 审批信息 -->
        <h4 class="detail-member-title">审批信息</h4>
        <el-table :data="currentRow.approvals || []" size="small" border v-if="(currentRow.approvals || []).length">
          <el-table-column prop="approver" label="审批人" />
          <el-table-column prop="role" label="职务" />
          <el-table-column prop="status" label="审批状态" width="100">
            <template #default="{ row }">
              <el-tag :type="row.status === 'approved' ? 'success' : row.status === 'rejected' ? 'danger' : 'info'" size="small">
                {{ row.status === 'approved' ? '已通过' : row.status === 'rejected' ? '已驳回' : '待审批' }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="date" label="审批日期" width="120" />
          <el-table-column prop="comment" label="审批意见" />
        </el-table>
        <div v-else class="no-member-tip">暂无审批信息</div>

        <!-- 关联施工工艺 -->
        <h4 class="detail-member-title">关联施工工艺</h4>
        <el-table :data="currentRow.processes || []" size="small" border v-if="(currentRow.processes || []).length">
          <el-table-column prop="code" label="工艺编号" width="120" />
          <el-table-column prop="name" label="工艺名称" />
          <el-table-column prop="category" label="类别" width="120" />
        </el-table>
        <div v-else class="no-member-tip">暂无关联施工工艺</div>
      </template>
    </el-dialog>

    <!-- 编辑弹窗 -->
    <el-dialog v-model="editVisible" title="编辑工程" width="800px" top="3vh">
      <el-form ref="editFormRef" :model="editForm" :rules="editRules" label-width="110px" v-if="editVisible">
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="项目编号" prop="project_code">
              <el-input v-model="editForm.project_code" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="项目名称" prop="project_name">
              <el-input v-model="editForm.project_name" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="工程类别" prop="project_type">
          <el-checkbox-group v-model="editForm.project_type">
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
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="工程地点" prop="location">
              <el-input v-model="editForm.location" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="分包单位" prop="subcontractor">
              <el-input v-model="editForm.subcontractor" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="电压等级">
              <el-input v-model="editForm.voltage_level" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="线路名称">
              <el-input v-model="editForm.line_name" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="编制单位">
              <el-input v-model="editForm.preparation_unit" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="工作任务">
              <el-input v-model="editForm.work_task" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="16">
          <el-col :span="12">
            <el-form-item label="开工日期" prop="start_date">
              <el-date-picker v-model="editForm.start_date" type="date" value-format="YYYY-MM-DD" style="width: 100%" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="竣工日期" prop="end_date">
              <el-date-picker v-model="editForm.end_date" type="date" value-format="YYYY-MM-DD" style="width: 100%" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="工程概况" prop="description">
          <el-input v-model="editForm.description" type="textarea" :rows="3" />
        </el-form-item>

        <!-- 勘察信息折叠面板 -->
        <el-collapse v-model="editCollapseActive">
          <el-collapse-item title="勘察信息" name="survey">
            <el-row :gutter="16">
              <el-col :span="12">
                <el-form-item label="勘察负责人">
                  <el-input v-model="editForm.survey_leader" />
                </el-form-item>
              </el-col>
              <el-col :span="12">
                <el-form-item label="勘察人员">
                  <el-input v-model="editForm.survey_members" />
                </el-form-item>
              </el-col>
            </el-row>
            <el-row :gutter="16">
              <el-col :span="12">
                <el-form-item label="勘察日期">
                  <el-date-picker v-model="editForm.survey_date" type="date" value-format="YYYY-MM-DD" style="width: 100%" />
                </el-form-item>
              </el-col>
            </el-row>
            <el-form-item label="勘察结果">
              <el-input v-model="editForm.survey_result" type="textarea" :rows="2" />
            </el-form-item>
          </el-collapse-item>

          <!-- 交底信息 -->
          <el-collapse-item title="交底信息" name="briefing">
            <el-row :gutter="16">
              <el-col :span="12">
                <el-form-item label="交底主持人">
                  <el-input v-model="editForm.briefing_host" />
                </el-form-item>
              </el-col>
              <el-col :span="12">
                <el-form-item label="交底日期">
                  <el-date-picker v-model="editForm.briefing_date" type="date" value-format="YYYY-MM-DD" style="width: 100%" />
                </el-form-item>
              </el-col>
            </el-row>
            <el-form-item label="交底内容">
              <el-input v-model="editForm.briefing_content" type="textarea" :rows="2" />
            </el-form-item>
          </el-collapse-item>

          <!-- 施工机具清单 -->
          <el-collapse-item title="施工机具清单" name="equipments">
            <el-table :data="editEquipments" size="small" border>
              <el-table-column prop="name" label="机具名称" width="140">
                <template #default="{ row }"><el-input v-model="row.name" size="small" /></template>
              </el-table-column>
              <el-table-column prop="model" label="规格型号" width="140">
                <template #default="{ row }"><el-input v-model="row.model" size="small" /></template>
              </el-table-column>
              <el-table-column prop="quantity" label="数量" width="90">
                <template #default="{ row }"><el-input-number v-model="row.quantity" size="small" :min="0" controls-position="right" /></template>
              </el-table-column>
              <el-table-column prop="unit" label="单位" width="80">
                <template #default="{ row }"><el-input v-model="row.unit" size="small" /></template>
              </el-table-column>
              <el-table-column label="操作" width="60" align="center">
                <template #default="{ $index }">
                  <el-button type="danger" :icon="Delete" circle size="small" @click="editEquipments.splice($index, 1)" />
                </template>
              </el-table-column>
            </el-table>
            <el-button type="primary" plain size="small" style="margin-top: 8px" @click="editEquipments.push({ name: '', model: '', quantity: 0, unit: '' })">
              <el-icon><Plus /></el-icon>添加机具
            </el-button>
          </el-collapse-item>

          <!-- 质量控制点 -->
          <el-collapse-item title="质量控制点" name="quality">
            <el-table :data="editQualityControls" size="small" border>
              <el-table-column prop="control_point" label="控制点" width="130">
                <template #default="{ row }"><el-input v-model="row.control_point" size="small" /></template>
              </el-table-column>
              <el-table-column prop="inspection_method" label="检验方法" width="130">
                <template #default="{ row }"><el-input v-model="row.inspection_method" size="small" /></template>
              </el-table-column>
              <el-table-column prop="standard" label="质量标准" width="130">
                <template #default="{ row }"><el-input v-model="row.standard" size="small" /></template>
              </el-table-column>
              <el-table-column prop="responsible" label="责任人" width="100">
                <template #default="{ row }"><el-input v-model="row.responsible" size="small" /></template>
              </el-table-column>
              <el-table-column label="操作" width="60" align="center">
                <template #default="{ $index }">
                  <el-button type="danger" :icon="Delete" circle size="small" @click="editQualityControls.splice($index, 1)" />
                </template>
              </el-table-column>
            </el-table>
            <el-button type="primary" plain size="small" style="margin-top: 8px" @click="editQualityControls.push({ control_point: '', inspection_method: '', standard: '', responsible: '' })">
              <el-icon><Plus /></el-icon>添加控制点
            </el-button>
          </el-collapse-item>

          <!-- 审批信息 -->
          <el-collapse-item title="审批信息" name="approvals">
            <el-table :data="editApprovals" size="small" border>
              <el-table-column prop="approver" label="审批人" width="110">
                <template #default="{ row }"><el-input v-model="row.approver" size="small" /></template>
              </el-table-column>
              <el-table-column prop="role" label="职务" width="110">
                <template #default="{ row }"><el-input v-model="row.role" size="small" /></template>
              </el-table-column>
              <el-table-column prop="status" label="审批状态" width="110">
                <template #default="{ row }">
                  <el-select v-model="row.status" size="small">
                    <el-option label="待审批" value="pending" />
                    <el-option label="已通过" value="approved" />
                    <el-option label="已驳回" value="rejected" />
                  </el-select>
                </template>
              </el-table-column>
              <el-table-column prop="date" label="审批日期" width="140">
                <template #default="{ row }"><el-date-picker v-model="row.date" type="date" value-format="YYYY-MM-DD" size="small" style="width: 100%" /></template>
              </el-table-column>
              <el-table-column prop="comment" label="审批意见">
                <template #default="{ row }"><el-input v-model="row.comment" size="small" /></template>
              </el-table-column>
              <el-table-column label="操作" width="60" align="center">
                <template #default="{ $index }">
                  <el-button type="danger" :icon="Delete" circle size="small" @click="editApprovals.splice($index, 1)" />
                </template>
              </el-table-column>
            </el-table>
            <el-button type="primary" plain size="small" style="margin-top: 8px" @click="editApprovals.push({ approver: '', role: '', status: 'pending', date: '', comment: '' })">
              <el-icon><Plus /></el-icon>添加审批
            </el-button>
          </el-collapse-item>
        </el-collapse>

        <!-- 编辑管理人员 -->
        <el-divider content-position="left">项目管理人员</el-divider>
        <div v-for="(m, i) in editManagers" :key="'em-' + i" class="member-row">
          <el-input v-model="m.name" placeholder="姓名" style="width: 180px" />
          <el-input v-model="m.role" placeholder="职务" style="width: 220px; margin-left: 8px" />
          <el-button type="danger" :icon="Delete" circle size="small" @click="editManagers.splice(i, 1)" style="margin-left: 8px" />
        </div>
        <el-button type="primary" plain size="small" @click="editManagers.push({ name: '', member_type: 'manager', role: '' })">
          <el-icon><Plus /></el-icon>添加管理人员
        </el-button>

        <!-- 编辑施工人员 -->
        <el-divider content-position="left">施工人员</el-divider>
        <div v-for="(w, i) in editWorkers" :key="'ew-' + i" class="member-row">
          <el-input v-model="w.name" placeholder="姓名" style="width: 180px" />
          <el-input v-model="w.role" placeholder="工种" style="width: 220px; margin-left: 8px" />
          <el-button type="danger" :icon="Delete" circle size="small" @click="editWorkers.splice(i, 1)" style="margin-left: 8px" />
        </div>
        <el-button type="primary" plain size="small" @click="editWorkers.push({ name: '', member_type: 'worker', role: '' })">
          <el-icon><Plus /></el-icon>添加施工人员
        </el-button>
      </el-form>
      <template #footer>
        <el-button @click="editVisible = false">取消</el-button>
        <el-button type="primary" :loading="saving" @click="handleSaveEdit">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, computed, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Delete, Plus, Upload, Download, Document } from '@element-plus/icons-vue'
import {
  getProjects, updateProject, deleteProject, getProjectTypes,
  importProjects, exportProjects, downloadImportTemplate, autoSchedule
} from '../api'

const router = useRouter()

// 加载状态
const loading = ref(false)
const saving = ref(false)

// 搜索与筛选
const searchKeyword = ref('')
const filterType = ref('')

// 工程部位类别选项
const projectTypeOptions = ref([])

// 分页
const currentPage = ref(1)
const pageSize = ref(10)

// 全部数据
const allData = ref([])

// 筛选后的数据
const filteredData = computed(() => {
  let data = allData.value
  if (searchKeyword.value) {
    const kw = searchKeyword.value.toLowerCase()
    data = data.filter(
      (item) =>
        (item.project_code || '').toLowerCase().includes(kw) ||
        (item.project_name || '').toLowerCase().includes(kw)
    )
  }
  if (filterType.value) {
    data = data.filter((item) => (item.project_type || '').split(',').includes(filterType.value))
  }
  return data
})

// 当前页数据
const pagedData = computed(() => {
  const start = (currentPage.value - 1) * pageSize.value
  return filteredData.value.slice(start, start + pageSize.value)
})

// ========== 导入相关 ==========
const importDialogVisible = ref(false)
const importUploadRef = ref(null)
const importFile = ref(null)
const importResult = ref(null)
const importing = ref(false)

const handleImportFileChange = (file) => {
  importFile.value = file.raw
}

const resetImportState = () => {
  importFile.value = null
  importResult.value = null
  if (importUploadRef.value) {
    importUploadRef.value.clearFiles()
  }
}

const handleImportPreview = async () => {
  if (!importFile.value) return
  importing.value = true
  try {
    const res = await importProjects(importFile.value)
    importResult.value = res
    if (res.fail_count === 0) {
      ElMessage.success(`校验通过，共 ${res.success_count} 条数据`)
    } else {
      ElMessage.warning(`校验完成：${res.success_count} 条成功，${res.fail_count} 条失败`)
    }
  } catch (error) {
    ElMessage.error('导入校验失败')
    console.error('导入校验失败：', error)
  } finally {
    importing.value = false
  }
}

const handleConfirmImport = () => {
  importDialogVisible.value = false
  loadData()
  ElMessage.success('导入成功')
}

// ========== 导出 ==========
const handleExport = async () => {
  try {
    const res = await exportProjects()
    const blob = new Blob([res], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = window.URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `工程列表_${new Date().toISOString().slice(0, 10)}.xlsx`
    link.click()
    window.URL.revokeObjectURL(url)
    ElMessage.success('导出成功')
  } catch (error) {
    ElMessage.error('导出失败')
    console.error('导出工程失败：', error)
  }
}

// ========== 下载模板 ==========
const handleDownloadTemplate = async () => {
  try {
    const res = await downloadImportTemplate()
    const blob = new Blob([res], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = window.URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = '工程导入模板.xlsx'
    link.click()
    window.URL.revokeObjectURL(url)
    ElMessage.success('模板下载成功')
  } catch (error) {
    ElMessage.error('下载模板失败')
    console.error('下载模板失败：', error)
  }
}

// ========== 自动排期 ==========
const handleAutoSchedule = async (row) => {
  try {
    await ElMessageBox.confirm(`确定对工程"${row.project_name}"进行自动排期吗？`, '自动排期', {
      confirmButtonText: '确定',
      cancelButtonText: '取消',
      type: 'info'
    })
    loading.value = true
    await autoSchedule(row.id)
    ElMessage.success('自动排期完成，即将跳转到横道图页面')
    router.push({ name: 'GanttChart', query: { project_id: row.id } })
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('自动排期失败')
      console.error('自动排期失败：', error)
    }
  } finally {
    loading.value = false
  }
}

// 详情弹窗
const detailVisible = ref(false)
const currentRow = ref(null)
const detailManagers = computed(() => (currentRow.value?.members || []).filter(m => m.member_type === 'manager'))
const detailWorkers = computed(() => (currentRow.value?.members || []).filter(m => m.member_type === 'worker'))

// 编辑弹窗
const editVisible = ref(false)
const editFormRef = ref(null)
const editCollapseActive = ref([])
const editForm = reactive({
  id: '',
  project_code: '',
  project_name: '',
  project_type: [],
  location: '',
  subcontractor: '',
  voltage_level: '',
  line_name: '',
  preparation_unit: '',
  work_task: '',
  start_date: '',
  end_date: '',
  description: '',
  survey_leader: '',
  survey_members: '',
  survey_date: '',
  survey_result: '',
  briefing_host: '',
  briefing_date: '',
  briefing_content: ''
})
const editManagers = reactive([])
const editWorkers = reactive([])
const editEquipments = reactive([])
const editQualityControls = reactive([])
const editApprovals = reactive([])
const editRules = {
  project_code: [{ required: true, message: '请输入项目编号', trigger: 'blur' }],
  project_name: [{ required: true, message: '请输入项目名称', trigger: 'blur' }],
  project_type: [{ required: true, type: 'array', min: 1, message: '请选择工程类别', trigger: 'change' }],
  location: [{ required: true, message: '请输入工程地点', trigger: 'blur' }],
  subcontractor: [{ required: true, message: '请输入分包单位', trigger: 'blur' }]
}

// 加载数据
const loadData = async () => {
  loading.value = true
  try {
    const res = await getProjects()
    allData.value = res.items || []
  } catch (error) {
    console.error('加载工程列表失败：', error)
  } finally {
    loading.value = false
  }
}

// 搜索
const handleSearch = () => {
  currentPage.value = 1
}

// 查看详情
const viewDetail = (row) => {
  currentRow.value = row
  detailVisible.value = true
}

// 编辑
const editProject = (row) => {
  Object.assign(editForm, {
    id: row.id,
    project_code: row.project_code,
    project_name: row.project_name,
    project_type: (row.project_type || '').split(',').filter(Boolean),
    location: row.location,
    subcontractor: row.subcontractor,
    voltage_level: row.voltage_level || '',
    line_name: row.line_name || '',
    preparation_unit: row.preparation_unit || '',
    work_task: row.work_task || '',
    start_date: row.start_date,
    end_date: row.end_date,
    description: row.description,
    survey_leader: row.survey_leader || '',
    survey_members: row.survey_members || '',
    survey_date: row.survey_date || '',
    survey_result: row.survey_result || '',
    briefing_host: row.briefing_host || '',
    briefing_date: row.briefing_date || '',
    briefing_content: row.briefing_content || ''
  })
  // 填充人员
  const members = row.members || []
  editManagers.splice(0, editManagers.length, ...members.filter(m => m.member_type === 'manager').map(m => ({ ...m })))
  editWorkers.splice(0, editWorkers.length, ...members.filter(m => m.member_type === 'worker').map(m => ({ ...m })))
  // 填充机具、质量控制点、审批
  editEquipments.splice(0, editEquipments.length, ...(row.equipments || []).map(e => ({ ...e })))
  editQualityControls.splice(0, editQualityControls.length, ...(row.quality_controls || []).map(q => ({ ...q })))
  editApprovals.splice(0, editApprovals.length, ...(row.approvals || []).map(a => ({ ...a })))
  editCollapseActive.value = []
  editVisible.value = true
}

// 保存编辑
const handleSaveEdit = async () => {
  if (!editFormRef.value) return
  try {
    await editFormRef.value.validate()
  } catch {
    return
  }
  saving.value = true
  try {
    const members = [...editManagers, ...editWorkers].filter(m => m.name.trim())
    await updateProject(editForm.id, {
      ...editForm,
      project_type: editForm.project_type.join(','),
      members,
      equipments: editEquipments.filter(e => e.name.trim()),
      quality_controls: editQualityControls.filter(q => q.control_point.trim()),
      approvals: editApprovals.filter(a => a.approver.trim())
    })
    ElMessage.success('更新成功')
    editVisible.value = false
    loadData()
  } catch (error) {
    ElMessage.error('更新失败')
    console.error('更新工程失败：', error)
  } finally {
    saving.value = false
  }
}

// 删除
const handleDelete = async (row) => {
  try {
    await ElMessageBox.confirm(`确定删除工程"${row.project_name}"吗？`, '提示', {
      confirmButtonText: '确定',
      cancelButtonText: '取消',
      type: 'warning'
    })
    await deleteProject(row.id)
    ElMessage.success('删除成功')
    loadData()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败')
      console.error('删除工程失败：', error)
    }
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

onMounted(() => {
  loadData()
  loadProjectTypes()
})
</script>

<style scoped>
.project-list {
  max-width: 1200px;
}

.search-card {
  margin-bottom: 16px;
  border-radius: 8px;
}

.table-card {
  border-radius: 8px;
}

.pagination-wrap {
  display: flex;
  justify-content: flex-end;
  margin-top: 16px;
}

.detail-member-title {
  font-size: 14px;
  font-weight: 600;
  color: #303133;
  margin: 16px 0 8px;
}

.no-member-tip {
  color: #909399;
  font-size: 13px;
}

.member-row {
  display: flex;
  align-items: center;
  margin-bottom: 8px;
}
</style>
