<template>
  <div class="dashboard">
    <!-- 统计卡片 -->
    <el-row :gutter="20" class="stat-row">
      <el-col :span="8">
        <el-card shadow="hover" class="stat-card">
          <div class="stat-content">
            <div class="stat-info">
              <div class="stat-label">工程总数</div>
              <div class="stat-value">{{ stats.projectCount }}</div>
            </div>
            <el-icon size="48" color="#409EFF"><FolderOpened /></el-icon>
          </div>
        </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" class="stat-card">
          <div class="stat-content">
            <div class="stat-info">
              <div class="stat-label">已生成文档数</div>
              <div class="stat-value">{{ stats.documentCount }}</div>
            </div>
            <el-icon size="48" color="#67C23A"><Document /></el-icon>
          </div>
        </el-card>
      </el-col>
      <el-col :span="8">
        <el-card shadow="hover" class="stat-card">
          <div class="stat-content">
            <div class="stat-info">
              <div class="stat-label">本月新增</div>
              <div class="stat-value">{{ stats.monthlyCount }}</div>
            </div>
            <el-icon size="48" color="#E6A23C"><TrendCharts /></el-icon>
          </div>
        </el-card>
      </el-col>
    </el-row>

    <!-- 快捷操作 -->
    <el-card shadow="hover" class="section-card">
      <template #header>
        <span class="section-title">快捷操作</span>
      </template>
      <el-space wrap :size="12">
        <el-button type="primary" @click="$router.push('/project/new')">
          <el-icon><DocumentAdd /></el-icon>新建工程
        </el-button>
        <el-button type="success" @click="$router.push('/construction-design')">
          <el-icon><Notebook /></el-icon>生成施工组织设计
        </el-button>
        <el-button type="warning" @click="$router.push('/survey')">
          <el-icon><Search /></el-icon>生成勘察单
        </el-button>
        <el-button type="info" @click="$router.push('/tech-briefing')">
          <el-icon><Reading /></el-icon>生成技术交底
        </el-button>
        <el-button type="danger" @click="$router.push('/safety-briefing')">
          <el-icon><Warning /></el-icon>生成安全交底
        </el-button>
        <el-button type="warning" @click="$router.push('/gantt-chart')">
          <el-icon><Grid /></el-icon>施工进度横道图
        </el-button>
      </el-space>
    </el-card>

    <!-- 最近工程列表 -->
    <el-card shadow="hover" class="section-card">
      <template #header>
        <span class="section-title">最近工程</span>
      </template>
      <el-table :data="recentProjects" stripe style="width: 100%">
        <el-table-column prop="project_code" label="项目编号" width="160" />
        <el-table-column prop="project_name" label="项目名称" />
        <el-table-column prop="subcontractor" label="分包单位" />
        <el-table-column prop="project_type" label="工程类型" width="120" />
        <el-table-column prop="start_date" label="开工日期" width="120" />
        <el-table-column label="操作" width="100">
          <template #default="{ row }">
            <el-button type="primary" link size="small" @click="viewProject(row)">查看</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { getProjects } from '../api'
import useAppStore from '../store'

const router = useRouter()
const store = useAppStore()

// 统计数据
const stats = ref({
  projectCount: 0,
  documentCount: 0,
  monthlyCount: 0
})

// 最近工程列表
const recentProjects = ref([])

// 查看工程详情
const viewProject = (row) => {
  store.setCurrentProject(row)
  router.push('/project/list')
}

// 加载统计数据
const loadStats = async () => {
  try {
    const res = await getProjects()
    const projects = res.items || []
    recentProjects.value = projects.slice(0, 10)
    stats.value.projectCount = res.total || projects.length
    // 计算本月新增
    const now = new Date()
    const thisMonth = projects.filter((p) => {
      const d = new Date(p.created_at || p.start_date)
      return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear()
    })
    stats.value.monthlyCount = thisMonth.length
  } catch (error) {
    console.error('加载统计数据失败：', error)
  }
}

onMounted(() => {
  loadStats()
})
</script>

<style scoped>
.dashboard {
  max-width: 1200px;
}

.stat-row {
  margin-bottom: 20px;
}

.stat-card {
  border-radius: 8px;
}

.stat-content {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.stat-info {
  display: flex;
  flex-direction: column;
}

.stat-label {
  font-size: 14px;
  color: #909399;
  margin-bottom: 8px;
}

.stat-value {
  font-size: 32px;
  font-weight: 700;
  color: #303133;
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
</style>