<template>
  <el-container class="layout-container">
    <!-- 左侧侧边栏 -->
    <el-aside width="220px" class="layout-aside">
      <!-- 系统名称 -->
      <div class="aside-title">
        <el-icon size="24"><OfficeBuilding /></el-icon>
        <span>施工资料管理系统</span>
      </div>
      <!-- 导航菜单 -->
      <el-menu
        :default-active="activeMenu"
        class="aside-menu"
        background-color="#304156"
        text-color="#bfcbd9"
        active-text-color="#409EFF"
        router
      >
        <el-menu-item index="/dashboard">
          <el-icon><DataBoard /></el-icon>
          <span>首页总览</span>
        </el-menu-item>
        <el-menu-item index="/project/new">
          <el-icon><DocumentAdd /></el-icon>
          <span>新建工程</span>
        </el-menu-item>
        <el-menu-item index="/project/list">
          <el-icon><FolderOpened /></el-icon>
          <span>工程查阅</span>
        </el-menu-item>
        <el-menu-item index="/construction-design">
          <el-icon><Notebook /></el-icon>
          <span>施工组织设计</span>
        </el-menu-item>
        <el-menu-item index="/survey">
          <el-icon><Search /></el-icon>
          <span>项目勘察单</span>
        </el-menu-item>
        <el-menu-item index="/tech-briefing">
          <el-icon><Reading /></el-icon>
          <span>技术交底</span>
        </el-menu-item>
        <el-menu-item index="/safety-briefing">
          <el-icon><Warning /></el-icon>
          <span>安全交底</span>
        </el-menu-item>
        <el-menu-item index="/data-foundation">
          <el-icon><Coin /></el-icon>
          <span>数据底座</span>
        </el-menu-item>
        <el-menu-item index="/gantt-chart">
          <el-icon><TrendCharts /></el-icon>
          <span>施工进度横道图</span>
        </el-menu-item>
      </el-menu>
    </el-aside>

    <!-- 右侧内容区 -->
    <el-container class="layout-main">
      <!-- 顶部栏 -->
      <el-header class="layout-header">
        <span class="header-title">{{ pageTitle }}</span>
      </el-header>
      <!-- 主内容区 -->
      <el-main class="layout-content">
        <router-view />
      </el-main>
    </el-container>
  </el-container>
</template>

<script setup>
import { computed } from 'vue'
import { useRoute } from 'vue-router'

const route = useRoute()

// 当前激活的菜单项
const activeMenu = computed(() => route.path)

// 页面标题
const pageTitle = computed(() => route.meta?.title || '施工资料管理系统')
</script>

<style scoped>
.layout-container {
  height: 100vh;
}

.layout-aside {
  background-color: #304156;
  overflow-y: auto;
}

.aside-title {
  height: 60px;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  color: #fff;
  font-size: 16px;
  font-weight: 600;
  border-bottom: 1px solid rgba(255, 255, 255, 0.1);
  white-space: nowrap;
}

.aside-menu {
  border-right: none;
}

.layout-main {
  background-color: #f0f2f5;
}

.layout-header {
  background-color: #fff;
  display: flex;
  align-items: center;
  box-shadow: 0 1px 4px rgba(0, 21, 41, 0.08);
  z-index: 1;
}

.header-title {
  font-size: 18px;
  font-weight: 600;
  color: #303133;
}

.layout-content {
  padding: 20px;
  overflow-y: auto;
}
</style>
