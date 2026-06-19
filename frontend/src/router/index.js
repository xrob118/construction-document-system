import { createRouter, createWebHistory } from 'vue-router'
import Layout from '../views/Layout.vue'

// 路由配置
const routes = [
  {
    path: '/',
    component: Layout,
    redirect: '/dashboard',
    children: [
      {
        path: 'dashboard',
        name: 'Dashboard',
        component: () => import('../views/Dashboard.vue'),
        meta: { title: '首页总览' }
      },
      {
        path: 'project/new',
        name: 'ProjectForm',
        component: () => import('../views/ProjectForm.vue'),
        meta: { title: '新建工程' }
      },
      {
        path: 'project/list',
        name: 'ProjectList',
        component: () => import('../views/ProjectList.vue'),
        meta: { title: '工程查阅' }
      },
      {
        path: 'construction-design',
        name: 'ConstructionDesign',
        component: () => import('../views/ConstructionDesign.vue'),
        meta: { title: '施工组织设计' }
      },
      {
        path: 'survey',
        name: 'SurveyForm',
        component: () => import('../views/SurveyForm.vue'),
        meta: { title: '项目勘察单' }
      },
      {
        path: 'tech-briefing',
        name: 'TechBriefing',
        component: () => import('../views/TechBriefing.vue'),
        meta: { title: '技术交底' }
      },
      {
        path: 'safety-briefing',
        name: 'SafetyBriefing',
        component: () => import('../views/SafetyBriefing.vue'),
        meta: { title: '安全交底' }
      },
      {
        path: 'data-foundation',
        name: 'DataFoundation',
        component: () => import('../views/DataFoundation.vue'),
        meta: { title: '数据底座' }
      },
      {
        path: 'gantt-chart',
        name: 'GanttChart',
        component: () => import('../views/GanttChart.vue'),
        meta: { title: '施工进度横道图' }
      }
    ]
  }
]

const router = createRouter({
  history: createWebHistory(),
  routes
})

export default router
