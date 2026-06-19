import { defineStore } from 'pinia'
import { getProjects } from '../api'

// 应用全局状态管理
const useAppStore = defineStore('app', {
  state: () => ({
    // 当前选中的工程
    currentProject: null,
    // 工程列表
    projects: []
  }),

  actions: {
    // 设置当前工程
    setCurrentProject(project) {
      this.currentProject = project
    },

    // 从后端获取工程列表
    async fetchProjects() {
      try {
        const res = await getProjects()
        this.projects = res.data || []
      } catch (error) {
        console.error('获取工程列表失败：', error)
      }
    }
  }
})

export default useAppStore
