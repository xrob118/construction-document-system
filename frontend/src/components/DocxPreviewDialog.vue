<template>
  <el-dialog
    v-model="visible"
    title="文档预览"
    width="90%"
    top="5vh"
    :close-on-click-modal="false"
    @close="handleClose"
    class="docx-preview-dialog"
  >
    <template #header>
      <div class="preview-header">
        <span class="title">文档预览</span>
        <div class="actions">
          <el-button-group>
            <el-button size="small" @click="zoomOut" :disabled="zoomLevel <= 0.5">
              <el-icon><ZoomOut /></el-icon>
            </el-button>
            <el-button size="small" disabled>{{ Math.round(zoomLevel * 100) }}%</el-button>
            <el-button size="small" @click="zoomIn" :disabled="zoomLevel >= 2">
              <el-icon><ZoomIn /></el-icon>
            </el-button>
          </el-button-group>
          <el-button size="small" @click="handleDownload">
            <el-icon><Download /></el-icon>下载
          </el-button>
          <el-button size="small" @click="handlePrint">
            <el-icon><Printer /></el-icon>打印
          </el-button>
        </div>
      </div>
    </template>

    <div class="preview-container" v-loading="loading" element-loading-text="正在解析文档...">
      <div v-if="errorMsg" class="preview-error">
        <el-empty :description="errorMsg" />
      </div>
      <div v-else class="pages-wrapper" :style="{ transform: `scale(${zoomLevel})`, transformOrigin: 'top center' }">
        <div
          v-for="(page, idx) in paginatedPages"
          :key="idx"
          class="a4-page"
          :class="{
            'is-cover': idx === 0 && props.coverPages >= 1,
            'is-approval': idx === 1 && props.coverPages >= 2,
            'is-section-header': idx >= props.coverPages && isSectionHeaderPage(idx)
          }"
        >
          <div class="page-content" v-html="page"></div>
          <div class="page-number">- {{ idx + 1 }} -</div>
        </div>
      </div>
      <!-- 隐藏的测量容器：用于计算各子元素实际高度 -->
      <div
        ref="measureRef"
        class="measure-container"
        v-html="fullHtml"
      ></div>
    </div>
  </el-dialog>
</template>

<script setup>
import { ref, watch, nextTick } from 'vue'
import { ElMessage } from 'element-plus'
import { ZoomIn, ZoomOut, Download, Printer } from '@element-plus/icons-vue'
import mammoth from 'mammoth'
import JSZip from 'jszip'

// 4 个 docx 模块通用的预览组件：施工组织设计 / 项目勘察单 / 技术交底 / 安全交底
// 通过 pageBreaks prop 自定义"哪些标题段需要强制分页到下一页"
// 留空 = 不强制分页（适用于内容较少的勘察单/交底单）
// coverPages：前 N 页应用"封面/审批表"上下分布样式（默认 2 = 封面 + 审批表）
//   施工组织设计：2（封面 + 审批表）
//   勘察单/技术交底/安全交底：0（没有独立封面页，只有抬头表格）
const props = defineProps({
  modelValue: { type: Boolean, default: false },
  docxBlob: { type: Blob, default: null },
  filename: { type: String, default: 'document.docx' },
  pageBreaks: { type: Array, default: () => [] },
  coverPages: { type: Number, default: 2 },
})

const emit = defineEmits(['update:modelValue', 'close'])

const visible = ref(props.modelValue)
const loading = ref(false)
const errorMsg = ref('')
const fullHtml = ref('')
const paginatedPages = ref([])
const zoomLevel = ref(0.85)
const measureRef = ref(null)
let imageBlobUrls = []

watch(() => props.modelValue, (v) => { visible.value = v })
watch(visible, (v) => emit('update:modelValue', v))

const handleClose = () => {
  visible.value = false
  emit('close')
}

const zoomIn = () => { zoomLevel.value = Math.min(2, +(zoomLevel.value + 0.1).toFixed(2)) }
const zoomOut = () => { zoomLevel.value = Math.max(0.5, +(zoomLevel.value - 0.1).toFixed(2)) }

const handleDownload = () => {
  if (!props.docxBlob) return
  const url = URL.createObjectURL(props.docxBlob)
  const a = document.createElement('a')
  a.href = url
  a.download = props.filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
  ElMessage.success('已下载原文件')
}

const handlePrint = () => {
  const w = window.open('', '_blank')
  if (!w) { ElMessage.warning('请允许弹窗以使用打印功能'); return }
  w.document.write(`
    <html><head><title>${props.filename}</title>
    <style>${printCss}</style>
    </head><body>${fullHtml.value}</body></html>
  `)
  w.document.close()
  w.focus()
  setTimeout(() => w.print(), 500)
}

// 1mm = 3.7795275591 px @ 96dpi
const MM_TO_PX = 3.7795275591
// A4 内容区域高度 = 297mm - 25mm 顶 - 25mm 底 = 247mm
const A4_CONTENT_HEIGHT_PX = 247 * MM_TO_PX
// A4 内容区域宽度 = 210mm - 25mm 左 - 25mm 右 = 160mm（统一 2.5cm 页边距）
const A4_CONTENT_WIDTH_PX = 160 * MM_TO_PX
// 单行高度 = fontSize(14) * lineHeight(1.5) = 21 px（与 .a4-page line-height: 1.5 一致）
const LINE_HEIGHT_PX = 14 * 1.5

/**
 * 检查某页是不是"只有章节标题几乎没有内容"的空白页
 * （如 六、安全措施 等 force-new-page 后第一个 chunk 还不够塞满一页）
 */
const isSectionHeaderPage = (idx) => {
  if (!paginatedPages.value || !paginatedPages.value[idx]) return false
  const html = paginatedPages.value[idx]
  // 去掉 <div class="page-filler"> 和 force-new-page 等辅助元素
  // 估算：除 <p> 外几乎没有可见块，且 p 数量少
  const pCount = (html.match(/<p[\s>]/g) || []).length
  const hasTable = /<table[\s>]/i.test(html)
  // 章节标题页：只有 1-2 个 <p> 且没有表格
  return pCount <= 2 && !hasTable
}

/**
 * 把 measureRef 容器中的 <table> 拆开为多个 <div class="table-row"> 顶层块，
 * 这样分页器就能像对待普通 <p> 一样逐行切分、跨页。
 * 适用场景：勘察单/技术交底/安全交底模板只有 1 张表，但表内某 <td> 塞了 200 行的"安全措施"，整张表 > 247mm。
 * 表头 + 数据行 1 行 = 1 个 div.table-row；如果该 <td> 内的 <p> 还是太高，再走 splitGiantBlocks 切分。
 *
 * 只拆"包含巨型 <p>"的表：small table（如审批表 9 行小数据）保留原表格结构，
 * 只在表内某 <p> 高度 > 247mm 时才拆开成 div.table-row 序列。
 */
const unpackTables = () => {
  if (!measureRef.value) return
  const tables = [...measureRef.value.querySelectorAll('table')]
  for (const table of tables) {
    // 检测 1：表内是否有 <p> 高度 > A4 内容高（巨型段落，可能换行成 50+ 行）
    let hasGiantP = false
    for (const p of table.querySelectorAll('p')) {
      if ((p.offsetHeight || 0) > A4_CONTENT_HEIGHT_PX) {
        hasGiantP = true
        break
      }
    }
    // 检测 2：整张表高度 > A4 内容高（9 行小数据表 + 5cm 行间距 = 1214px 也要拆）
    // 关键修复：force-new-page 页（封面/审批表）如果整表超高，必须拆行后才能跨页
    const tableH = table.offsetHeight || 0
    if (!hasGiantP && tableH <= A4_CONTENT_HEIGHT_PX) {
      // 短表：保留原表格结构（视觉更友好）
      continue
    }
    const rows = [...table.querySelectorAll('tr')]
    if (rows.length === 0) {
      table.remove()
      continue
    }
    const parent = table.parentNode
    const insertBeforeNode = table
    for (const tr of rows) {
      const wrap = document.createElement('div')
      wrap.className = 'table-row'
      const tds = [...tr.querySelectorAll('td, th')]
      for (const td of tds) {
        const cellDiv = document.createElement('div')
        cellDiv.className = 'table-cell'
        // 关键：把 td 的 inline 宽度继承到 cellDiv（保证拆表后列宽仍对齐 docx）
        const tdWidth = td.style.width || td.getAttribute('width') || ''
        if (tdWidth) {
          const px = tdWidth.endsWith('px') ? tdWidth : (parseFloat(tdWidth) || 0) + 'px'
          if (px && px !== '0px' && px !== 'NaNpx') {
            cellDiv.style.width = px
            cellDiv.style.minWidth = px
            cellDiv.style.maxWidth = px
          }
        }
        // 继承 colspan（用于后续拆分时判断）
        const colspan = td.getAttribute('colspan')
        if (colspan && parseInt(colspan) > 1) {
          cellDiv.setAttribute('colspan', colspan)
        }
        while (td.firstChild) {
          cellDiv.appendChild(td.firstChild)
        }
        wrap.appendChild(cellDiv)
      }
      parent.insertBefore(wrap, insertBeforeNode)
    }
    table.remove()
    console.log(`[DocxPreview] 拆表（因含巨型 <p>）: ${rows.length} 行 -> ${rows.length} 个 div.table-row`)
  }
}

/**
 * 把超 A4 的 div.table-row 进一步"炸开"为多个独立的 div.table-row
 * 触发条件：splitGiantBlocks 之后，某些 table-row 的高度仍 > 247mm（因为 1 个 cell 塞了 10+ 段 <p>）
 * 处理：把那个 cell 的多个 <p> 拉出来，每个 <p> 包成一个独立的 div.table-row（含外框），原 table-row 只保留"小数据行"内容
 */
const explodeOversizedTableRows = () => {
  if (!measureRef.value) return
  const rows = [...measureRef.value.querySelectorAll(':scope > .table-row')]
  let explodedCount = 0
  for (const row of rows) {
    const height = row.offsetHeight || 0
    if (height <= A4_CONTENT_HEIGHT_PX) continue
    const cells = [...row.querySelectorAll('.table-cell')]
    // 找到含 <p> 数量最多的那个 cell
    let bigCell = null
    let maxP = 0
    for (const cell of cells) {
      const ps = cell.querySelectorAll('p')
      if (ps.length > maxP) {
        maxP = ps.length
        bigCell = cell
      }
    }
    if (!bigCell || maxP < 2) {
      // 没法炸开，整行单独占一页
      console.log(`[DocxPreview] 警告: div.table-row 高度=${height} 仍 > 247mm, 整行放一页`)
      continue
    }
    // 把 bigCell 的 <p> 逐个抽出
    const ps = [...bigCell.querySelectorAll('p')]
    const parent = row.parentNode
    const insertBeforeNode = row
    // 在原 row 前插入多个新 table-row（每个含 1 个 <p>）
    for (const p of ps) {
      const newRow = document.createElement('div')
      newRow.className = 'table-row exploded'
      const newCell = document.createElement('div')
      newCell.className = 'table-cell'
      newCell.appendChild(p.cloneNode(true))
      newRow.appendChild(newCell)
      parent.insertBefore(newRow, insertBeforeNode)
      explodedCount++
    }
    // 清空原 row 的 bigCell，保留"标签"段
    bigCell.innerHTML = ''
    console.log(`[DocxPreview] 炸开 div.table-row: ${ps.length} <p> -> ${ps.length} 个独立 div.table-row`)
  }
  if (explodedCount > 0) {
    console.log(`[DocxPreview] 共炸开 ${explodedCount} 个独立 div.table-row`)
  }
}

/**
 * 把超 A4 的巨型 <p> 段落（一般是 mammoth 把多行文本输出为单个 <p><br/><br/>...）按 <br> 拆成多个小 <p>。
 * 拆分后每个小 <p> 的高度都 < 247mm，能正常分页。
 * 递归处理：包括表格单元格内的巨型段落（勘察单/交底单的 mammoth 输出经常把内容塞进单个 <td>）。
 * 修改 measureRef 容器中的 DOM 元素，原地替换。
 */
const splitGiantBlocks = () => {
  if (!measureRef.value) return
  // 实际经验值：<p> 内的 <br> 分割的"逻辑行"在 160mm 内容宽度下经常换行成 2-3 个视觉行；
  // lineHeight=25.2px, 一页 933px ≈ 37 个单倍行 / 12 个三倍行。
  // 用更激进的拆分行数（除以 3 而不是 2），每段 6 行 ≈ 12 个三倍行 ≈ 302px，留 ~600px 余量
  // —— 即便某些段全部按 3 倍行换行（极少数），仍不会超 933px。
  const linesPerPage = Math.max(2, Math.floor(A4_CONTENT_HEIGHT_PX / LINE_HEIGHT_PX / 3) - 1)
  // 收集所有 <p>，包括 <table> 内的（递归遍历）
  const allPs = measureRef.value.querySelectorAll('p')
  for (const el of allPs) {
    const height = el.offsetHeight || 0
    if (height <= A4_CONTENT_HEIGHT_PX) continue
    // 收集按 <br> 分割的"行"
    const lines = []
    let currentLine = []
    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType === 1 && child.tagName === 'BR') {
        if (currentLine.length > 0) {
          lines.push(currentLine)
          currentLine = []
        }
      } else {
        currentLine.push(child)
      }
    }
    if (currentLine.length > 0) lines.push(currentLine)
    if (lines.length <= linesPerPage) continue
    // 拆分为多段
    const fragments = []
    for (let i = 0; i < lines.length; i += linesPerPage) {
      const chunkLines = lines.slice(i, i + linesPerPage)
      const newEl = document.createElement(el.tagName)
      // 复制属性（包括 class）
      for (const attr of el.attributes) {
        newEl.setAttribute(attr.name, attr.value)
      }
      // 重要：所有 chunk 都不能保留 force-new-page（否则首个 chunk 又会触发分页）
      // 续段不保留 docx-heading 样式（改为普通正文缩进）
      newEl.classList.remove('force-new-page')
      if (i > 0) {
        newEl.classList.remove('docx-heading')
      }
      for (let j = 0; j < chunkLines.length; j++) {
        for (const node of chunkLines[j]) {
          newEl.appendChild(node.cloneNode(true))
        }
        if (j < chunkLines.length - 1) {
          newEl.appendChild(document.createElement('br'))
        }
      }
      fragments.push(newEl)
    }
    // 在原父节点中替换
    const parent = el.parentNode
    for (const frag of fragments) {
      parent.insertBefore(frag, el)
    }
    parent.removeChild(el)
    console.log(`[DocxPreview] 拆分巨型 <p> (in ${parent.tagName}.${parent.className || ''}): ${lines.length} 行 -> ${fragments.length} 段 (每段约 ${linesPerPage} 行)`)
  }
}

/**
 * 等待 measureRef 中所有 <img> 加载完成
 */
const waitImagesLoaded = async () => {
  if (!measureRef.value) return
  const imgs = measureRef.value.querySelectorAll('img')
  if (imgs.length === 0) return
  await Promise.all([...imgs].map(img => {
    if (img.complete && img.naturalHeight > 0) return Promise.resolve()
    return new Promise(resolve => {
      img.onload = () => resolve()
      img.onerror = () => resolve()
    })
  }))
}

/**
 * 把超 A4 的"块 outerHTML"按 <br> 拆成多个小段（fallback 给 splitIntoPages 用）
 * 与 splitGiantBlocks 不同：本函数接收 outerHTML 字符串，输出多个 outerHTML 字符串
 * 用于在 splitIntoPages 主循环里处理"虽然 splitGiantBlocks 处理过、但仍超高的单块"。
 *
 * @param {string} outerHtml - 块的完整 outerHTML
 * @param {number} maxHeightPx - 目标单段最大高度
 * @returns {string[]} - 拆分后的多段 outerHTML（若无 <br> 可拆，返回原数组）
 */
const splitBlockByBr = (outerHtml, maxHeightPx) => {
  try {
    const container = document.createElement('div')
    container.style.cssText = 'position:absolute;visibility:hidden;left:0;top:0;width:160mm;'
    container.innerHTML = outerHtml
    document.body.appendChild(container)
    const linesPerPage = Math.max(2, Math.floor(maxHeightPx / LINE_HEIGHT_PX / 3) - 1)
    const result = []
    // 处理 <p> 内的 <br> 切分
    const ps = [...container.querySelectorAll('p')]
    for (const p of ps) {
      const h = p.offsetHeight || 0
      if (h <= maxHeightPx) {
        result.push(p.outerHTML)
        continue
      }
      // 收集按 <br> 分割的"行"
      const lines = []
      let cur = []
      for (const child of Array.from(p.childNodes)) {
        if (child.nodeType === 1 && child.tagName === 'BR') {
          if (cur.length > 0) { lines.push(cur); cur = [] }
        } else {
          cur.push(child)
        }
      }
      if (cur.length > 0) lines.push(cur)
      if (lines.length <= 1) {
        result.push(p.outerHTML)
        continue
      }
      for (let i = 0; i < lines.length; i += linesPerPage) {
        const chunkLines = lines.slice(i, i + linesPerPage)
        const newP = document.createElement('p')
        for (let j = 0; j < chunkLines.length; j++) {
          for (const node of chunkLines[j]) newP.appendChild(node.cloneNode(true))
          if (j < chunkLines.length - 1) newP.appendChild(document.createElement('br'))
        }
        result.push(newP.outerHTML)
      }
    }
    // 处理 div.table-row 的拆分（按内部 p 的 br 切）
    const rows = [...container.querySelectorAll('div.table-row')]
    for (const row of rows) {
      const h = row.offsetHeight || 0
      if (h <= maxHeightPx) {
        result.push(row.outerHTML)
        continue
      }
      // 找含 <p> 最多的 cell
      let bigCell = null
      let maxP = 0
      for (const cell of row.querySelectorAll('.table-cell')) {
        const ps2 = cell.querySelectorAll('p')
        if (ps2.length > maxP) { maxP = ps2.length; bigCell = cell }
      }
      if (!bigCell || maxP < 2) {
        result.push(row.outerHTML)
        continue
      }
      // 抽出该 cell 的所有 p，每个 p 一个独立 row
      const cellPs = [...bigCell.querySelectorAll('p')]
      for (const p of cellPs) {
        const newRow = document.createElement('div')
        newRow.className = 'table-row exploded'
        const newCell = document.createElement('div')
        newCell.className = 'table-cell'
        newCell.appendChild(p.cloneNode(true))
        newRow.appendChild(newCell)
        result.push(newRow.outerHTML)
      }
    }
    document.body.removeChild(container)
    return result
  } catch (e) {
    console.warn('[DocxPreview] splitBlockByBr 失败:', e)
    return [outerHtml]
  }
}

/**
 * 测量一段 outerHTML 字符串的视觉高度（临时渲染到隐藏容器）
 */
const measureFragmentHeight = (outerHtml) => {
  try {
    const c = document.createElement('div')
    c.style.cssText = 'position:absolute;visibility:hidden;left:0;top:0;width:160mm;'
    c.innerHTML = outerHtml
    document.body.appendChild(c)
    const h = c.offsetHeight || 0
    document.body.removeChild(c)
    return h
  } catch {
    return 0
  }
}

/**
 * 从 docx blob 中提取所有 <w:tblGrid> 列宽（twips）
 * mammoth 默认不保留 tblGrid 列宽到 HTML，必须从原始 docx XML 中读取
 * 返回数组：[grid_for_table_1, grid_for_table_2, ...]，每个 grid 是 twips 数字数组
 */
const extractTableGrids = async (arrayBuffer) => {
  try {
    const zip = await JSZip.loadAsync(arrayBuffer)
    const docFile = zip.file('word/document.xml')
    if (!docFile) return []
    const docXml = await docFile.async('string')
    const parser = new DOMParser()
    const xmlDoc = parser.parseFromString(docXml, 'application/xml')
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    const tbls = xmlDoc.getElementsByTagNameNS(W, 'tbl')
    const grids = []
    for (const tbl of tbls) {
      const grid = tbl.getElementsByTagNameNS(W, 'tblGrid')[0]
      if (!grid) { grids.push([]); continue }
      const cols = grid.getElementsByTagNameNS(W, 'gridCol')
      const widths = []
      for (const col of cols) {
        const w = col.getAttributeNS(W, 'w') || '0'
        widths.push(parseInt(w, 10))
      }
      grids.push(widths)
    }
    return grids
  } catch (e) {
    console.warn('[DocxPreview] 提取 tblGrid 失败:', e)
    return []
  }
}

/**
 * 把 docx 的 tblGrid 列宽按比例映射到预览内容宽度（605px），注入到 mammoth HTML 的 <td>/<th> 上
 * - 处理 colspan（合并多列宽）
 * - 处理 rowspan（用 layout 跟踪被占用的列）
 * - 对每张表独立处理，按表格的实际 total twips 缩放
 *
 * @param {string} html - mammoth 输出的 HTML
 * @param {number[][]} grids - 每张表的列宽（twips）
 * @param {number} contentWidthPx - 预览内容区宽度（默认 605 = 160mm @ 96dpi）
 */
const injectTableWidths = (html, grids, contentWidthPx = 605) => {
  if (!grids || grids.length === 0) return html
  const wrap = document.createElement('div')
  wrap.innerHTML = html
  const tables = [...wrap.querySelectorAll('table')]
  for (let i = 0; i < tables.length && i < grids.length; i++) {
    const tbl = tables[i]
    const grid = grids[i]
    if (!grid || grid.length === 0) continue
    const totalTwips = grid.reduce((a, b) => a + b, 0)
    if (totalTwips <= 0) continue
    const pxPerTwip = contentWidthPx / totalTwips
    // 用 layout 矩阵跟踪 rowspan 占用
    const rows = [...tbl.querySelectorAll('tr')]
    const layout = rows.map(() => new Array(grid.length).fill(false))
    for (let r = 0; r < rows.length; r++) {
      const cells = [...rows[r].querySelectorAll('td, th')]
      let colIdx = 0
      for (const cell of cells) {
        // 跳过被 rowspan 占据的列
        while (colIdx < grid.length && layout[r][colIdx]) colIdx++
        if (colIdx >= grid.length) break
        const colspan = parseInt(cell.getAttribute('colspan') || '1', 10) || 1
        const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10) || 1
        // 计算该 cell 跨越的列总宽
        let widthTwips = 0
        for (let c = 0; c < colspan && colIdx + c < grid.length; c++) {
          widthTwips += grid[colIdx + c] || 0
        }
        const widthPx = Math.max(20, Math.round(widthTwips * pxPerTwip))
        // 标记 layout 占用
        for (let dr = 0; dr < rowspan && r + dr < rows.length; dr++) {
          for (let dc = 0; dc < colspan && colIdx + dc < grid.length; dc++) {
            layout[r + dr][colIdx + dc] = true
          }
        }
        // 注入 inline 宽度（style 优先于 attribute，CSS 才会真的生效）
        cell.setAttribute('width', widthPx)
        cell.style.width = widthPx + 'px'
        cell.style.minWidth = widthPx + 'px'
        cell.style.maxWidth = widthPx + 'px'
        colIdx += colspan
      }
    }
  }
  return wrap.innerHTML
}

/**
 * 把 mammoth HTML 拆分成多页 A4
 * 规则：
 * 1. 逐个遍历段落/表格
 * 2. 累计高度超过 A4 内容高度时分页
 * 3. 单独的段落/表格如果超过整页高度，按 <br> 拆碎后分多页（避免腰斩小段落 + 避免被 overflow:hidden 裁掉）
 * 4. 不可分割的块（如图片/标题/小段落）用 break-inside: avoid
 */
const splitIntoPages = async () => {
  if (!fullHtml.value || !measureRef.value) {
    paginatedPages.value = []
    return
  }

  // 1. 等待所有图片加载（确保 offsetHeight 准确）
  await waitImagesLoaded()
  // 等一帧让浏览器应用布局
  await nextTick()

  // 1.4. 拆开 <table> 为多个顶层 <div class="table-row">
  // 适用场景：勘察单/技术交底/安全交底模板的"安全措施"塞在单个 <td> 内，整张表 > 247mm
  // 拆开后每个 <tr> 变成顶层块，分页器就能逐行切分跨页
  unpackTables()
  await nextTick()

  // 1.5. 拆分巨型 <p> 段落（高度 > 247mm）到多个小 <p>
  splitGiantBlocks()
  // 拆分后浏览器需要重新布局，再等一帧拿新高度
  await nextTick()
  await new Promise(r => requestAnimationFrame(r))
  await nextTick()

  // 1.6. 炸开仍超 A4 的 div.table-row（cell 塞了 10+ 段 <p>，整行仍 > 247mm）
  explodeOversizedTableRows()
  await nextTick()
  await new Promise(r => requestAnimationFrame(r))
  await nextTick()

  // 调试：检查 measure 容器状态
  const mc = measureRef.value
  console.log(`[DocxPreview] measure 容器: 自身高度=${mc.offsetHeight}, 子元素数=${mc.children.length}, display=${getComputedStyle(mc).display}, visibility=${getComputedStyle(mc).visibility}, width=${mc.offsetWidth}`)

  // 2. 把 mammoth 输出的子元素克隆到一个临时测量器
  // 收集所有顶层元素（<p>/<table>）
  const blocks = []
  const containerChildren = [...measureRef.value.children]
  let skipCount = 0
  for (const el of containerChildren) {
    const h = el.offsetHeight || 0
    const isForceNew = el.classList && el.classList.contains('force-new-page')
    const isHeading = el.classList && el.classList.contains('docx-heading')
    // 关键修复1：过滤掉"完全空白的 <p>"——mammoth 会把 docx 里用作行间距的 <w:p/> 输出为 <p></p>，
    // 这些段没有文字、没有图片、没有 <br>，只会制造空白页和大段空白（p41 7个空段就是这种）
    if (el.tagName === 'P') {
      const text = (el.textContent || '').trim()
      const hasImg = el.querySelector('img')
      const hasBr = el.querySelector('br')
      const hasForcedBreak = el.classList && el.classList.contains('force-new-page')
      if (!text && !hasImg && !hasBr && !hasForcedBreak) {
        skipCount++
        continue
      }
    }
    if (h > A4_CONTENT_HEIGHT_PX) {
      console.log(`[DocxPreview] 警告: 块 ${el.tagName}.${el.className} 高度=${h} 仍 > 247mm!`)
    }
    if (isHeading || isForceNew) {
      console.log(`[DocxPreview] 块: ${el.tagName} h=${h} forceNew=${isForceNew} heading=${isHeading} text="${el.textContent.substring(0, 30).replace(/\n/g, ' ')}"`)
    }
    blocks.push({
      outerHTML: el.outerHTML,
      height: h,
      forceNewPage: isForceNew,
    })
  }
  console.log(`[DocxPreview] 过滤空白段: 跳过 ${skipCount} 个`)

  if (blocks.length === 0) {
    paginatedPages.value = []
    return
  }

  // 3. 拆分
  const pages = []
  let current = []
  let currentHeight = 0
  console.log(`[DocxPreview] === 开始分页: ${blocks.length} 个块 ===`)

  for (const block of blocks) {
    // 强制分页点（封面/审批表/正文分界）：先把当前页 flush
    if (block.forceNewPage && current.length > 0) {
      // 跳过全空白页（mammoth 输出的空块/失败图片占位）
      if (currentHeight < 20) {
        current = [block.outerHTML]
        currentHeight = block.height
        console.log(`[DocxPreview] force-new-page 跳过空白: block=${block.outerHTML.substring(0,40)} h=${block.height}`)
        continue
      }
      // 封面/审批表等"有上下分布语义"的页用 flex：第一个元素顶对齐、最后一个元素底对齐
      console.log(`[DocxPreview] force-new-page flush: currentHeight=${currentHeight}, blocks=${current.length}`)
      pages.push(`<div class="page-flex">${current.join('')}</div>`)
      current = [block.outerHTML]
      currentHeight = block.height
      continue
    }
    // 块自身就超过一页：单放一页（仍可能溢出，但避免腰斩小段落）
    if (block.height > A4_CONTENT_HEIGHT_PX) {
      // 先 flush 当前页
      if (current.length > 0) {
        if (currentHeight < 20) {
          // 全空白，跳过
        } else {
          const remaining = A4_CONTENT_HEIGHT_PX - currentHeight
          if (remaining > 40) {
            current.push(`<div class="page-filler" style="height:${remaining}px"></div>`)
          }
          pages.push(`<div class="page-flex">${current.join('')}</div>`)
        }
        current = []
        currentHeight = 0
      }
      // 关键：单块超高时不再"整块塞一页被裁掉"，
      // 而是把该块按 <br> 拆成多个小段后，递归走正常的分页流程
      const splitFragments = splitBlockByBr(block.outerHTML, A4_CONTENT_HEIGHT_PX)
      if (splitFragments.length > 1) {
        // 把拆出的小段当独立块重新走分页
        for (const frag of splitFragments) {
          const fragHeight = measureFragmentHeight(frag)
          if (currentHeight + fragHeight > A4_CONTENT_HEIGHT_PX && current.length > 0) {
            const remaining = A4_CONTENT_HEIGHT_PX - currentHeight
            if (remaining > 40) {
              current.push(`<div class="page-filler" style="height:${remaining}px"></div>`)
            }
            pages.push(`<div class="page-flex">${current.join('')}</div>`)
            current = [frag]
            currentHeight = fragHeight
          } else {
            current.push(frag)
            currentHeight += fragHeight
          }
        }
        continue
      }
      // 真的拆不开（<br> 少于 1 个）：整块放一页（兜底）
      pages.push(`<div class="page-flex">${block.outerHTML}</div>`)
      continue
    }
    // 当前块追加后会超出一页：换页
    if (currentHeight + block.height > A4_CONTENT_HEIGHT_PX && current.length > 0) {
      // 换页前，撑开当前页（短页填充）
      if (currentHeight >= 20) {
        const remaining = A4_CONTENT_HEIGHT_PX - currentHeight
        if (remaining > 40) {
          current.push(`<div class="page-filler" style="height:${remaining}px"></div>`)
        }
        pages.push(`<div class="page-flex">${current.join('')}</div>`)
        console.log(`[DocxPreview] 换页 flush: currentHeight=${currentHeight}, blocks=${current.length}`)
      }
      // 跳过空白页：currentHeight < 20 时直接丢弃旧 current
      current = [block.outerHTML]
      currentHeight = block.height
    } else {
      current.push(block.outerHTML)
      currentHeight += block.height
    }
  }
  if (current.length > 0 && currentHeight >= 20) {
    // 最后一页：与循环内的换页逻辑保持一致的"短页填充"处理
    const remaining = A4_CONTENT_HEIGHT_PX - currentHeight
    if (remaining > 40) {
      current.push(`<div class="page-filler" style="height:${remaining}px"></div>`)
    }
    pages.push(`<div class="page-flex">${current.join('')}</div>`)
    console.log(`[DocxPreview] 最后一页 flush: currentHeight=${currentHeight}, blocks=${current.length}`)
  }

  // 4. post-check: 检查每页内部是否仍超出 A4（拆分/换页逻辑没处理干净的兜底）
  //    对溢出页：把 page-flex 内的所有顶层 <p>/.table-row 重新按 measureFragmentHeight 测量，
  //    并强制拆碎超高单块，直到每页都不溢出
  const finalPages = []
  for (const pageHtml of pages) {
    const tempDiv = document.createElement('div')
    tempDiv.style.cssText = 'position:absolute;visibility:hidden;left:0;top:0;width:160mm;font-family:"SimSun","宋体",sans-serif;font-size:14px;line-height:1.5;'
    tempDiv.innerHTML = pageHtml
    document.body.appendChild(tempDiv)
    const overflow = (tempDiv.offsetHeight || 0) > A4_CONTENT_HEIGHT_PX
    if (!overflow) {
      document.body.removeChild(tempDiv)
      finalPages.push(pageHtml)
      continue
    }
    // 溢出：把 page-flex 内容拆出来重新分页
    const innerBlocks = [...tempDiv.querySelectorAll(':scope > div > *')]  // page-flex 内的子元素
    document.body.removeChild(tempDiv)
    console.log(`[DocxPreview] post-check: 页面内 overflow=${overflow}px, 重新拆 ${innerBlocks.length} 个块`)
    let newPage = []
    let newHeight = 0
    for (const blk of innerBlocks) {
      let blockHtml = blk.outerHTML
      let blockH = measureFragmentHeight(blockHtml)
      // 单块超高：拆碎
      if (blockH > A4_CONTENT_HEIGHT_PX) {
        const frags = splitBlockByBr(blockHtml, A4_CONTENT_HEIGHT_PX)
        for (const f of frags) {
          const fH = measureFragmentHeight(f)
          if (newHeight + fH > A4_CONTENT_HEIGHT_PX && newPage.length > 0) {
            // 换页
            finalPages.push(`<div class="page-flex">${newPage.join('')}</div>`)
            newPage = [f]
            newHeight = fH
          } else {
            newPage.push(f)
            newHeight += fH
          }
        }
      } else {
        if (newHeight + blockH > A4_CONTENT_HEIGHT_PX && newPage.length > 0) {
          finalPages.push(`<div class="page-flex">${newPage.join('')}</div>`)
          newPage = [blockHtml]
          newHeight = blockH
        } else {
          newPage.push(blockHtml)
          newHeight += blockH
        }
      }
    }
    if (newPage.length > 0) {
      // 关键修复：跳过纯 page-filler 的空白页（用户反馈"43页是空白页要删除"）
      const realContent = newPage.filter(html => !html.includes('class="page-filler"')).length
      if (realContent > 0) {
        finalPages.push(`<div class="page-flex">${newPage.join('')}</div>`)
      } else {
        console.log(`[DocxPreview] 跳过空白页 (仅含 page-filler)`)
      }
    }
  }

  paginatedPages.value = finalPages
  console.log(`[DocxPreview] 分页完成：${blocks.length} 个块 -> ${pages.length} 页 -> post-check -> ${finalPages.length} 页`)
}

// mammoth 自定义样式映射
const styleMap = [
  "p[style-name='Title'] => h1.docx-title:fresh",
  "p[style-name='Heading 1'] => h2.docx-h1:fresh",
  "p[style-name='Heading 2'] => h3.docx-h2:fresh",
  "p[style-name='Heading 3'] => h4.docx-h3:fresh",
  "p[style-name='heading 1'] => h2.docx-h1:fresh",
  "p[style-name='heading 2'] => h3.docx-h2:fresh",
  "p[style-name='heading 3'] => h4.docx-h3:fresh",
]

const printCss = `
  body { font-family: "SimSun", "宋体", "Microsoft YaHei", "微软雅黑", sans-serif; padding: 40px; line-height: 1.5; color: #000; }
  p { margin: 4px 0; text-indent: 2em; text-align: justify; }
  p.docx-heading { text-align: left !important; font-weight: bold; font-size: 15px; text-indent: 0; margin: 12px 0 8px; }
  h1, h2, h3, h4 { text-align: left; font-weight: bold; margin: 12px 0 8px; }
  table { border-collapse: collapse; width: 100%; margin: 10px 0; page-break-inside: avoid; }
  td, th { border: 1px solid #000; padding: 6px 10px; vertical-align: top; }
  table p { text-indent: 0; }
  @page { size: A4; margin: 25mm 20mm; }
`

// 解析 DOCX Blob -> HTML
const parseDocx = async (blob) => {
  loading.value = true
  errorMsg.value = ''
  fullHtml.value = ''
  paginatedPages.value = []
  // 释放之前的 blob URL
  imageBlobUrls.forEach(u => URL.revokeObjectURL(u))
  imageBlobUrls = []
  try {
    const arrayBuffer = await blob.arrayBuffer()
    // 1. 提取 docx tblGrid 列宽（在 mammoth 转换前，从原始 docx XML 读取）
    const grids = await extractTableGrids(arrayBuffer)
    console.log(`[DocxPreview] 提取到 ${grids.length} 张表的列宽:`, grids.map(g => `${g.length}cols[${g.join(',')}]`))
    const result = await mammoth.convertToHtml(
      { arrayBuffer },
      {
        styleMap,
        // 关键优化：把图片转成 blob URL（避免 base64 内联导致 HTML 体积爆炸）
        convertImage: (image) => {
          return image.read('arrayBuffer').then((imageBuffer) => {
            const blob = new Blob([imageBuffer], { type: image.contentType })
            const url = URL.createObjectURL(blob)
            imageBlobUrls.push(url)
            return { src: url, alt: image.altText || '图片' }
          })
        },
      }
    )
    if (result.messages && result.messages.length > 0) {
      console.warn('mammoth 警告:', result.messages)
    }
    let html = result.value || '<p style="text-align:center;color:#999;">文档为空</p>'
    // 2. 注入列宽到 mammoth HTML 的 <td>/<th>
    html = injectTableWidths(html, grids, A4_CONTENT_WIDTH_PX)
    // 后处理：把"整段只有 <strong> 文本"的段落标记为标题
    // 左对齐 + 加粗 + 较大字号，不缩进
    html = html.replace(
      /<p>\s*<strong>([\s\S]*?)<\/strong>\s*<\/p>/g,
      '<p class="docx-heading">$1</p>'
    )
    // 模板分页点：mammoth 压缩了空段（Word 模板用空段撑出独立页面），
    // 在以下标题段落加 force-new-page 标记，splitIntoPages 遇到就强制分页
    // 默认使用施工组织设计的 3 个分页点；其它模块可通过 pageBreaks prop 覆盖
    const PAGE_BREAK_TEXTS = props.pageBreaks && props.pageBreaks.length > 0
      ? props.pageBreaks
      : [
          '组织技术安全措施审批表',  // 封面结束 → 审批表独立一页
          '一、工程内容',            // 审批表结束 → 正文从新页开始
          '五、技术措施',            // 一~四 节结束 → 技术措施独立一页
        ]
    for (const txt of PAGE_BREAK_TEXTS) {
      const esc = txt.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      // 匹配 <p class="docx-heading">X</p> 或 <p>X</p> 形式
      const re1 = new RegExp(`<p\\s+class="docx-heading">(\\s*${esc}\\s*)</p>`, 'g')
      html = html.replace(re1, `<p class="docx-heading force-new-page">$1</p>`)
      const re2 = new RegExp(`<p>(\\s*${esc}\\s*)</p>`, 'g')
      html = html.replace(re2, `<p class="force-new-page">$1</p>`)
    }
    fullHtml.value = html
    // 等 mammoth 输出渲染到测量容器
    await nextTick()
    // 执行分页
    await splitIntoPages()
  } catch (e) {
    console.error('解析 DOCX 失败:', e)
    errorMsg.value = '文档解析失败：' + (e.message || e)
  } finally {
    loading.value = false
  }
}

watch(() => props.docxBlob, (blob) => {
  if (blob) parseDocx(blob)
}, { immediate: true })

// 缩放变化时不需要重新分页（CSS scale 不影响真实高度）
</script>

<style scoped>
.docx-preview-dialog :deep(.el-dialog__body) {
  padding: 0;
  background: #f0f2f5;
  height: 80vh;
  display: flex;
  flex-direction: column;
}

.preview-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 8px;
  width: 100%;
}
.preview-header .title {
  font-size: 16px;
  font-weight: 600;
}
.preview-header .actions {
  display: flex;
  gap: 12px;
  align-items: center;
}

.preview-container {
  flex: 1;
  overflow: auto;
  padding: 20px;
  background: #e8eaed;
  position: relative;
}

.pages-wrapper {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 16px;
  width: 100%;
  transition: transform 0.2s ease;
}

/* A4 纸张样式：页边距 2.5cm 统一（上下左右都是 25mm） */
.a4-page {
  width: 210mm;
  height: 297mm;  /* 用 height 而非 min-height，避免 flex 把高度算成 295mm */
  padding: 25mm;  /* 统一 2.5cm：与 docx 模板 <w:pgMar> 1418/1418/1418/1418 dxa 一致 */
  background: #fff;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
  font-family: "SimSun", "宋体", "Microsoft YaHei", "微软雅黑", sans-serif;
  font-size: 14px;
  line-height: 1.5;  /* 固定 1.5 倍行距（与 docx 模板 line_spacing=1.5 一致） */
  color: #000;
  position: relative;
  page-break-after: always;
  display: flex;
  flex-direction: column;
  box-sizing: border-box;
  overflow: hidden;
}
/* 封面/审批表：docx 模板审批表 line=223/240≈0.93 倍（紧密排版），封面无显式行距用 Normal 默认 */
.a4-page.is-cover {
  line-height: 1.5;
}
.a4-page.is-approval {
  line-height: 1.0;  /* docx 审批表 line=223 rule=auto ≈ 0.93 倍，CSS 最小取 1.0 */
}
.page-content {
  flex: 1;
  display: flex;
  flex-direction: column;
  min-height: 0;  /* 允许 flex 子元素收缩到内容尺寸以下 */
}
/* 上下分布：把"标题"推到顶部，把"末尾段"推到底部（封面/审批表）
   用 height 而非 min-height，避免内容溢出时被裁 */
.page-content :deep(.page-flex) {
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  flex: 1;
  min-height: 0;
}

/* 测量容器：不可见但保留布局 */
.measure-container {
    position: absolute;
    top: 0;
    left: 0;
    visibility: hidden;
    width: 160mm;  /* 等于 A4 内容区宽度 210mm - 2*25mm padding（统一 2.5cm 后） */
    font-family: "SimSun", "宋体", "Microsoft YaHei", "微软雅黑", sans-serif;
    font-size: 14px;
    line-height: 1.5;  /* 关键：与 .a4-page 一致（除封面/审批表外都是 1.5） */
    pointer-events: none;
    box-sizing: border-box;
  }
  .measure-container p {
    margin: 4px 0;
    line-height: 1.5;
    font-size: 14px;
    /* 关键：与 .a4-page 内的 :deep(p) 完全一致，否则测量高度不准（measure 时是 517px，渲染时是 1072px） */
    text-indent: 2em;
    text-align: justify;
    word-break: break-all;
  }
  .measure-container :deep(p) {
    /* scoped CSS 默认不生效于动态 HTML，必须用 :deep() 穿透 */
    margin: 4px 0;
    line-height: 1.5;
    font-size: 14px;
    text-indent: 2em;
    text-align: justify;
    word-break: break-all;
  }
  .measure-container :deep(p.docx-heading) {
    font-size: 15px;
    line-height: 1.6;
    margin: 12px 0 8px;
    font-weight: bold;
    text-indent: 0;
  }
  .measure-container :deep(table) {
    border-collapse: collapse;
    width: 100%;
    margin: 10px 0;
  }
  /* 关键：scoped CSS 必须用 :deep() 才能影响 mammoth 动态输出的 td/th
     （不穿透的话 measure 容器 td padding = 1px，渲染时 = 1.5cm，导致高度差 555px） */
  .measure-container :deep(table td),
  .measure-container :deep(table th) {
    border: 1px solid #000;
    padding: 1.5cm 6px;
    vertical-align: top;
  }
  .measure-container :deep(table p) {
    text-indent: 0;
    margin: 2px 0;
  }
  .measure-container :deep(.page-flex) {
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    min-height: 247mm;
  }
  .measure-container :deep(.page-filler) {
    flex: 0 0 auto;
  }

/* 标题样式：左对齐 + 加粗 + 不缩进 + 较大字号 */
.page-content :deep(h1),
.page-content :deep(h2),
.page-content :deep(h3),
.page-content :deep(h4) {
  text-align: left;
  font-weight: bold;
  margin: 12px 0 8px;
  text-indent: 0;
}
.page-content :deep(h1) { font-size: 20px; }
.page-content :deep(h2) { font-size: 18px; }
.page-content :deep(h3) { font-size: 16px; }
.page-content :deep(h4) { font-size: 15px; }

/* 段落：左对齐 + 中文首行缩进 2em + 两端对齐
   用 !important 覆盖 mammoth 输出可能带的 inline style */
.page-content :deep(p) {
  margin: 2px 0 !important;  /* 用户反馈"行间距太大"：4px→2px */
  text-indent: 2em !important;
  text-align: justify !important;
  word-break: break-all;
}

/* 整段加粗的"标题"段落：左对齐 + 加粗 + 不缩进 + 略大字号 */
.page-content :deep(p.docx-heading) {
  text-align: left !important;
  font-weight: bold !important;
  font-size: 15px !important;
  margin: 12px 0 8px !important;
  text-indent: 0 !important;
  line-height: 1.6;
}

/* 表格 */
.page-content :deep(table) {
  border-collapse: collapse;
  width: 100%;
  margin: 10px 0;
  font-size: 13px;
  /* 关键：列宽由 inline style 控制，必须用 fixed 才能真正生效 */
  table-layout: fixed;
}
/* 表格 cell 内边距：上下 1.5cm（与 docx 模板 tblCellMar top/bottom=850 dxa 一致，
   实现"13 页表格上下页边距固定 1.5cm"），左右 6px */
.page-content :deep(table td),
.page-content :deep(table th) {
  border: 1px solid #000;
  padding: 1.5cm 6px;
  vertical-align: top;
  /* 关键：width 由 injectTableWidths 注入的 inline style 控制，
     不要用 flex/grid 等会改变 width 解释的属性 */
  box-sizing: border-box;
  word-break: break-all;
  overflow-wrap: break-word;
  line-height: 1.5;
}
/* 表格内段落不缩进 */
.page-content :deep(table p) {
  text-indent: 0 !important;
  margin: 2px 0 !important;
  line-height: 1.5;
  font-size: 13px;
}

/* 拆开后的"伪表格行"——把 <tr> 内容序列化为普通块，
   用虚线框模拟表格行，便于跨页切分；table-cell 间用竖线分隔，宽度由 inline style 控制 */
.page-content :deep(.table-row) {
  display: flex;
  flex-wrap: nowrap;
  align-items: stretch;
  border: 1px solid #000;
  border-bottom: none;
  padding: 0;
  margin: 0;
  page-break-inside: avoid;
  /* 关键：让整行占满 content 区，避免出现"短表未铺满宽度"的右空白 */
  width: 100%;
  box-sizing: border-box;
}
.page-content :deep(.table-row:last-child) {
  border-bottom: 1px solid #000;
}
.page-content :deep(.table-cell) {
  /* width 由 inline style 从 docx tblGrid 注入；不要用 flex-grow，否则等宽分布 */
  flex: 0 0 auto;
  border-right: 1px solid #000;
  border-bottom: 1px solid #000;
  padding: 4px 6px;  /* 用户反馈"表格行距太大"：6px 8px→4px 6px */
  vertical-align: top;
  box-sizing: border-box;
  word-break: break-all;
  overflow-wrap: break-word;
  line-height: 1.25;  /* 用户反馈"表格行距太大"：默认1.5→1.25 */
  font-size: 13px;
  text-align: left;
}
/* 末行末格不要底部边框（外框已有） */
.page-content :deep(.table-row:last-child .table-cell) {
  border-bottom: none;
}
/* 末列不要右边框（外框已有） */
.page-content :deep(.table-cell:last-child) {
  border-right: none;
}
.page-content :deep(.table-cell p) {
  text-indent: 0 !important;
  margin: 0 !important;  /* 表格 cell 内段落无 margin，避免行间距过大 */
  /* 与表格 cell 一致：不缩进、不加段间距 */
  line-height: 1.25;  /* 用户反馈"表格行距太大"：1.5→1.25 */
  font-size: 13px;
  text-align: left;
}
/* 单元格内的章节标题（被识别为 docx-heading 的 <p>）也按正文处理 */
.page-content :deep(.table-cell p.docx-heading) {
  text-align: left !important;
  font-weight: bold !important;
  font-size: 14px !important;
  margin: 0 !important;
  text-indent: 0 !important;
  line-height: 1.25;
}

/* 封面页：让"施工方案"4 行大标题占据 A4 中部，项目名顶部、落款底部，
   整页视觉填满不留空白。8 段结构：1=项目名 2=组织技术安全措施 3-6=施/工/方/案 7=公司 8=日期
   使用 justify-content: center 让 8 段自然居中分布：上半页项目名+施工方案，下半页落款 */
.a4-page.is-cover :deep(.page-flex) {
  justify-content: center;
}
.a4-page.is-cover :deep(.page-flex p) {
  text-align: center !important;
  text-indent: 0 !important;
}
/* 项目名 + "组织技术安全措施" 副标题：docx 模板用 宋体 22pt bold
   注意：必须加 !important，因为 .page-content p.docx-heading { margin: 12px 0 8px !important } 会覆盖 */
.a4-page.is-cover :deep(.page-flex p:nth-child(1)),
.a4-page.is-cover :deep(.page-flex p:nth-child(2)) {
  font-size: 22px !important;
  line-height: 1.3;
  margin: 4px 0 !important;
}
/* "施工方案" 4 行大标题：docx 模板用 宋体 26pt bold，4 字使用 24px margin 让 4 字之间有视觉间隔 */
.a4-page.is-cover :deep(.page-flex p:nth-child(3)),
.a4-page.is-cover :deep(.page-flex p:nth-child(4)),
.a4-page.is-cover :deep(.page-flex p:nth-child(5)),
.a4-page.is-cover :deep(.page-flex p:nth-child(6)) {
  font-size: 35px !important;
  font-weight: bold !important;
  line-height: 1.3;
  letter-spacing: 8px;
  margin: 24px 0 !important;
}
/* 公司名：docx 模板用 仿宋 18pt bold */
.a4-page.is-cover :deep(.page-flex p:nth-child(7)) {
  font-size: 18px !important;
  line-height: 1.3;
  margin-top: 60px !important;
  margin-bottom: 12px !important;
}
/* 日期：docx 模板用 仿宋 18pt bold */
.a4-page.is-cover :deep(.page-flex p:nth-child(8)) {
  font-size: 18px !important;
  line-height: 1.3;
  margin-bottom: 20px !important;
}

/* 审批表页：把表格撑高，让单元格均匀铺满 A4 内容区
   8 行 × 28mm = 224mm ≈ 撑满 247mm 内容区 */
.a4-page.is-approval :deep(.page-flex) {
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  flex: 1;
}
.a4-page.is-approval :deep(.page-flex > p:first-child) {
  margin: 0 0 12px 0 !important;
  font-weight: bold !important;
  font-size: 22px !important;  /* docx 模板：宋体 21.5pt bold */
  line-height: 1.0 !important;  /* docx 模板：line=223 ≈ 0.93 倍 */
}
.a4-page.is-approval :deep(.page-flex table) {
  font-size: 14px;
  width: 100%;
  flex: 1;
  table-layout: fixed;
}
.a4-page.is-approval :deep(.page-flex table td),
.a4-page.is-approval :deep(.page-flex table th) {
  /* 审批表 cell 内边距：上下 1.2cm，左右 8px */
  padding: 1.2cm 8px;
  font-size: 14px;
  line-height: 1.0;  /* docx 模板审批表 line=223 ≈ 0.93 倍 */
  vertical-align: middle;
  text-align: left !important;  /* 关键：表格内文字左对齐（覆盖 .page-content p 可能的居中） */
}
/* 关键：审批表被拆解为 div.table-row 后也要保持 1.0 倍行距（docx line=223）
   + 完整四边框 + 文字左对齐 + 表格内单倍行距（用户反馈"表格行距太大"） */
.a4-page.is-approval :deep(.page-flex .table-cell) {
  line-height: 1.0;
  font-size: 14px;
  padding: 1.2cm 8px;
  vertical-align: middle;
  /* 关键：明确指定四边框，避免被通用 .table-cell 的右侧 border:0 覆盖 */
  border-top: 1px solid #000 !important;
  border-right: 1px solid #000 !important;
  border-bottom: 1px solid #000 !important;
  border-left: 1px solid #000 !important;
  text-align: left !important;
}
.a4-page.is-approval :deep(.page-flex .table-cell p) {
  line-height: 1.0;
  font-size: 14px;
  margin: 0 !important;
  text-indent: 0 !important;
  text-align: left !important;
}

/* 只有章节标题/大段空白页：把内容顶置 */
.a4-page.is-section-header :deep(.page-flex) {
  justify-content: flex-start;
}

.page-content :deep(ul),
.page-content :deep(ol) {
  padding-left: 30px;
  margin: 6px 0;
}
.page-content :deep(li) {
  margin: 4px 0;
}
.page-content :deep(strong),
.page-content :deep(b) {
  font-weight: bold;
}
.page-content :deep(img) {
  max-width: 100%;
  height: auto;
}

.page-number {
  position: absolute;
  bottom: 10mm;
  left: 0;
  right: 0;
  text-align: center;
  font-size: 12px;
  color: #666;
}

.preview-error {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 100%;
}

/* 打印样式 */
@media print {
  .preview-header { display: none !important; }
  .preview-container { padding: 0; background: #fff; }
  .pages-wrapper { transform: none !important; gap: 0; }
  .a4-page {
    box-shadow: none;
    margin: 0;
    page-break-after: always;
  }
}
</style>
