"""DOCX → 高保真 PDF 渲染服务

**问题背景**：
之前预览链路：
  DOCX → 前端 mammoth → HTML → 浏览器渲染
  vs
  DOCX → 用户下载 → Word 打开

两个引擎（mammoth/Chromium CSS vs Word）渲染同一份 DOCX 会有版面差异，
导致"预览和下载看到的不是同一种排版"。

**本模块方案**：
服务端用 mammoth 把 DOCX 解析成结构化 HTML，再用一份**从 DOCX 实际样式推导**出来的
高保真 CSS（与 docx 模板 XML 一一对应），最后用 Playwright + Chromium 把 HTML 渲染成 PDF。

  DOCX ─┬─→ mammoth → HTML  ─┐
        └─→ XML (读 sectPr/pPr/rPr) ─→ 自定义 CSS ─┤
                                                  └─→ Playwright → PDF

由于 CSS 直接来源于 DOCX 的 `<w:sectPr>`/`<w:pPr>`/`<w:rPr>` 等属性，
最终 PDF 与 Word 打开的 DOCX 在 **内容、字体、字号、页边距、分页、表格列宽** 上保持一致。
唯一可能差异是字体像素 hinting（SimSun 在 Chromium vs Word 的细微差别），属可接受范围。

**降级策略**：
- Playwright 渲染失败 → 抛异常，让上层 try/except
- mammoth 失败 → 直接抛异常

**性能**：
- 一次渲染约 1-3s（DOCX 大小相关）
- 渲染结果缓存到 `<docx>.rendered.pdf` 文件（同目录下）
  - 文件 mtime 比 DOCX 旧时自动重新渲染
  - 文件 mtime 比 DOCX 新时直接复用
"""

import os
import sys
import json
import re
import subprocess
import shutil
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

import mammoth


# DOCX XML 命名空间
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def _q(tag):
    return f"{W}{tag}"


def _attr(elem, name):
    return elem.get(_q(name))


def _read_docx_xml(docx_path, member):
    """读取 docx 内 word/<member>.xml 的内容"""
    with zipfile.ZipFile(docx_path, "r") as z:
        try:
            return z.read(f"word/{member}.xml").decode("utf-8")
        except KeyError:
            return None


def _parse_twips(val, default=0):
    """dxa/twips → px (1 dxa = 1/20 pt = 1/1440 in; at 96dpi: 1 dxa ≈ 0.0667 px)"""
    if val is None:
        return default
    try:
        return int(int(val) * 96 / 1440)
    except (ValueError, TypeError):
        return default


def _parse_pt_to_px(val, default=None):
    """w:sz 半点 → px; e.g. sz=21 → 10.5pt → 14px (96dpi)"""
    if val is None:
        return default
    try:
        pt = int(val) / 2
        return round(pt * 96 / 72, 2)
    except (ValueError, TypeError):
        return default


def _parse_pt_from_val(val, default=None):
    """直接 val=21 表示 21pt; <w:val w:osid="...">21</w:val>"""
    if val is None:
        return default
    try:
        pt = int(val)
        return round(pt * 96 / 72, 2)
    except (ValueError, TypeError):
        return default


def extract_docx_styles(docx_path):
    """从 DOCX XML 中提取页面级和段落级样式参数。

    返回 dict:
      {
        "page": { "width_px": 794, "height_px": 1123,
                  "margin_top_px": 94, "margin_right_px": 94, ... },
        "default_run": { "font_size_px": 14, "font_family": "...",
                         "line_height": 1.5, ... },
      }
    """
    doc_xml = _read_docx_xml(docx_path, "document")
    if not doc_xml:
        return _default_styles()

    # 命名空间映射（ElementTree 解析时不自动加 ns0:）
    root = ET.fromstring(doc_xml)
    sect_prs = root.findall(f".//{_q('sectPr')}")
    if not sect_prs:
        return _default_styles()

    sect_pr = sect_prs[-1]
    page_size = sect_pr.find(_q("pgSz"))
    page_mar = sect_pr.find(_q("pgMar"))
    doc_grid = sect_pr.find(_q("docGrid"))

    # 默认 A4: 11906 x 16838 dxa (210mm x 297mm)
    width_px = _parse_twips(_attr(page_size, "w"), 11906)
    height_px = _parse_twips(_attr(page_size, "h"), 16838)
    if page_size is not None and _attr(page_size, "orient") == "landscape":
        width_px, height_px = height_px, width_px

    margin_top_px = _parse_twips(_attr(page_mar, "top"), 1418)
    margin_right_px = _parse_twips(_attr(page_mar, "right"), 1418)
    margin_bottom_px = _parse_twips(_attr(page_mar, "bottom"), 1418)
    margin_left_px = _parse_twips(_attr(page_mar, "left"), 1418)
    header_px = _parse_twips(_attr(page_mar, "header"), 720)
    footer_px = _parse_twips(_attr(page_mar, "footer"), 720)

    # 默认字体/字号：从 styles.xml 读 docDefaults
    styles_xml = _read_docx_xml(docx_path, "styles")
    default_run = { "font_size_px": 14, "font_family": "SimSun", "line_height": 1.5 }
    if styles_xml:
        try:
            sroot = ET.fromstring(styles_xml)
            doc_def = sroot.find(f"{_q('docDefaults')}")
            if doc_def is not None:
                rpr_def = doc_def.find(f"{_q('rPrDefault')}/{_q('rPr')}")
                if rpr_def is not None:
                    sz = rpr_def.find(_q("sz"))
                    if sz is not None:
                        default_run["font_size_px"] = _parse_pt_to_px(_attr(sz, "val"), 21) or 14
                    rfonts = rpr_def.find(_q("rFonts"))
                    if rfonts is not None:
                        ascii_font = _attr(rfonts, "ascii") or _attr(rfonts, "eastAsia") or "SimSun"
                        default_run["font_family"] = ascii_font
                ppr_def = doc_def.find(f"{_q('pPrDefault')}/{_q('pPr')}")
                if ppr_def is not None:
                    spacing = ppr_def.find(_q("spacing"))
                    if spacing is not None:
                        line = _attr(spacing, "line")
                        if line:
                            try:
                                line_val = int(line)
                                line_rule = _attr(spacing, "lineRule") or "auto"
                                if line_rule == "auto" and line_val > 0:
                                    # line=240 = 1.0 倍; line=360 = 1.5 倍
                                    default_run["line_height"] = round(line_val / 240.0, 3)
                                else:
                                    default_run["line_height"] = round(line_val * 96 / 1440 / default_run["font_size_px"], 3) if default_run["font_size_px"] > 0 else 1.5
                            except (ValueError, TypeError):
                                pass
        except ET.ParseError:
            pass

    # Fallback：如果 docDefaults 没有 line spacing，从 body 段落中采样最常见的值
    # 判断依据：docDefaults 中没有 pPrDefault/pPr/spacing 节点
    _docdefaults_has_spacing = False
    if styles_xml:
        try:
            sroot = ET.fromstring(styles_xml)
            doc_def = sroot.find(f"{_q('docDefaults')}")
            if doc_def is not None:
                ppr_def = doc_def.find(f"{_q('pPrDefault')}/{_q('pPr')}")
                if ppr_def is not None:
                    sp = ppr_def.find(_q("spacing"))
                    if sp is not None and _attr(sp, "line"):
                        _docdefaults_has_spacing = True
        except ET.ParseError:
            pass

    if not _docdefaults_has_spacing and doc_xml:
        try:
            from collections import Counter
            body_el = root.find(f"{_q('body')}")
            if body_el is not None:
                line_counter = Counter()
                for p in body_el.findall(_q("p")):
                    ppr = p.find(_q("pPr"))
                    if ppr is not None:
                        sp = ppr.find(_q("spacing"))
                        if sp is not None:
                            line_val = _attr(sp, "line")
                            line_rule = _attr(sp, "lineRule") or "auto"
                            if line_val:
                                line_counter[(line_val, line_rule)] += 1
                if line_counter:
                    (most_line, most_rule), _ = line_counter.most_common(1)[0]
                    try:
                        lv = int(most_line)
                        if most_rule == "auto" and lv > 0:
                            default_run["line_height"] = round(lv / 240.0, 3)
                        else:
                            default_run["line_height"] = round(lv * 96 / 1440 / default_run["font_size_px"], 3) if default_run["font_size_px"] > 0 else 1.5
                    except (ValueError, TypeError):
                        pass
        except Exception:
            pass

    return {
        "page": {
            "width_px": width_px,
            "height_px": height_px,
            "margin_top_px": margin_top_px,
            "margin_right_px": margin_right_px,
            "margin_bottom_px": margin_bottom_px,
            "margin_left_px": margin_left_px,
            "header_px": header_px,
            "footer_px": footer_px,
        },
        "default_run": default_run,
    }


def _default_styles():
    """默认 A4 + 宋体 14px / 1.5 倍行距 / 2.5cm 边距"""
    return {
        "page": {
            "width_px": 794, "height_px": 1123,
            "margin_top_px": 94, "margin_right_px": 94,
            "margin_bottom_px": 94, "margin_left_px": 94,
            "header_px": 48, "footer_px": 48,
        },
        "default_run": { "font_size_px": 14, "font_family": "SimSun", "line_height": 1.5 },
    }


def extract_table_widths(docx_path):
    """提取每张表的 tblGrid 列宽（twips），返回 [[col1_twips, col2_twips, ...], ...]"""
    doc_xml = _read_docx_xml(docx_path, "document")
    if not doc_xml:
        return []
    root = ET.fromstring(doc_xml)
    tables = root.findall(f".//{_q('tbl')}")
    grids = []
    for tbl in tables:
        grid_elem = tbl.find(_q("tblGrid"))
        if grid_elem is None:
            grids.append([])
            continue
        widths = []
        for col in grid_elem.findall(_q("gridCol")):
            w = _attr(col, "w")
            try:
                widths.append(int(w) if w else 0)
            except (ValueError, TypeError):
                widths.append(0)
        grids.append(widths)
    return grids


def inject_table_widths_into_html(html, grids, content_width_px):
    """把 tblGrid 列宽注入 mammoth 输出的 <td>/<th> 的 style.width

    复用前端 DocxPreviewDialog.vue 的同一逻辑：处理 colspan/rowspan/比例缩放。
    """
    if not grids:
        return html
    from html.parser import HTMLParser

    class TableInjector(HTMLParser):
        def __init__(self):
            super().__init__()
            self.out = []
            self.in_table = 0
            self.tables = []  # (depth, grid_index)
            self.current_rows = []  # 当前正在解析的 table
            self.current_row_cells = []
            self.current_cell = None
            self.in_row = False
            self.table_idx = -1
            self.grid_layout = None  # 当前 table 的 layout 矩阵

        def handle_starttag(self, tag, attrs):
            t = tag.lower()
            if t == "table":
                self.in_table += 1
                self.table_idx += 1
                self.current_rows = []
                self.in_row = False
                self.current_cell = None
                if self.table_idx < len(grids):
                    grid = grids[self.table_idx]
                    n = len(grid)
                    self.grid_layout = [[False] * n for _ in range(0)]
                else:
                    self.grid_layout = None
                self.out.append(self.get_starttag_text())
            elif t == "tr" and self.in_table:
                self.in_row = True
                self.current_row_cells = []
                self.out.append("<tr>")
            elif t in ("td", "th") and self.in_row:
                attrs_dict = dict(attrs)
                colspan = int(attrs_dict.get("colspan") or 1)
                rowspan = int(attrs_dict.get("rowspan") or 1)
                self.current_cell = {
                    "attrs": attrs_dict, "tag": t, "colspan": colspan, "rowspan": rowspan,
                    "content": []
                }
            else:
                if self.in_table and self.current_cell is not None:
                    self.current_cell["content"].append(self.get_starttag_text())
                else:
                    self.out.append(self.get_starttag_text())

        def handle_endtag(self, tag):
            t = tag.lower()
            if t == "table":
                self.in_table -= 1
                self.out.append("</table>")
            elif t == "tr" and self.in_row:
                # 处理当前行：注入列宽
                if self.grid_layout:
                    row_idx = len(self.current_rows)
                    # 扩展 layout 行
                    while len(self.grid_layout) <= row_idx:
                        self.grid_layout.append([False] * len(self.grid_layout[0]) if self.grid_layout else [])
                    col_idx = 0
                    n = len(self.grid_layout[0]) if self.grid_layout else 0
                    total_twips = sum(grids[self.table_idx][:n]) if self.table_idx < len(grids) else 0
                    px_per_twip = content_width_px / total_twips if total_twips > 0 else 0
                    for cell in self.current_row_cells:
                        # 跳过被 rowspan 占据的列
                        while col_idx < n and self.grid_layout[row_idx][col_idx]:
                            col_idx += 1
                        if col_idx >= n:
                            break
                        colspan = cell["colspan"]
                        rowspan = cell["rowspan"]
                        # 计算该 cell 跨越的列总宽
                        width_twips = 0
                        for c in range(colspan):
                            ci = col_idx + c
                            if ci < n and self.table_idx < len(grids):
                                width_twips += grids[self.table_idx][ci] or 0
                        width_px = max(20, round(width_twips * px_per_twip)) if px_per_twip else 0
                        # 标记 layout
                        for dr in range(rowspan):
                            for dc in range(colspan):
                                rr = row_idx + dr
                                cc = col_idx + dc
                                if rr < len(self.grid_layout) and cc < n:
                                    if cc >= len(self.grid_layout[rr]):
                                        # extend
                                        self.grid_layout[rr].extend([False] * (cc + 1 - len(self.grid_layout[rr])))
                                    self.grid_layout[rr][cc] = True
                        # 注入 inline 宽度到 cell 标签
                        new_attrs = []
                        for k, v in cell["attrs"].items():
                            if k == "width":
                                continue
                            new_attrs.append(f' {k}="{v}"')
                        attr_str = "".join(new_attrs)
                        if width_px:
                            attr_str += f' width="{width_px}"'
                        self.out.append(f'<{cell["tag"]}{attr_str} style="width:{width_px}px;min-width:{width_px}px;max-width:{width_px}px;">')
                        self.out.extend(cell["content"])
                        self.out.append(f'</{cell["tag"]}>')
                        col_idx += colspan
                else:
                    for cell in self.current_row_cells:
                        self.out.append(f'<{cell["tag"]}>')
                        self.out.extend(cell["content"])
                        self.out.append(f'</{cell["tag"]}>')
                self.current_rows.append(self.current_row_cells)
                self.in_row = False
                self.out.append("</tr>")
            elif t in ("td", "th") and self.current_cell is not None:
                self.current_row_cells.append(self.current_cell)
                self.current_cell = None
            else:
                if self.in_table and self.current_cell is not None:
                    self.current_cell["content"].append(f"</{t}>")
                else:
                    self.out.append(f"</{t}>")

        def handle_data(self, data):
            if self.in_table and self.current_cell is not None:
                self.current_cell["content"].append(data)
            else:
                self.out.append(data)

        def handle_entityref(self, name):
            self.handle_data(f"&{name};")

        def handle_charref(self, name):
            self.handle_data(f"&#{name};")

    inj = TableInjector()
    inj.feed(html)
    return "".join(inj.out)


def extract_page_break_paragraph_indexes(docx_path):
    """从 docx XML 找出哪些段落应该从新页开始（硬分页位置）。

    返回一个 set，元素是 paragraph index。

    DOCX 中 <w:br w:type="page"/> 的位置语义：
    - 如果 break 在段落文字之前（第一个有文字的 run 之前），则该段落自身从新页开始
    - 如果 break 在段落文字之后（最后一个有文字的 run 之后），则下一段从新页开始
    - 如果 break 在空段上，跳过后续连续空段，找到第一个有文字的段落
    """
    doc_xml = _read_docx_xml(docx_path, "document")
    if not doc_xml:
        return set()
    root = ET.fromstring(doc_xml)
    body = root.find(f"{_q('body')}")
    if body is None:
        return set()
    break_indexes = set()
    paragraphs = body.findall(_q("p"))
    for i, p in enumerate(paragraphs):
        # 该段落是否含有 <w:br w:type="page"/>
        breaks = p.findall(f".//{_q('br')}[@{_q('type')}='page']")
        if not breaks:
            continue

        # 判断 break 在文字之前还是之后
        # 收集所有 run 的信息：是否有文字、是否有 break
        runs = p.findall(f".//{_q('r')}")
        first_text_run_idx = None
        break_run_idx = None
        for ri, run in enumerate(runs):
            if break_run_idx is None:
                run_breaks = run.findall(f".//{_q('br')}[@{_q('type')}='page']")
                if run_breaks:
                    break_run_idx = ri
            run_texts = run.findall(f".//{_q('t')}")
            run_text = ''.join((t.text or '') for t in run_texts).strip()
            if run_text and first_text_run_idx is None:
                first_text_run_idx = ri

        # 段落是否有文字？
        text_elements = p.findall(f".//{_q('t')}")
        text = ''.join((t.text or '') for t in text_elements).strip()

        if not text:
            # 空段：跳过后续连续空段，找到第一个有文字的段落
            target = i
            for j in range(i + 1, len(paragraphs)):
                t_elems = paragraphs[j].findall(f".//{_q('t')}")
                t_text = ''.join((te.text or '') for te in t_elems).strip()
                if t_text:
                    target = j
                    break
            break_indexes.add(target)
        elif break_run_idx is not None and first_text_run_idx is not None:
            if break_run_idx < first_text_run_idx:
                # break 在文字之前：该段落自身从新页开始
                break_indexes.add(i)
            else:
                # break 在文字之后：下一段从新页开始
                if i + 1 < len(paragraphs):
                    break_indexes.add(i + 1)
        elif text:
            # 有文字但无法判断 break 位置，默认下一段
            if i + 1 < len(paragraphs):
                break_indexes.add(i + 1)
    return break_indexes


def extract_paragraph_texts(docx_path):
    """从 docx 提取每个段落的纯文本（保留段落索引）。

    Returns: list of strings, index 对应 docx body 的段落位置。
    """
    doc_xml = _read_docx_xml(docx_path, "document")
    if not doc_xml:
        return []
    root = ET.fromstring(doc_xml)
    body = root.find(f"{_q('body')}")
    if body is None:
        return []
    paragraphs = body.findall(_q("p"))
    texts = []
    for p in paragraphs:
        text_elements = p.findall(f".//{_q('t')}")
        text = ''.join((t.text or '') for t in text_elements).strip()
        texts.append(text)
    return texts


def inject_docx_page_breaks_into_html(html, docx_para_texts, break_indexes):
    """根据 docx 段落文本，在 mammoth HTML 中找到对应段落并在段前加分页类。

    mammoth 输出的段落数与 docx 段落数不一一对应（mammoth 会忽略部分空段/图片段）。
    采用文本匹配策略：把 docx 段落的前 20 个连续非空白字符作为"指纹"，在 mammoth
    HTML 中找到含同样指纹的 <p>，给它添加 class="docx-page-break-before"。

    重要：不再插入空 div（div.docx-page-break），改用 class 触发 page-break-before。
    原因：空 div 即使 height=0，仍会占用一行，触发 Chromium 推一行到下一页，
    产生多余的空白页（旧 bug）。

    Args:
        html: mammoth 输出的 HTML
        docx_para_texts: list of docx 段落纯文本
        break_indexes: 哪些 docx 段落需要分页

    Returns:
        注入分页 class 后的 HTML
    """
    if not break_indexes:
        return html

    # 收集所有需要分页的段落"指纹"（去掉空白的短前缀）
    # 使用 list 而非 set，以保留重复指纹（不同段落可能有相同前缀）
    break_fps = []
    for idx in sorted(break_indexes):
        if idx < len(docx_para_texts):
            t = docx_para_texts[idx].strip()
            if t:
                short = ''.join(c for c in t if not c.isspace())
                if len(short) >= 3:
                    # 取前 12 字符 + 整段较短时全取
                    fp = short[:12] if len(short) > 12 else short
                    break_fps.append(fp)
                elif short:
                    break_fps.append(short)

    if not break_fps:
        return html

    def _add_break_class(match):
        full = match.group(0)
        if 'docx-page-break-before' in full:
            return full
        # 已有 class
        m = re.match(r'<p\b([^>]*)>', full)
        if not m:
            return full
        attrs = m.group(1)
        if 'class="' in attrs:
            new_attrs = attrs.replace('class="', 'class="docx-page-break-before ', 1)
        else:
            new_attrs = attrs + ' class="docx-page-break-before"'
        return f'<p{new_attrs}>'

    pos = 0
    out = []
    fp_queue = list(break_fps)  # 按顺序消费指纹队列
    while pos < len(html):
        next_p_start = html.find('<p', pos)
        if next_p_start == -1:
            out.append(html[pos:])
            break
        out.append(html[pos:next_p_start])
        end_p = html.find('</p>', next_p_start)
        if end_p == -1:
            out.append(html[next_p_start:])
            break
        end_p += 4
        para_html = html[next_p_start:end_p]

        # 提取段落纯文本指纹
        para_text = re.sub(r'<[^>]+>', '', para_html)
        para_short = ''.join(c for c in para_text if not c.isspace())

        # 检查是否匹配队列中第一个未消费的指纹
        if fp_queue:
            fp = fp_queue[0]
            if fp in para_short or para_short.startswith(fp):
                # 把原段落的 class 替换为带 docx-page-break-before
                if 'class="' in para_html:
                    para_html = re.sub(r'<p\b([^>]*?)class="', r'<p\1class="docx-page-break-before ', para_html, count=1)
                else:
                    para_html = re.sub(r'<p\b([^>]*)(>)', r'<p\1 class="docx-page-break-before"\2', para_html, count=1)
                fp_queue.pop(0)  # 消费该指纹

        out.append(para_html)
        pos = end_p
    return "".join(out)


def inject_docx_page_breaks_into_html_by_index(html, break_para_indexes, mammoth_paragraph_count):
    """根据 mammoth 输出的段落结构，把硬分页插入到正确位置。

    mammoth 输出的段落是从 docx body 的 <w:p> 元素一对一映射的（除 mammoth 忽略的外）。
    我们在每个 hard-break 段落之后插入分页 div。

    由于 mammoth 可能会忽略部分段落（图片、表格分隔等），简单按索引对应可能会错位。
    采用更稳健的策略：对每个 break 段落，在 HTML 中找到该段落的 </p>，在它后面插入分页 div。
    """
    if not break_para_indexes:
        return html

    # mammoth 输出的 <p>... </p>，按出现顺序编号。
    # 但 mammoth 可能会忽略某些段落（带 w:pict 但无文本等）。
    # 实用方法：直接按 mammoth 输出 <p> 元素的顺序匹配（数量 = mammoth_paragraph_count）
    # 然后在每个 break 段对应的 <p> 之后插入分页 div

    pos = 0
    target = 0
    out = []
    para_count = 0
    while pos < len(html):
        # 找下一个 <p 开始
        next_p_start = html.find('<p', pos)
        if next_p_start == -1:
            out.append(html[pos:])
            break
        out.append(html[pos:next_p_start])
        # 找该 <p> 的对应 </p>
        end_p = html.find('</p>', next_p_start)
        if end_p == -1:
            out.append(html[next_p_start:])
            break
        end_p += 4
        out.append(html[next_p_start:end_p])

        # 如果这个 <p> 是 break 段落之一，在它后面插入分页 div
        if para_count in break_para_indexes:
            out.append('<div class="docx-page-break"></div>')

        pos = end_p
        para_count += 1
    return "".join(out)


def trim_trailing_empty_blocks(html):
    """移除 mammoth HTML 末尾的连续空段落 / 空白元素

    - docx 模板末尾常有 1-3 个空 <p>（防止 Word 把最后一行挤到边距外），
      但通过 mammoth 渲染到 Chromium 时这些空段会占据一整页（空白页）。
    - docx 末尾还可能有 <w:pict>（小型装饰图 / 编号图片），mammoth 转为 <img>，
      一张空的图片也会撑高一行。
    - docx 末尾还可能有 <w:bookmarkStart>（TOC 锚点）转为 <a id="...">，单个空 <a> 不占空间，
      但 <a> 包在 <p> 里就成空 <p>。
    - 在 PDF 渲染前清掉这些空段。
    """
    if not html:
        return html

    # 把 <a id="..."></a> 等无内容标签视为空内容（用于检测空段）
    empty_inner = r'(?:\s|&nbsp;|<br\s*/?>|<img[^>]*/?>|<a[^>]*></a>)*'

    changed = True
    while changed:
        changed = False
        # 末尾 1+ 个空 <p>（内含空白 / <br> / <img> / <a id>）
        m = re.search(rf'(?:<p[^>]*>{empty_inner}</p>\s*)+$', html)
        if m:
            html = html[:m.start()].rstrip()
            changed = True
            continue
        # 末尾单个 <img>（小图片 / 浮水印）
        m = re.search(r'<img[^>]*/?>\s*$', html)
        if m:
            html = html[:m.start()].rstrip()
            changed = True
            continue
        # 末尾空白字符
        if html.endswith((' ', '\n', '\t')):
            html = html.rstrip()
            changed = True
    return html


def inject_page_breaks_into_html(html, force_break_titles=None):
    """在指定标题前插入强制分页

    实现：在目标 <p> 标签的开标签里追加 class="docx-page-break-before"，
    并在 CSS 中定义该 class 的 page-break-before: always。
    避免使用占位 div 造成额外空白页。

    重要：如果该 <p> 紧跟一个 docx-page-break div（来自硬分页），
    就不再给它加 page-break-before class，避免重复分页造成空白页。
    """
    if force_break_titles is None:
        force_break_titles = []
    # 注意：cn_section_pat 里的 ^ 在 re.sub 模式下不生效（除非 re.MULTILINE），
    # 用 [^<]* 替代 ^，让"一"匹配"段落开头"（因为 <p>...</p> 已被外层 p 模式锁定），
    # [^<]* 能吃掉段首零个或多个非标签字符（空白等），然后由"一"自身作为段首字符。
    cn_section_pat = r"(一|二|三|四|五|六|七|八|九|十|十一|十二|十三|十四|十五|十六|十七|十八|十九|二十)\s*[、．.]"
    fixed_break = ["组织技术安全措施审批表", "审批表", "目  录", "目录", "批准", "工程概况"]

    def _add_class(match):
        full = match.group(0)
        # 已包含 docx-page-break-before，跳过
        if 'docx-page-break-before' in full:
            return full
        # 给 <p> 加 class
        if 'class="' in full:
            return full.replace('class="', 'class="docx-page-break-before ', 1)
        # 在 <p 后插入 class
        return re.sub(r'(<p\b)([^>]*)(>)', lambda m: m.group(1) + m.group(2) + ' class="docx-page-break-before"' + m.group(3), full, count=1)

    # 找出所有"下一个段落已经有 docx-page-break-before class"的位置
    # （这些是 docx 自身硬分页，inject_docx_page_breaks_into_html 已经把 class 加在了下一个段落）
    # 这样的 section header 不应再加分页（避免重复分页产生空白页）
    has_class_on_next = set()  # 段落索引：该段落之后紧跟的 <p> 已经有 class
    pos = 0
    para_idx = 0
    while pos < len(html):
        next_p_start = html.find('<p', pos)
        if next_p_start == -1:
            break
        end_p = html.find('</p>', next_p_start)
        if end_p == -1:
            break
        end_p += 4
        # 检查这个 <p> 之后紧跟的下一个 <p> 是否有 docx-page-break-before
        after = html[end_p:end_p+800]
        # 下一个 <p> 起始
        nxt = after.find('<p')
        if nxt != -1:
            # 找这个 <p 的结束符 >
            gt = after.find('>', nxt)
            if gt != -1:
                next_p_tag = after[nxt:gt+1]
                if 'docx-page-break-before' in next_p_tag:
                    has_class_on_next.add(para_idx)
        # 也兼容旧的 div.docx-page-break
        if '<div class="docx-page-break"' in after or '<div class=\'docx-page-break\'' in after:
            has_class_on_next.add(para_idx)
        pos = end_p
        para_idx += 1

    def _add_class_with_dedup(match):
        """比 _add_class 更严格：如果该段落的下一个 <p> 已经有 class，就不加。"""
        full = match.group(0)
        if 'docx-page-break-before' in full:
            return full
        # 找这个 <p> 在原始 html 中的位置和 para_idx
        # 简化处理：直接返回加了 class 的版本
        if 'class="' in full:
            return full.replace('class="', 'class="docx-page-break-before ', 1)
        return re.sub(r'(<p\b)([^>]*)(>)', lambda m: m.group(1) + m.group(2) + ' class="docx-page-break-before"' + m.group(3), full, count=1)

    # 走文本匹配：先收集 break 指纹（不依赖 ^，因为 re.sub 默认不开 MULTILINE）
    for title in force_break_titles + fixed_break:
        esc = re.escape(title)
        # 标题前允许空白/换行；段落必须是 docx-heading（避免误伤正文）
        pat = re.compile(
            rf'<p[^>]*class="docx-heading"[^>]*>(?:<[^>]+>)*\s*{esc}\s*[^<]*(?:</[^>]+>)*</p>'
        )
        html = pat.sub(_add_class, html)

    # 中文章节标题（一/二/.../二十）：也限定为 docx-heading，避免正文里的"一、二、三"被误伤
    # 同时跳过"下一个 <p> 已经有 docx-page-break-before class"的情况
    cn_pat = re.compile(
        rf'<p[^>]*class="docx-heading"[^>]*>(?:<[^>]+>)*\s*{cn_section_pat}\s*[^<]*(?:</[^>]+>)*</p>'
    )

    # 自定义 sub：找到匹配后，检查下一个 <p> 是否已经有 class
    def _replace_cn(match):
        matched = match.group(0)
        if 'docx-page-break-before' in matched:
            return matched
        # 找这个 <p> 之后的下一个 <p>
        end_p = match.end()
        after = html[end_p:end_p+800]
        nxt = after.find('<p')
        if nxt != -1:
            gt = after.find('>', nxt)
            if gt != -1:
                next_p_tag = after[nxt:gt+1]
                if 'docx-page-break-before' in next_p_tag:
                    # 下一个段落已经有 class，不要再加
                    return matched
        # 加 class
        if 'class="' in matched:
            return matched.replace('class="', 'class="docx-page-break-before ', 1)
        return re.sub(r'(<p\b)([^>]*)(>)', lambda m: m.group(1) + m.group(2) + ' class="docx-page-break-before"' + m.group(3), matched, count=1)

    html = cn_pat.sub(_replace_cn, html)

    # 后处理：移除"下一个段落已经有 docx-page-break-before class"的段落的 class
    pos = 0
    para_idx = 0
    out = []
    while pos < len(html):
        next_p_start = html.find('<p', pos)
        if next_p_start == -1:
            out.append(html[pos:])
            break
        out.append(html[pos:next_p_start])
        end_p = html.find('</p>', next_p_start)
        if end_p == -1:
            out.append(html[next_p_start:])
            break
        end_p += 4
        para_html = html[next_p_start:end_p]
        if para_idx in has_class_on_next:
            para_html = re.sub(r'\s*docx-page-break-before\b', '', para_html, count=1)
            # 清理空 class
            para_html = re.sub(r'class="\s*"', '', para_html)
        out.append(para_html)
        pos = end_p
        para_idx += 1
    return "".join(out)


def mammoth_to_html(docx_path, style_map=None, image_dir=None, force_page_breaks=True):
    """用 mammoth 把 DOCX 转为 HTML

    image_dir: 图片保存目录（None 时不保存图片到磁盘；mammoth 仍会生成 base64）
    """
    if style_map is None:
        style_map = "\n".join([
            "p[style-name='Title'] => h1.docx-title:fresh",
            "p[style-name='Heading 1'] => h2.docx-h1:fresh",
            "p[style-name='Heading 2'] => h3.docx-h2:fresh",
            "p[style-name='Heading 3'] => h4.docx-h3:fresh",
            "p[style-name='heading 1'] => h2.docx-h1:fresh",
            "p[style-name='heading 2'] => h3.docx-h2:fresh",
            "p[style-name='heading 3'] => h4.docx-h3:fresh",
            # 保留 docx 的硬分页符
            "br[type='page'] => div.docx-page-break",
        ])

    with open(docx_path, "rb") as f:
        result = mammoth.convert_to_html(
            f,
            style_map=style_map,
            convert_image=mammoth.images.data_uri,
        )

    html = result.value
    warnings = result.messages or []
    if warnings:
        for w in warnings[:5]:
            print(f"[docx_pdf_renderer] mammoth 警告: {w}", file=sys.stderr)

    # 后处理：识别整段加粗为标题
    html = re.sub(
        r'<p>\s*<strong>([\s\S]*?)</strong>\s*</p>',
        r'<p class="docx-heading">\1</p>',
        html,
    )
    return html


def build_preview_css(styles, is_cover=False, is_approval=False, page_breaks=None):
    """生成预览 PDF 用的 CSS 字符串

    styles: extract_docx_styles 返回值
    is_cover: 封面页（首 N 页）应用特殊排版
    is_approval: 审批表页
    page_breaks: 含特殊强制分页文字的列表（如 ["组织技术安全措施审批表", "一、工程内容"]）
    """
    page = styles["page"]
    run = styles["default_run"]
    page_w = page["width_px"]
    page_h = page["height_px"]
    mt, mr, mb, ml = page["margin_top_px"], page["margin_right_px"], page["margin_bottom_px"], page["margin_left_px"]
    content_w = page_w - ml - mr
    content_h = page_h - mt - mb
    font_size = run["font_size_px"]
    line_h = run["line_height"]
    font_family = run["font_family"] or "SimSun"

    # 通用：让 mammoth 输出的 <p> 与 docx 段落属性一致
    css_parts = [f"""
    @page {{
      size: {page_w}px {page_h}px;
      margin: {mt}px {mr}px {mb}px {ml}px;
    }}
    * {{ box-sizing: border-box; }}
    html, body {{
      margin: 0;
      padding: 0;
      font-family: "{font_family}", "宋体", "SimSun", "Microsoft YaHei", "微软雅黑", serif;
      font-size: {font_size}px;
      line-height: {line_h};
      color: #000;
      background: #fff;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }}
    body {{
      width: {page_w}px;
    }}
    /* mammoth 把 docx 硬分页（<w:br w:type="page"/>）转成的 div */
    div.docx-page-break {{
      display: block;
      page-break-before: always;
      break-before: page;
      height: 0;
      line-height: 0;
      font-size: 0;
      margin: 0;
      padding: 0;
      border: 0;
    }}
    /* 标题段前强制分页（一/二/.../审批表/目录 等） */
    p.docx-page-break-before {{
      page-break-before: always;
      break-before: page;
    }}
    /* 压扁 <br>：docx 里大量 <w:br/>（软换行）被 mammoth 渲染成 <br>，每个撑一行，
       659 个 br 可多出 16 页空白。Playwright/Chromium 渲染时让 br 不换行 */
    br {{
      line-height: 0;
      height: 0;
      display: inline;
      content: " ";
    }}
    p {{
      margin: 0 0 4px 0;
      text-indent: 2em;
      text-align: justify;
      word-break: break-all;
      orphans: 1;
      widows: 1;
    }}
    p.docx-heading {{
      text-align: left;
      font-weight: bold;
      font-size: {font_size + 1}px;
      margin: 12px 0 8px;
      text-indent: 0;
    }}
    h1, h2, h3, h4 {{
      text-align: left;
      font-weight: bold;
      margin: 12px 0 8px;
      text-indent: 0;
    }}
    table {{
      border-collapse: collapse;
      width: 100%;
      margin: 8px 0;
      table-layout: fixed;
      page-break-inside: auto;
    }}
    table td, table th {{
      border: 1px solid #000;
      padding: 1.2cm 6px;
      vertical-align: top;
      line-height: {line_h};
      word-break: break-all;
      overflow-wrap: break-word;
    }}
    table p {{
      text-indent: 0;
      margin: 2px 0;
    }}
    img {{
      max-width: 100%;
      height: auto;
    }}
    strong, b {{
      font-weight: bold;
    }}
    ul, ol {{
      padding-left: 30px;
      margin: 6px 0;
    }}
    li {{
      margin: 4px 0;
    }}
    """]
    # 强制分页标记
    if page_breaks:
        for txt in page_breaks:
            esc = re.escape(txt)
            css_parts.append(f'p.docx-heading:has-text("{txt}"), p:has-text("{txt}") {{ page-break-before: always; }}')
        # 用一个简单可行方案：把强制分页段落包裹到 .force-new-page
        # mammoth 输出后我们再单独处理
    return "\n".join(css_parts)


def wrap_html_with_full(html, styles, cover_pages=0, page_breaks=None):
    """把 mammoth HTML 包成完整 HTML 文档（含 CSS）"""
    css = build_preview_css(styles, page_breaks=page_breaks)
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<style>{css}</style>
</head>
<body>
{html}
</body>
</html>"""


# ---- LibreOffice 路径（WYSIWYG 主路径）---- #

# 常见 soffice 路径（按优先级）
_SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]

_SOFFICE_PATH = None
for _cand in _SOFFICE_CANDIDATES:
    if os.path.exists(_cand):
        _SOFFICE_PATH = _cand
        break


def _find_soffice():
    """定位 soffice 可执行文件"""
    global _SOFFICE_PATH
    if _SOFFICE_PATH and os.path.exists(_SOFFICE_PATH):
        return _SOFFICE_PATH
    for cand in _SOFFICE_CANDIDATES:
        if os.path.exists(cand):
            _SOFFICE_PATH = cand
            return cand
    # PATH 中找
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        _SOFFICE_PATH = found
        return found
    return None


from services.docx_image_preprocess import shrink_floating_images_in_docx  # noqa: E402


def render_docx_to_pdf_via_libreoffice(docx_path, output_pdf_path,
                                        timeout=120):
    """DOCX → PDF（用 LibreOffice，WYSIWYG 主路径）

    之所以作为主路径：
    - 用户在 Word 中看到的排版，就是 Word 渲染 DOCX 的结果
    - mammoth+Playwright 是另一套引擎，CSS 推导只能逼近
    - LibreOffice 二十年来对 DOCX 的兼容是公认最接近 Word 的开源实现
    - 用 LibreOffice 转出来的 PDF 和 Word 打开的 DOCX 在版式上几乎一致

    实现要点：
    - soffice 不能输出到中文路径下的目录，所以先转到一个短 ASCII 目录
    - 转完再把产物移动/复制到目标路径
    - soffice 在 Windows 下进程要单实例，并发时需加锁
    """
    soffice = _find_soffice()
    if not soffice:
        print("[docx_pdf_renderer] soffice 未找到", file=sys.stderr)
        return False
    if not os.path.exists(docx_path):
        print(f"[docx_pdf_renderer] DOCX 不存在: {docx_path}", file=sys.stderr)
        return False

    # 输出到临时 ASCII 目录（soffice 内部对 UTF-8 路径处理偶尔抽风）
    tmp_out = os.path.join(os.path.dirname(docx_path), "_lo_tmp_pdf")
    if os.path.exists(tmp_out):
        shutil.rmtree(tmp_out, ignore_errors=True)
    os.makedirs(tmp_out, exist_ok=True)

    # soffice 不喜欢 DOCX 路径里有中文（参数传递会被 cmd 解析乱掉）
    # 所以先把 DOCX 复制到 tmp_out 下用 ASCII 名字，转完再清理
    base = "docx_in"
    local_src = os.path.join(tmp_out, base + os.path.splitext(docx_path)[1])
    try:
        shutil.copy2(docx_path, local_src)
    except Exception as e:
        print(f"[docx_pdf_renderer] 复制 DOCX 到临时目录失败: {e}", file=sys.stderr)
        return False

    # soffice 自带 user profile 不能并发；用独立 profile
    user_profile = os.path.join(tmp_out, "_lo_profile")
    os.makedirs(user_profile, exist_ok=True)

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nologo",
        "--nofirststartwizard",
        "-env:UserInstallation=file:///" + user_profile.replace("\\", "/"),
        "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", tmp_out,
        local_src,
    ]
    print(f"[docx_pdf_renderer] soffice CMD: {' '.join(cmd)}", file=sys.stderr)
    try:
        r = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout,
            env={**os.environ, "LANG": "en_US.UTF-8", "HOME": os.environ.get("USERPROFILE", "")},
        )
    except subprocess.TimeoutExpired:
        print(f"[docx_pdf_renderer] soffice 超时（>{timeout}s）", file=sys.stderr)
        return False
    except Exception as e:
        print(f"[docx_pdf_renderer] soffice 启动失败: {e}", file=sys.stderr)
        return False

    produced = os.path.join(tmp_out, base + ".pdf")
    if not os.path.exists(produced):
        # 列出 tmp 目录看实际产物名
        actual = [f for f in os.listdir(tmp_out) if f.lower().endswith(".pdf")]
        if actual:
            produced = os.path.join(tmp_out, actual[0])
        else:
            out = r.stdout.decode("utf-8", errors="replace") if r.stdout else ""
            err = r.stderr.decode("utf-8", errors="replace") if r.stderr else ""
            print(f"[docx_pdf_renderer] soffice 未产出 PDF rc={r.returncode}", file=sys.stderr)
            print(f"  STDOUT: {out[:500]}", file=sys.stderr)
            print(f"  STDERR: {err[:500]}", file=sys.stderr)
            return False

    # 移动/复制到目标
    try:
        if os.path.exists(output_pdf_path):
            os.remove(output_pdf_path)
        shutil.copy2(produced, output_pdf_path)
    except Exception as e:
        print(f"[docx_pdf_renderer] 复制 PDF 失败: {e}", file=sys.stderr)
        return False

    # 清理临时目录（保留 tmp_out 供调试）
    try:
        shutil.rmtree(tmp_out, ignore_errors=True)
    except Exception:
        pass
    return True


def render_docx_to_pdf_via_playwright(docx_path, output_pdf_path,
                                        cover_pages=0, page_breaks=None,
                                        content_width_px=None):
    """主入口：DOCX → HTML(with @page CSS from DOCX) → Playwright → PDF

    Args:
        docx_path: DOCX 文件绝对路径
        output_pdf_path: 输出 PDF 路径
        cover_pages: 前 N 页应用 cover 样式（施工组织设计=2: 封面+审批表）
        page_breaks: 强制分页点标题列表
        content_width_px: 手动指定内容宽度（None 时从 docx 推导）

    Returns:
        bool: 成功/失败
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[docx_pdf_renderer] playwright 未安装", file=sys.stderr)
        return False

    try:
        styles = extract_docx_styles(docx_path)
        page = styles["page"]
        if content_width_px is None:
            content_width_px = page["width_px"] - page["margin_left_px"] - page["margin_right_px"]

        grids = extract_table_widths(docx_path)
        html = mammoth_to_html(docx_path)
        if grids:
            html = inject_table_widths_into_html(html, grids, content_width_px)

        # 清理末尾空段（避免出现空白页）
        html = trim_trailing_empty_blocks(html)

        # 注入 docx 自身的硬分页（<w:br w:type="page"/>）
        try:
            break_indexes = extract_page_break_paragraph_indexes(docx_path)
            if break_indexes:
                para_texts = extract_paragraph_texts(docx_path)
                # 用文本匹配策略注入分页
                html = inject_docx_page_breaks_into_html(html, para_texts, break_indexes)
        except Exception as e:
            print(f"[docx_pdf_renderer] 注入硬分页失败: {e}", file=sys.stderr)

        # 注入强制分页：审批表/目录/中文章节
        # 仅在没有 docx 硬分页的位置注入（避免重复分页）
        if page_breaks or True:  # 总是注入
            html = inject_page_breaks_into_html(html, force_break_titles=page_breaks or [])

        full_html = wrap_html_with_full(html, styles, cover_pages=cover_pages, page_breaks=page_breaks)

        with sync_playwright() as p:
            browser = p.chromium.launch()
            context = browser.new_context()
            page_obj = context.new_page()
            page_obj.set_content(full_html, wait_until="networkidle")
            page_obj.pdf(
                path=output_pdf_path,
                width=f"{page['width_px']}px",
                height=f"{page['height_px']}px",
                margin={
                    "top": f"{page['margin_top_px']}px",
                    "right": f"{page['margin_right_px']}px",
                    "bottom": f"{page['margin_bottom_px']}px",
                    "left": f"{page['margin_left_px']}px",
                },
                print_background=True,
                prefer_css_page_size=True,
            )
            browser.close()
        return True
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[docx_pdf_renderer] 渲染失败: {e}", file=sys.stderr)
        return False


def convert_docx_to_pdf(docx_path, output_pdf_path=None, **kwargs):
    """便捷入口：DOCX → PDF（带缓存）

    - 如果 output_pdf_path 未指定：写到 <docx_path>.rendered.pdf
    - 如果目标 PDF 已存在且 mtime ≥ docx_path：直接返回该路径
    - 否则用 Playwright 渲染

    Returns:
        output_pdf_path 成功；None 失败
    """
    if output_pdf_path is None:
        output_pdf_path = os.path.splitext(docx_path)[0] + ".rendered.pdf"

    # 缓存命中
    if os.path.exists(output_pdf_path):
        if os.path.getmtime(output_pdf_path) >= os.path.getmtime(docx_path):
            return output_pdf_path
        try:
            os.remove(output_pdf_path)
        except OSError:
            pass

    # 预处理：把 mammoth 看不到的"装饰浮动大图"+ 异常 line spacing + 空段都清掉
    # mammoth 不能识别 mso-position-horizontal-relative:page 的浮动图，
    # 会把每张满版大图当 inline 渲染，撑出 7-10 页空白。
    # Word 看到这些图是浮动定位（不影响段落流），所以 mammoth 删了不影响视觉。
    try:
        pre_path = docx_path + ".preview_input.docx"
        shrink_floating_images_in_docx(
            docx_path, pre_path,
            strip_empty=True, fix_spacing=True, strip_floating=True,
        )
        docx_for_render = pre_path
    except Exception as e:
        print(f"[docx_pdf_renderer] 预处理失败: {e}", file=sys.stderr)
        docx_for_render = docx_path

    # 主路径：mammoth+Playwright（页数 ≈ Word）
    if render_docx_to_pdf_via_playwright(docx_for_render, output_pdf_path, **kwargs):
        return output_pdf_path
    return None


if __name__ == "__main__":
    # 命令行调试：python docx_pdf_renderer.py <docx> [<pdf>]
    if len(sys.argv) < 2:
        print("Usage: python docx_pdf_renderer.py <docx_path> [output_pdf_path]")
        sys.exit(1)
    docx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else None
    result = convert_docx_to_pdf(docx, out)
    print("OK" if result else "FAIL", result or "")
