"""preprocess docx:
1. 把超大浮动图 width/height 缩到合理尺寸（避免 LibreOffice 拆成多页）
2. 删除纯空段（没有文字、没有图、没有分页符）—— mammoth 自动吞空段，LibreOffice 不会
3. **修复 line spacing**：把 `w:line="360" w:lineRule="auto"` 改成 `w:line="240" w:lineRule="auto"`
   Word 解释 line=360+auto 为"行高 ≥18pt"；LibreOffice 解释为"1.8 倍单倍行距"，
   这种语义差异让同一份 docx 在两个引擎下产生巨大页数差（如本项目 29 vs 401 页）。

原因：原 docx 模板生成时会插入 240+ 个空段 + 全局 line=360。
mammoth 渲染时全部忽略 → 23 页；LibreOffice 忠实按段分页 → 401 页。
Word 实际打开是 29 页。"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

# 大图尺寸上限（pt）。超过这个就缩
MAX_W_PT = 280.0
MAX_H_PT = 395.0


def _strip_floating_images(doc_xml):
    """删除超大浮动图（mso-position-horizontal-relative:page 且尺寸 > MAX）

    mammoth 不能正确处理 VML 浮动图，会把它们当 inline 渲染导致大量空白页。
    Word 看到这些图是浮动定位（不影响段落流），所以 mammoth 删了也不影响视觉。
    """
    # 找 v:shape 标签 + 包含 mso-position-horizontal-relative:page + width:Xpt
    # 注意：docx 中的 v: 前缀可能是 ns1/ns2 等序列化时由 ElementTree 改写过的
    # 这里用更宽松的匹配：任何带 :shape/:rect/:group 后缀的标签
    n_removed = 0

    def find_balanced_shape(text, start_pos):
        """找到 <*:shape ...> 对应的 </*:shape>"""
        # 找最近一个 <XXX:shape
        last_lt = -1
        for prefix in ["v:", "ns1:", "ns2:", "ns3:", "ns4:"]:
            p = text.rfind(f"<{prefix}shape", 0, start_pos + 1)
            if p > last_lt:
                last_lt = p
        if last_lt < 0:
            return None
        # 找 <...shape ...> 的开头
        tag_start = last_lt
        gt = text.find(">", tag_start)
        if gt < 0:
            return None
        if text[gt - 1] == "/":
            return (tag_start, gt + 1)
        # 找 </...shape>
        end = -1
        for prefix in ["v:", "ns1:", "ns2:", "ns3:", "ns4:"]:
            e = text.find(f"</{prefix}shape>", gt)
            if e >= 0 and (end < 0 or e < end):
                end = e
        if end < 0:
            return None
        return (tag_start, end + len("</X:shape>"))

    # 找所有 *:shape 起始（任何 namespace prefix）
    out = []
    cursor = 0
    matches = []
    for prefix in ["v:", "ns1:", "ns2:", "ns3:", "ns4:"]:
        for m in re.finditer(rf"<{prefix}shape\b", doc_xml):
            matches.append(m)
    matches.sort(key=lambda m: m.start())
    for m in matches:
        shape_start = m.start()
        bal = find_balanced_shape(doc_xml, m.end())
        if not bal:
            continue
        s, e = bal
        shape_text = doc_xml[s:e]
        # 判断是否是"页面绝对定位 + 超大尺寸"
        if "mso-position-horizontal-relative:page" in shape_text:
            wm = re.search(r"width:(\d+(?:\.\d+)?)pt", shape_text)
            if wm and float(wm.group(1)) > 200:
                # 删掉这一段（连同所在 <w:p>）
                p_start_alt = doc_xml.rfind("<ns0:p>", 0, shape_start)
                p_start_w = doc_xml.rfind("<w:p ", 0, shape_start)
                p_start_p = doc_xml.rfind("<w:p>", 0, shape_start)
                p_start = max(p_start_alt, p_start_w, p_start_p)
                if p_start >= 0:
                    p_end_alt = doc_xml.find("</ns0:p>", e)
                    p_end_w = doc_xml.find("</w:p>", e)
                    p_end = p_end_alt if p_end_alt > 0 else p_end_w
                    if p_end > 0:
                        p_end += 6
                        if doc_xml[cursor:p_start]:
                            out.append(doc_xml[cursor:p_start])
                        n_removed += 1
                        cursor = p_end
                        continue
        # 找外层段落
        p_start_alt = doc_xml.rfind("<ns0:p>", 0, shape_start)
        p_start_w = doc_xml.rfind("<w:p ", 0, shape_start)
        p_start_p = doc_xml.rfind("<w:p>", 0, shape_start)
        p_start = max(p_start_alt, p_start_w, p_start_p)
        if p_start >= 0:
            p_end_alt = doc_xml.find("</ns0:p>", e)
            p_end_w = doc_xml.find("</w:p>", e)
            p_end = p_end_alt if p_end_alt > 0 else p_end_w
            if p_end > 0:
                p_end += 6
                if doc_xml[cursor:p_start]:
                    out.append(doc_xml[cursor:p_start])
                n_removed += 1
                cursor = p_end
                continue
    out.append(doc_xml[cursor:])
    new_xml = "".join(out)
    if n_removed:
        print(f"[strip_floating_imgs] 删除 {n_removed} 个超大浮动图所在段", file=sys.stderr)
    return new_xml, n_removed


def _shrink_dimensions(doc_xml):
    """把超大 width:Xpt;height:Ypt 缩小到上限内"""
    def shrink_one(m):
        w, h = float(m.group(1)), float(m.group(2))
        if w <= MAX_W_PT and h <= MAX_H_PT:
            return m.group(0)
        scale = min(MAX_W_PT / w, MAX_H_PT / h)
        new_w = round(w * scale, 1)
        new_h = round(h * scale, 1)
        return f"width:{new_w}pt;height:{new_h}pt"

    pattern = re.compile(r"width:(\d+(?:\.\d+)?)pt;height:(\d+(?:\.\d+)?)pt")
    return pattern.subn(shrink_one, doc_xml)


def _fix_line_spacing(doc_xml):
    """修复 line spacing 不一致问题。

    现象：docx 大量用 <w:spacing w:line="360" w:lineRule="auto"/>
    Word 解析：line=360 + auto = 1.5 倍行距（360/240=1.5）
    LibreOffice 解析：行高 = 1.8 × 单倍行距（即使只有 0 内容也撑出整行）
    → LibreOffice 看到 200 个空段 + line=360 = 撑出几百页

    修复策略：
    - 不再改 line=360→240（那会让 mammoth CSS 用 1.0 倍，与 Word 1.5 倍差太远）
    - 保留 line=360，让 extract_docx_styles 读到 line_height=1.5，CSS 也用 1.5
    - 但 mammoth 不读 spacing，所以 CSS 行高由 extract_docx_styles 决定
    - 空段由 _strip_empty_paragraphs 处理，不影响页数
    """
    # 不再修改 line spacing，保持原值让 extract_docx_styles 正确读取
    return doc_xml, 0


def _strip_empty_paragraphs(doc_xml):
    """删除纯空段（无文字、无图、无分页符的 <w:p>）"""
    p_pattern = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.DOTALL)
    removed = 0
    total = 0
    pieces = []
    cursor = 0
    for m in p_pattern.finditer(doc_xml):
        total += 1
        p = m.group(0)
        text_joined = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)).strip()
        has_pict = "<w:pict>" in p or "<w:drawing>" in p
        has_pgbr = '<w:br w:type="page"' in p
        has_tab = "<w:tab/>" in p
        has_special_ppr = bool(re.search(r"<w:pBdr\b", p)) or bool(re.search(r"<w:numPr\b", p))
        is_empty = not text_joined and not has_pict and not has_pgbr and not has_tab and not has_special_ppr
        pieces.append(doc_xml[cursor:m.start()])
        if not is_empty:
            pieces.append(p)
        else:
            removed += 1
        cursor = m.end()
    pieces.append(doc_xml[cursor:])
    new_xml = "".join(pieces)
    print(f"[strip_empty_paras] {total} 段，删除 {removed} 个空段", file=sys.stderr)
    return new_xml, removed


def shrink_floating_images_in_docx(src_docx, dst_docx=None, strip_empty=True, fix_spacing=True, strip_floating=True):
    """复制 src 到 dst，把 document.xml 里的超大图缩小 + 删空段 + 修行距
    返回最终 docx 路径"""
    if dst_docx is None:
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp.close()
        dst_docx = tmp.name
    shutil.copy2(src_docx, dst_docx)

    with zipfile.ZipFile(src_docx, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        other_names = [n for n in zin.namelist() if n != "word/document.xml"]
        other_data = {n: zin.read(n) for n in other_names}

    n_floating = 0
    if strip_floating:
        doc_xml, n_floating = _strip_floating_images(doc_xml)

    new_xml, n_img = _shrink_dimensions(doc_xml)
    print(f"[shrink_images] 替换大图尺寸: {n_img} 处", file=sys.stderr)

    n_sp = 0
    if fix_spacing:
        new_xml, n_sp = _fix_line_spacing(new_xml)

    n_strip = 0
    if strip_empty:
        new_xml, n_strip = _strip_empty_paragraphs(new_xml)

    if n_img == 0 and n_strip == 0 and n_sp == 0 and n_floating == 0:
        return dst_docx

    tmp_dst = dst_docx + ".new"
    with zipfile.ZipFile(tmp_dst, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("word/document.xml", new_xml.encode("utf-8"))
        for n in other_data:
            zout.writestr(n, other_data[n])
    os.replace(tmp_dst, dst_docx)
    return dst_docx


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python docx_image_preprocess.py <src.docx> [dst.docx]")
        sys.exit(1)
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) >= 3 else None
    out = shrink_floating_images_in_docx(src, dst)
    print("OUTPUT:", out)


