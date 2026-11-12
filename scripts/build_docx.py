# -*- coding: utf-8 -*-
"""
课程标准生成器 —— 依据岗位能力分析，套邯郸科技职业学院《课程标准》模板生成 .docx

用法:
    python build_docx.py --config 课标内容.json --template 模板.docx --out 输出.docx

    --config 里是结构化课标内容（由 AI 根据岗位分析 + 课程名生成），
    本脚本只负责「套模板排版」，不做内容创作。

模板说明:
    模板 .doc 需先用 WPS/Word COM 转为 .docx 作填充母版（保留封面、字体、版式）。
    若未提供 --template，则使用内置 fallback 版式（无校徽，标准公文样式）。
"""
import argparse
import json
import os
import copy

try:
    import docx
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False


# ---------------------------------------------------------------------------
# 文本辅助
# ---------------------------------------------------------------------------

# 正文规范（对标成品样例）：宋体、小四(12pt)、1.5倍行距、首行缩进2字符
BODY_FONT = "宋体"
BODY_SIZE = 12          # pt（小四）
BODY_LINE_SPACING = 1.5
BODY_FIRST_INDENT = Pt(24)   # 2 字符 ≈ 24pt


def _set_run_font(run, font_name=BODY_FONT, size=BODY_SIZE, bold=None):
    """统一设置 run 字体（含中文字体 eastAsia），避免落入 Word 默认等线字体。"""
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def _fmt_body_paragraph(p, indent=True):
    """给段落设置正文格式：1.5倍行距 + 首行缩进2字符。"""
    pf = p.paragraph_format
    pf.line_spacing = BODY_LINE_SPACING
    if indent:
        pf.first_line_indent = BODY_FIRST_INDENT


def set_cell_text(cell, text, font_name="宋体", size=None, bold=None, align=None):
    """清空单元格并写入文本（保留首段格式），可指定字体/字号/加粗/对齐。"""
    # 保留第一个段落，清掉其余段落
    paras = cell.paragraphs
    p = paras[0]
    for extra in paras[1:]:
        extra._element.getparent().remove(extra._element)
    # 清空 runs
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(text)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if align is not None:
        p.alignment = align
    return p


def replace_paragraph_text(p, text):
    """保留段落格式，只替换文字。"""
    # 保留第一个 run 的格式，清掉其余 run
    runs = p.runs
    if runs:
        first = runs[0]
        for r in runs[1:]:
            r._element.getparent().remove(r._element)
        first.text = text
    else:
        p.add_run(text)


def find_paragraph(doc, keyword):
    """按关键词定位段落（精确匹配优先，前缀匹配兜底）。"""
    for p in doc.paragraphs:
        if p.text.strip() == keyword:
            return p
    for p in doc.paragraphs:
        if p.text.strip().startswith(keyword):
            return p
    return None


def get_col_text(header_cells):
    return [c.text.strip() for c in header_cells]


# ---------------------------------------------------------------------------
# 表1 动态重建
# ---------------------------------------------------------------------------

def _set_tc_text(tc, text, bold=None, center=False):
    """直接往物理 <w:tc> 写文本（绕过 row.cells 的 gridSpan/vMerge 索引错乱）。"""
    from docx.text.paragraph import Paragraph
    from docx.table import _Cell
    # 取 tc 的第一个 <w:p>
    p_el = tc.find(qn('w:p'))
    if p_el is None:
        from docx.oxml import OxmlElement
        p_el = OxmlElement('w:p')
        tc.append(p_el)
    p = Paragraph(p_el, None)
    # 清空已有 run
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(text)
    _set_run_font(run, size=12)   # 表1 内容统一宋体小四
    if bold is not None:
        run.font.bold = bold
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def rebuild_table1(table, projects):
    """
    用 projects 列表重填表1（直接操作底层 tc，兼容模板 gridSpan/vMerge 脏结构）。

    模板表1 真实结构（5 列宽，表头因 gridSpan 只有 4 个物理 tc）：
      表头(4tc):  [学习情境名称(跨2列)] [介绍] [学时建议] [目标]
      数据(5tc):  [序号] [项目名] [介绍] [学时] [目标]
    projects: [{no, name, intro, hours, goal}]
    """
    n = len(projects)
    rows = table.rows

    # 0) 清除数据行的残留 vMerge 标记（模板脏数据：每行都标 vMerge=restart，
    #    会导致序号列等被纵向合并，先清干净变成正常独立行）
    for r in rows[1:]:
        for tc in r._tr.tc_lst:
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                vm = tcPr.find(qn('w:vMerge'))
                if vm is not None:
                    tcPr.remove(vm)

    # 1) 表头：修正第一格为「学习情境名称」（跨2列合并格，成品即如此）
    hdr_tcs = rows[0]._tr.tc_lst
    if len(hdr_tcs) >= 1:
        _set_tc_text(hdr_tcs[0], "学习情境名称", bold=True, center=True)

    # 2) 数据行数量对齐：模板默认 8 个数据行
    data_rows = rows[1:]
    while len(data_rows) < n:
        from copy import deepcopy
        last = data_rows[-1]._tr
        new_tr = deepcopy(last)
        last.addnext(new_tr)
        data_rows = table.rows[1:]

    # 3) 填数据：数据行 5 个物理 tc = [序号, 项目名, 介绍, 学时, 目标]
    for idx, proj in enumerate(projects):
        tr = data_rows[idx]._tr
        tcs = tr.tc_lst
        if len(tcs) < 5:
            continue
        _set_tc_text(tcs[0], str(proj.get("no", idx + 1)), center=True)
        _set_tc_text(tcs[1], proj.get("name", ""))
        _set_tc_text(tcs[2], proj.get("intro", ""))
        _set_tc_text(tcs[3], str(proj.get("hours", "")), center=True)
        _set_tc_text(tcs[4], proj.get("goal", ""))

    # 4) 多余旧数据行清空
    for r in data_rows[n:]:
        for tc in r._tr.tc_lst:
            _set_tc_text(tc, "")


# ---------------------------------------------------------------------------
# 简介页元数据填充
# ---------------------------------------------------------------------------

def fill_intro(doc, meta):
    """填充简介页：值插入到标签冒号之后，替换下划线/占位文字。"""
    # 每个标签的填法：找到「标签[:：]」段，清掉冒号后的所有 run，追加「 值」
    fields = [
        ("课程名称", meta.get("course_name", "")),
        ("课程代码", meta.get("course_code", "")),
        ("课程类别", meta.get("course_type", "")),
        ("适用专业", meta.get("major", "")),
        ("学分", meta.get("credit", "")),
        ("编写执笔人", meta.get("author", "")),
        ("审定负责人", meta.get("reviewer", "")),
    ]
    started = False
    for p in doc.paragraphs:
        t = p.text.strip()
        if t == "简   介" or t == "简介":
            started = True
            continue
        if started and t.startswith("一、"):
            break
        if not started:
            continue
        # 处理各字段
        for label, val in fields:
            if val and t.startswith(label):
                _fill_intro_value(p, label, val)
                break
        # 参考学时 / 理论学时：替换下划线 run 为值
        if t.startswith("参考学时") and meta.get("total_hours"):
            _fill_hours(p, meta["total_hours"], meta.get("total_hours", ""))
        if t.startswith("理论学时") and meta.get("theory_hours"):
            _fill_hours(p, meta["theory_hours"], meta.get("practice_hours", ""))


def _fill_hours(p, first_val, second_val):
    """替换学时段落里带下划线的 run（值套宋体小四）。

    参考学时段落: 下划线 run = [参考学时值位, 总学时值位1, 总学时值位2]
    理论学时段落: 下划线 run = [理论学时值位, 课内实践学时值位]
    策略：第一个下划线 run 填 first_val，其余下划线 run 清空，再在「总学时/课内实践学时」冒号后填 second_val。
    """
    underline_runs = [r for r in p.runs if r.font.underline]
    if not underline_runs:
        return
    # 第一个下划线 run → first_val
    underline_runs[0].text = " " + str(first_val) + " "
    underline_runs[0].font.underline = False
    _set_run_font(underline_runs[0])
    # 其余下划线 run 清空（避免残留下划线）
    for r in underline_runs[1:]:
        r.text = ""
        r.font.underline = False
    # 在「总学时/课内实践学时」冒号后的第一个下划线 run 填 second_val
    # 但已清空，改为：定位次级标签后的 run 填值
    if second_val:
        full = p.text
        for sub in ["总学时", "课内实践学时"]:
            idx = full.find(sub)
            if idx != -1:
                # 找冒号
                colon = -1
                for ch in [':', '：']:
                    pos = full.find(ch, idx + len(sub))
                    if pos != -1:
                        colon = pos
                        break
                if colon != -1:
                    # 直接重建该次级标签后的内容太复杂，改为在段落末尾补
                    # 用简单方式：找到 sub 标签对应 run，在其后 run 填值
                    for ri, r in enumerate(p.runs):
                        if sub in r.text or (ri > 0 and sub in (p.runs[ri-1].text + r.text)):
                            # 下一个非空 run 填值
                            for r2 in p.runs[ri+1:]:
                                if r2.text.strip() == "" or r2.text.strip() in (" ", ""):
                                    r2.text = " " + str(second_val) + " "
                                    _set_run_font(r2)
                                    break
                            break
                break


def _fill_intro_value(p, label, val):
    """把值填到标签冒号后：保留标签原有格式（楷体加粗），只清掉冒号后的 run 追加「 值」。"""
    full = p.text
    # 找到冒号位置（标签之后）
    colon_idx = -1
    for ch in [':', '：']:
        pos = full.find(ch, len(label))
        if pos != -1:
            colon_idx = pos
            break
    if colon_idx == -1:
        return
    # 定位冒号所在的 run，保留它及之前的 run（标签部分），清掉其后的 run
    runs = p.runs
    # 计算每个 run 的文本累计长度，找到冒号所在的 run 索引
    acc = 0
    colon_run_idx = -1
    for ri, r in enumerate(runs):
        if acc <= colon_idx < acc + len(r.text):
            colon_run_idx = ri
            break
        acc += len(r.text)
    if colon_run_idx == -1:
        return
    # 冒号 run 内冒号之后的部分截掉，只留到冒号为止
    colon_r = runs[colon_run_idx]
    off = colon_idx - acc
    colon_r.text = colon_r.text[:off + 1]
    # 删掉冒号 run 之后的所有 run
    for r in runs[colon_run_idx + 1:]:
        r._element.getparent().remove(r._element)
    # 追加值（宋体小四）
    r_val = p.add_run(" " + val)
    _set_run_font(r_val)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def build(config_path, template_path, out_path):
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    use_template = bool(template_path and os.path.exists(template_path))
    if use_template:
        doc = Document(template_path)
    else:
        # fallback 版式已生成完整文档，直接保存返回
        doc = build_fallback_doc(cfg)
        doc.save(out_path)
        return out_path

    meta = cfg.get("meta", {})

    # ---- 封面表格：课程名称/所属教研室/所属系部 ----
    if doc.tables:
        cover = doc.tables[0]
        cover_vals = [
            meta.get("course_name", ""),
            meta.get("department", ""),      # 所属教研室
            meta.get("faculty", ""),         # 所属系部
        ]
        for i, val in enumerate(cover_vals):
            if i < len(cover.rows):
                set_cell_text(cover.rows[i].cells[1], val, size=22)

    # ---- 简介页 ----
    fill_intro(doc, meta)

    # ---- 一、课程性质（正文另起一页，更贴近正式文档）----
    p = find_paragraph(doc, "一、课程性质")
    if p:
        _add_page_break_before(p)
        # 在「一、课程性质」标题后插入若干段
        insert_after_paragraph(doc, p, cfg.get("nature", []))

    # ---- 二、课程目标 ----
    fill_objectives(doc, cfg.get("objectives", {}))

    # ---- 三、教学内容与学时分配 ----
    fill_section3(doc, cfg.get("projects", []), cfg.get("section3_intro", ""))

    # ---- 四、考核方式 ----
    fill_simple_section(doc, "四、考核方式", cfg.get("assessment", ""))

    # ---- 五、参考教材 ----
    fill_simple_section(doc, "五、参考教材", cfg.get("textbooks", ""))

    # ---- 六、教学建议 ----
    fill_simple_section(doc, "六、教学建议", cfg.get("suggestions", ""))

    doc.save(out_path)
    return out_path


def _add_page_break_before(p):
    """在段落前插入分页符（pageBreakBefore），让正文另起一页。"""
    from docx.oxml import OxmlElement
    pPr = p._element.get_or_add_pPr()
    # 移除已有的 pageBreakBefore
    for el in pPr.findall(qn('w:pageBreakBefore')):
        pPr.remove(el)
    pb = OxmlElement('w:pageBreakBefore')
    pPr.append(pb)


def _new_empty_paragraph_after(anchor_p):
    """在 anchor_p 后插入一个空段落，返回新 Paragraph（格式同 anchor_p 的正文样式）。"""
    from docx.text.paragraph import Paragraph
    from docx.oxml import OxmlElement
    new_p_el = OxmlElement('w:p')
    anchor_p._element.addnext(new_p_el)
    np_ = Paragraph(new_p_el, anchor_p._parent)
    # 复制锚点段落属性（对齐、缩进等），避免新段落裸奔
    pPr = anchor_p._element.find(qn('w:pPr'))
    if pPr is not None:
        import copy as _copy
        pPr_copy = _copy.deepcopy(pPr)
        # 剔除分页属性，避免正文段被误分页
        for tag in ('w:pageBreakBefore', 'w:pageBreakAfter'):
            for el in pPr_copy.findall(qn(tag)):
                pPr_copy.remove(el)
        np_._element.insert(0, pPr_copy)
    return np_


def _append_paragraphs_after(anchor_p, lines):
    """在 anchor_p 后依次插入 lines 段，套正文格式（宋体小四+1.5倍行距+首行缩进）。"""
    result = []
    cur = anchor_p
    for line in lines:
        np_ = _new_empty_paragraph_after(cur)
        run = np_.add_run(line)
        _set_run_font(run)
        _fmt_body_paragraph(np_)
        result.append(np_)
        cur = np_
    return result


def insert_after_paragraph(doc, anchor_p, lines):
    """在锚点段落后插入多段文本。"""
    _append_paragraphs_after(anchor_p, lines)


def fill_objectives(doc, objectives):
    """填充课程目标：能力目标(A)/知识目标(K)/素质目标(Q)，在各标签后插入。"""
    sections = [
        ("（1）能力目标", objectives.get("ability", [])),
        ("（2）知识目标", objectives.get("knowledge", [])),
        ("（3）素质目标", objectives.get("quality", [])),
    ]
    for label, items in sections:
        p = find_paragraph(doc, label)
        if not p:
            continue
        _append_paragraphs_after(p, items)


def fill_section3(doc, projects, intro_text):
    """填充三、教学内容与学时分配 + 表1。"""
    p = find_paragraph(doc, "三、教学内容与学时分配")
    if p and intro_text:
        # 修改「本课程共 N 个递进式教学项目」段（标题后第一段）
        from docx.text.paragraph import Paragraph
        parent = p._element.getparent()
        idx = list(parent).index(p._element)
        nxt = Paragraph(parent[idx + 1], p._parent) if idx + 1 < len(parent) else None
        if nxt is not None and "教学项目" in nxt.text:
            replace_paragraph_text(nxt, intro_text)

    # 表1 是第 2 张表（index 1）
    if len(doc.tables) >= 2:
        rebuild_table1(doc.tables[1], projects)


def fill_simple_section(doc, title, content):
    """在标题段落后插入正文（支持多行，用 \\n 分隔）。"""
    p = find_paragraph(doc, title)
    if not p or not content:
        return
    lines = content.split("\n")
    _append_paragraphs_after(p, lines)


def build_fallback_doc(cfg):
    """无模板时的兜底：标准公文样式生成。"""
    doc = Document()
    meta = cfg.get("meta", {})
    # 标题
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(meta.get("course_name", "") + " 课程标准")
    r.font.size = Pt(22)
    r.bold = True

    doc.add_paragraph("一、课程性质")
    for line in cfg.get("nature", []):
        doc.add_paragraph(line)

    doc.add_paragraph("二、课程目标")
    for label, key in [("（1）能力目标", "ability"), ("（2）知识目标", "knowledge"), ("（3）素质目标", "quality")]:
        doc.add_paragraph(label)
        for item in cfg.get("objectives", {}).get(key, []):
            doc.add_paragraph(item)

    doc.add_paragraph("三、教学内容与学时分配")
    doc.add_paragraph(cfg.get("section3_intro", ""))
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    for i, h in enumerate(["序号", "学习情境名称", "教学内容介绍", "学时建议", "目标"]):
        hdr[i].text = h
    for proj in cfg.get("projects", []):
        cells = t.add_row().cells
        cells[0].text = str(proj.get("no", ""))
        cells[1].text = proj.get("name", "")
        cells[2].text = proj.get("intro", "")
        cells[3].text = str(proj.get("hours", ""))
        cells[4].text = proj.get("goal", "")

    for title, key in [("四、考核方式", "assessment"), ("五、参考教材", "textbooks"), ("六、教学建议", "suggestions")]:
        doc.add_paragraph(title)
        for line in cfg.get(key, "").split("\n"):
            doc.add_paragraph(line)

    return doc


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True)
    ap.add_argument("--template", default="")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    if not HAS_DOCX:
        raise SystemExit("python-docx 未安装，请先: pip install python-docx")
    out = build(args.config, args.template, args.out)
    print("OK", out)
