#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
wordfmt.py —— Word(.docx) 批量格式引擎（一次调用完成全文档排版）

解决痛点：MCP/编辑器逐段改样式极慢（几十次调用、每次只改一小部分）。
本脚本一次性遍历 样式定义 + docDefaults + 全部段落/表格/Run，秒级完成。

子命令
  inspect   FILE            格式体检（字号/行距/字体/页边距分布，定位直设格式）
  apply     FILE [选项]     按预设批量应用格式
  md2docx   FILE [选项]     Markdown -> docx（含标题层级、上标引注、三线表）
  replace   FILE OLD NEW    保格式查找替换（run 级最小跨度编辑，正文+表格+页眉页脚）
  revisions FILE [选项]     修订记录：列出 / 全部接受 / 全部拒绝（可按作者过滤）
  headers   FILE [选项]     页眉页脚：查看 / 设置文本（各节 + 首页 + 奇偶页）

apply 选项
  --preset NAME            预设名（presets/NAME.json）或内置 official/academic
  --set a.b=v              临时覆盖，如 --set body.size=12 --set page.margin_left_cm=3.18
  --out FILE               输出到新文件（默认原地覆盖，先备份）
  --scope body|headings|tables|style|all   作用范围，默认 all
  --force/--no-force       是否强制改写 run/段落直设格式（默认 force，防直设覆盖）
  --report                 输出修改统计

replace 选项
  --scope all|body|tables|headers|footers  搜索范围（默认 all 含页眉页脚）
  --regex                  OLD 为正则、NEW 支持 \\1 反向引用
  --count N                每段最多替换 N 处（默认不限）
  --dry-run                只统计不写入
  --out FILE               输出到新文件

revisions 选项
  --action list|accept|reject   默认 list（accept=接受全部修订出干净版）
  --author NAME            只处理该作者的修订
  --out FILE               输出到新文件

headers 选项
  --set-header TEXT        设置各节默认页眉文本
  --set-footer TEXT        设置各节默认页脚文本
  --align left|center|right|both   设置时对齐，默认 center
  --out FILE               输出到新文件

示例
  wordfmt.py apply a.docx --preset cn-formal --report
  wordfmt.py inspect a.docx
  wordfmt.py replace a.docx "北京市朝阳区XX路1号" "北京市海淀区YY大街2号" --dry-run
  wordfmt.py replace a.docx "20(\\d{2})年" "公元\\1年" --regex
  wordfmt.py revisions a.docx                          # 列出修订
  wordfmt.py revisions a.docx --action accept          # 接受全部 -> 干净终版
  wordfmt.py revisions a.docx --action reject --author 张三
  wordfmt.py headers a.docx --set-header "XX大学教务处"
"""
import argparse
import copy
import json
import os
import re
import shutil
import sys
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
PRESET_DIR = os.path.join(os.path.dirname(HERE), "presets")

# 中文字号 -> pt
CN_SIZE = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}
# 常见中文字体 -> eastAsia / ascii 配对
FONT_EN_DEFAULT = "Times New Roman"

BUILTIN = {
    # 杨老师教材立项申报书 / 学术申报材料规范
    "cn-formal": {
        "page": {"size": "A4", "margin_left_cm": 3.18, "margin_right_cm": 3.18,
                 "margin_top_cm": 2.54, "margin_bottom_cm": 2.54},
        "body": {"font_cn": "仿宋_GB2312", "font_en": "Times New Roman", "size": 12,
                 "line_spacing_multiple": 1.5, "first_line_indent_chars": 2,
                 "align": "both", "space_after_pt": 0},
        "headings": {
            "1": {"font_cn": "黑体", "size": 14, "bold": True, "align": "center",
                  "line_spacing_multiple": 1.5},
            "2": {"font_cn": "黑体", "size": 14, "bold": True, "align": "left",
                  "line_spacing_multiple": 1.5, "space_before_pt": 6, "space_after_pt": 3},
            "3": {"font_cn": "仿宋_GB2312", "size": 14, "bold": True, "align": "left",
                  "line_spacing_multiple": 1.5},
            "4": {"font_cn": "仿宋_GB2312", "size": 12, "bold": True, "align": "left",
                  "line_spacing_multiple": 1.5},
        },
        "table": {"font_cn": "仿宋_GB2312", "size": 12, "line_spacing_multiple": 1.5,
                  "align": "center"},
    },
    # 党政机关公文 GB/T 9704-2012
    "official": {
        "page": {"size": "A4", "margin_left_cm": 2.8, "margin_right_cm": 2.6,
                 "margin_top_cm": 3.7, "margin_bottom_cm": 3.5},
        "body": {"font_cn": "仿宋_GB2312", "font_en": "Times New Roman", "size": 16,
                 "line_spacing_exact_pt": 28, "first_line_indent_chars": 2, "align": "both"},
        "headings": {
            "1": {"font_cn": "黑体", "size": 16, "bold": False, "align": "left",
                  "line_spacing_exact_pt": 28, "first_line_indent_chars": 2},
            "2": {"font_cn": "楷体_GB2312", "size": 16, "bold": False, "align": "left",
                  "line_spacing_exact_pt": 28, "first_line_indent_chars": 2},
            "3": {"font_cn": "仿宋_GB2312", "size": 16, "bold": True, "align": "left",
                  "line_spacing_exact_pt": 28, "first_line_indent_chars": 2},
        },
        "table": {"font_cn": "仿宋_GB2312", "size": 12, "line_spacing_exact_pt": 18},
    },
    # 学术论文（期刊常见）
    "academic": {
        "page": {"size": "A4", "margin_left_cm": 3.0, "margin_right_cm": 3.0,
                 "margin_top_cm": 2.5, "margin_bottom_cm": 2.5},
        "body": {"font_cn": "宋体", "font_en": "Times New Roman", "size": 10.5,
                 "line_spacing_multiple": 1.5, "first_line_indent_chars": 2, "align": "both"},
        "headings": {
            "1": {"font_cn": "黑体", "size": 14, "bold": False, "align": "center",
                  "line_spacing_multiple": 1.5},
            "2": {"font_cn": "黑体", "size": 12, "bold": False, "align": "left",
                  "line_spacing_multiple": 1.5},
            "3": {"font_cn": "宋体", "size": 10.5, "bold": True, "align": "left",
                  "line_spacing_multiple": 1.5},
        },
        "table": {"font_cn": "宋体", "size": 9, "line_spacing_multiple": 1.0, "align": "center"},
    },
}


def load_preset(name):
    if name in BUILTIN:
        return copy.deepcopy(BUILTIN[name])
    p = os.path.join(PRESET_DIR, f"{name}.json")
    if os.path.isfile(p):
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data.get("preset", data)
    p2 = name if os.path.isfile(str(name)) else None
    if p2:
        with open(p2, encoding="utf-8") as f:
            data = json.load(f)
        return data.get("preset", data)
    raise SystemExit(f"[错误] 未找到预设: {name}（内置: {', '.join(BUILTIN)}）")


def apply_overrides(preset, overrides):
    for item in overrides or []:
        if "=" not in item:
            continue
        path, val = item.split("=", 1)
        try:
            val = json.loads(val)
        except Exception:
            if val.lower() in ("true", "false"):
                val = val.lower() == "true"
        node = preset
        keys = path.split(".")
        for k in keys[:-1]:
            node = node.setdefault(k, {})
        node[keys[-1]] = val
    return preset


def to_pt(size):
    """数字=pt；中文名（三号/小四）转 pt。"""
    if size is None:
        return None
    if isinstance(size, (int, float)):
        return float(size)
    s = str(size).strip()
    if s in CN_SIZE:
        return float(CN_SIZE[s])
    m = re.match(r"^([\d.]+)\s*(pt)?$", s, re.I)
    if m:
        return float(m.group(1))
    raise SystemExit(f"[错误] 无法识别字号: {size}")


ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT, "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def set_run_font(run, font_cn=None, font_en=None, size_pt=None, bold=None):
    if font_cn:
        run.font.name = font_en or font_cn
        rpr = run._element.get_or_add_rPr()
        rf = rpr.get_or_add_rFonts()
        rf.set(qn("w:eastAsia"), font_cn)
        if font_en:
            rf.set(qn("w:ascii"), font_en)
            rf.set(qn("w:hAnsi"), font_en)
    if size_pt:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold


def set_para_format(p, fmt, size_pt):
    pf = p.paragraph_format
    if fmt.get("align") in ALIGN_MAP:
        pf.alignment = ALIGN_MAP[fmt["align"]]
    if "line_spacing_multiple" in fmt and fmt["line_spacing_multiple"]:
        pf.line_spacing = float(fmt["line_spacing_multiple"])
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    elif fmt.get("line_spacing_exact_pt"):
        pf.line_spacing = Pt(float(fmt["line_spacing_exact_pt"]))
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    if fmt.get("first_line_indent_chars") is not None and size_pt:
        pf.first_line_indent = Pt(size_pt * float(fmt["first_line_indent_chars"]))
    if fmt.get("space_before_pt") is not None:
        pf.space_before = Pt(float(fmt["space_before_pt"]))
    if fmt.get("space_after_pt") is not None:
        pf.space_after = Pt(float(fmt["space_after_pt"]))


def heading_level_of(p):
    """返回 1-9 或 None。兼容 style.name 与 outlineLvl。"""
    name = (p.style.name or "").lower()
    m = re.match(r"heading (\d)", name)
    if m:
        return int(m.group(1))
    if name.startswith("title"):
        return 1
    ppr = p._element.pPr
    if ppr is not None:
        ol = ppr.find(qn("w:outlineLvl"))
        if ol is not None and ol.get(qn("w:val")) is not None:
            try:
                return int(ol.get(qn("w:val"))) + 1
            except ValueError:
                pass
    return None


def iter_block_paragraphs(doc):
    """遍历正文段落 + 表格内段落（含嵌套表）。"""
    for p in doc.paragraphs:
        yield p, "body"
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p, "table"


def set_page(doc, page):
    if not page:
        return
    for sec in doc.sections:
        if str(page.get("size", "")).upper() == "A4":
            sec.page_width = Cm(21)
            sec.page_height = Cm(29.7)
        for key, attr in (("margin_left_cm", "left_margin"), ("margin_right_cm", "right_margin"),
                          ("margin_top_cm", "top_margin"), ("margin_bottom_cm", "bottom_margin")):
            if page.get(key) is not None:
                setattr(sec, attr, Cm(float(page[key])))


def style_definitions(doc, preset, force=True):
    """改样式定义（Normal / Heading 1-9），让后续新增内容自动合规。"""
    body = preset.get("body", {})
    heads = preset.get("headings", {})
    size_pt = to_pt(body.get("size"))
    try:
        st = doc.styles["Normal"]
        if body.get("font_cn"):
            st.font.name = body.get("font_en", FONT_EN_DEFAULT)
            st.element.rPr.rFonts.set(qn("w:eastAsia"), body["font_cn"])
        if size_pt:
            st.font.size = Pt(size_pt)
        if body.get("line_spacing_multiple"):
            st.paragraph_format.line_spacing = float(body["line_spacing_multiple"])
    except KeyError:
        pass
    for lvl, fmt in heads.items():
        for nm in (f"Heading {lvl}", f"标题 {lvl}"):
            try:
                st = doc.styles[nm]
            except KeyError:
                continue
            sz = to_pt(fmt.get("size")) or size_pt
            if fmt.get("font_cn"):
                st.font.name = fmt.get("font_en", FONT_EN_DEFAULT)
                st.element.rPr.rFonts.set(qn("w:eastAsia"), fmt["font_cn"])
            if sz:
                st.font.size = Pt(sz)
            if fmt.get("bold") is not None:
                st.font.bold = fmt["bold"]
            if fmt.get("line_spacing_multiple"):
                st.paragraph_format.line_spacing = float(fmt["line_spacing_multiple"])


def cmd_apply(args):
    preset = apply_overrides(load_preset(args.preset), args.set)
    src = args.file
    if not os.path.isfile(src):
        raise SystemExit(f"[错误] 文件不存在: {src}")
    out = args.out
    if not out:
        bak = src.replace(".docx", "") + ".bak-" + datetime.now().strftime("%H%M%S") + ".docx"
        shutil.copy2(src, bak)
        out = src
    doc = Document(src)

    scope = args.scope
    stats = {"para": 0, "run": 0, "table_para": 0, "head": 0, "page": 0}
    body = preset.get("body", {})
    heads = preset.get("headings", {})
    tbl = preset.get("table", {})
    body_size = to_pt(body.get("size"))

    if scope in ("all", "page"):
        set_page(doc, preset.get("page"))
        stats["page"] = len(doc.sections)
    if scope in ("all", "style"):
        style_definitions(doc, preset, force=args.force)

    for p, kind in iter_block_paragraphs(doc):
        lvl = heading_level_of(p)
        if kind == "table":
            fmt = tbl or body
            size_pt = to_pt(fmt.get("size")) or body_size
            stats["table_para"] += 1
        elif lvl and str(lvl) in heads:
            fmt = heads[str(lvl)]
            size_pt = to_pt(fmt.get("size")) or body_size
            stats["head"] += 1
        else:
            if scope not in ("all", "body"):
                continue
            fmt = body
            size_pt = body_size
            stats["para"] += 1
        if scope == "headings" and not (lvl and str(lvl) in heads):
            continue
        if scope == "tables" and kind != "table":
            continue
        if scope == "body" and (kind == "table" or lvl):
            continue

        set_para_format(p, fmt, size_pt)
        if args.force:
            for r in p.runs:
                if not r.text:
                    continue
                # 上标引注 [n] 不放大：保持原有 vertical_align
                set_run_font(r, fmt.get("font_cn"), fmt.get("font_en", FONT_EN_DEFAULT),
                             size_pt, fmt.get("bold") if "bold" in fmt else None)
                stats["run"] += 1

    doc.save(out)
    print(f"[完成] 已排版 -> {out}")
    if not args.out:
        print(f"[备份] 原文件已备份为 {bak}")
    if args.report:
        print("[统计] " + " | ".join(f"{k}={v}" for k, v in stats.items()))
    return 0


def cmd_inspect(args):
    doc = Document(args.file)
    sizes, spacing, fonts, aligns = {}, {}, {}, {}
    for p, _ in iter_block_paragraphs(doc):
        lvl = heading_level_of(p)
        tag = f"H{lvl}" if lvl else "正文"
        for r in p.runs:
            if not r.text.strip():
                continue
            sz = r.font.size.pt if r.font.size else None
            sizes.setdefault(tag, {}).setdefault(sz, 0)
            sizes[tag][sz] += 1
            fn = r.font.name or "(继承)"
            rpr = r._element.rPr
            ea = None
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None:
                    ea = rf.get(qn("w:eastAsia"))
            label = f"{ea or fn}" + (f"/{fn}" if ea and fn != ea else "")
            fonts.setdefault(tag, {}).setdefault(label, 0)
            fonts[tag][label] += 1
        pf = p.paragraph_format
        key = (pf.line_spacing, str(pf.line_spacing_rule))
        spacing.setdefault(tag, {}).setdefault(key, 0)
        spacing[tag][key] += 1
        aligns.setdefault(tag, {}).setdefault(str(pf.alignment), 0)
        aligns[tag][str(pf.alignment)] += 1
    sec = doc.sections[0]
    print("=" * 56)
    print(" 格式体检报告:", os.path.basename(args.file))
    print("=" * 56)
    print(f"页面: 宽{sec.page_width.cm:.2f}cm 高{sec.page_height.cm:.2f}cm | "
          f"页边距 左{sec.left_margin.cm:.2f} 右{sec.right_margin.cm:.2f} "
          f"上{sec.top_margin.cm:.2f} 下{sec.bottom_margin.cm:.2f} cm")
    print(f"段落总数: {len(doc.paragraphs)} | 表格数: {len(doc.tables)}")
    for tag in sorted(sizes):
        print(f"\n[{tag}]")
        print("  字号(pt):", ", ".join(f"{k}×{v}" for k, v in sorted(sizes[tag].items(), key=lambda x: -x[1])))
        print("  行距    :", ", ".join(f"{k[0]}/{k[1].split(' ')[0]}×{v}" for k, v in spacing.get(tag, {}).items()))
        print("  字体    :", ", ".join(f"{k}×{v}" for k, v in sorted(fonts[tag].items(), key=lambda x: -x[1])[:4]))
    return 0


SUP_RE = re.compile(r"\[(\d+(?:[-,]\d+)*)\]")


# ============================================================
# 2.0 新增：保格式替换 / 修订记录 / 页眉页脚
# 核心原则（吸收自 word-docx skill 的方法论）：
#   1. 最小跨度编辑——只动受影响的文字 span，不重写整段，防止格式漂移
#   2. 文字可能跨 run 拆分——替换前先重建全段文本再映射回 run
#   3. 修订记录是 OOXML 层结构（w:ins/w:del），python-docx 不暴露，须直接操作 XML
# ============================================================

def _iter_scope_paragraphs(doc, scope):
    """按范围遍历段落，yield (paragraph, where)。scope: all/body/tables/headers/footers。"""
    if scope in ("all", "body"):
        for p in doc.paragraphs:
            yield p, "正文"
    if scope in ("all", "tables"):
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p, "表格"
    if scope in ("all", "headers", "footers"):
        for si, sec in enumerate(doc.sections, 1):
            parts = (("页眉", sec.header), ("页脚", sec.footer),
                     ("页眉(首页)", sec.first_page_header), ("页脚(首页)", sec.first_page_footer),
                     ("页眉(奇偶)", sec.even_page_header), ("页脚(奇偶)", sec.even_page_footer))
            for label, hf in parts:
                if scope == "headers" and not label.startswith("页眉"):
                    continue
                if scope == "footers" and not label.startswith("页脚"):
                    continue
                try:
                    if hf.is_linked_to_previous:
                        continue
                except Exception:
                    pass
                for p in hf.paragraphs:
                    yield p, f"{label}#节{si}"


def replace_in_paragraph(p, old, new, regex=False, max_count=0):
    """run 级保格式替换：返回替换次数。

    实现：拼出全段文本 -> 定位匹配 span -> 从后往前逐 span 修改受影响 run。
    替换文字继承 span 起点 run 的格式（最小跨度编辑，其余文字/格式零扰动）。
    """
    runs = list(p.runs)
    if not runs:
        return 0
    full = "".join(r.text for r in runs)
    if regex:
        try:
            spans = [(m.start(), m.end(), m.expand(new)) for m in re.finditer(old, full)]
        except re.error as e:
            raise SystemExit(f"[错误] 正则无效: {e}")
    else:
        spans = []
        idx = full.find(old)
        while idx >= 0:
            spans.append((idx, idx + len(old), new))
            idx = full.find(old, idx + len(old))
    if not spans:
        return 0
    if max_count:
        spans = spans[:max_count]
    # run 边界表：(run, start, end)
    bounds, pos = [], 0
    for r in runs:
        L = len(r.text)
        bounds.append((r, pos, pos + L))
        pos += L
    # 从后往前处理：早期 span 的字符位置不受后期编辑影响
    n = 0
    for s, e, repl in reversed(spans):
        affected = [b for b in bounds if b[1] < e and b[2] > s]
        if not affected:
            continue
        r0, s0, _ = affected[0]
        r1, s1, _ = affected[-1]
        off0, off1 = s - s0, e - s1
        if r0 is r1:
            r0.text = r0.text[:off0] + repl + r0.text[off1:]
        else:
            r0.text = r0.text[:off0] + repl
            for r, _, _ in affected[1:-1]:
                r.text = ""
            r1.text = r1.text[off1:]
        n += 1
    return n


def _save_with_backup(doc, src, out, extra_note=""):
    if out and out != src:
        doc.save(out)
        print(f"[完成] 已保存 -> {out}")
    else:
        bak = src.replace(".docx", "") + ".bak-" + datetime.now().strftime("%H%M%S") + ".docx"
        shutil.copy2(src, bak)
        doc.save(src)
        print(f"[完成] 已保存 -> {src}")
        print(f"[备份] 原文件已备份为 {bak}")
    if extra_note:
        print(extra_note)


def cmd_replace(args):
    doc = Document(args.file)
    total, detail = 0, {}
    for p, where in _iter_scope_paragraphs(doc, args.scope):
        n = replace_in_paragraph(p, args.old, args.new, regex=args.regex,
                                 max_count=args.count)
        if n:
            total += n
            detail[where] = detail.get(where, 0) + n
    scope_note = f"范围={args.scope}"
    if detail:
        print("[明细] " + " | ".join(f"{k}×{v}" for k, v in sorted(detail.items())))
    if args.dry_run:
        print(f"[预览] 共 {total} 处可替换（{scope_note}，未写入）")
        return 0
    if total == 0:
        print(f"[提示] 未找到匹配（{scope_note}），文件未修改")
        return 1
    _save_with_backup(doc, args.file, args.out, f"[统计] 共替换 {total} 处（{scope_note}）")
    return 0


# ---------- 修订记录（OOXML 层，python-docx 不支持） ----------

def _revision_roots(doc):
    """文档主体 + 各节页眉页脚的 XML 根（修订可能藏在任何 part）。"""
    roots = [doc.element]
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header,
                   sec.first_page_footer, sec.even_page_header, sec.even_page_footer):
            try:
                if not hf.is_linked_to_previous:
                    roots.append(hf._element)
            except Exception:
                pass
    return roots


def _unwrap(el):
    """把 el 的子元素提升到 el 的位置，删除 el 本身。"""
    parent = el.getparent()
    idx = parent.index(el)
    for child in list(el):
        parent.insert(idx, child)
        idx += 1
    parent.remove(el)


def _del_to_text(el):
    """w:del 内的 w:delText 转回普通 w:t（拒绝删除=恢复文字）。"""
    for dt in el.iter(qn("w:delText")):
        dt.tag = qn("w:t")


def _count_comments(doc):
    """统计批注数（word/comments.xml）。"""
    try:
        n = 0
        for part in doc.part.package.iter_parts():
            if str(part.partname).endswith("comments.xml"):
                import lxml.etree as etree
                root = etree.fromstring(part.blob)
                n += len(root.findall(qn("w:comment")))
        return n
    except Exception:
        return -1


def _text_preview(el, limit=30):
    txt = "".join(t.text or "" for t in el.iter() if t.tag in (qn("w:t"), qn("w:delText")))
    return (txt[:limit] + "…") if len(txt) > limit else txt


def cmd_revisions(args):
    doc = Document(args.file)
    roots = _revision_roots(doc)

    if args.action == "list":
        rows = []
        for root in roots:
            for el in root.iter():
                if el.tag == qn("w:ins"):
                    rows.append(("插入", el.get(qn("w:author")) or "?",
                                 el.get(qn("w:date")) or "", _text_preview(el)))
                elif el.tag == qn("w:del"):
                    rows.append(("删除", el.get(qn("w:author")) or "?",
                                 el.get(qn("w:date")) or "", _text_preview(el)))
        n_fmt = sum(len(root.findall(f".//{qn('w:rPrChange')}")) +
                    len(root.findall(f".//{qn('w:pPrChange')}")) for root in roots)
        n_cmt = _count_comments(doc)
        print("=" * 56)
        print(" 修订记录报告:", os.path.basename(args.file))
        print("=" * 56)
        if not rows:
            print("（无插入/删除修订）")
        for kind, author, date, preview in rows:
            print(f"  [{kind}] {author} {date[:10]}  “{preview}”")
        print(f"格式修订(rPrChange/pPrChange): {n_fmt} 处 | 批注: "
              f"{'?' if n_cmt < 0 else n_cmt} 条")
        print(f"共 {len(rows)} 处文字修订。接受全部: --action accept；拒绝: --action reject")
        return 0

    # accept / reject
    author = args.author
    def mine(el):
        return author is None or (el.get(qn("w:author")) or "") == author

    n_ins = n_del = n_fmt = 0
    if args.action == "accept":
        # 顺序关键：先接受删除（连同 w:ins 内嵌套的 w:del 一起清掉），再展开插入
        for root in roots:
            for el in list(root.iter(qn("w:del"))):
                if mine(el):
                    el.getparent().remove(el)
                    n_del += 1
        for root in roots:
            for el in list(root.iter(qn("w:ins"))):
                if mine(el):
                    _unwrap(el)
                    n_ins += 1
        # 接受格式修订 = 丢弃旧格式标记
        for root in roots:
            for tag in ("w:rPrChange", "w:pPrChange"):
                for el in list(root.iter(qn(tag))):
                    el.getparent().remove(el)
                    n_fmt += 1
        note = f"[统计] 接受修订：插入展开 {n_ins} | 删除移除 {n_del} | 格式修订 {n_fmt}"
    else:  # reject
        for root in roots:
            for el in list(root.iter(qn("w:ins"))):
                if mine(el):
                    el.getparent().remove(el)
                    n_ins += 1
        for root in roots:
            for el in list(root.iter(qn("w:del"))):
                if mine(el):
                    _del_to_text(el)
                    _unwrap(el)
                    n_del += 1
        # 拒绝格式修订 = 恢复 rPrChange/pPrChange 里存的旧属性
        for root in roots:
            for tag, oldtag in (("w:rPrChange", "w:rPr"), ("w:pPrChange", "w:pPr")):
                for el in list(root.iter(qn(tag))):
                    parent = el.getparent()
                    old = el.find(qn(oldtag))
                    for c in list(parent):
                        parent.remove(c)
                    if old is not None:
                        for c in list(old):
                            parent.append(c)
                    n_fmt += 1
        note = f"[统计] 拒绝修订：插入撤销 {n_ins} | 删除恢复 {n_del} | 格式回滚 {n_fmt}"
    _save_with_backup(doc, args.file, args.out, note)
    return 0


# ---------- 页眉页脚 ----------

def cmd_headers(args):
    doc = Document(args.file)
    if args.set_header is None and args.set_footer is None:
        print("=" * 56)
        print(" 页眉页脚报告:", os.path.basename(args.file))
        print("=" * 56)
        for i, sec in enumerate(doc.sections, 1):
            print(f"[第{i}节]")
            for label, hf in (("默认页眉", sec.header), ("默认页脚", sec.footer),
                              ("首页页眉", sec.first_page_header),
                              ("首页页脚", sec.first_page_footer),
                              ("奇偶页页眉", sec.even_page_header),
                              ("奇偶页页脚", sec.even_page_footer)):
                try:
                    linked = hf.is_linked_to_previous
                except Exception:
                    linked = False
                if linked:
                    print(f"  {label}: [继承上一节]")
                else:
                    txt = " ¦ ".join(p.text for p in hf.paragraphs if p.text.strip())
                    print(f"  {label}: {txt or '(空)'}")
        return 0

    align = ALIGN_MAP.get(args.align or "center", WD_ALIGN_PARAGRAPH.CENTER)
    for sec in doc.sections:
        if args.set_header is not None:
            _set_hf_text(sec.header, args.set_header, align)
        if args.set_footer is not None:
            _set_hf_text(sec.footer, args.set_footer, align)
    _save_with_backup(doc, args.file, args.out,
                      f"[统计] 已设置各节页眉/页脚（对齐={args.align or 'center'}）")
    return 0


def _set_hf_text(hf, text, align):
    """设置页眉/页脚文本：清空原 run，写入新 run（继承该节现有段落格式）。"""
    try:
        if hf.is_linked_to_previous:
            hf.is_linked_to_previous = False  # 创建独立定义
    except Exception:
        pass
    paras = hf.paragraphs
    p = paras[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.add_run(text)
    p.alignment = align


def _add_runs(p, text, fmt, size_pt):
    """把 [n] 渲染为上标，其余按正文格式。"""
    pos = 0
    for m in SUP_RE.finditer(text):
        if m.start() > pos:
            r = p.add_run(text[pos:m.start()])
            set_run_font(r, fmt.get("font_cn"), fmt.get("font_en", FONT_EN_DEFAULT), size_pt,
                         fmt.get("bold") if "bold" in fmt else None)
        r = p.add_run(m.group(0))
        set_run_font(r, fmt.get("font_cn"), fmt.get("font_en", FONT_EN_DEFAULT), size_pt)
        r.font.superscript = True
        pos = m.end()
    if pos < len(text):
        r = p.add_run(text[pos:])
        set_run_font(r, fmt.get("font_cn"), fmt.get("font_en", FONT_EN_DEFAULT), size_pt,
                     fmt.get("bold") if "bold" in fmt else None)


def cmd_md2docx(args):
    preset = apply_overrides(load_preset(args.preset), args.set)
    body = preset.get("body", {})
    heads = preset.get("headings", {})
    tbl_fmt = preset.get("table", {})
    body_size = to_pt(body.get("size"))
    doc = Document()
    set_page(doc, preset.get("page"))
    style_definitions(doc, preset)

    with open(args.file, encoding="utf-8") as f:
        lines = f.read().splitlines()

    i, table_buf = 0, []
    while i < len(lines):
        line = lines[i].rstrip()
        if line.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s:\-|]+\|$", lines[i + 1].strip()):
            table_buf = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if not re.match(r"^\|[\s:\-|]+\|$", lines[i].strip()):
                    table_buf.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            if table_buf:
                t = doc.add_table(rows=len(table_buf), cols=len(table_buf[0]))
                t.style = "Table Grid"
                for ri, row in enumerate(table_buf):
                    for ci, val in enumerate(row):
                        cell = t.cell(ri, ci)
                        cell.text = ""
                        p = cell.paragraphs[0]
                        _add_runs(p, val, tbl_fmt, to_pt(tbl_fmt.get("size")) or body_size)
                        set_para_format(p, tbl_fmt, to_pt(tbl_fmt.get("size")) or body_size)
            continue
        if not line.strip():
            i += 1
            continue
        mimg = re.match(r"^!\[(.*?)\]\((.*?)\)\s*$", line)
        if mimg:
            img_path = mimg.group(2)
            if os.path.isfile(img_path):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(2)
                p.add_run().add_picture(img_path, width=Cm(14.5))
                cap = doc.add_paragraph()
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cr = cap.add_run(mimg.group(1))
                set_run_font(cr, body.get("font_cn"), body.get("font_en", FONT_EN_DEFAULT),
                             to_pt(tbl_fmt.get("size")) or 10.5)
                cap.paragraph_format.space_after = Pt(8)
                i += 1
                continue
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            lvl = min(len(m.group(1)), 4)
            fmt = heads.get(str(lvl), body)
            size_pt = to_pt(fmt.get("size")) or body_size
            p = doc.add_paragraph()
            try:
                p.style = doc.styles[f"Heading {lvl}"]
            except KeyError:
                pass
            r = p.add_run(m.group(2).strip())
            set_run_font(r, fmt.get("font_cn"), fmt.get("font_en", FONT_EN_DEFAULT), size_pt,
                         fmt.get("bold") if "bold" in fmt else None)
            set_para_format(p, fmt, size_pt)
            i += 1
            continue
        plain = re.sub(r"\*\*(.+?)\*\*", r"\1", line)
        # 表题/图题：以 表x-x / 图x-x 开头 → 居中加粗
        if re.match(r"^[表图]\d+[-.．]", plain):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(plain)
            set_run_font(r, body.get("font_cn"), body.get("font_en", FONT_EN_DEFAULT),
                         body_size, True)
            set_para_format(p, {"align": "center", "space_before_pt": 4,
                                "space_after_pt": 4}, body_size)
            i += 1
            continue
        p = doc.add_paragraph()
        _add_runs(p, plain, body, body_size)
        set_para_format(p, body, body_size)
        i += 1

    out = args.out or os.path.splitext(args.file)[0] + ".docx"
    doc.save(out)
    print(f"[完成] Markdown -> {out}")
    return 0


def main():
    ap = argparse.ArgumentParser(description="Word 批量格式引擎")
    sub = ap.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("apply", help="按预设批量应用格式")
    a.add_argument("file")
    a.add_argument("--preset", default="cn-formal")
    a.add_argument("--set", action="append", default=[])
    a.add_argument("--out")
    a.add_argument("--scope", default="all",
                   choices=["all", "body", "headings", "tables", "style", "page"])
    a.add_argument("--force", dest="force", action="store_true", default=True)
    a.add_argument("--no-force", dest="force", action="store_false")
    a.add_argument("--report", action="store_true")
    a.set_defaults(func=cmd_apply)

    b = sub.add_parser("inspect", help="格式体检")
    b.add_argument("file")
    b.set_defaults(func=cmd_inspect)

    c = sub.add_parser("md2docx", help="Markdown 转 docx")
    c.add_argument("file")
    c.add_argument("--preset", default="cn-formal")
    c.add_argument("--set", action="append", default=[])
    c.add_argument("--out")
    c.set_defaults(func=cmd_md2docx)

    d = sub.add_parser("replace", help="保格式查找替换（run 级最小跨度编辑）")
    d.add_argument("file")
    d.add_argument("old")
    d.add_argument("new")
    d.add_argument("--scope", default="all",
                   choices=["all", "body", "tables", "headers", "footers"])
    d.add_argument("--regex", action="store_true")
    d.add_argument("--count", type=int, default=0, help="每段最多替换 N 处，0=不限")
    d.add_argument("--dry-run", action="store_true")
    d.add_argument("--out")
    d.set_defaults(func=cmd_replace)

    e = sub.add_parser("revisions", help="修订记录：列出/接受/拒绝")
    e.add_argument("file")
    e.add_argument("--action", default="list", choices=["list", "accept", "reject"])
    e.add_argument("--author", help="只处理该作者的修订")
    e.add_argument("--out")
    e.set_defaults(func=cmd_revisions)

    f = sub.add_parser("headers", help="页眉页脚：查看/设置")
    f.add_argument("file")
    f.add_argument("--set-header", dest="set_header")
    f.add_argument("--set-footer", dest="set_footer")
    f.add_argument("--align", choices=["left", "center", "right", "both"])
    f.add_argument("--out")
    f.set_defaults(func=cmd_headers)

    args = ap.parse_args()
    sys.exit(args.func(args) or 0)


if __name__ == "__main__":
    main()
