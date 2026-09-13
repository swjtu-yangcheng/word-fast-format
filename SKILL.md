---
name: word-fast-format
description: "Word(.docx) 高速批量排版与生成引擎。当用户要求修改 Word 格式（调字号、调行距、换字体、改页边距、统一标题样式、批量重排整篇文档）、或对已有 docx 做格式体检/问题诊断、或从 Markdown/纯文本快速生成符合中文规范的 docx 时使用。2.0 新增：保格式查找替换（跨 run 最小跨度编辑）、修订记录处理（列出/接受/拒绝 track changes，可按作者过滤）、页眉页脚读写。替代逐段 MCP 调用，一次命令秒级完成全文档排版。触发词：改 Word 格式、统一字号、行距、字体、页边距、排版、格式体检、markdown 转 word、批量替换、接受修订、拒绝修订、页眉页脚、生成申报书/报告/公文 docx。"
agent_created: true
version: "2.0.0"
---

# Word 高速排版引擎（word-fast-format）

## 为什么用这个而不是编辑器 MCP

| 方式 | 整篇改格式的耗时 | 准确度 |
|---|---|---|
| 编辑器 MCP（逐段 doc_modify_paragraph / doc_update_text_property） | 数十次调用、分钟级 | 易漏：直设格式覆盖、表格内段落、空段落 |
| **本引擎（python-docx 一次性遍历）** | **约 1 秒** | 样式定义 + docDefaults + 每个 Run 三层同时改，无遗漏 |

实测（64 段申报书）：inspect 体检 0.76s，apply 全篇排版 0.93s（523 个文字片段 + 14 标题 + 12 表格段），md2docx 出稿 1.07s。

## 环境

专用 venv（已装 python-docx）：`<你的 python>`
脚本：`scripts/wordfmt.py`

## 标准工作流

```bash
PY="<你的 python>"
WF="scripts/wordfmt.py"

# 1) 先体检：看字号/行距/字体/页边距分布，判断问题在哪
"$PY" "$WF" inspect 文件.docx

# 2) 套预设排版（默认原地覆盖并自动备份 .bak-HHMMSS.docx）
"$PY" "$WF" apply 文件.docx --preset cn-formal --report

# 3) 不改原文件、输出到副本
"$PY" "$WF" apply 文件.docx --preset cn-formal --out 新文件.docx
```

## 预设

| 预设 | 用途 | 规格 |
|---|---|---|
| `cn-formal`（默认） | 中文正式文档（申报书/论文/报告） | A4；页边距 左3.18/右3.18/上2.54/下2.54 cm；正文 仿宋_GB2312 小四 12pt、1.5倍行距、首行缩进2字符、两端对齐；H1/H2 黑体四号、H3 仿宋四号、H4 仿宋小四；表格 仿宋小四 1.5倍 |
| `official` | 党政机关公文 GB/T 9704-2012 | 三号仿宋_GB2312、固定行距28磅、页边距 上3.7/下3.5/左2.8/右2.6 cm；一级标题黑体、二级楷体_GB2312 |
| `academic` | 期刊论文 | 宋体五号(10.5pt)、1.5倍、黑体标题 |

自定义预设：`presets/名称.json`，结构见下；也可用 `--set` 临时覆盖任意字段。

```bash
--set body.size=12 --set body.line_spacing_multiple=1.5 --set page.margin_left_cm=3.18
--set body.size=三号          # 支持中文号数：三号/小四/四号/五号…
--set headings.2.size=14
```

预设结构（JSON）：
```json
{
  "page": {"size":"A4","margin_left_cm":3.18,"margin_right_cm":3.18,"margin_top_cm":2.54,"margin_bottom_cm":2.54},
  "body": {"font_cn":"仿宋_GB2312","font_en":"Times New Roman","size":12,
           "line_spacing_multiple":1.5,"first_line_indent_chars":2,"align":"both"},
  "headings": {"1":{...},"2":{...},"3":{...}},
  "table": {"font_cn":"仿宋_GB2312","size":12,"line_spacing_multiple":1.5}
}
```
`line_spacing_multiple`（倍数，如 1.5）与 `line_spacing_exact_pt`（固定磅值，如 28）二选一，前者优先。

## Markdown → docx（写 Word 的快路径）

```bash
"$PY" "$WF" md2docx 稿件.md --preset cn-formal --out 成稿.docx
```

- `#/##/###/####` → Heading 1-4（Word 大纲层级，可直接生成目录）
- 正文中的 `[1]`、`[2-3]` 自动渲染为**上标引注**
- Markdown 表格 → 带框线表格，套用表格样式
- 套用预设的字体/字号/行距/页边距，一次成型

**何时不用它**：需要复杂图文混排、脚注、图表题注时，仍走完整文档流水线；生成后可用本引擎 apply 统一格式。

## 保格式查找替换（2.0 新增：replace）

改现有文档**只动受影响的文字**，不重写整段——防"改一个字、整段格式漂移"：

```bash
# 先预览再写入（默认搜正文+表格+页眉页脚）
"$PY" "$WF" replace 合同.docx "北京市朝阳区XX路1号" "北京市海淀区YY大街2号" --dry-run
"$PY" "$WF" replace 合同.docx "北京市朝阳区XX路1号" "北京市海淀区YY大街2号"

# 正则 + 反向引用
"$PY" "$WF" replace 报告.docx "(\d{4})年" "公元\1年" --regex

# 只搜表格 / 页眉页脚
"$PY" "$WF" replace 报告.docx 旧词 新词 --scope tables
```

要点：文字在 XML 里可能**跨 run 拆分**，本引擎先重建全段文本再映射回 run；替换文字并入 span 起点 run（继承其格式），其余文字零扰动。原地写入自动备份。

## 修订记录处理（2.0 新增：revisions，类 Git）

多人修订过的合同/论文，一键出干净版或回退（python-docx 不支持修订，本引擎直接操作 OOXML 层 `w:ins`/`w:del`/`w:rPrChange`）：

```bash
"$PY" "$WF" revisions 合同.docx                        # 列出：谁在何时插入/删除了什么
"$PY" "$WF" revisions 合同.docx --action accept        # 接受全部 -> 干净终版
"$PY" "$WF" revisions 合同.docx --action reject        # 拒绝全部 -> 回到修订前
"$PY" "$WF" revisions 合同.docx --action accept --author 李四   # 只接受李四的修订
```

- **accept**：展开插入、移除删除、丢弃格式修订标记
- **reject**：撤销插入、把 `w:delText` 恢复为正常文字、从 `rPrChange` 回滚格式
- 注意：python-docx 的 `paragraph.text` **不读取** `w:ins` 内的 run，验证修订效果须 XML 层取全文

## 页眉页脚读写（2.0 新增：headers）

```bash
"$PY" "$WF" headers 文件.docx                                  # 查看：各节默认/首页/奇偶页页眉页脚
"$PY" "$WF" headers 文件.docx --set-header "XX大学教务处"       # 设置各节默认页眉
"$PY" "$WF" headers 文件.docx --set-footer "第 1 页" --align center
```

显示 `[继承上一节]` 表示该页眉页脚无独立定义；设置时会自动创建独立定义。

## 关键实现要点（避免踩坑）

1. **三层同时改**：只改样式定义无效——Word 段落/Run 上的**直设格式**会覆盖样式。本引擎同时改 `styles[]` 定义 + 每个段落属性 + 每个 Run 的 rPr（含 `w:eastAsia` 中文字体），`--no-force` 可关闭 Run 层改写。
2. **中文字体必须写 eastAsia**：只设 `run.font.name` 中文不生效，须 `rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')`。
3. **首行缩进**：python-docx 无"字符"单位，用 `Pt(字号 × 字符数)`（小四2字符 = 24pt）。
4. **行距**：倍数用 `line_spacing=1.5` + `WD_LINE_SPACING.MULTIPLE`；固定值用 `Pt(30)` + `EXACTLY`。混用会残留旧值。
5. **表格内段落**：`doc.paragraphs` 不包含表格内文字，必须遍历 `table.rows→cells→paragraphs`。
6. **改前备份**：apply 原地模式自动生成 `.bak-HHMMSS.docx`。

## 相关资源

- 已装第三方 skill `document-format-skills`（中文公文排版工具包，228★，MIT）：
  `"<你的 python>" （第三方 skill）process.py analyze 文件.docx`
  擅长标点/空格清理与公文 GB/T 9704 预设。
  **注意**：其 `punctuation` 会把英文半角标点转全角，含 DOI、英文参考文献的学术文档**不要**跑标点清理，只用 `analyze` 诊断。
- 文档打开预览仍用 editor_sdk（`present_files` + `tencent-local-office-edit`），但**批量格式改动一律先走本引擎**，不要在编辑器里逐段改。
