# word-fast-format

> Word（.docx）**高速批量排版与生成引擎**。
> 改字号、调行距、换字体、统一标题、改页边距、整篇重排——一次命令，约 1 秒完成。
> 也支持从 Markdown 直接生成符合中文规范的 docx。
>
> **v2.0 新增**：保格式查找替换（跨 run 最小跨度编辑）、修订记录处理（类 Git 的
> 接受/拒绝 track changes）、页眉页脚读写。

---

## 为什么不用编辑器逐个改

AI 改 Word 格式时，常见做法是调用文档编辑接口**逐段修改**。这在长文档上是灾难：

| 方式 | 整篇改格式的耗时 | 准确度 |
|---|---|---|
| 编辑器接口逐段改（`doc_modify_paragraph` / `doc_update_text_property`） | 数十次调用、分钟级 | **易漏**：直设格式、表格内段落、空段落都改不到 |
| **本引擎（python-docx 一次性遍历）** | **约 1 秒** | 样式定义 + 文档默认值 + 每个 Run 三层同改，无遗漏 |

实测（64 段文档）：

```
inspect 格式体检   0.76s
apply   全篇排版   0.93s   （523 个文字片段 + 14 个标题 + 12 个表格段）
md2docx 出稿       1.07s
```

> 曾有过一次教训：用逐段方式改完，体检发现**仍有 10 段是固定 30 磅行距、12 段是单倍行距**——
> 压根没改到。这就是为什么"改完必须体检"。

---

## 快速开始

```bash
pip install -r requirements.txt
python examples/run_demo.py
```

示例会演示完整流程：Markdown 出稿 → 体检 → 一键换排版 → 复检。

---

## 三条命令

```bash
# 1. 体检：看清字号/行距/字体/页边距的实际分布，问题在哪一目了然
python scripts/wordfmt.py inspect 文档.docx

# 2. 套预设排版（默认原地覆盖，自动备份 .bak-HHMMSS.docx）
python scripts/wordfmt.py apply 文档.docx --preset cn-formal --report

# 3. 不改原文件，输出到副本
python scripts/wordfmt.py apply 文档.docx --preset cn-formal --out 新文档.docx
```

---

## 从 Markdown 生成 Word（写文档的快路径）

```bash
python scripts/wordfmt.py md2docx 稿件.md --preset cn-formal --out 成稿.docx
```

自动处理：

- `#` `##` `###` `####` → Heading 1–4（Word 大纲层级，可直接生成目录）
- 正文里的 `[1]`、`[2-3]` → **上标引注**
- Markdown 表格 → 带框线表格
- 字体/字号/行距/页边距一次成型

**何时不用它**：需要复杂图文混排、脚注、图表题注时，生成后可用本引擎 `apply` 统一格式。

---

## v2.0 新增三命令

### replace —— 保格式查找替换

改现有文档**只动受影响的文字**，不重写整段，防止"改一个字、整段格式漂移"：

```bash
# 先预览再写入（默认搜正文 + 表格 + 页眉页脚）
python scripts/wordfmt.py replace 合同.docx "北京市朝阳区XX路1号" "北京市海淀区YY大街2号" --dry-run
python scripts/wordfmt.py replace 合同.docx "北京市朝阳区XX路1号" "北京市海淀区YY大街2号"

# 正则 + 反向引用
python scripts/wordfmt.py replace 报告.docx "(\d{4})年" "公元\1年" --regex

# 限定范围
python scripts/wordfmt.py replace 报告.docx 旧词 新词 --scope tables   # 只搜表格
```

文字在 docx 的 XML 里可能**跨 run 拆分**（一个词存在多个节点里），本引擎先重建全段
文本再映射回 run；替换文字并入起点 run（继承其格式），其余文字零扰动。
实测：目标串拆在 3 个不同格式的 run 里，替换后其余文字格式不变。

### revisions —— 修订记录处理（类 Git）

多人修订过的合同/论文，一键出干净版或回退。python-docx 不支持修订，
本引擎直接操作 OOXML 层的 `w:ins` / `w:del` / `w:rPrChange`：

```bash
python scripts/wordfmt.py revisions 合同.docx                        # 列出：谁在何时插入/删除了什么
python scripts/wordfmt.py revisions 合同.docx --action accept        # 接受全部 -> 干净终版
python scripts/wordfmt.py revisions 合同.docx --action reject        # 拒绝全部 -> 回到修订前
python scripts/wordfmt.py revisions 合同.docx --action accept --author 李四   # 只接受李四的
```

- **accept**：展开插入、移除删除、丢弃格式修订标记
- **reject**：撤销插入、把 `w:delText` 恢复为正常文字、从 `rPrChange` 回滚旧格式
- 处理顺序有讲究：accept 先删后展开、reject 先撤销插入后恢复删除（正确处理嵌套修订）

### headers —— 页眉页脚读写

```bash
python scripts/wordfmt.py headers 文档.docx                                  # 查看（各节默认/首页/奇偶页）
python scripts/wordfmt.py headers 文档.docx --set-header "XX大学教务处"       # 设置页眉
python scripts/wordfmt.py headers 文档.docx --set-footer "第 1 页" --align center
```

---

## 预设

| 预设 | 用途 | 规格 |
|---|---|---|
| `cn-formal`（默认） | 中文正式文档（申报书/论文/报告） | A4；页边距 左3.18/右3.18/上2.54/下2.54 cm；正文 仿宋_GB2312 小四 12pt、1.5 倍行距、首行缩进 2 字符、两端对齐；H1/H2 黑体四号，H3 仿宋四号，H4 仿宋小四；表格 仿宋小四 1.5 倍 |
| `official` | 党政机关公文 GB/T 9704-2012 | 三号仿宋_GB2312、固定行距 28 磅、页边距 上3.7/下3.5/左2.8/右2.6 cm；一级标题黑体、二级楷体_GB2312 |
| `academic` | 期刊论文 | 宋体五号（10.5pt）、1.5 倍、黑体标题 |

自定义：在 `presets/` 下放同名 JSON 即可；也可用 `--set` 临时覆盖任意字段。

```bash
--set body.size=12 --set body.line_spacing_multiple=1.5 --set page.margin_left_cm=3.18
--set body.size=三号            # 支持中文号数：三号/小四/四号/五号…
--set headings.2.size=14
```

预设结构：

```json
{
  "page": {"size":"A4","margin_left_cm":3.18,"margin_right_cm":3.18,
           "margin_top_cm":2.54,"margin_bottom_cm":2.54},
  "body": {"font_cn":"仿宋_GB2312","font_en":"Times New Roman","size":12,
           "line_spacing_multiple":1.5,"first_line_indent_chars":2,"align":"both"},
  "headings": {"1":{...},"2":{...},"3":{...}},
  "table": {"font_cn":"仿宋_GB2312","size":12,"line_spacing_multiple":1.5}
}
```

`line_spacing_multiple`（倍数，如 1.5）与 `line_spacing_exact_pt`（固定磅值，如 28）二选一，前者优先。

---

## 六个实现要点（都是踩过的坑）

1. **必须三层同时改** —— 只改样式定义无效，Word 段落/Run 上的**直设格式会覆盖样式**。
   本引擎同时改 `styles[]` 定义 + 每个段落属性 + 每个 Run 的 `rPr`（含 `w:eastAsia`）。
   加 `--no-force` 可关闭 Run 层改写。
2. **中文字体要写 `eastAsia`** —— 只设 `run.font.name` 中文不生效，
   必须 `rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')`。
3. **首行缩进没有"字符"单位** —— python-docx 只有磅值，用 `Pt(字号 × 字符数)`（小四 2 字符 = 24pt）。
4. **行距别混用** —— 倍数用 `line_spacing=1.5` + `MULTIPLE`；固定值用 `Pt(30)` + `EXACTLY`。混用会残留旧值。
5. **表格内段落是隐藏的** —— `doc.paragraphs` **不包含**表格里的文字，必须遍历 `table.rows → cells → paragraphs`。
   这是逐段改法最容易漏的一块。
6. **改前自动备份** —— `apply` 原地模式会生成 `.bak-HHMMSS.docx`。

---

## 目录结构

```
word-fast-format/
├── SKILL.md                  # Agent Skill 定义（供 AI 助手加载）
├── README.md
├── requirements.txt
├── LICENSE                   # MIT
├── scripts/wordfmt.py        # 排版引擎（inspect / apply / md2docx / replace / revisions / headers）
├── tests/test_v2.py          # 24 项自动化测试（覆盖全部 6 个命令）
├── presets/cn-formal.json    # 中文正式文档预设
└── examples/
    ├── sample.md             # 示例稿件
    └── run_demo.py           # 一键演示
```

跑测试：`python tests/test_v2.py`（24 项断言，含跨 run 替换、修订 accept/reject、页眉页脚读写）。

---

## 环境要求

- Python 3.9+
- `python-docx`

---

## 作为 AI Agent Skill 使用

本仓库同时是 Agent Skills 格式的能力包，复制到 skills 目录即可被 AI 助手自动加载：

```bash
cp -r word-fast-format ~/.workbuddy/skills/
```

`SKILL.md` 的 `description` 声明了触发时机（"改 Word 格式""统一字号行距""markdown 转 word"等）。

---

## 许可

MIT License — 见 [LICENSE](LICENSE)。
