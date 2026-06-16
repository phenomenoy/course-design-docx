# Skill: course-design-docx
> 面向所有 AI 代理的中文课程设计 Word 文档生成标准规范。基于 python-docx 库，定义从封面到附录的完整样式体系、排版规则、图片/表格处理标准和自动目录生成流程。适用于课程设计报告、商业计划书、实习报告、实验报告等中文正规文档。
> 使用场景：当用户要求「写一份课程设计文档 / 商业计划书 / 实验报告 / 实习报告」并输出 .docx 格式时调用。

---

## 1. 工具栈与前置条件

### 必需库
```python
from docx import Document
from docx.shared import Pt, Inches, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import re, copy
```

### 版本要求
- `python-docx >= 0.8.11`
- `matplotlib >= 3.5`（用于自动生成图表）
- `numpy >= 1.21`（图表数据处理）
- Python >= 3.8

### 环境约束
- 所有中文内容使用 **宋体 (SimSun)** 为正文字体，**黑体 (SimHei)** 为标题字体
- 英文和数字使用 **Times New Roman**（正文）/ **Arial**（表格）
- 字体大小单位为 **pt**（磅），python-docx 中通过 `Pt(n)` 设置
- 行距值：Word 内部行距单位为"1/20 磅"（twentieths-of-a-point），单倍=240（对应12pt字号的1倍行高），1.5倍=360，2倍=480。也可以通过 `paragraph_format.line_spacing = 1.5` 直接用倍数设置
- 首行缩进值：约2字符 = 480（1/20磅单位），通过 `paragraph_format.first_line_indent = Pt(24)` 或 `Cm(0.85)` 设置

---

## 2. 文档整体架构

课程设计文档必须是 **三段式分节结构**：

```
┌─────────────────────────────────┐
│  第一节：封面（无页眉页脚页码）      │  ← 分节符（下一页）
├─────────────────────────────────┤
│  第二节：目录（无页眉页脚页码）      │  ← 分节符（下一页）
├─────────────────────────────────┤
│  第三节：正文（含页眉页脚，页码从1开始）│
│    ├─ 摘要                       │
│    ├─ 第X章（一级标题：1, 2, 3…）   │
│    ├─ X.Y节（二级标题）           │
│    ├─ 图片与表格                  │
│    ├─ 参考文献                    │
│    └─ 附录                       │
└─────────────────────────────────┘
```

> **关键规则**：目录的生成必须发生在**正文全部写完、保存文档之前**的最后一个环节。绝对不能在正文之前或正文写到一半时生成目录。

### 2.1 文档类型识别

AI 代理在生成文档前，必须根据用户输入判断文档类型，然后加载对应的结构模板和图表配置。不同类型的文档有完全不同的章节结构和图表要求。

| 文档类型 | 识别关键词 | 必备章节 | 图数量 |
|---------|-----------|---------|--------|
| **商业计划书** | 商业计划、创业、商业模式、市场分析、财务预测 | 市场分析、商业模式、管理团队、财务分析、风险分析 | ≥8 |
| **实验报告** | 实验、实验原理、实验步骤、实验数据、结果分析 | 实验目的、实验原理、实验步骤、实验结果、讨论 | ≥8 |
| **软件课程设计** | 软件、系统、架构、数据库、UML、需求分析 | 需求分析、系统设计、数据库设计、实现、测试 | ≥8 |
| **实习报告** | 实习、实践、工作岗位、实习内容、实习总结 | 实习概况、工作内容、技能分析、总结反思 | ≥8 |
| **毕业论文/设计** | 论文、研究、文献综述、方法、实验 | 文献综述、研究方法、实验/实现、结果、结论 | ≥8 |

> **文档类型决定图表规则**：不同文档类型使用 §8.5 中对应的配图规则表。AI 代理不得混用或自行编造图类型。

---

## 3. 页面设置规范

### 3.1 整体设置
```python
doc = Document()

# 获取默认节并设置
section = doc.sections[0]
section.page_width  = Cm(21.0)   # A4
section.page_height = Cm(29.7)
section.top_margin    = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)
```

| 参数 | 推荐值 | 说明 |
|------|--------|------|
| 纸张 | A4 (21.0×29.7 cm) | 中文规范文档标准 |
| 上边距 | 2.54 cm (1 in) | 标准装订留白 |
| 下边距 | 2.54 cm (1 in) | 标准装订留白 |
| 左边距 | 2.5 cm | 国标 GB/T 7713.1-2006 订口≥25mm |
| 右边距 | 2.5 cm | 国标 GB/T 7713.1-2006 切口≥20mm，取推荐值 |

### 3.2 分节设置
正文节需要**独立控制页码**：

```python
# 在封面和目录之后添加新节（正文节）
new_section = doc.add_section()
new_section.page_width  = Cm(21.0)
new_section.page_height = Cm(29.7)
new_section.top_margin    = Cm(2.54)
new_section.bottom_margin = Cm(2.54)
new_section.left_margin   = Cm(2.5)
new_section.right_margin  = Cm(2.5)

# 正文节：页码从1开始
new_section.start_type = 2  # WD_SECTION_START.NEW_PAGE
sectPr = new_section._sectPr
pgNumType = parse_xml('<w:pgNumType {} w:fmt="decimal" w:start="1"/>'.format(nsdecls('w')))
sectPr.append(pgNumType)
```

---

## 4. 完整样式系统

### 4.1 必须自定义的样式

**绝对禁止**所有段落使用 `Normal` 样式。必须通过 `doc.styles.add_style()` 创建以下命名样式：

| 样式名称 | 类型 | 用途 | 中文 | 英文 | 字号 | 加粗 | 对齐 | 行距 | 段前/段后 | 首行缩进 |
|----------|------|------|------|------|------|------|------|------|-----------|---------|
| `CoverTitle` | 段落 | 封面大标题 | 黑体 | Arial | 22pt | 是 | 居中 | 1.5倍 | 0/0 | 无 |
| `CoverInfo` | 段落 | 封面信息行 | 黑体 | Arial | 14pt | 是 | 居中 | 1.5倍 | 0/0 | 无 |
| `Heading1_Course` | 段落 | 一级标题（1, 2, 3…） | 黑体 | Arial | 18pt | 是 | 左对齐（顶格） | 1.5倍 | 12pt/6pt | 无 |
| `Heading2_Course` | 段落 | 二级标题（1.1, 2.1…） | 黑体 | Arial | 15pt | 是 | 左对齐 | 1.5倍 | 12pt/6pt | 无 |
| `Heading3_Course` | 段落 | 三级标题（1.1.1…） | 黑体 | Arial | 13pt | 是 | 左对齐 | 1.5倍 | 6pt/3pt | 无 |
| `BodyText_Course` | 段落 | 正文段落 | 宋体 | Times New Roman | 12pt | 否 | 两端对齐 | 1.5倍 | 0pt/6pt | 2字符 |
| `Caption_Course` | 段落 | 图注/表注 | 宋体 | Times New Roman | 10.5pt | 否 | 居中 | 1.25倍 | 3pt/6pt | 无 |
| `TableHeader` | 段落 | 表格表头 | 黑体 | Arial | 10pt | 是 | 居中 | 单倍 | 3pt/3pt | 无 |
| `TableData` | 段落 | 表格数据行 | 宋体 | Times New Roman | 10pt | 否 | 居中 | 单倍 | 3pt/3pt | 无 |
| `TOCHeading` | 段落 | 目录标题 | 黑体 | Arial | 18pt | 是 | 居中 | 1.5倍 | 12pt/12pt | 无 |
| `TOCEntry` | 段落 | 目录条目 | 宋体 | Times New Roman | 12pt | 否 | 左对齐 | 1.5倍 | 0/0 | 无 |
| `RefEntry` | 段落 | 参考文献条目 | 宋体 | Times New Roman | 10.5pt | 否 | 两端对齐 | 单倍 | 0/3pt | 无 |
| `AbstractTitle` | 段落 | 摘要标题 | 黑体 | Arial | 18pt | 是 | 居中 | 1.5倍 | 12pt/6pt | 无 |
| `AbstractBody` | 段落 | 摘要正文 | 宋体 | Times New Roman | 12pt | 否 | 两端对齐 | 1.5倍 | 0/6pt | 无 |

### 4.2 样式定义代码模板

```python
def setup_styles(doc):
    """创建课程设计文档所需的所有自定义样式"""
    
    # ---------- 封面样式 ----------
    style = doc.styles.add_style('CoverTitle', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Arial'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    style.font.size = Pt(22)
    style.font.bold = True
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.paragraph_format.line_spacing = 1.5
    
    style = doc.styles.add_style('CoverInfo', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Arial'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    style.font.size = Pt(14)
    style.font.bold = True
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.paragraph_format.line_spacing = 1.5
    
    # ---------- 标题样式 ----------
    # 注意：一级标题使用阿拉伯数字编号（国标 GB/T 7713.2-2022 §5.2.2）
    # 例如 "1 项目背景"、"2 市场分析"（编号后空1汉字间隙，无点号）
    # 同时设置 outlineLvl 以确保 Word 目录可识别自定义标题
    style = doc.styles.add_style('Heading1_Course', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Arial'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    style.font.size = Pt(18)
    style.font.bold = True
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT  # 顶格（国标 GB/T 7713.2-2022 §5.2.2）
    style.paragraph_format.space_before = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.5
    # 设置大纲级别为 1（章），Word TOC 据此识别
    pPr = style.element.get_or_add_pPr()
    outlineLvl = parse_xml('<w:outlineLvl {} w:val="0"/>'.format(nsdecls('w')))
    pPr.append(outlineLvl)
    
    style = doc.styles.add_style('Heading2_Course', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Arial'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    style.font.size = Pt(15)
    style.font.bold = True
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    style.paragraph_format.space_before = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.5
    # 设置大纲级别为 2（节）
    pPr = style.element.get_or_add_pPr()
    outlineLvl = parse_xml('<w:outlineLvl {} w:val="1"/>'.format(nsdecls('w')))
    pPr.append(outlineLvl)
    
    # ---------- 正文样式 ----------
    style = doc.styles.add_style('BodyText_Course', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(12)
    style.font.bold = False
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style.paragraph_format.first_line_indent = Pt(24)  # 约2字符
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.5
    
    # ---------- 图注/表注样式 ----------
    style = doc.styles.add_style('Caption_Course', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(10.5)
    style.font.bold = False
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.paragraph_format.space_before = Pt(3)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.line_spacing = 1.25
    
    # ---------- 目录标题 ----------
    style = doc.styles.add_style('TOCHeading', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Arial'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    style.font.size = Pt(18)
    style.font.bold = True
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.paragraph_format.space_before = Pt(12)
    style.paragraph_format.space_after = Pt(12)
    style.paragraph_format.line_spacing = 1.5
```

### 4.3 应用样式的方法

```python
def add_heading1(doc, text):
    """添加一级标题，编号格式为 "1 项目背景"（阿拉伯数字，编号后空1汉字间隙）"""
    p = doc.add_paragraph(style='Heading1_Course')
    p.add_run(text)
    return p

def add_heading2(doc, text):
    """添加二级标题"""
    p = doc.add_paragraph(style='Heading2_Course')
    p.add_run(text)
    return p

def add_body(doc, text):
    """添加正文段落"""
    p = doc.add_paragraph(style='BodyText_Course')
    p.add_run(text)
    return p

def add_caption(doc, text):
    """添加图注/表注"""
    p = doc.add_paragraph(style='Caption_Course')
    p.add_run(text)
    return p
```

> **注意**：python-docx 中通过 `add_run()` 添加文本时，如果该段已有样式，`run` 上不要再重复设置字体/字号——除非需要局部格式变化。样式系统的意义在于：样式统一定义格式，段落只需指定样式名称。

---

## 5. 封面页设计

### 5.1 封面结构

封面由以下元素组成（从上到下）：
1. **校名 / 课程名称**（可选，大字号突出显示）
2. **文档大标题**（如"课程设计报告"、"商业计划书"等）
3. **分隔线**（使用段落底边框，**绝对禁止**使用 `━━━━` 或 `———` 字符拼凑）
4. **项目信息**（使用无边框表格对齐，**绝对禁止**使用空格对齐）
5. **日期**

### 5.2 分隔线（段落底边框）

```python
def add_separator_line(doc):
    """使用段落底边框作为分隔线，替代字符拼凑"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        '<w:pBdr {}>'
        '  <w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/>'
        '</w:pBdr>'.format(nsdecls('w'))
    )
    pPr.append(pBdr)
    return p
```

### 5.3 信息行（无边框表格）

```python
def add_cover_info_table(doc, items):
    """
    使用无边框表格对齐封面信息。
    items: [(label, value), ...] 例如
        [("项 目 名 称", "时尚两元精品店"),
         ("创 作 小 组", "第6小组"), ...]
    返回表格对象。
    """
    table = doc.add_table(rows=len(items), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 去掉所有边框
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        '<w:tblBorders {}>'
        '  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'.format(nsdecls('w'))
    )
    tblPr.append(tblBorders)
    
    for i, (label, value) in enumerate(items):
        # 标签列：右对齐
        cell_label = table.rows[i].cells[0]
        cell_label.width = Inches(2.0)
        p = cell_label.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(label)
        run.font.name = 'Arial'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        run.font.size = Pt(14)
        run.font.bold = True
        p.paragraph_format.line_spacing = 1.5
        
        # 值列：左对齐
        cell_value = table.rows[i].cells[1]
        cell_value.width = Inches(3.0)
        p = cell_value.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(value)
        run.font.name = 'Arial'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        run.font.size = Pt(14)
        run.font.bold = True
        p.paragraph_format.line_spacing = 1.5
    
    return table
```

### 5.4 参考文献真实性验证规则

AI 代理生成参考文献时必须遵守以下真实性规则：

1. **可检索来源**：每篇参考文献必须来自以下至少一种可检索数据库：Crossref（含DOI）、CNKI（知网）、万方、维普、Google Scholar
2. **禁止虚构字段**：禁止虚构 DOI、作者姓名、期刊名称、卷期号、页码
3. **格式校验**：
   - DOI 格式：`10.xxxx/xxxxx`（以 `10.` 开头）
   - ISSN 格式：`XXXX-XXXX`（8位数字，中间连字符）
   - 期刊文章必须包含：作者、标题[J]、期刊名、年份、卷(期): 页码
4. **参考文献数量**：商业计划书 ≥8 篇，实验报告 ≥5 篇，课程设计 ≥6 篇，毕业论文 ≥15 篇
5. **真实性标记**：如无法确认某篇参考文献的真实性，必须在条目末尾标注 `[来源待确认]`

---

## 6. 自动目录生成

### 6.1 核心规则

> **⚠️ 绝对关键**：目录生成是文档构建的**倒数第二步**（仅早于保存）。必须在**正文全部写完、所有标题和内容就绪**之后，在**保存文档之前**执行。顺序：先写封面→分节→先写目录标题→再分节→写正文→所有内容完成后→回到目录节→插入TOC字段码→保存。

### 6.2 目录标题

```python
def add_toc_heading(doc):
    """添加目录页的标题「目  录」，居中，18pt黑体加粗"""
    p = doc.add_paragraph(style='TOCHeading')
    run = p.add_run('目  录')
    # 确保字体设置
    run.font.name = 'Arial'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    run.font.size = Pt(18)
    run.font.bold = True
    return p
```

### 6.3 插入 TOC 字段码

**注意**：python-docx 不提供高级 TOC API，必须直接操作底层 XML 插入字段码。

```python
def insert_toc_field(doc):
    """
    在当前位置插入自动目录字段。
    必须在正文内容全部写入后调用。
    目录会捕获所有应用了 Heading*_Course 样式的标题。
    """
    paragraph = doc.add_paragraph()
    
    # 设置 TOC 段落格式
    paragraph.paragraph_format.line_spacing = 1.5
    
    # 创建 TOC 字段的三个部分
    # 1. 字段开始标记
    run_begin = paragraph.add_run()
    fldChar_begin = parse_xml('<w:fldChar {} w:fldCharType="begin"/>'.format(nsdecls('w')))
    run_begin._r.append(fldChar_begin)
    
    # 2. TOC 指令（告诉 Word 生成目录）
    run_instr = paragraph.add_run()
    instrText = parse_xml(
        '<w:instrText {} xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'.format(nsdecls('w'))
    )
    run_instr._r.append(instrText)
    
    # 3. 字段结束标记
    run_sep = paragraph.add_run()
    fldChar_sep = parse_xml('<w:fldChar {} w:fldCharType="separate"/>'.format(nsdecls('w')))
    run_sep._r.append(fldChar_sep)
    
    # 4. 占位文本（当在代码中预览时会看到这个）
    run_text = paragraph.add_run('（请在 Word 中右键点击此处，选择"更新域"以生成目录）')
    run_text.font.color.rgb = RGBColor(128, 128, 128)
    run_text.font.size = Pt(10)
    
    # 5. 字段结束
    run_end = paragraph.add_run()
    fldChar_end = parse_xml('<w:fldChar {} w:fldCharType="end"/>'.format(nsdecls('w')))
    run_end._r.append(fldChar_end)
    
    return paragraph
```

### 6.4 目录样式调整

Word 会为 TOC 条目标记预定义的 `TOC1`（一级标题）、`TOC2`（二级标题）、`TOC3`（三级标题）样式。在生成目录**之前**调整这些样式的格式：

```python
def setup_toc_entry_styles(doc):
    """设置目录条目格式：右对齐点线前导符，页码"""
    # TOC1 — 一级目录条目
    style = doc.styles['TOC1']
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(12)
    style.font.bold = False
    style.paragraph_format.line_spacing = 1.5
    
    # 设置制表位：右对齐 + 点线前导符
    pPr = style.element.get_or_add_pPr()
    tabs = parse_xml(
        '<w:tabs {}>'
        '  <w:tab w:val="right" w:leader="dot" w:pos="9072"/>'
        '</w:tabs>'.format(nsdecls('w'))
    )
    pPr.append(tabs)
    
    # TOC2 — 二级目录条目（缩进1字符）
    style = doc.styles['TOC2']
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(12)
    style.paragraph_format.left_indent = Cm(0.75)
    style.paragraph_format.line_spacing = 1.5
    
    pPr = style.element.get_or_add_pPr()
    tabs = parse_xml(
        '<w:tabs {}>'
        '  <w:tab w:val="right" w:leader="dot" w:pos="9072"/>'  
        '</w:tabs>'.format(nsdecls('w'))
    )
    pPr.append(tabs)
```

---

## 7. 正文排版细则

### 7.1 标题层级对比表

| 层级 | 样式名 | 字体 | 字号 | 加粗 | 对齐 | 段前 | 段后 | 行距 |
|------|--------|------|------|------|------|------|------|------|
| 一级（1, 2, 3…） | `Heading1_Course` | 黑体 | 18pt (三号) | 是 | 左对齐（顶格） | 12pt | 6pt | 1.5倍 |
| 二级（1.1, 2.1…） | `Heading2_Course` | 黑体 | 15pt (四号) | 是 | 左对齐 | 12pt | 6pt | 1.5倍 |
| 三级（1.1.1…） | `Heading3_Course` | 黑体 | 13pt (小四) | 是 | 左对齐 | 6pt | 3pt | 1.5倍 |

> **国标依据**：根据 GB/T 7713.2-2022 §5.2.2，章节编号应使用阿拉伯数字，不同层次间用下圆点相隔，末位数字后不加点号。例如："1 项目背景与经营理念"（一级）、"1.1 项目背景"（二级）。编号与标题正文之间空1个汉字的间隙。**禁止**使用"一、二、三…"等中文数字作为章节编号。
>
> **对齐要求**：GB/T 7713.2-2022 §5.2.2 同时规定"各层次章节编号全部顶格排"——即所有层级标题必须**左对齐**（不缩进），不可居中。标题末尾不加标点。只有文档大标题（封面）和目录标题允许居中。

### 7.2 正文段落格式

| 属性 | 值 | 说明 |
|------|-----|------|
| 中文字体 | 宋体 (SimSun) | 中文正文标准 |
| 英文字体 | Times New Roman | 英文/数字标准 |
| 字号 | 12pt (小四) | 中文规范要求 |
| 行距 | 1.5倍 | 360缇，保证阅读舒适 |
| 段后 | 6pt | 段落间呼吸感 |
| 首行缩进 | 2字符（约24pt） | 中文段落标准 |
| 对齐 | 两端对齐 (JUSTIFY) | 左右边缘整齐 |

### 7.3 列表项格式

列表（如"第一…"、"（1）…"等）应与正文样式一致，保持首行缩进和行距：

```python
def add_list_item(doc, text, indent_level=0):
    """添加列表项，保持与正文相同的字体和行距"""
    p = doc.add_paragraph(style='BodyText_Course')
    p.paragraph_format.first_line_indent = Pt(0)  # 列表项不需要首行缩进
    if indent_level > 0:
        p.paragraph_format.left_indent = Cm(0.75 * indent_level)
    run = p.add_run(text)
    return p
```

---

## 7.3 参考文献格式（GB/T 7714-2015）

参考文献著录应遵循 GB/T 7714-2015《信息与文献 参考文献著录规则》。以下提供五种常见文献类型的著录格式模板：

```python
def add_reference_entry(doc, ref_type, authors, title, **kwargs):
    """
    添加参考文献条目（GB/T 7714-2015 格式）。
    
    参数：
        doc: Document 对象
        ref_type: 文献类型：'J'(期刊) / 'M'(专著) / 'D'(学位论文) / 'R'(报告) / 'S'(标准)
        authors: 作者列表，如 ['帅满, 葛雅南']
        title: 文献标题
        **kwargs: 不同类型需要的额外字段
    
    返回：段落对象
    """
    p = doc.add_paragraph(style='RefEntry')
    
    # 作者部分
    author_str = ', '.join(authors) + '. '
    
    if ref_type == 'J':  # 期刊论文 [J]
        journal = kwargs.get('journal', '')
        year = kwargs.get('year', '')
        volume = kwargs.get('volume', '')
        issue = kwargs.get('issue', '')
        pages = kwargs.get('pages', '')
        full_text = (
            f"{author_str}{title}[J]. {journal}, "
            f"{year}{f', {volume}' if volume else ''}"
            f"{f'({issue})' if issue else ''}"
            f"{f': {pages}' if pages else ''}."
        )
    elif ref_type == 'M':  # 专著 [M]
        translator = kwargs.get('translator', '')
        publisher = kwargs.get('publisher', '')
        city = kwargs.get('city', '')
        year = kwargs.get('year', '')
        trans_str = f'{translator}, 译. ' if translator else ''
        city_str = f'{city}: ' if city else ''
        full_text = f"{author_str}{title}[M]. {trans_str}{city_str}{publisher}, {year}."
    elif ref_type == 'D':  # 学位论文 [D]
        institution = kwargs.get('institution', '')
        year = kwargs.get('year', '')
        full_text = f"{author_str}{title}[D]. {institution}, {year}."
    elif ref_type == 'R':  # 报告 [R]
        institution = kwargs.get('institution', '')
        city = kwargs.get('city', '')
        year = kwargs.get('year', '')
        full_text = f"{author_str}{title}[R]. {city}: {institution}, {year}."
    elif ref_type == 'S':  # 标准 [S]
        std_code = kwargs.get('std_code', '')
        year = kwargs.get('year', '')
        full_text = f"{author_str}{std_code} {title}[S]. {year}."
    else:
        full_text = f"{author_str}{title}."
    
    p.add_run(full_text)
    return p


# 使用示例
# add_reference_entry(doc, 'J',
#     authors=['帅满, 葛雅南'],
#     title='大学生消费的面子补偿机制研究',
#     journal='青年研究', year='2024', issue='1', pages='78-92')
#
# add_reference_entry(doc, 'M',
#     authors=['菲利普·科特勒, 凯文·莱恩·凯勒'],
#     title='营销管理（第16版）',
#     translator='何佳讯',
#     city='北京', publisher='中信出版集团', year='2022')
#
# add_reference_entry(doc, 'R',
#     authors=['艾媒咨询'],
#     title='2024年中国大学生消费行为调查研究报告',
#     city='广州', institution='艾媒咨询', year='2024')
#
# add_reference_entry(doc, 'S',
#     authors=['财政部, 国家税务总局'],
#     title='关于小微企业和个体工商户发展有关税费政策的公告',
#     std_code='财政部 税务总局公告2023年第12号', year='2023')
```

### 7.4 参考文献在文档中的位置

参考文献应位于正文之后、附录之前，单独成章。一级标题为"参考文献"（使用 `Heading1_Course` 样式），条目按正文引用顺序排列。

### 7.5 英文摘要（Abstract）

**【规则】**：所有文档**必须**包含英文摘要。正文超过 6000 字时，英文摘要不少于 150 词。英文摘要位于中文摘要之后、第一章之前。

```python
def add_english_abstract(doc, abstract_text, keywords, is_short=True):
    """
    添加英文摘要（Abstract + Keywords）。
    
    参数：
        doc: Document 对象
        abstract_text: 英文摘要内容
        keywords: 英文关键词（逗号分隔）
        is_short: True=简短摘要(≥60词)，False=完整摘要(≥150词，正文>6000字时)
    """
    # Abstract 标题
    p = doc.add_paragraph(style='TOCHeading')  # 复用目录标题样式（居中18pt）
    run = p.add_run('Abstract')
    run.font.name = 'Arial'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    run.font.size = Pt(18)
    run.font.bold = True
    
    # Abstract 正文（Times New Roman 12pt, italic）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(abstract_text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.font.italic = True
    
    # Keywords 行
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing = 1.5
    run_label = p.add_run('Keywords: ')
    run_label.font.name = 'Times New Roman'
    run_label.font.size = Pt(12)
    run_label.font.bold = True
    run_values = p.add_run(keywords)
    run_values.font.name = 'Times New Roman'
    run_values.font.size = Pt(12)
    run_values.font.italic = True
```

### 7.6 反查重写作指南

AI 生成课程设计文档存在被查重系统检测的风险。以下规则可有效降低重复率：

1. **句式变换**：避免"本文介绍了…""首先…其次…最后…"等模板化开头
2. **同义替换**：常见高频词用同义词替换（如"分析→剖析""重要→关键""提升→增强"）
3. **段落重组**：同一内容可用不同逻辑顺序表达（因果→并列或递进）
4. **长度控制**：每段控制在 **150-350 字**之间，避免长段落（>400字）和短段落（<100字）
5. **禁止列表替代段落**：不得使用 `(1)(2)(3)(4)` 或 `① ② ③ ④` 编号列表替代正文段落。列表仅用于枚举型内容（如技术参数、版本号、成员名单），且每个列表项不超过 2 行，前后必须有过渡段落。
5. **增加个人见解**：每章至少加入1-2句基于数据的个人判断（非模板化语言）
6. **连续13字不重复**：确保任意连续13个字不与常见模板库中的表述相同
7. **数据差异化**：表格中的具体数值使用文档上下文计算出的真实值，不使用"约""大概""左右"

### 7.7 章节字数控制与篇幅要求

**【强制规则】**：最终文档总字数不得低于下表对应的最低要求。每章实际字数不得低于该章目标字数的 60%，也不得超过 150%（即 ±40% 浮动范围）。

| 文档类型 | 总字数下限 | 段落字数范围 | 最短段落 | 最长段落 |
|---------|-----------|-------------|---------|---------|
| 商业计划书 | 8,000 | 150-350 | 100 | 400 |
| 软件课程设计 | 6,000 | 150-350 | 100 | 400 |
| 实验报告 | 5,000 | 150-300 | 100 | 350 |
| 实习报告 | 5,000 | 150-300 | 100 | 350 |
| 毕业论文/设计 | 10,000 | 200-400 | 150 | 500 |

#### 商业计划书各章字数分配表

| 章节 | 目标字数 | 浮动范围（±40%） |
|------|---------|-----------------|
| 第1章 项目背景与经营理念 | 1,500 | 900-2,100 |
| 第2章 市场分析 | 1,200 | 720-1,680 |
| 第3章 产品与服务 | 1,000 | 600-1,400 |
| 第4章 商业模式 | 800 | 480-1,120 |
| 第5章 管理团队 | 600 | 360-840 |
| 第6章 财务分析 | 1,000 | 600-1,400 |
| 第7章 风险分析 | 600 | 360-840 |
| 其他（摘要/过渡/附录） | 1,300 | 780-1,820 |

#### 软件课程设计各章字数分配表

| 章节 | 目标字数 | 浮动范围（±40%） |
|------|---------|-----------------|
| 第1章 需求分析 | 1,200 | 720-1,680 |
| 第2章 系统设计 | 1,000 | 600-1,400 |
| 第3章 数据库设计 | 800 | 480-1,120 |
| 第4章 详细设计 | 800 | 480-1,120 |
| 第5章 系统实现 | 1,000 | 600-1,400 |
| 第6章 测试 | 600 | 360-840 |
| 其他（摘要/过渡/附录） | 600 | 360-840 |

#### 实验报告各章字数分配表

| 章节 | 目标字数 | 浮动范围（±40%） |
|------|---------|-----------------|
| 第1章 实验目的 | 600 | 360-840 |
| 第2章 实验原理 | 800 | 480-1,120 |
| 第3章 实验步骤 | 600 | 360-840 |
| 第4章 实验结果 | 1,200 | 720-1,680 |
| 第5章 讨论与分析 | 1,000 | 600-1,400 |
| 第6章 结论 | 400 | 240-560 |
| 其他（摘要/过渡/附录） | 400 | 240-560 |

#### 实习报告各章字数分配表

| 章节 | 目标字数 | 浮动范围（±40%） |
|------|---------|-----------------|
| 第1章 实习概况 | 800 | 480-1,120 |
| 第2章 工作内容 | 1,200 | 720-1,680 |
| 第3章 技能分析 | 800 | 480-1,120 |
| 第4章 问题与反思 | 600 | 360-840 |
| 第5章 总结 | 600 | 360-840 |
| 其他（摘要/过渡/附录） | 1,000 | 600-1,400 |

#### 字数检查工具函数

```python
def check_document_length(doc, doc_type='business_plan'):
    """
    检查文档总字数是否达标，返回 (pass, details)。
    
    参数：
        doc: Document 对象
        doc_type: 文档类型
    
    返回：
        (True/False, 报告字符串)
    """
    # 各类型最低字数
    min_words = {
        'business_plan': 8000,
        'software_course': 6000,
        'experiment': 5000,
        'internship': 5000,
        'thesis': 10000,
    }
    
    # 统计所有正文段落字数和短段落数量
    total = 0
    short_paras = []
    for p in doc.paragraphs:
        if p.style and p.style.name == 'BodyText_Course':
            text = p.text.strip()
            if text:
                word_count = len(text)
                total += word_count
                if word_count < 100:
                    short_paras.append((p.style.name, word_count, text[:50]))
    
    threshold = min_words.get(doc_type, 6000)
    passed = total >= threshold
    report_parts = [
        f'文档类型: {doc_type}',
        f'总字数: {total}',
        f'目标字数: {threshold}',
        f'通过: {"✅" if passed else "❌"}',
        f'短段落(<100字): {len(short_paras)}个',
    ]
    if short_paras:
        report_parts.append('--- 短段落详情 ---')
        for style_name, wc, preview in short_paras[:10]:
            report_parts.append(f'  [{wc}字] {preview}...')
    
    return passed, '\n'.join(report_parts)
```

### 7.8 写作质量规则

#### 章节过渡规则

每章结尾必须有 **1-2 句承上启下的过渡段落**，将当前章节的结论自然引导到下一章。禁止章节间"硬切换"（即上一章最后一句话直接结束，无任何过渡）。

```python
def add_transition(doc, from_chapter, to_chapter, summary):
    """
    添加章节过渡段落。
    
    示例：add_transition(doc, '市场分析', '产品定位',
        '综合以上市场数据，本报告接下来将基于这些发现进行产品定位分析。')
    """
    p = doc.add_paragraph(style='BodyText_Course')
    run = p.add_run(
        f'综合以上对{from_chapter}的剖析，'
        f'下文将进一步探讨{to_chapter}。'
    )
```

#### 图表数据解读规则

每个图表之后必须紧跟 **至少 1 段（≥150 字）的数据分析和解读段落**。禁止出现"有图无析"（呈现了图表但正文只字不提）或"有表无析"（呈现了表格但无分析文字）。

数据解读段落应包含：
1. **总体趋势**：图表呈现的整体规律或方向
2. **关键数据点**：标注最大值、最小值、增长率等关键数值
3. **业务含义**：数据背后的商业/技术/实验意义
4. **与前文呼应**：指向章节论点，证明或佐证前文观点

#### 图注/表题格式统一规则

所有图注必须遵循以下固定格式：
- 图注：`图X-Y  标题（数据来源：xxx）`
- 表题：`表X-Y  标题（数据来源：xxx）`

其中：
- `X` = 章号（阿拉伯数字），`Y` = 该章内图/表序号
- 标题与编号之间空 **2 个汉字字符**（`  `）
- 数据来源不可省略，不可使用"笔者整理""自行绘制"等模糊来源
- 数据来源应明确到具体机构或数据库，如"国家统计局""艾媒咨询""Wind 数据库"

---

## 8. 图片插入标准

### 8.1 图片处理规则

| 规则 | 值 | 说明 |
|------|-----|------|
| 最大宽度 | 4.5 英寸 (11.43 cm) | A4正文区宽度约16cm（边距2.5cm），留出边距 |
| 对齐 | 居中 | `paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER` |
| 图注位置 | 图片下方，单独段落 | 用 `Caption_Course` 样式 |
| **图注格式** | `图X-Y  标题（数据来源：xxx）` | 固定格式，必须包含数据来源 |
| 图片格式 | PNG / JPG | 优先PNG保证清晰度 |
| 防止Markdown残留 | 用 re 替换 `![...](...)` | 确保原数据不残留 |

### 8.2 图片插入函数

```python
from docx.shared import Inches
import re

def add_image(doc, image_path, caption=None, max_width=Inches(4.5)):
    """
    插入图片并自动添加图注。
    
    参数：
        image_path: 图片文件路径
        caption: 图注文本，如"图1  中国小商品市场规模"
        max_width: 图片最大宽度，默认4.5英寸
    """
    # 创建图片段落的空段落
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 插入图片
    run = paragraph.add_run()
    run.add_picture(image_path, width=max_width)
    
    # 添加图注
    if caption:
        add_caption(doc, caption)
    
    return paragraph


def sanitize_markdown_images(text):
    """
    清洗文本中的 Markdown 图片语法残留。
    例如：![](media/image1.png) → 空字符串
    """
    import re
    text = re.sub(r'!\[.*?\]\(.*?\)', '', text)
    return text
```

### 8.3 图片随文布局

图片段落前后各空一行（空段落），以确保图片不会与正文粘连：

```python
def add_image_with_spacing(doc, image_path, caption, max_width=Inches(4.5)):
    """在图片上下各留一行空白"""
    doc.add_paragraph(style='BodyText_Course').add_run('')  # 上方留白
    add_image(doc, image_path, caption, max_width)
    doc.add_paragraph(style='BodyText_Course').add_run('')  # 下方留白
```

---

## 8.4 图表自动生成引擎（ChartEngine）

**【强制规则】**：文档正文超过10页时，必须自动生成本地图表（PNG格式）并插入docx。图表总数不少于8张，且图类型不得少于4种。

### 8.4.1 ChartEngine 框架

```python
import matplotlib
matplotlib.use('Agg')  # 非交互模式，确保无GUI
import matplotlib.pyplot as plt
from matplotlib import font_manager
import numpy as np
import os, tempfile

# ──────────────────────────────────────────────
# 全局 matplotlib 中文字体配置
# 在所有绘图函数执行前调用，确保图中中文正常显示
# ──────────────────────────────────────────────
def _setup_chinese_font():
    """动态查找系统中文字体并设置为 matplotlib 默认字体"""
    # 常见中文字体名称（按优先级排列）
    preferred = ['SimHei', 'SimSun', 'Microsoft YaHei', 'Heiti SC', 'Heiti TC',
                 'PingFang SC', 'WenQuanYi Micro Hei', 'WenQuanYi Zen Hei',
                 'Noto Sans CJK SC', 'Noto Sans CJK JP', 'AR PL UMing CN']
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in preferred:
        if name in available:
            matplotlib.rcParams['font.family'] = 'sans-serif'
            matplotlib.rcParams['font.sans-serif'] = [name] + matplotlib.rcParams.get('font.sans-serif', [])
            matplotlib.rcParams['axes.unicode_minus'] = False
            return True
    # 兜底：搜索名字中包含中文字体关键词的字体
    fallback = [f.name for f in font_manager.fontManager.ttflist
                if any(k in f.name for k in ['SimHei','SimSun','YaHei','Heiti','Song','Hei','Kai','Fang'])]
    if fallback:
        matplotlib.rcParams['font.family'] = 'sans-serif'
        matplotlib.rcParams['font.sans-serif'] = [fallback[0]] + matplotlib.rcParams.get('font.sans-serif', [])
        matplotlib.rcParams['axes.unicode_minus'] = False
        return True
    return False

# 启动时执行一次
_setup_chinese_font()


class ChartEngine:
    """图表生成引擎，根据文档类型自动生成并插入图表"""
    
    def __init__(self, doc, doc_type='business_plan', temp_dir='./chart_temp'):
        # 再次确保中文字体已加载（子进程中可能丢失 rcParams）
        _setup_chinese_font()
        self.doc = doc
        self.doc_type = doc_type
        self.temp_dir = temp_dir
        self.chart_counter = 0  # 全局图序号
        self.table_counter = 0  # 全局表序号
        self.current_chapter = 0
        self.generated_types = set()  # 记录已生成的图类型
        os.makedirs(temp_dir, exist_ok=True)
    
    def _next_figure_num(self, chapter=None):
        """生成图编号（章节级），如 图2-1"""
        ch = chapter or self.current_chapter
        self.chart_counter += 1
        return f'{ch}-{self.chart_counter}'
    
    def _next_table_num(self, chapter=None):
        ch = chapter or self.current_chapter
        self.table_counter += 1
        return f'{ch}-{self.table_counter}'
    
    def _save_and_insert(self, fig, caption, chapter=None):
        """保存 matplotlib figure 为 PNG 并插入 docx"""
        fig_num = self._next_figure_num(chapter)
        filename = f'fig_{fig_num}.png'
        filepath = os.path.join(self.temp_dir, filename)
        fig.savefig(filepath, dpi=200, bbox_inches='tight')
        plt.close(fig)
        
        # 插入图片（inline with text，前后空段）
        add_image_with_spacing(self.doc, filepath, f'图{fig_num}  {caption}')
        return filepath
    
    def set_chapter(self, num):
        """切换到新章节，重置二级计数器"""
        self.current_chapter = num
        self.chart_counter = 0
        self.table_counter = 0
```

### 8.4.2 支持的图表类型（7种 matplotlib 函数模板）

```python
def draw_line_chart(data, title, xlabel, ylabel, labels=None):
    """折线图 — 趋势/时间序列"""
    fig, ax = plt.subplots(figsize=(6, 3.5))
    if isinstance(data[0], (list, tuple)):
        for i, series in enumerate(data):
            label = labels[i] if labels and i < len(labels) else f'系列{i+1}'
            ax.plot(series, marker='o', linewidth=1.5, label=label)
    else:
        ax.plot(data, marker='o', linewidth=1.5, color='#2E86AB')
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel(xlabel, fontsize=10); ax.set_ylabel(ylabel, fontsize=10)
    ax.grid(True, alpha=0.3); ax.legend()
    fig.tight_layout()
    return fig

def draw_bar_chart(data, categories, title, xlabel='', ylabel=''):
    """柱状图 — 对比/分布"""
    fig, ax = plt.subplots(figsize=(6, 3.5))
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#3B8C5A', '#C73E3E']
    bars = ax.bar(categories, data, color=colors[:len(categories)], edgecolor='white', linewidth=0.5)
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel(xlabel, fontsize=10); ax.set_ylabel(ylabel, fontsize=10)
    ax.grid(axis='y', alpha=0.3)
    # 在柱状图上标注数值
    for bar, val in zip(bars, data):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(data)*0.02,
                f'{val}', ha='center', va='bottom', fontsize=9)
    fig.tight_layout()
    return fig

def draw_pie_chart(data, labels, title):
    """饼图 — 占比/结构"""
    fig, ax = plt.subplots(figsize=(5, 4))
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#3B8C5A', '#C73E3E', '#6C4B9E']
    wedges, texts, autotexts = ax.pie(
        data, labels=labels, autopct='%1.1f%%',
        colors=colors[:len(labels)], startangle=90,
        textprops={'fontsize': 9}
    )
    ax.set_title(title, fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig

def draw_radar_chart(data, dimensions, title, series_names=None):
    """雷达图 — 多维度对比"""
    angles = np.linspace(0, 2*np.pi, len(dimensions), endpoint=False).tolist()
    angles += angles[:1]
    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    colors = ['#2E86AB', '#A23B72', '#F18F01']
    
    if isinstance(data[0], (list, tuple)):
        for i, values in enumerate(data):
            values_plot = values + values[:1]
            label = series_names[i] if series_names and i < len(series_names) else f'S{i+1}'
            ax.plot(angles, values_plot, 'o-', linewidth=1.5, label=label, color=colors[i % len(colors)])
            ax.fill(angles, values_plot, alpha=0.1, color=colors[i % len(colors)])
    else:
        data_plot = data + data[:1]
        ax.plot(angles, data_plot, 'o-', linewidth=1.5, color='#2E86AB')
        ax.fill(angles, data_plot, alpha=0.15, color='#2E86AB')
    
    ax.set_xticks(angles[:-1]); ax.set_xticklabels(dimensions, fontsize=9)
    ax.set_title(title, fontsize=12, fontweight='bold', pad=20)
    ax.legend(loc='upper right', fontsize=9)
    fig.tight_layout()
    return fig

def draw_horizontal_bar(data, labels, title, xlabel='', ylabel=''):
    """横向柱状图 — 排名"""
    fig, ax = plt.subplots(figsize=(6, 4))
    colors = plt.cm.Blues(np.linspace(0.4, 0.9, len(labels)))
    bars = ax.barh(labels, data, color=colors, edgecolor='white', linewidth=0.5)
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel(xlabel, fontsize=10); ax.set_ylabel(ylabel, fontsize=10)
    ax.grid(axis='x', alpha=0.3)
    for bar, val in zip(bars, data):
        ax.text(bar.get_width() + max(data)*0.02, bar.get_y() + bar.get_height()/2,
                f'{val}', ha='left', va='center', fontsize=9)
    fig.tight_layout()
    return fig

def draw_grouped_bar(data_dict, categories, title, xlabel='', ylabel=''):
    """分组柱状图 — 多维度对比"""
    fig, ax = plt.subplots(figsize=(7, 4))
    n_groups = len(categories); n_series = len(data_dict)
    width = 0.8 / n_series
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#3B8C5A']
    
    for i, (label, values) in enumerate(data_dict.items()):
        x = np.arange(n_groups) + i * width - (n_series - 1) * width / 2
        bars = ax.bar(x, values, width, label=label, color=colors[i % len(colors)], edgecolor='white')
    
    ax.set_xticks(np.arange(n_groups)); ax.set_xticklabels(categories, fontsize=9)
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel(xlabel, fontsize=10); ax.set_ylabel(ylabel, fontsize=10)
    ax.legend(fontsize=9); ax.grid(axis='y', alpha=0.3)
    fig.tight_layout()
    return fig

def draw_multi_line(data_dict, title, xlabel='', ylabel='', x_labels=None):
    """多系列折线图 — 多条趋势线对比"""
    fig, ax = plt.subplots(figsize=(6.5, 3.8))
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#3B8C5A', '#C73E3E']
    markers = ['o', 's', '^', 'D', 'v']
    for i, (label, values) in enumerate(data_dict.items()):
        x = range(len(values))
        ax.plot(x, values, marker=markers[i % len(markers)], linewidth=1.5,
                label=label, color=colors[i % len(colors)])
    if x_labels:
        ax.set_xticks(range(len(x_labels))); ax.set_xticklabels(x_labels, fontsize=8, rotation=30)
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel(xlabel, fontsize=10); ax.set_ylabel(ylabel, fontsize=10)
    ax.legend(fontsize=8); ax.grid(True, alpha=0.3)
    fig.tight_layout()
    return fig
```

## 8.5 按文档类型的智能配图规则

AI 代理根据用户输入判断文档类型，从下表中选择对应的 ≥8 张图配置。生成时**禁止**全篇使用同一种图表类型，且同类型图不可连续出现超过 2 次。

### 商业计划书类

| # | 图表名 | 图类型 | 插入位置 | 数据内容 |
|---|--------|--------|---------|---------|
| 1 | 市场规模趋势 | 折线图 | 1.1 项目背景后 | 2023-2027年市场规模预测 |
| 2 | 消费结构分布 | 饼图 | 2.2 目标市场分析后 | 各类消费占比 |
| 3 | 客群画像雷达 | 雷达图 | 2.2 目标市场分析后 | 多维度客群特征对比 |
| 4 | 竞品多维对比 | 分组柱状图 | 2.4 竞争格局分析后 | 本店 vs 精品屋 vs 超市 vs 电商 |
| 5 | 营收预测趋势 | 多系列折线图 | 4.1 盈利模式后 | 5年营收/成本/利润趋势 |
| 6 | 成本结构分析 | 饼图 | 4.1 盈利模式后 | 各项成本占比 |
| 7 | 盈亏平衡分析 | 折线图 | 6.4 盈亏平衡分析后 | 固定成本/变动成本/总收入 |
| 8 | 财务指标对比 | 横向柱状图 | 6.3 盈利能力分析后 | 各年ROE/净利润对比 |

### 实验报告类

| # | 图表名 | 图类型 | 插入位置 | 数据内容 |
|---|--------|--------|---------|---------|
| 1 | 实验流程 | 折线图（节点） | 实验原理后 | 实验步骤流程图 |
| 2 | 数据对比 | 柱状图 | 实验结果后 | 各组实验数据 |
| 3 | 对照组对比 | 分组柱状图 | 实验结果后 | 对照组 vs 实验组 |
| 4 | 误差分析 | 柱状图 | 实验分析后 | 误差棒分布 |
| 5 | 变化趋势 | 折线图 | 实验分析后 | 随时间变化趋势 |
| 6 | 数据分布 | 饼图 | 实验分析后 | 数据类别占比 |
| 7 | 相关性分析 | 多系列折线图 | 讨论后 | 多变量关系 |
| 8 | 结论对比 | 横向柱状图 | 结论后 | 结论汇总对比 |

### 软件课程设计类

| # | 图表名 | 图类型 | 插入位置 | 数据内容 |
|---|--------|--------|---------|---------|
| 1 | 系统架构 | 折线图（树形节点） | 系统设计后 | 系统架构层级图 |
| 2 | 功能模块对比 | 柱状图 | 需求分析后 | 各模块功能点 |
| 3 | 用户角色分布 | 饼图 | 需求分析后 | 用户类型占比 |
| 4 | 性能基准测试 | 分组柱状图 | 测试后 | 响应时间/吞吐量 |
| 5 | 并发能力 | 横向柱状图 | 测试后 | 并发用户数对比 |
| 6 | 响应时间趋势 | 折线图 | 测试后 | 多次测试响应时间 |
| 7 | 技术栈对比 | 雷达图 | 技术选型后 | 各技术多维对比 |
| 8 | 数据量增长 | 多系列折线图 | 数据库设计后 | 数据记录增长趋势 |

### 实习报告类

| # | 图表名 | 图类型 | 插入位置 | 数据内容 |
|---|--------|--------|---------|---------|
| 1 | 实习时间轴 | 横向柱状图 | 实习概况后 | 各阶段时间分布 |
| 2 | 工作内容分布 | 饼图 | 工作内容后 | 不同类型工作占比 |
| 3 | 技能成长雷达 | 雷达图 | 技能分析后 | 入职前 vs 入职后 |
| 4 | 成长趋势 | 折线图 | 成长分析后 | 各月能力评分 |
| 5 | 任务难度对比 | 柱状图 | 工作内容后 | 各类任务难度评分 |
| 6 | 成果类型分布 | 饼图 | 成果展示后 | 成果类型占比 |
| 7 | 团队结构 | 分组柱状图 | 团队介绍后 | 各团队人数/项目数 |
| 8 | 反思总结 | 多系列折线图 | 总结后 | 多维度成长曲线 |

## 8.6 图表生成与插入约束规则

**生成约束：**
1. 全文图类型不得少于 4 种（如全用折线图属于违规）
2. 同类型图不可连续出现超过 2 次（如连续3张柱状图属于违规）
3. 每张图必须生成独立图注，格式为 `图X-Y  标题（数据来源：xxx）`
4. 图注使用 `Caption_Course` 样式（宋体 10.5pt 居中）
5. 每张图在正文中必须有交叉引用（如"如图X-Y所示"）
6. 图表数据必须基于正文内容计算生成，不得编造

**插入约束：**
1. 所有图片使用 `run.add_picture()` 在空段落中（inline with text 模式）
2. 图片**禁止浮动、禁止文字环绕**（不设置 `anchor` 或 `wrap` 属性）
3. 图片前后各插入一个空段（使用 `BodyText_Course` 样式，内容为空）
4. 图片最大宽度 4.5 英寸（`Inches(4.5)`），高度按比例缩放
5. 图片段落居中对齐（`paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER`）
6. 生成的临时 PNG 保存在 `./chart_temp/` 目录，最终交付时清理

## 8.7 图表编号系统

所有图和表使用**章节级编号**，格式为 `图X-Y` / `表X-Y`，其中 X 为章号，Y 为序号。

```python
class FigureCounter:
    """全局图表编号管理器"""
    def __init__(self):
        self.chapter = 0
        self.fig_seq = 0
        self.tbl_seq = 0
    
    def new_chapter(self, num):
        self.chapter = num
        self.fig_seq = 0
        self.tbl_seq = 0
    
    def next_fig(self):
        self.fig_seq += 1
        return f'{self.chapter}-{self.fig_seq}'
    
    def next_tbl(self):
        self.tbl_seq += 1
        return f'{self.chapter}-{self.tbl_seq}'

# 使用示例
counter = FigureCounter()
counter.new_chapter(2)
# 图2-1, 图2-2, 表2-1 ...
```

## 8.8 交叉引用规范

正文中引用图/表时，使用统一的引用格式：

```python
def add_figure_ref(doc, fig_num):
    """在正文中添加图的交叉引用，如：如图2-1所示"""
    p = doc.add_paragraph(style='BodyText_Course')
    p.add_run(f'如图{fig_num}所示')

def add_table_ref(doc, tbl_num):
    """在正文中添加表的交叉引用，如：如表3-1所示"""
    p = doc.add_paragraph(style='BodyText_Course')
    p.add_run(f'如表{tbl_num}所示')
```

> **交叉引用规则**：① 每个图和表在正文中至少被引用一次 ② 引用位置在图表出现之前（先引用后展示） ③ 引用格式统一为"如图X-Y所示""如表X-Y所示" ④ 图表编号在最终定稿前确认无误

---

## 9. 表格格式化标准

### 9.1 三线表创建函数（默认标准）

**【强制规则】**：课程设计文档的表格**默认必须使用三线表**（顶线+栏目线+底线，无竖线，无彩色表头，无隔行底纹）。全边框彩色表格仅限商业演示等非学术场景使用。三线表符合 CY/T 170 和 GB/T 7713.2-2022 附录 B 推荐规范。

```python
def create_three_line_table(doc, headers, data, col_widths=None):
    """
    创建标准学术三线表（顶线+栏目线+底线，无竖线）。
    
    三线表规范（CY/T 170, GB/T 7713.2-2022 附录B）：
    ═══════════════════════════    ← 顶线 粗 1.5pt
      项目  │ 2024 │ 2025         ← 表头（白底黑字黑体加粗）
    ────────────                   ← 栏目线 细 0.75pt（仅表头与数据行之间）
      数据A   xx     xx
      数据B   xx     xx           ← 数据行（宋体，无内部横线）
    ═══════════════════════════    ← 底线 粗 1.5pt
    * 禁止：竖线、彩色表头、深蓝背景、隔行底纹、数据行间横线
    
    参数：
        doc: Document 对象
        headers: 表头列表，如 ['指标', '2023年', '2024年']
        data: 数据行列表，如 [['在校大学生', '4350', '4420'], ...]
        col_widths: 列宽百分比列表（可选），如 [0.2, 0.16, ...]，总和为1
    
    返回：表格对象
    """
    table = doc.add_table(rows=1 + len(data), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    tblPr = table._tbl.tblPr

    # ===== 三线表边框设置（关键修正） =====
    # 正确方法：表格级只设顶线+底线（粗），所有竖线和内部横线为 none。
    # 栏目线（表头下方的细线）通过为表头行每个单元格添加 tcBorders 底部边框实现。
    # 这样数据行之间不会有任何横线，真正符合三线表规范。
    tblBorders = parse_xml(
        '<w:tblBorders {}>'
        '  <w:top w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
        '  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:bottom w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
        '  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'.format(nsdecls('w'))
    )
    tblPr.append(tblBorders)

    # ===== 表头行（白底黑字黑体加粗）+ 栏目线（单元格底部边框） =====
    header_row = table.rows[0]
    for i, header_text in enumerate(headers):
        cell = header_row.cells[i]
        tcPr = cell._tc.get_or_add_tcPr()

        # 垂直居中
        vAlign = parse_xml('<w:vAlign {} w:val="center"/>'.format(nsdecls('w')))
        tcPr.append(vAlign)

        # 为表头单元格添加底部边框（栏目线：细 0.75pt）
        tcBorders = parse_xml(
            '<w:tcBorders {}>'
            '  <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
            '</w:tcBorders>'.format(nsdecls('w'))
        )
        tcPr.append(tcBorders)

        # 无背景色（白底）
        if col_widths and i < len(col_widths):
            cell.width = Inches(col_widths[i] * 6.0)

        cell.paragraphs[0].clear()
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)

        run = p.add_run(header_text)
        run.font.name = 'Arial'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # 黑色文字

    # ===== 数据行（宋体 10pt，无背景色，无隔行底纹） =====
    for row_idx, row_data in enumerate(data):
        row = table.rows[row_idx + 1]
        for col_idx, cell_text in enumerate(row_data):
            cell = row.cells[col_idx]
            tcPr = cell._tc.get_or_add_tcPr()

            # 垂直居中
            vAlign = parse_xml('<w:vAlign {} w:val="center"/>'.format(nsdecls('w')))
            tcPr.append(vAlign)

            if col_widths and col_idx < len(col_widths):
                cell.width = Inches(col_widths[col_idx] * 6.0)

            cell.paragraphs[0].clear()
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)

            run = p.add_run(cell_text)
            run.font.name = 'Times New Roman'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(10)

    return table


def add_table_title(doc, caption_text):
    """
    在表格上方添加表序+表题（GB/T 7713.2-2022 要求）。
    
    格式：表X-Y  标题（宋体 10.5pt 居中）
    位置：表格上方，紧跟表格
    """
    p = doc.add_paragraph(style='Caption_Course')
    run = p.add_run(caption_text)
    # 表题后跟一个空段，确保与表格之间有一点间距
    doc.add_paragraph(style='BodyText_Course').add_run('')


def add_table_note(doc, note_text):
    """
    在表格下方添加注释（"注：xxx"）。
    
    格式：注：xxx（宋体 9pt 左对齐）
    位置：表格下方，表后留空段后
    """
    # 表格与注释之间留空段
    doc.add_paragraph(style='BodyText_Course').add_run('')
    p = doc.add_paragraph(style='BodyText_Course')
    p.paragraph_format.first_line_indent = Pt(0)  # 注释不需要首行缩进
    run = p.add_run(f'注：{note_text}')
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(9)
    run.font.italic = True
    return p
```

### 9.2 三线表属性速查表

| 属性 | 值 | 说明 |
|------|-----|------|
| 对齐 | 居中 | `table.alignment = WD_TABLE_ALIGNMENT.CENTER` |
| 宽度 | auto | 由内容自动分配 |
| 顶线 | 粗 1.5pt 黑色 | 表格最上方 |
| 栏目线 | 粗 0.75pt 黑色 | 仅表头与数据行之间 |
| 底线 | 粗 1.5pt 黑色 | 表格最下方 |
| 竖线 | **禁止** | 左/右/内部竖线均为 none |
| 数据行间横线 | **禁止** | 数据行之间无横线 |
| 栏目线实现 | 单元格 tcBorders | 表头行每个单元格单独添加底部边框，不依赖表格级 insideH |
| 表头背景 | **禁止彩色** | 纯白底（无 shading） |
| 表头文字 | 黑体 10pt 黑色 加粗 | 居中 |
| 数据文字 | 宋体 10pt | 居中 |
| 垂直对齐 | 居中 | `vAlign="center"` |
| 单元格边距 | 上下各 3pt | `space_before` / `space_after` |
| 隔行底纹 | **禁止** | 所有数据行纯白底 |

> **✅ 已修复**：旧版代码使用 `insideH=single` 导致数据行间出现横线。现已改用表头行单元格 `tcBorders` 实现栏目线，表格级 `insideH=none`，数据行间完全无横线，符合 CY/T 170 三线表规范。

### 9.3 全边框表格（备选，仅限非学术场景）

全边框彩色表格（表头 `#2E86AB` 深蓝背景+白色文字+隔行底纹）仅在**商业路演PPT**、**内部报表**等非学术场景使用。课程设计、实验报告、毕业论文**禁止使用**。如需使用，参考以下简化代码：

```python
# 备选：全边框彩色表格 — 仅限商业演示场景
def create_fancy_table(doc, headers, data):
    table = doc.add_table(rows=1+len(data), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # ... （完整代码见原 create_course_table）
    return table
```

---

## 10. 页眉页脚与页码

### 10.1 正文节页脚（页码）

```python
def add_page_number_footer(section):
    """
    在指定节的页脚添加页码。
    格式：居中，宋体 10pt
    """
    footer = section.footer
    footer.is_linked_to_previous = False
    
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 创建 PAGE 字段
    run = p.add_run()
    fldChar_begin = parse_xml('<w:fldChar {} w:fldCharType="begin"/>'.format(nsdecls('w')))
    run._r.append(fldChar_begin)
    
    run2 = p.add_run()
    instrText = parse_xml('<w:instrText {} xml:space="preserve"> PAGE </w:instrText>'.format(nsdecls('w')))
    run2._r.append(instrText)
    
    run3 = p.add_run()
    fldChar_sep = parse_xml('<w:fldChar {} w:fldCharType="separate"/>'.format(nsdecls('w')))
    run3._r.append(fldChar_sep)
    
    run4 = p.add_run('1')  # 占位数字
    run4.font.name = 'Times New Roman'
    run4.font.size = Pt(10)
    
    run5 = p.add_run()
    fldChar_end = parse_xml('<w:fldChar {} w:fldCharType="end"/>'.format(nsdecls('w')))
    run5._r.append(fldChar_end)
    
    return footer
```

### 10.2 分节与页码控制完整示例

```python
def setup_sections_with_pagination(doc):
    """
    完整的分节与页码设置流程。
    假设 doc 已有封面和目录内容（在第一节中）。
    """
    # 第一节（封面+目录）：无页码
    first_section = doc.sections[0]
    first_section.top_margin = Cm(2.54)
    first_section.bottom_margin = Cm(2.54)
    first_section.left_margin = Cm(2.5)
    first_section.right_margin = Cm(2.5)
    
    # 封面+目录节不要页眉页脚
    first_section.header.is_linked_to_previous = False
    first_section.footer.is_linked_to_previous = False
    # 清空页眉页脚
    for p in first_section.header.paragraphs:
        p.clear()
    for p in first_section.footer.paragraphs:
        p.clear()
    
    # 添加第二节（正文节）
    new_section = doc.add_section()
    new_section.top_margin = Cm(2.54)
    new_section.bottom_margin = Cm(2.54)
    new_section.left_margin = Cm(2.5)
    new_section.right_margin = Cm(2.5)
    
    # 正文节添加页码
    add_page_number_footer(new_section)
    
    # 正文节页码从1开始
    sectPr = new_section._sectPr
    pgNumType = parse_xml('<w:pgNumType {} w:fmt="decimal" w:start="1"/>'.format(nsdecls('w')))
    sectPr.append(pgNumType)
    
    return new_section


### 10.3 正文节页眉设计

多数高校要求页眉显示学校名称和文档类型，例如"XX大学课程设计报告"。

```python
def add_header(section, header_text='课程设计报告'):
    """
    在指定节的页眉添加学校名称+文档标题。
    格式：宋体 五号(10.5pt) 居中。
    
    参数：
        section: 节对象
        header_text: 页眉文字，例如"XX大学课程设计报告"
    """
    header = section.header
    header.is_linked_to_previous = False
    
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    run = p.add_run(header_text)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)
    
    return header
```

> **页眉规则**：仅正文节需要页眉，封面节和目录节应清空页眉内容。默认情况下 Word 新节的页眉会继承上一节，必须设置 `is_linked_to_previous = False` 断开链接。

### 10.4 前置部分罗马数字页码

根据 GB/T 7713.2-2022 要求，正文前的目录页应使用罗马数字单独编连续码。

```python
def setup_page_numbering(doc):
    """
    完整的三节页码控制：
    - 第一节（封面）：无页码
    - 第二节（目录）：罗马数字页码（i, ii, iii…）
    - 第三节（正文）：阿拉伯数字页码从1开始
    """
    sections = doc.sections
    
    # ===== 防御性检查 =====
    if len(sections) < 3:
        raise ValueError(
            f'setup_page_numbering() 需要至少 3 个节（封面/目录/正文），'
            f'当前只有 {len(sections)} 个节。'
            f'请确认在调用前已通过 add_section() 创建了目录节和正文节。'
        )
    
    # === 第一节（封面）：无页码 ===
    if len(sections) >= 1:
        sections[0].header.is_linked_to_previous = False
        sections[0].footer.is_linked_to_previous = False
    
    # === 第二节（目录）：罗马数字 ===
    if len(sections) >= 2:
        toc_section = sections[1]
        toc_section.header.is_linked_to_previous = False
        toc_section.footer.is_linked_to_previous = False
        
        # 目录节页码格式：小写罗马数字
        tocSectPr = toc_section._sectPr
        tocPgNumType = parse_xml('<w:pgNumType {} w:fmt="romanLower" w:start="1"/>'.format(nsdecls('w')))
        tocSectPr.append(tocPgNumType)
        
        # 添加罗马数字页码
        footer = toc_section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # PAGE 字段（格式设置为 romanLower）
        run = p.add_run()
        fldChar_begin = parse_xml('<w:fldChar {} w:fldCharType="begin"/>'.format(nsdecls('w')))
        run._r.append(fldChar_begin)
        
        run2 = p.add_run()
        # 注意：Python 字符串中 \\ 表示一个反斜杠，最终 XML 内容为 PAGE \* roman
        instrText = parse_xml('<w:instrText {} xml:space="preserve"> PAGE \\\\* roman </w:instrText>'.format(nsdecls('w')))
        run2._r.append(instrText)
        
        run3 = p.add_run()
        fldChar_sep = parse_xml('<w:fldChar {} w:fldCharType="separate"/>'.format(nsdecls('w')))
        run3._r.append(fldChar_sep)
        
        run4 = p.add_run('i')
        run4.font.name = 'Times New Roman'
        run4.font.size = Pt(10)
        
        run5 = p.add_run()
        fldChar_end = parse_xml('<w:fldChar {} w:fldCharType="end"/>'.format(nsdecls('w')))
        run5._r.append(fldChar_end)
    
    # === 第三节（正文）：阿拉伯数字从1开始 ===
    if len(sections) >= 3:
        body_section = sections[2]
        body_section.header.is_linked_to_previous = False
        body_section.footer.is_linked_to_previous = False
        
        # 添加阿拉伯数字页码
        add_page_number_footer(body_section)
        
        # 页码从1开始
        sectPr = body_section._sectPr
        pgNumType = parse_xml('<w:pgNumType {} w:fmt="decimal" w:start="1"/>'.format(nsdecls('w')))
        sectPr.append(pgNumType)
```
```

---

## 11. 生成顺序工作流

这是**最重要的部分**。严格遵循以下顺序，不得跳步或颠倒：

```
第1步：接收用户需求，判断文档类型（§2.1）
第2步：创建 Document 对象
第3步：定义全局页面设置（纸张、边距）
第4步：调用 setup_styles() 创建所有自定义样式（含 outlineLvl）
第5步：初始化 ChartEngine，按文档类型选择 §8.5 配图表
第6步：写封面内容（校名、标题、分隔线、信息表格）
第7步：调用 add_section() → 进入目录节（新页开始）
第8步：写目录标题「目  录」（TOCHeading样式，居中）
第9步：调用 insert_toc_field() 插入TOC字段码
第10步：调用 add_section() → 进入正文节（新页开始）
第11步：调用 setup_page_numbering() 设置三节页码（封面无→目录罗马→正文阿拉伯）
第12步：调用 add_header() 为正文节添加页眉
第13步：写正文全部内容，逐章执行：
       • counter.new_chapter(chapter_num)
       • 按 §8.5 规则在对应位置调用绘图函数 → 插入图表
       • 表格使用 create_three_line_table() 三线表，表题在前（add_table_title），注释在后（add_table_note）
       • 图注使用章节级编号 图X-Y/表X-Y
第14步：调用 check_document_length() 验证总字数达标，短段落检测
第15步：始终调用 add_english_abstract() 生成英文摘要（正文>6000字时≥150词）
第16步：写参考文献（GB/T 7714-2015，验证真实性）
第17步：保存文档 (.docx)
第18步：清理临时PNG文件，自我质量检查
```

> **为什么 TOC 字段可以在正文之前插入？**
> TOC 字段码（`TOC \o "1-3" \h \z \u`）只是一个指令，Word 在打开文档并"更新域"时才会实际计算目录内容。因此在正文之前插入字段码是安全的——当用户在 Word 中点击右键→更新域，Word 会扫描全文的所有标题，生成完整的目录条目和正确的页码。

### 完整伪代码

```python
def generate_course_document(content_data):
    doc = Document()
    
    # ===== 第1~3步：初始化 =====
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    
    setup_styles(doc)
    setup_toc_entry_styles(doc)
    
    # ===== 第4步：封面 =====
    # ... 写校名、课程名称、分隔线 ...
    add_cover_info_table(doc, [
        ("项 目 名 称", "xxx"),
        ("创 作 小 组", "xxx"),
        ("主 要 负 责 人", "xxx"),
        ("小 组 成 员", "xxx"),
        ("呈 报 日 期", "xxx"),
    ])
    
    # ===== 第5步：分节 → 目录节 =====
    doc.add_section()
    toc_section = doc.sections[-1]
    # 首页节（封面+目录）不要页码
    toc_section.header.is_linked_to_previous = False
    toc_section.footer.is_linked_to_previous = False
    
    # ===== 第6~7步：目录标题 + TOC字段 =====
    add_toc_heading(doc)
    insert_toc_field(doc)
    
    # ===== 第8步：分节 → 正文节 =====
    doc.add_section()
    body_section = doc.sections[-1]
    
    # ===== 第9步：正文节页码从1开始 + 页脚 =====
    body_section.header.is_linked_to_previous = False
    body_section.footer.is_linked_to_previous = False
    add_page_number_footer(body_section)
    
    sectPr = body_section._sectPr
    pgNumType = parse_xml('<w:pgNumType {} w:fmt="decimal" w:start="1"/>'.format(nsdecls('w')))
    sectPr.append(pgNumType)
    
    # ===== 第10~11步：主体内容 =====
    for chapter in content_data['chapters']:
        if chapter.get('page_break', True):
            # 一级标题前分页
            p_break = doc.add_paragraph()
            p_break.paragraph_format.page_break_before = True
        
        add_heading1(doc, chapter['title'])
        for section_data in chapter.get('sections', []):
            add_heading2(doc, section_data['title'])
            for para in section_data.get('paragraphs', []):
                add_body(doc, para)
        
        if 'tables' in chapter:
            for t in chapter['tables']:
                add_table_title(doc, t['caption'])            # 表题在上方
                create_three_line_table(doc, t['headers'], t['data'])
                if 'note' in t and t['note']:
                    add_table_note(doc, t['note'])          # 注释在下方
        
        if 'images' in chapter:
            for img in chapter['images']:
                add_image_with_spacing(doc, img['path'], img['caption'])
    
    # 参考文献
    add_heading1(doc, '参考文献')
    for ref in content_data.get('references', []):
        p = doc.add_paragraph(style='RefEntry')
        p.add_run(ref)
    
    # ===== 第12步：保存 =====
    doc.save('课程设计报告.docx')
    return doc
```

> **关于第5步的分节说明**：`doc.add_section()` 会自动插入一个分节符（下一页），并从当前文档位置开始新节。所以封面写完后调用 → 目录节开始 → 目录标题+TOC字段 → 再调用 → 正文节开始。每次新节的页眉页脚默认继承上一节，必须设置 `is_linked_to_previous = False` 来断开链接。

---

## 12. 质量检查清单

在交付最终文档前，逐项检查以下内容：

- [ ] **样式完整性**：文档中所有段落都使用了自定义样式（Normal 样式未被直接使用）
- [ ] **导航窗格**：在 Word 中打开后，按 Ctrl+F 打开导航窗格，应能通过标题层级浏览整篇文档
- [ ] **章节编号合规**：一级标题使用阿拉伯数字（1, 2, 3…），非中文数字（一、二、三…）
- [ ] **标题顶格对齐**：所有层级标题左对齐顶格（不缩进），非居中
- [ ] **大纲级别存在**：自定义标题样式已设置 outlineLvl，导航窗格可识别
- [ ] **目录可更新**：右键点击目录区域，选择"更新域"→"更新整个目录"，应能正确生成目录和页码
- [ ] **前置页码**：目录页使用罗马数字（i, ii, iii…），正文使用阿拉伯数字（1, 2, 3…）
- [ ] **封面无页码**：封面页不应显示任何页码；页码无 "PAGE 1" 等异常文本
- [ ] **正文页码从1开始**：正文第一页应为页码1，与目录页页码连续
- [ ] **页眉正确**：正文节有页眉（如"XX大学课程设计报告"），封面/目录节无页眉
- [ ] **页边距合规**：左右边距 2.5cm，上下边距 2.54cm
- [ ] **三线表强制**：所有表格使用三线表（顶线+栏目线+底线），禁止竖线/彩色表头/隔行底纹
- [ ] **图表数量充足**：正文≥10页时，图数量≥8张，表数量合理搭配
- [ ] **图类型多样**：全文图类型≥4种（折线/柱状/饼图/雷达/横向柱/分组柱/多系列），禁止只用同一种
- [ ] **图类型不连续重复**：同类型图不连续出现超过2次
- [ ] **图片嵌入方式**：所有图片 inline with text（非浮动/非环绕），前后有空段
- [ ] **图片未溢出**：图片宽度≤4.5英寸，不超出正文区域
- [ ] **图片居中**：所有图片段落居中对齐
- [ ] **图片无Markdown残留**：文档中不应出现 `![](...)` 文本
- [ ] **图表编号系统**：图和表使用章节级编号（图2-1、表3-2），非全文流水号
- [ ] **交叉引用存在**：每张图和表在正文中被交叉引用至少一次
- [ ] **图表真实**：图表数据基于正文内容计算生成，未编造数据
- [ ] **模板三线表规范**：仅顶线粗、栏目线细、底线粗，无竖线，无数据行间横线
- [ ] **分隔线为段落边框**：不是 `━━━━` 字符
- [ ] **封面信息对齐**：使用无边框表格（不是空格）
- [ ] **行距统一**：正文 1.5 倍，标题 1.5 倍
- [ ] **页眉页脚干净**：封面/目录节无页码，正文节有页码
- [ ] **英文摘要**（始终需要）：英文 Abstract + Keywords，正文>6000字时摘要≥150词
- [ ] **文献真实**：参考文献可检索（标注了DOI/CNKI等来源），无虚构信息
- [ ] **反查重**：无连续13字模板化表达，句式有变化
- [ ] **段落长度合规**：每段 150-350 字（最短≥100，最长≤400），无 (1)(2)(3)(4) 列表替代正文
- [ ] **章节字数均衡**：每章字数在目标值的 ±40% 范围内，无头重脚轻
- [ ] **总字数达标**：商业计划书≥8000，软件课设≥6000，实验/实习报告≥5000，毕业论文≥10000
- [ ] **表题位置正确**：表序+表题在表格上方（GB/T 7713.2-2022），注释"注：xxx"在表格下方
- [ ] **图注格式统一**：所有图注格式为"图X-Y  标题（数据来源：xxx）"，含数据来源
- [ ] **图表后有数据解读**：每个图表后紧跟 ≥1 段数据分析和解读段落
- [ ] **章节间过渡**：每章结尾有 1-2 句承上启下的过渡段落

---

## 13. 常见错误速查表

| # | 错误表现 | 根本原因 | 正确做法 | 后果 |
|---|---------|---------|---------|------|
| 1 | 所有段落都使用 Normal 样式 | 未创建自定义样式 | 通过 `doc.styles.add_style()` 创建 Heading1_Course 等 | 无法导航、无法自动目录 |
| 2 | 标题层级字体/字号不一致 | 未统一定义标题样式 | 定义 Heading1_Course (18pt黑体顶格) / Heading2_Course (15pt黑体左对齐) | 读者无法识别文档结构 |
| 3 | 正文无字体、无字号、无行距 | 未设置 run.font 和 paragraph_format | 宋体12pt、1.5倍行距、段后6pt、首行缩进2字符 | 阅读体验差、显得拥挤 |
| 4 | 图片中出现 `![](media/...png)` | Markdown语法未解析 | 用 `re.sub(r'!\\[.*?\\]\\(.*?\\)', '', text)` 清洗 | 文档中出现无效文本 |
| 5 | 图片宽度撑满页面 | 未限制图片宽度 | `run.add_picture('img.png', width=Inches(4.5))` | 图片溢出边距 |
| 6 | 图片未居中 | 未设置 paragraph.alignment | `paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER` | 图片偏左 |
| 7 | 表格宽度参差不齐 | 未设置 table.alignment | `table.alignment = WD_TABLE_ALIGNMENT.CENTER` | 表格不美观 |
| 8 | 表头无背景色或颜色异常 | 未设置 shd 或 XML 方法错误 | `<w:shd w:fill="2E86AB" w:val="clear"/>` | 表头不突出 |
| 9 | 封面使用 `━━━━` 分隔线 | 字符拼凑 | 使用段落底边框 `w:pBdr/w:bottom` | 字体不同时显示不一致 |
| 10 | 封面信息用空格对齐 | 手动空格填充 | 使用无边框表格或制表位 | 对齐不精确 |
| 11 | 目录在正文前生成 | 生成顺序错误 | 正文全部写完后，最后步骤插入 TOC 字段码 | 目录缺少内容 |
| 12 | 页边距过大导致正文区拥挤 | 未按国标设置页边距 | 左=右=2.5cm（国标GB/T 7713.1-2006），正文区宽约16cm | 表格和图片容易溢出 |
| 13 | 页眉页脚有默认内容 | 未清除 | 封面/目录节：`header.is_linked_to_previous = False` 并清空 | 页码出现在封面 |
| 14 | 章节间未分页 | 缺少分页控制 | 一级标题前使用：`doc.add_paragraph().paragraph_format.page_break_before = True` | 标题出现在页面底部 |
| 15 | 字体设置不生效（中文） | 只设置了 font.name 未设置 eastAsia | 必须同时设置 `rFonts.set(qn('w:eastAsia'), '黑体')` 和 `font.name = 'Arial'` | 中文显示为默认字体 |
| 16 | 一级标题使用"一、二、三…"而非"1, 2, 3…" | 违反 GB/T 7713.2-2022 §5.2.2 | 编号用阿拉伯数字，编号后空1汉字间隙接排标题 | 学术审查时被直接退回 |
| 17 | 一级标题居中对齐 | 违反 GB/T 7713.2-2022 顶格排要求 | 所有层级标题左对齐（顶格），不可居中 | 不符合学术文档规范 |
| 18 | 表格使用全边框而非三线表 | 不符合学术出版规范 | 使用 `create_three_line_table()` 生成三线表（仅顶线+栏目线+底线） | 表格风格不够学术化 |
| 19 | 缺少页眉 | 未设置页眉 | 使用 `add_header()` 添加"XX大学课程设计报告" | 部分高校会扣分 |
| 20 | 目录页未使用罗马数字页码 | 违反 GB/T 7713.2-2022 前置部分编码要求 | 目录节使用 `PAGE \* roman` 罗马数字页码 | 不符合国家标准 |
| 21 | 参考文献格式不完整 | 未使用 GB/T 7714-2015 著录格式 | 使用 `add_reference_entry()` 按文献类型选择 J/M/D/R/S 模板 | 参考文献不规范 |
| 22 | 三线表含彩色表头/竖线/隔行底纹 | 违反三线表规范（CY/T 170） | 使用重写后的 `create_three_line_table()`，白底黑字无竖线无隔行 | 不符合学术规范 |
| 23 | 自定义标题未被目录识别 | 未设置 outlineLvl | 在 setup_styles() 中为 Heading1_Course 等添加 `<w:outlineLvl>` | 导航窗格为空，目录不完整 |
| 24 | 全文图表类型单一（全折线/全柱状） | 未强制图类型多样化 | 遵循 §8.5 配图规则表，保证 ≥4 种图类型 | 文档单调，可读性差 |
| 25 | 图片浮动/覆盖正文/文字环绕 | 未设置 inline with text | 使用 `run.add_picture()` 在空段落中（默认 inline），禁止 anchor | 图片与正文重叠错位 |
| 26 | 缺少英文摘要或摘要过短 | 未调用 add_english_abstract() 或摘要<60词 | 始终调用 `add_english_abstract()`，正文>6000字时摘要≥150词 | 部分高校扣分项 |
| 27 | 缺少交叉引用 | 图表无"如图X-Y所示"引用 | 在正文中使用 `add_figure_ref()`/`add_table_ref()` 引用图表 | 图表与正文脱节 |
| 28 | 参考文献虚构（编造DOI/作者/卷期） | 未做真实性验证 | 遵循 §5.4 规则：标注DOI/CNKI来源，禁止虚构 | 学术不端风险 |
| 29 | 图片中文显示为方框乱码 | matplotlib 未配置中文字体 | ChartEngine 自动调用 _setup_chinese_font() 动态查找字体 | 所有图表中文报废 |
| 30 | 三线表数据行间有横线 | 误用表格级 insideH | 使用 tcBorders 为表头行独立添加栏目线，数据行间无横线 | 表格不符合 CY/T 170 规范 |
| 31 | 表题在表格下方 | 违反 GB/T 7713.2-2022 | 使用 add_table_title() 在上方，add_table_note() 在下方 | 格式扣分 |
| 32 | 段落过短碎片化（<100字） | 用列表项替代正文段落 | 每段 150-350 字，禁止 (1)(2)(3)(4) 列表替代正文 | 内容空洞被退回 |
| 33 | 章节字数头重脚轻 | 未遵守字数分配表 | 参考 §7.7 各章目标字数 ±40% | 文档分布不均衡 |

---

## 14. 完整工作流速查（AI 代理清单）

当 AI 代理需要生成课程设计文档时，严格按此清单执行，每完成一项打勾：

```
[ ] 接收用户需求，判断文档类型（§2.1：商业计划书/实验报告/软件课设/实习报告/毕业论文）
[ ] 创建 Document 对象
[ ] 设置页面：A4 + 标准边距（左=右=2.5cm）
[ ] 调用 setup_styles() 和 setup_toc_entry_styles()
[ ] 写封面（校名/标题/分隔线/无边框表格）
[ ] add_section() 新建目录节
[ ] 写目录标题「目  录」
[ ] 调用 insert_toc_field() 插入TOC字段码
[ ] add_section() 新建正文节
[ ] 调用 setup_page_numbering() 设置三节页码
[ ] 调用 add_header() 添加页眉
[ ] 初始化 ChartEngine（matplotlib + 中文字体），按文档类型选择 §8.5 配图表
[ ] 写正文（逐章逐节，用自定义样式，同步生成图表并插入）
[ ]   每章开始时调用 counter.new_chapter(chapter_num)
[ ]   按配图规则在对应位置调用绘图函数 → _save_and_insert()
[ ]   每段控制在 150-350 字，禁止 (1)(2)(3)(4) 列表替代正文段落
[ ]   表格使用 create_three_line_table()（三线表），表题在上（add_table_title），注释在下（add_table_note）
[ ]   每个图表后紧跟 ≥1 段数据分析和解读
[ ]   每章结尾写 1-2 句承上启下的过渡段落
[ ]   图表编号使用章节级（图X-Y/表X-Y），交叉引用使用 add_figure_ref()
[ ] 调用 check_document_length() 验证总字数达标
[ ] 写参考文献（用 GB/T 7714-2015 格式，验证真实性）
[ ] 调用 add_english_abstract() 生成英文摘要（正文>6000字时摘要≥150词）
[ ] 调用 save() 保存文档
[ ] 自我质量检查（对照§12清单：图≥8/类型≥4种/三线表/字数达标/页码等）
[ ] 清理 chart_temp/ 目录中的临时PNG文件
```

---

## 15. 关于页面设置的补充说明

### 15.1 页边距参考

中国高校课程设计文档的页边距推荐值：上=2.54cm、下=2.54cm、左=2.5cm、右=2.5cm。这是基于 GB/T 7713.1-2006 的规定（天头≥25mm，订口≥25mm，地角≥20mm，切口≥20mm）。如遇特殊要求，可适当调整，但左/右边距不宜超过 3.0cm。

### 15.2 关于 A4 与 Letter

中国高校课程设计文档标准为 **A4 纸张**（21.0×29.7 cm）。如遇特殊要求（如美国高校），可切换到 US Letter（21.59×27.94 cm）。默认使用 A4。

### 15.3 首页不同设置

封面节和目录节的页眉页脚应链接关闭（`is_linked_to_previous = False`）并清空内容，避免默认页眉横线或页码污染封面页。

---

## 16. 高级技巧：python-docx 避坑指南

### 16.1 中文字体设置的正确姿势

```python
# ❌ 错误：只设置 font.name（这只设置了拉丁字体）
run.font.name = '黑体'

# ✅ 正确：同时设置 Latin 和 EastAsia 字体
run.font.name = 'Arial'  # 拉丁字体
run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')  # 东亚字体
```

### 16.2 段落样式 vs 直接格式

```python
# ❌ 错误：创建空段落然后手动设置每个格式
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run(text)
run.font.size = Pt(18)
run.font.bold = True
# ... 每个段落都这样写 → 不可维护

# ✅ 正确：使用样式
p = doc.add_paragraph(style='Heading1_Course')
p.add_run(text)
```

### 16.3 表格列宽设置

python-docx 的表格列宽设置较为复杂。推荐**不在 python-docx 层面强制固定列宽**，而是让 Word 的 `auto` 模式分配。如果必须固定列宽，使用：

```python
# 每个单元格设置宽度必须在填充内容之前
cell.width = Inches(1.5)
```

### 16.4 TOC 字段不自动更新

python-docx 生成的 TOC 字段在代码执行时**不会自动更新**（因为没有 Word 渲染引擎）。这是正常现象。最终用户需要在 Word 中打开文档后：
1. 右键点击目录区域
2. 选择"更新域"
3. 选择"更新整个目录"

在文档中可以用说明文字提示用户。

---

*本规范基于对标准课程设计文档的逆向工程分析，结合中国高校学术文档排版惯例制定。适用于所有基于 python-docx 的课程设计、实验报告、实习报告、商业计划书等中文文档生成任务。*
