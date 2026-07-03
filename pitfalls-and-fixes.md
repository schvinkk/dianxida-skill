# 踩坑记录与修复方案

本文档记录了在使用 OOXML 操作滇西应用技术大学毕业论文模板时遇到的所有问题，按严重程度排序。

## P0 — 擅自修改用户文字内容（致命错误）

### 现象

排版后用户发现论文正文被改写、润色或删减，原有表述被替换成"更好的"说法。

### 原因

AI 在处理文本时习惯性地"优化"措辞，擅自修改了用户的原文。

### 铁律

**排版只改格式，绝不动内容。** 用户写了什么就是什么。

允许的文字变更**仅限**：
- 标题空格规范化（"摘要"→"摘  要"、"目录"→"目  录"）
- 封面信息从草稿中提取填入
- 参考文献标点规范化（全角→半角逗号句号）

发现错别字、语法问题、内容薄弱 — **只提醒，不代改**。以清单形式输出让用户自行决定。

### 教训

- 修改用户论文内容等同于篡改学术成果，是最严重的错误
- 排版脚本中不应有任何文本替换/润色逻辑
- AI 率优化只给建议方向，不直接改写原文

## P0 — 硬编码样式 ID 导致格式全面错误（致命错误）

### 现象

输出的文档目录条目、页眉页脚、参考文献等格式全部错乱，与学校模板不一致。

### 原因

最初版本的 Skill 把样式 ID 写死在代码里：
```python
# 错误！假设 toc 1 = ID 10, header = ID 9, footer = ID 8
pStyle = '10'  # toc 1
pStyle = '9'   # header
pStyle = '8'   # footer
```

实测发现不同文档的样式 ID 完全不同：

| 样式名称 | 模板(.docx) | 例文 | 论文(1.毕业论文) |
|----------|:-----------:|:----:|:----------------:|
| toc 1 | 8 | 10 | 9 |
| toc 2 | 7 | 11 | 8 |
| header | 11 | 9 | 12 |
| footer | 10 | 8 | 11 |

### 修复

```python
# 正确！从 styles.xml 动态查找
style_name_map = {}
for style in styles_root.findall(f'{wns}style'):
    name_el = style.find(f'{wns}name')
    if name_el is not None:
        style_name_map[name_el.get(f'{wns}val')] = style.get(f'{wns}styleId')

toc1_id = style_name_map.get('toc 1')
header_style_id = style_name_map.get('header')
footer_style_id = style_name_map.get('footer')
```

### 教训

- **样式名称是稳定的，样式 ID 是不稳定的**
- 每次处理新文档时，第一步必须是动态分析 styles.xml
- 代码中绝不允许出现硬编码的样式 ID 数字

## P0 — 页眉错位一节（致命错误）

### 现象

每个章节的页眉显示的是**下一个**章节的标题。例如目录页显示"引言"，引言页显示"小檗碱概述"。

### 原因

错误逻辑：
```python
# 错误！遇到新章节标题时，创建新的空白段落放 sectPr
blank = ET.Element(f'{wns}p')
set_sect_break(blank, hdr_rid, ftr_rid, pgfmt, pgstart)
review_content.append((-1, blank))
# 这个空白段成了当前 section 的最后一段
# 但它的 header 用的是下一节的 header → 页眉错位
```

OOXML 规则：sectPr 所在段落属于该 section 的**结尾**。把 sectPr 放在新空白段上，这个空白段成了当前 section 的末段，但 header 引用指向了下一节的页眉。

### 修复

```python
# 正确！sectPr 放在当前 section 的最后一段上，header 属于当前 section
if next_para_index in sect_end_map:
    _, hdr_rid, ftr_rid, pgfmt, pgstart = sect_end_map[next_para_index]
    target = review_content[-1]  # 当前 section 的最后一段
    set_sect_break(target, hdr_rid, ftr_rid, pgfmt, pgstart)
    # header/footer ID 属于当前（正在结束的）section
```

### 教训

- 插入分节符前，**必须先用脚本分析例文**，搞清楚 sectPr 到底放在哪些段落上
- sectPr 的 header 永远属于"到此段落为止的 section"，不属于"从下一段开始的 section"
- 不要删除段落间的空白段落（它们是 section 间的分隔符），不要自己插入新空白段落

## P0 — 命名空间写错导致页眉不显示

### 现象

Word 打开后所有页眉页脚为空白。

### 原因

```python
# 错误！用了 wordprocessingml 命名空间
hr.set(f'{wns}type', 'default')  # 生成 w:type="default"
```

Word 要求 headerReference/footerReference 的 `type` 属性在 relationships 命名空间下：

```python
# 正确！
hr.set(f'{rns}type', 'default')  # 生成 r:type="default"
```

### 影响

`w:type` 被 Word 忽略，导致 header/footer 引用无效，页眉页脚不显示。

### 教训

- OOXML 中 `r:` 前缀的属性全部属于 relationships 命名空间
- `r:id` 和 `r:type` 是最常见的两个，都要用 `rns`

## P1 — 封面标题为空

### 现象

封面页中英文标题位置为空白。

### 原因

初始代码只清空了标题段落的文本（准备让用户自己填），没有填入论文的标题。

### 修复

```python
cover_titles = {
    9:  ('论文中文标题', None),
    10: ('English Title', 'Times New Roman'),
}
for idx, (title_text, ascii_font) in cover_titles.items():
    p = ex_para_to_child[idx]
    texts = p.findall('.//w:t', ns)
    if texts:
        texts[0].text = title_text
        for t in texts[1:]:
            t.text = ''
```

### 教训

- 封面标题是论文的重要信息，必须自动填入
- 英文标题需要设置 `rFonts ascii/hAnsi="Times New Roman"`
- **注意**：封面标题段落位置因文档而异（模板 para 9/11，例文 para 9/10），必须通过分析确定

## P1 — 目录是纯文本而非域代码

### 现象

目录条目是普通文本加手动点号（......），在 Word 中无法自动更新页码。

### 原因

直接保留了草稿中的手动目录，没有替换为域代码格式。

### 修复

将草稿中的手动目录段落（通常几十个 `<w:p>`）全部跳过（加入 skip_indices），在目录标题后插入域代码格式的目录条目。

### 教训

- 例文的目录是域代码格式（TOC + HYPERLINK + PAGEREF）
- 必须匹配例文的目录格式，不能保留手动目录

## P1 — 页眉/页脚 pStyle 写错

### 现象

页眉文字不居中，或页脚页码位置不对。

### 原因

header/footer XML 中的 `<w:pStyle w:val="9"/>` 写死了 ID 9，但实际例文中 header 样式 ID 可能是 11 或 12。

### 修复

从 styles.xml 动态查找 header 和 footer 样式的实际 ID：
```python
header_style_id = style_name_map.get('header')  # 例文中可能是 9/11/12
footer_style_id = style_name_map.get('footer')  # 例文中可能是 8/10/11
```

### 教训

- 同 P0-硬编码样式 ID，所有 pStyle 值都必须动态获取

## P2 — 多出空白页

### 现象

章节之间出现多余的空白页。

### 原因

1. 分节符（next page）与手动分页符叠加
2. 删除空白段落时打破了原有的分页逻辑

### 修复

- 保留原有的空白段落（它们通常是 section 间的分隔符）
- 不要自己插入新的空白段落
- 如果 section 末尾有多余空白，在设置 sectPr 前适当删除尾部空白

## P2 — 自创格式而非复制例文

### 现象

输出文档的样式与学校模板不一致（字体、间距、行高等细节差异）。

### 原因

从零开始构建 XML，没有以例文为基底。

### 修复

以例文的 unpacked 目录为基底，复制其 styles.xml、header/footer 文件、样式定义，只替换内容段落。

### 教训

**这是最基本也最重要的原则：复制粘贴，不要自创。** 例文包含了所有正确的样式定义，任何自行编写的格式都可能出错。

## P3 — 英文标题段落位置不对

### 现象

英文标题出现在了错误的位置（与中文标题间距不对或跑到了其他页面）。

### 原因

模板文档和例文的段落结构不同。模板中英文标题在 para 11，例文中在 para 10。

### 修复

必须先分析例文结构，确定标题段落的精确位置，不能假设模板和例文的位置相同。

## P3 — 模板 .doc 文件格式不支持

### 现象

学校提供的模板文件是 `.doc` 格式（非标准 docx），解压后不是标准的 `word/document.xml` 结构，而是 `drs/e2oDoc.xml` 的 Flat OPC 格式。

### 修复

- 使用桌面 "例文+模板" 文件夹中的 `.docx` 格式模板（`滇西应用技术大学毕业论文（设计）模板格式.docx`）
- 如果只有 `.doc` 文件，先用 WPS/Word 另存为 `.docx` 再操作

## P1 — 样式定义与 run 覆盖不一致导致字号错误

### 现象

heading 1 标题显示为 15pt 而非正确的 16pt（三号），heading 2 显示为 18pt 而非 14pt（四号）。

### 原因

只看 styles.xml 中的样式定义来设置字号，没有注意到论文中 run 级别有 `<w:sz>` 覆盖：

```python
# 错误！只看样式定义
# heading 1 样式定义 sz=30 → 设为 30 (15pt)
# 但论文实际 run 级别覆盖为 sz=32 (16pt = 三号)
```

### 修复

排版时以**例文中实际 run 级别的值**为准，不能只看样式定义。分析脚本应同时提取 run 级别的覆盖值：

| 样式 | 定义 sz | run 覆盖 sz | 应用值 |
|------|:------:|:----------:|:-----:|
| heading 1 | 30 | **32** | 32 |
| heading 2 | 36 | **28** | 28 |
| heading 3 | 28 | **24** | 24 |
| toc 1/2/3 | 30/28/— | **24** | 24 |
| header | 18 | **21** | 21 |

### 教训

- 样式定义是"默认值"，run 覆盖是"最终值"
- 分析例文时，不仅要看 styles.xml，还要检查 document.xml 中实际段落的 run 属性
- 生成新段落时，显式设置 run 级别的 `<w:sz>` 和 `<w:szCs>`，不要依赖样式继承

## P2 — TOC 制表位 pos 值硬编码

### 现象

目录条目的页码对齐位置偏左或偏右，与例文不一致。

### 原因

TOC 条目中的制表位 `<w:tab w:pos="8306"/>` 硬编码了一个值，但不同文档的 toc 样式定义中制表位不同：

- 论文(1.毕业论文)中 toc 1 的 tab pos = **8777**
- 例文中可能是其他值

### 修复

从例文的 styles.xml 中动态提取 toc 1 样式的制表位：

```python
def get_toc_tab_pos(styles_root, toc_style_name='toc 1'):
    """从 toc 样式定义中提取制表位 pos"""
    for style in styles_root.findall(f'{wns}style'):
        name_el = style.find(f'{wns}name')
        if name_el is not None and name_el.get(f'{wns}val') == toc_style_name:
            ppr = style.find(f'{wns}pPr')
            if ppr is not None:
                tabs = ppr.find(f'{wns}tabs')
                if tabs is not None:
                    tab = tabs.find(f'{wns}tab')
                    if tab is not None:
                        return tab.get(f'{wns}pos', '8306')
    return '8306'  # 默认回退值
```

### 教训

- TOC 制表位影响页码的对齐位置，必须与例文一致
- 不同文档的页面设置可能略有差异，导致制表位不同
- 应从样式定义中动态提取，不要硬编码

## P2 — 遗漏附录 section 导致页眉缺失

### 现象

论文有附录部分，但排版后附录页没有页眉，或页眉显示错误。

### 原因

脚本只处理了固定的章节列表（引言到结论+参考文献+致谢），没有考虑"附录"可能出现在结论和参考文献之间。

### 修复

在分析例文结构时，检测所有 heading 1 级别的标题，包括"附录"。如果草稿中存在一级标题"附录"，需要为其创建独立的 section 并设置页眉。

模板格式要求15规定：附录标题"字间空2格，字体字号同第一层次题序和标题"（即 heading 1 格式）。

### 教训

- 不要假设论文的章节结构是固定的
- 应动态检测所有 heading 1 标题，为每个创建独立 section
- 附录可能出现在结论与参考文献之间

## P3 — TOC 域代码缺少 \z 标志

### 现象

在 Word 的 Web Layout 视图中，目录条目显示了页码和制表符（不应该显示）。

### 原因

TOC 域代码缺少 `\z` 标志：

```python
# 错误！缺少 \z
instrText = 'TOC \\o "1-3" \\h \\u'

# 正确！模板和例文都使用 \z
instrText = 'TOC \\o "1-3" \\h \\z \\u'
```

`\z` 标志使 Word 在 Web Layout 视图中隐藏页码和前导符。

### 修复

在 TOC 域代码中始终包含 `\z` 标志。

### 教训

- 域代码的每个标志都有特定用途，不能遗漏
- 应以例文的实际域代码为准进行复制

## P3 — 参考文献标点使用全角

### 现象

参考文献条目中的逗号、句号显示为中文全角标点（，。），与模板要求不一致。

### 原因

从草稿中直接复制了参考文献文本，没有转换标点符号。

模板格式要求16明确规定："逗号句号用半角"，即使用 Times New Roman 的半角逗号(,)和句号(.)。

### 修复

处理参考文献条目时，将全角标点替换为半角：

```python
def fix_ref_punctuation(text):
    """参考文献条目：逗号句号改为半角 Times New Roman"""
    return text.replace('，', ',').replace('。', '.').replace('：', ':').replace('；', ';')
```

### 教训

- 参考文献的标点规范是半角，这是模板明确要求的
- 参考文献中的西文部分应全部使用 Times New Roman

## 通用排查清单

当输出结果不符合预期时，按以下顺序排查：

1. **检查样式 ID**：运行动态分析脚本，确认当前文档的实际样式 ID 映射
2. **检查例文结构**：运行分析脚本，确认每个 sectPr 的位置和对应 header
3. **检查 rId 映射**：确认 header_map 中每个 rId 对应的 header 文件内容正确
4. **检查命名空间**：所有 `r:type` 和 `r:id` 都用 `rns` 而非 `wns`
5. **检查 Content_Types**：所有新 header/footer 文件都在 `[Content_Types].xml` 中注册
6. **检查 rels**：所有新 rId 都在 `document.xml.rels` 中注册
7. **对比输出**：用分析脚本检查输出文档的 sectPr 位置，与例文对比
8. **检查 pStyle**：header/footer 文件中的 pStyle 使用例文中 header/footer 样式的实际 ID
9. **检查 run 级别覆盖**：确认生成的段落 run 级别 `<w:sz>` 与例文实际值一致，不只依赖样式继承
10. **检查 TOC 制表位**：确认目录条目的 `tab pos` 值与例文 toc 样式定义一致
11. **检查参考文献标点**：确认逗号句号为半角 Times New Roman
