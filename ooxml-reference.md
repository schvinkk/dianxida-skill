# OOXML 技术参考

## .docx 文件结构

.docx 本质是 ZIP 包，内含 XML 文件：

```
docx.zip/
├── [Content_Types].xml          # 声明所有部件的 MIME 类型
├── _rels/.rels                  # 包级关系
├── word/
│   ├── document.xml             # 主文档内容（段落、分节符）
│   ├── styles.xml               # 样式定义
│   ├── _rels/document.xml.rels  # 文档级关系（header/footer 引用）
│   ├── header1.xml ... headerN.xml  # 页眉文件
│   ├── footer1.xml ... footerN.xml  # 页脚文件
│   ├── settings.xml
│   └── ...
```

## 核心命名空间

| 前缀 | URI | 用途 |
|------|-----|------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | 主文档标记 |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | 关系引用 |
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | 兼容性标记 |
| `wps` | `http://schemas.microsoft.com/office/word/2010/wordprocessingShape` | 形状（文本框） |
| `v` | `urn:schemas-microsoft-com:vml` | VML 后备 |

Python 中定义：
```python
ns  = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
wns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
```

注册所有命名空间以避免序列化时出现 `ns0:`, `ns1:` 等前缀：
```python
for prefix, uri in [
    ('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('w14', 'http://schemas.microsoft.com/office/word/2010/wordml'),
    ('w15', 'http://schemas.microsoft.com/office/word/2012/wordml'),
    ('w10', 'urn:schemas-microsoft-com:office:word'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
    ('wpsCustomData', 'http://www.wps.cn/officeDocument/2013/wpsCustomData'),
    # ... 其他按需添加
]:
    ET.register_namespace(prefix, uri)
```

## 样式名称映射（动态查找）

**绝不能硬编码样式 ID！** 必须从 `styles.xml` 中按样式名称动态查找。

### 样式名称（稳定的，跨文档一致）

| 样式名称 | 用途 | 格式要点 |
|----------|------|---------|
| Normal | 正文 | 宋体/Times New Roman, 五号(sz=21), 两端对齐 |
| heading 1 | 一级标题 | 黑体, 三号(sz=32), 居中, 段前12磅段后6磅 |
| heading 2 | 二级标题 | 黑体, 四号(sz=28), 左对齐, 加粗, 段前6磅 |
| heading 3 | 三级标题 | 黑体, 小四(sz=24), 加粗 |
| toc 1 | 一级目录条目 | 黑体, 小四(sz=24) |
| toc 2 | 二级目录条目 | 宋体, 小四(sz=24), 首行缩进2字符 |
| toc 3 | 三级目录条目 | 宋体, 小四(sz=24), 首行缩进4字符 |
| header | 页眉样式 | 五号(sz=18/21), 居中 |
| footer | 页脚样式 | 九号(sz=18), 左对齐 |
| Plain Text / Normal (Web) | 致谢正文 | 宋体, 小四(sz=24) |

### 实测样式 ID 对照表（警告：不同文档 ID 不同）

| 样式名称 | 模板(.docx) | 例文 | 1.毕业论文+正确格式 |
|----------|:-----------:|:----:|:----------------:|
| Normal | 1 | 1 | 1 |
| heading 1 | 2 | 2 | 2 |
| heading 2 | 3 | 3 | 3 |
| heading 3 | 4 | 4 | 4 |
| toc 1 | **8** | **10** | **9** |
| toc 2 | **7** | **11** | **8** |
| toc 3 | **6** | **7** | **7** |
| header | **11** | **9** | **12** |
| footer | **10** | **8** | **11** |
| Title | 12 | — | **14** |
| Plain Text | 9 | 12 | **10** |
| Body Text Indent | 5 | 6 | **6** |
| TOC 标题1（自定义） | 15 | — | **19** |

### 关键：样式定义 ≠ 实际 run 覆盖

论文中的 run 级别 `<w:sz>` 会覆盖样式定义中的值。排版时必须以 run 级别为准：

| 样式 | 定义 sz | run 覆盖 sz | 实际字号 |
|------|:------:|:----------:|:------:|
| heading 1 | 30 (15pt) | **32 (16pt)** | 三号 |
| heading 2 | 36 (18pt) | **28 (14pt)** | 四号 |
| heading 3 | 28 (14pt) | **24 (12pt)** | 小四 |
| toc 1/2/3 | 30/28/— | **24 (12pt)** | 小四 |
| header | 18 (9pt) | **21 (10.5pt)** | 五号 |
| footer | 18 (9pt) | **21 (10.5pt)** | 五号 |

### 动态分析脚本（Step 0 必须运行）

```python
import zipfile
from xml.etree import ElementTree as ET
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
ns  = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def analyze_styles(docx_path):
    """从 docx 中提取样式名称到 ID 的映射"""
    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read('word/styles.xml'))
    
    style_map = {}  # name → id
    for style in root.findall(f'{wns}style'):
        sid = style.get(f'{wns}styleId', '')
        stype = style.get(f'{wns}type', '')
        name_el = style.find(f'{wns}name')
        name = name_el.get(f'{wns}val', '') if name_el is not None else ''
        if stype == 'paragraph' and name:
            style_map[name] = sid
    
    print(f"=== 样式映射 ({docx_path}) ===")
    for name, sid in sorted(style_map.items(), key=lambda x: int(x[1]) if x[1].isdigit() else 999):
        print(f"  {name:>30s} → ID={sid}")
    return style_map


def analyze_toc_tabs(docx_path):
    """从 toc 样式定义中提取制表位 pos 值"""
    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read('word/styles.xml'))
    
    toc_tabs = {}
    for style in root.findall(f'{wns}style'):
        name_el = style.find(f'{wns}name')
        if name_el is None:
            continue
        name = name_el.get(f'{wns}val', '')
        if name.startswith('toc '):
            ppr = style.find(f'{wns}pPr')
            if ppr is not None:
                tabs = ppr.find(f'{wns}tabs')
                if tabs is not None:
                    for tab in tabs.findall(f'{wns}tab'):
                        pos = tab.get(f'{wns}pos', '')
                        if pos:
                            toc_tabs[name] = int(pos)
    
    print(f"\n=== TOC 制表位 ({docx_path}) ===")
    for name, pos in sorted(toc_tabs.items()):
        print(f"  {name}: pos={pos}")
    return toc_tabs


def analyze_sections(docx_path):
    """分析文档的 section break 结构"""
    with zipfile.ZipFile(docx_path) as z:
        doc = ET.fromstring(z.read('word/document.xml'))
        rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
    
    body = doc.find(f'{wns}body')
    paras = list(body)
    rid_map = {r.get('Id'): r.get('Target') for r in rels}
    
    print(f"\n=== Section 结构 ({docx_path}) ===")
    print(f"总段落/元素数: {len(paras)}")
    
    sect_count = 0
    for i, p in enumerate(paras):
        if p.tag != f'{wns}p':
            if p.tag == f'{wns}sectPr':
                sect_count += 1
                _print_sectpr(p, rid_map, sect_count, f'body-level', z)
            continue
        
        ppr = p.find(f'{wns}pPr')
        sectpr = ppr.find(f'{wns}sectPr') if ppr is not None else None
        if sectpr is None:
            continue
        
        sect_count += 1
        text = ''.join(t.text or '' for t in p.iter(f'{wns}t'))[:60]
        _print_sectpr(sectpr, rid_map, sect_count, f'para {i}: [{text}]', z)


def _print_sectpr(sectpr, rid_map, sect_count, label, zf):
    """打印 sectPr 详细信息"""
    hdrs = sectpr.findall(f'{wns}headerReference')
    ftrs = sectpr.findall(f'{wns}footerReference')
    pgnum = sectpr.find(f'{wns}pgNumType')
    
    hdr_info = []
    for h in hdrs:
        rid = h.get(f'{rns}id', '')
        target = rid_map.get(rid, '')
        hdr_text = ''
        if target and target.startswith('header'):
            try:
                hdr_xml = ET.fromstring(zf.read(f'word/{target}'))
                hdr_text = ''.join(t.text for t in hdr_xml.iter(f'{wns}t') if t.text)
            except: pass
        hdr_info.append(f'{rid}→{target}:"{hdr_text}"')
    
    ftr_info = []
    for f in ftrs:
        rid = f.get(f'{rns}id', '')
        ftr_info.append(f'{rid}→{rid_map.get(rid, "")}')
    
    pg_info = ''
    if pgnum is not None:
        pg_info = f'{pgnum.get(f"{wns}fmt", "?")}/{pgnum.get(f"{wns}start", "?")}'
    
    print(f"  Section {sect_count}: {label}")
    print(f"    headers: {hdr_info}")
    print(f"    footers: {ftr_info}")
    if pg_info:
        print(f"    pgNumType: {pg_info}")


# 使用示例
# styles = analyze_styles('例文.docx')
# analyze_sections('例文.docx')
```

## 分节符（sectPr）详解

### 位置规则

`<w:sectPr>` 可以出现在两个地方：

1. **段落属性中** `<w:pPr><w:sectPr>...</w:sectPr></w:pPr>` — 定义**到此段落为止**的 section
2. **文档体直接子元素** `<w:body>...<w:sectPr>...</w:sectPr></w:body>` — 定义**最后一个** section

关键认知：**sectPr 所在的段落属于该 section 的结尾，不是下一个 section 的开始。**

```
段落 A1 ─┐
段落 A2  │── Section A（sectPr 在 A3 上）
段落 A3 ◄─┘  ← 此段携带 sectPr，header/footer 属于 Section A
段落 B1 ─┐
段落 B2  │── Section B（sectPr 在 B2 上）
段落 B2 ◄─┘  ← 此段携带 sectPr，header/footer 属于 Section B
段落 C1 ─┐
段落 C2  │── Section C（由 body 级 sectPr 定义）
<body sectPr> ← 定义最后一个 section
```

### sectPr XML 结构

```xml
<w:sectPr>
  <w:headerReference r:id="rId38" r:type="default"/>  <!-- 注意 r:type 不是 w:type -->
  <w:footerReference r:id="rId39" r:type="default"/>
  <w:pgSz w:w="11906" w:h="16838"/>                   <!-- A4 -->
  <w:pgMar w:top="1417" w:right="1134" w:bottom="1417"
           w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>
  <w:pgNumType w:fmt="decimal" w:start="1"/>           <!-- 可选 -->
  <w:cols w:space="425" w:num="1"/>
  <w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>
</w:sectPr>
```

### 页面设置（实测统一值）

所有 section 共用相同的页面设置：

```xml
<w:pgSz w:w="11906" w:h="16838"/>  <!-- A4 纵向 (210mm × 297mm) -->
<w:pgMar w:top="1417" w:right="1134" w:bottom="1417"
         w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>
<!-- 上2.5cm 右2.0cm 下2.5cm 左3.0cm 页眉1.5cm 页脚1.75cm -->
<w:cols w:space="425" w:num="1"/>
<w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>
```

换算：1cm = 567 twips（1英寸=1440twips，1cm=1440/2.54≈567）
- top/bottom: 2.5cm = 1417 twips
- right: 2.0cm = 1134 twips
- left: 3.0cm = 1701 twips
- header: 1.5cm = 851 twips
- footer: 1.75cm = 992 twips

### 页码格式（实测验证）

| Section | `pgNumType fmt` | `start` | 说明 |
|---------|-----------------|---------|------|
| 封面 section | `upperRoman` | `1` | 封面不显示页码 |
| 声明+授权 section | 不设 | — | 继承 decimal，不显示 |
| 摘要 section | `upperRoman` | `1` | 摘要 = I |
| Abstract section | `upperRoman` | `1` | Abstract = I（独立起编） |
| 目录 section | `upperRoman` | 不设 | 继承 = II, III... |
| 正文第一章 | `decimal` | `1` | 正文从 1 开始 |
| 后续正文各节 | `decimal` | 不设 | 连续编号 |
| 致谢（body级） | `decimal` | 不设 | 连续编号 |

## 页眉 XML 模板

**pStyle 的值必须动态获取**（例文中可能是 9/11/12）。以下用 `{HEADER_STYLE_ID}` 表示从 styles.xml 中按名称 "header" 查到的实际 ID。

### 有内容的页眉

```xml
<?xml version="1.0" encoding="utf-8"?>
<w:hdr xmlns:w="..." xmlns:r="..." ...>
  <w:p>
    <w:pPr>
      <w:pStyle w:val="{HEADER_STYLE_ID}"/>
      <w:pBdr>
        <w:bottom w:val="single" w:color="auto" w:sz="4" w:space="1"/>
      </w:pBdr>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:hint="eastAsia"/>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
      <w:t>{章节标题}</w:t>
    </w:r>
  </w:p>
</w:hdr>
```

要点：`pStyle` 使用例文中 header 样式的实际 ID、底部单线边框 `sz=4`、居中、五号字 `sz=21`。

**实测要点**：
- header 样式定义的 sz=18 (9pt)，但论文中 run 级别覆盖为 sz=21 (五号)
- 页眉下边框：`single, color=auto, sz=4, space=1`（由 header 样式定义自带）
- 居中由 `<w:jc w:val="center"/>` 实现
- 字体为宋体（eastAsia hint），数字/英文自动回退到 Times New Roman

### 空白页眉

```xml
<w:hdr ...>
  <w:p>
    <w:pPr><w:pStyle w:val="{HEADER_STYLE_ID}"/></w:pPr>
  </w:p>
</w:hdr>
```

## 页脚 XML 模板

页脚使用文本框包裹 PAGE 域代码，带 `mc:AlternateContent`（wps 主选 + VML 后备）。pStyle 使用例文中 footer 样式的实际 ID（`{FOOTER_STYLE_ID}`）。

```xml
<w:ftr ...>
  <w:p>
    <w:pPr><w:pStyle w:val="{FOOTER_STYLE_ID}"/></w:pPr>
    <w:r>
      <mc:AlternateContent>
        <mc:Choice Requires="wps">
          <w:drawing>
            <wp:anchor ...>
              <!-- wps:wsp 文本框 -->
              <!-- 内含 PAGE \* MERGEFORMAT 域代码 -->
            </wp:anchor>
          </w:drawing>
        </mc:Choice>
        <mc:Fallback>
          <w:pict>
            <!-- v:rect VML 后备文本框 -->
          </w:pict>
        </mc:Fallback>
      </mc:AlternateContent>
    </w:r>
  </w:p>
</w:ftr>
```

**直接复制例文的页脚文件即可**，所有 section 使用相同的页脚结构。

## 域代码目录条目结构

pStyle 值必须使用从 styles.xml 中动态查到的 toc 1 / toc 2 / toc 3 ID。以下用 `{TOC1_ID}`、`{TOC2_ID}`、`{TOC3_ID}` 表示。

每个一级目录条目（pStyle="{TOC1_ID}"）：

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="{TOC1_ID}"/>
    <w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>
    <w:spacing w:line="300" w:lineRule="auto"/>
  </w:pPr>
  <!-- 第一个条目还需 TOC 域开始 -->
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText>TOC \o "1-3" \h \z \u </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <!-- HYPERLINK 域 -->
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText> HYPERLINK \l _Toc{id} </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <!-- 标题文本 -->
  <w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>{标题}</w:t></w:r>
  <!-- 制表符 -->
  <w:r><w:tab/></w:r>
  <!-- PAGEREF 域 -->
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText> PAGEREF _Toc{id} \h </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>{页码}</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <!-- HYPERLINK 结束 -->
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
```

二级条目用 `pStyle="{TOC2_ID}"`，三级条目用 `pStyle="{TOC3_ID}"`，其余结构相同。

最后一个条目后加 TOC 域结束段落：
```xml
<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
```

**注意**：`\z` 标志在 Web Layout 视图中隐藏页码和制表符。模板和例文均使用此标志。

**制表位 pos 值**：必须从例文的 toc 1 样式定义中动态提取。论文中为 `pos="8777"`，不同文档可能不同。分析脚本应读取 styles.xml 中 toc 1 样式的 `<w:tabs>` 定义。

## relationships 和 Content_Types 更新

### document.xml.rels

每个 header/footer 文件需要一条 Relationship：

```xml
<Relationship Id="rId30"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
  Target="header7.xml"/>
```

### [Content_Types].xml

每个 header/footer 文件需要一个 Override：

```xml
<Override PartName="/word/header7.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
```

## rId 分配策略

建议从 rId30 开始分配，避免与例文已有的 rId 冲突：

```
rId30/rId31  → 封面 section 的 header/footer
rId32/rId33  → 摘要 section
rId34/rId35  → Abstract section
rId36/rId37  → 目录 section
rId38/rId39  → 引言 section（或第一章后的分节）
...依次递增...
最后两个 rId  → body 级 sectPr（最后一个 section）
```

header 文件命名从例文已有最大编号 +1 开始（如例文有 header1-header6，则从 header7.xml 开始），footer 同理。
