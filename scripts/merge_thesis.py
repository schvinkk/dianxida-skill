# -*- coding: utf-8 -*-
"""
Merge script v3: Section break + unique header for EVERY 一级标题.
Pattern (from example doc analysis):
  - Section break sits on the last paragraph of a section (or a blank para before the next heading)
  - The section break defines header/footer/pgnum for the section that ENDS there
  - The final body <w:sectPr> defines the last section's properties
"""
import xml.etree.ElementTree as ET
import copy, os, shutil, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
wns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# Register all namespaces to preserve them on write
for prefix, uri in [
    ('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'),
    ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('w14', 'http://schemas.microsoft.com/office/word/2010/wordml'),
    ('w15', 'http://schemas.microsoft.com/office/word/2012/wordml'),
    ('w10', 'urn:schemas-microsoft-com:office:word'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
    ('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'),
    ('wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk'),
    ('wne', 'http://schemas.microsoft.com/office/word/2006/wordml'),
    ('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'),
    ('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math'),
    ('o', 'urn:schemas-microsoft-com:office:office'),
    ('v', 'urn:schemas-microsoft-com:vml'),
    ('wpsCustomData', 'http://www.wps.cn/officeDocument/2013/wpsCustomData'),
]:
    ET.register_namespace(prefix, uri)


def analyze_styles(unpacked_dir):
    """从例文的 styles.xml 中动态查找样式名称到 ID 的映射。
    样式 ID 因文档而异，绝不能硬编码！"""
    styles_path = os.path.join(unpacked_dir, 'word', 'styles.xml')
    root = ET.parse(styles_path).getroot()
    style_map = {}
    for style in root.findall(f'{wns}style'):
        sid = style.get(f'{wns}styleId', '')
        name_el = style.find(f'{wns}name')
        name = name_el.get(f'{wns}val', '') if name_el is not None else ''
        if name:
            style_map[name] = sid
    # Print for verification
    print("=== 动态样式映射 ===")
    for key in ['heading 1', 'heading 2', 'heading 3', 'toc 1', 'toc 2', 'toc 3',
                'header', 'footer', 'Normal', 'Plain Text', 'Normal (Web)', 'Title']:
        print(f"  {key:>20s} → ID={style_map.get(key, 'NOT FOUND')}")
    return style_map


def get_toc_tab_pos(unpacked_dir, toc_style_name='toc 1'):
    """从 toc 样式定义中提取制表位 pos 值（动态，不硬编码）。
    论文中 toc 1 的 tab pos=8777，不同文档可能不同。"""
    styles_path = os.path.join(unpacked_dir, 'word', 'styles.xml')
    root = ET.parse(styles_path).getroot()
    for style in root.findall(f'{wns}style'):
        name_el = style.find(f'{wns}name')
        if name_el is not None and name_el.get(f'{wns}val') == toc_style_name:
            ppr = style.find(f'{wns}pPr')
            if ppr is not None:
                tabs = ppr.find(f'{wns}tabs')
                if tabs is not None:
                    tab = tabs.find(f'{wns}tab')
                    if tab is not None:
                        pos = tab.get(f'{wns}pos', '8306')
                        print(f"  TOC tab pos ({toc_style_name}): {pos}")
                        return pos
    print(f"  WARNING: TOC tab pos not found for '{toc_style_name}', using default 8306")
    return '8306'


def get_text(p):
    return ''.join(t.text or '' for t in p.findall('.//w:t', ns))


def get_style(p):
    ppr = p.find('w:pPr', ns)
    if ppr is not None:
        ps = ppr.find('w:pStyle', ns)
        if ps is not None:
            return ps.get(f'{wns}val', '')
    return ''


def get_jc(p):
    ppr = p.find('w:pPr', ns)
    if ppr is not None:
        jc = ppr.find('w:jc', ns)
        if jc is not None:
            return jc.get(f'{wns}val', '')
    return ''


def make_header_xml(title, header_style_id='9'):
    """Header: exact copy of example's header structure, just swap title text.
    header_style_id must be dynamically obtained from analyze_styles()."""
    ns_hdr = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
              'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
              'xmlns:o="urn:schemas-microsoft-com:office:office" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
              'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
              'xmlns:v="urn:schemas-microsoft-com:vml" '
              'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
              'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
              'xmlns:w10="urn:schemas-microsoft-com:office:word" '
              'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
              'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
              'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
              'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
              'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
              'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
              'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
              'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" '
              'mc:Ignorable="w14 w15 wp14"')
    if not title:
        # Blank header: header style, empty paragraph
        return (
            '<?xml version="1.0" encoding="utf-8"?>\n'
            f'<w:hdr {ns_hdr}>\n'
            '  <w:p>\n'
            '    <w:pPr>\n'
            f'      <w:pStyle w:val="{header_style_id}"/>\n'
            '    </w:pPr>\n'
            '  </w:p>\n'
            '</w:hdr>'
        )
    # Content header: exact copy of example header structure
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        f'<w:hdr {ns_hdr}>\n'
        '  <w:p>\n'
        '    <w:pPr>\n'
        f'      <w:pStyle w:val="{header_style_id}"/>\n'
        '      <w:pBdr>\n'
        '        <w:bottom w:val="single" w:color="auto" w:sz="4" w:space="1"/>\n'
        '      </w:pBdr>\n'
        '      <w:jc w:val="center"/>\n'
        '      <w:rPr>\n'
        '        <w:rFonts w:hint="default" w:eastAsia="宋体"/>\n'
        '        <w:lang w:val="en-US" w:eastAsia="zh-CN"/>\n'
        '      </w:rPr>\n'
        '    </w:pPr>\n'
        '    <w:r>\n'
        '      <w:rPr>\n'
        '        <w:rFonts w:hint="eastAsia"/>\n'
        '        <w:sz w:val="21"/>\n'
        '        <w:szCs w:val="21"/>\n'
        '        <w:lang w:val="en-US" w:eastAsia="zh-CN"/>\n'
        '      </w:rPr>\n'
        f'      <w:t>{title}</w:t>\n'
        '    </w:r>\n'
        '  </w:p>\n'
        '</w:hdr>'
    )


def make_footer_xml(footer_style_id='8'):
    """Footer: exact copy of example's footer1.xml — text-box based PAGE field.
    footer_style_id must be dynamically obtained from analyze_styles()."""
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" '
        'mc:Ignorable="w14 w15 wp14">\n'
        '  <w:p>\n'
        '    <w:pPr>\n'
        f'      <w:pStyle w:val="{footer_style_id}"/>\n'
        '    </w:pPr>\n'
        '    <w:r>\n'
        '      <w:rPr>\n'
        '        <w:sz w:val="18"/>\n'
        '      </w:rPr>\n'
        '      <mc:AlternateContent>\n'
        '        <mc:Choice Requires="wps">\n'
        '          <w:drawing>\n'
        '            <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">\n'
        '              <wp:simplePos x="0" y="0"/>\n'
        '              <wp:positionH relativeFrom="margin">\n'
        '                <wp:align>center</wp:align>\n'
        '              </wp:positionH>\n'
        '              <wp:positionV relativeFrom="paragraph">\n'
        '                <wp:posOffset>0</wp:posOffset>\n'
        '              </wp:positionV>\n'
        '              <wp:extent cx="1828800" cy="1828800"/>\n'
        '              <wp:effectExtent l="0" t="0" r="0" b="0"/>\n'
        '              <wp:wrapNone/>\n'
        '              <wp:docPr id="1" name="文本框 1"/>\n'
        '              <wp:cNvGraphicFramePr/>\n'
        '              <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        '                <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">\n'
        '                  <wps:wsp>\n'
        '                    <wps:cNvSpPr/>\n'
        '                    <wps:spPr>\n'
        '                      <a:xfrm>\n'
        '                        <a:off x="0" y="0"/>\n'
        '                        <a:ext cx="1828800" cy="1828800"/>\n'
        '                      </a:xfrm>\n'
        '                      <a:prstGeom prst="rect">\n'
        '                        <a:avLst/>\n'
        '                      </a:prstGeom>\n'
        '                      <a:ln w="12700">\n'
        '                        <a:noFill/>\n'
        '                      </a:ln>\n'
        '                    </wps:spPr>\n'
        '                    <wps:txbx>\n'
        '                      <w:txbxContent>\n'
        '                        <w:p>\n'
        '                          <w:pPr>\n'
        f'                            <w:pStyle w:val="{footer_style_id}"/>\n'
        '                          </w:pPr>\n'
        '                          <w:r>\n'
        '                            <w:fldChar w:fldCharType="begin"/>\n'
        '                          </w:r>\n'
        '                          <w:r>\n'
        '                            <w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText>\n'
        '                          </w:r>\n'
        '                          <w:r>\n'
        '                            <w:fldChar w:fldCharType="separate"/>\n'
        '                          </w:r>\n'
        '                          <w:r>\n'
        '                            <w:t>1</w:t>\n'
        '                          </w:r>\n'
        '                          <w:r>\n'
        '                            <w:fldChar w:fldCharType="end"/>\n'
        '                          </w:r>\n'
        '                        </w:p>\n'
        '                      </w:txbxContent>\n'
        '                    </wps:txbx>\n'
        '                    <wps:bodyPr rot="0" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0">\n'
        '                      <a:spAutoFit/>\n'
        '                    </wps:bodyPr>\n'
        '                  </wps:wsp>\n'
        '                </a:graphicData>\n'
        '              </a:graphic>\n'
        '            </wp:anchor>\n'
        '          </w:drawing>\n'
        '        </mc:Choice>\n'
        '        <mc:Fallback>\n'
        '          <w:pict>\n'
        '            <v:rect id="文本框 1" style="position:absolute;left:0pt;margin-top:0pt;'
        'height:144pt;width:144pt;mso-position-horizontal:center;'
        'mso-position-horizontal-relative:margin;mso-wrap-style:none;'
        'z-index:251659264;" filled="f" stroked="f">\n'
        '              <v:fill on="f" focussize="0,0"/>\n'
        '              <v:stroke on="f" weight="1pt"/>\n'
        '              <v:imagedata o:title=""/>\n'
        '              <o:lock v:ext="edit" aspectratio="f"/>\n'
        '              <v:textbox inset="0mm,0mm,0mm,0mm" style="mso-fit-shape-to-text:t;">\n'
        '                <w:txbxContent>\n'
        '                  <w:p>\n'
        '                    <w:pPr>\n'
        f'                      <w:pStyle w:val="{footer_style_id}"/>\n'
        '                    </w:pPr>\n'
        '                    <w:r>\n'
        '                      <w:fldChar w:fldCharType="begin"/>\n'
        '                    </w:r>\n'
        '                    <w:r>\n'
        '                      <w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText>\n'
        '                    </w:r>\n'
        '                    <w:r>\n'
        '                      <w:fldChar w:fldCharType="separate"/>\n'
        '                    </w:r>\n'
        '                    <w:r>\n'
        '                      <w:t>1</w:t>\n'
        '                    </w:r>\n'
        '                    <w:r>\n'
        '                      <w:fldChar w:fldCharType="end"/>\n'
        '                    </w:r>\n'
        '                  </w:p>\n'
        '                </w:txbxContent>\n'
        '              </v:textbox>\n'
        '            </v:rect>\n'
        '          </w:pict>\n'
        '        </mc:Fallback>\n'
        '      </mc:AlternateContent>\n'
        '    </w:r>\n'
        '  </w:p>\n'
        '</w:ftr>'
    )


def make_sectpr(header_rid, footer_rid, pgnum_fmt=None, pgnum_start=None):
    """Create a <w:sectPr> element."""
    sp = ET.Element(f'{wns}sectPr')
    if header_rid:
        hr = ET.SubElement(sp, f'{wns}headerReference')
        hr.set(f'{rns}id', header_rid)
        hr.set(f'{rns}type', 'default')
    if footer_rid:
        fr = ET.SubElement(sp, f'{wns}footerReference')
        fr.set(f'{rns}id', footer_rid)
        fr.set(f'{rns}type', 'default')
    # A4 page size
    pgSz = ET.SubElement(sp, f'{wns}pgSz')
    pgSz.set(f'{wns}w', '11906')
    pgSz.set(f'{wns}h', '16838')
    # Margins (matching example)
    pgMar = ET.SubElement(sp, f'{wns}pgMar')
    pgMar.set(f'{wns}top', '1417')
    pgMar.set(f'{wns}right', '1134')
    pgMar.set(f'{wns}bottom', '1417')
    pgMar.set(f'{wns}left', '1701')
    pgMar.set(f'{wns}header', '851')
    pgMar.set(f'{wns}footer', '992')
    pgMar.set(f'{wns}gutter', '0')
    # Page number format (only set when explicitly needed)
    if pgnum_fmt:
        pn = ET.SubElement(sp, f'{wns}pgNumType')
        pn.set(f'{wns}fmt', pgnum_fmt)
        if pgnum_start:
            pn.set(f'{wns}start', pgnum_start)
    # Columns
    cols = ET.SubElement(sp, f'{wns}cols')
    cols.set(f'{wns}space', '425')
    cols.set(f'{wns}num', '1')
    # Document grid
    dg = ET.SubElement(sp, f'{wns}docGrid')
    dg.set(f'{wns}type', 'lines')
    dg.set(f'{wns}linePitch', '312')
    dg.set(f'{wns}charSpace', '0')
    return sp


def set_sect_break(para, header_rid, footer_rid, pgnum_fmt=None, pgnum_start=None):
    """Set a section break on a paragraph's <w:pPr>."""
    ppr = para.find('w:pPr', ns)
    if ppr is None:
        ppr = ET.SubElement(para, f'{wns}pPr')
        para.remove(ppr)
        para.insert(0, ppr)
    existing = ppr.find('w:sectPr', ns)
    if existing is not None:
        ppr.remove(existing)
    sp = make_sectpr(header_rid, footer_rid, pgnum_fmt, pgnum_start)
    ppr.append(sp)


# TOC entry data: (text, level, toc_id, page_num_str)
toc_entries = [
    ('摘  要',                           1, 50001, 'I'),
    ('Abstract',                         1, 50002, 'II'),
    ('目  录',                           1, 50003, 'III'),
    ('1 引言',                           1, 50004, '1'),
    ('2 小檗碱概述',                     1, 50005, '2'),
    ('  2.1 化学结构与来源',             2, 50006, '2'),
    ('  2.2 药代动力学特征',             2, 50007, '3'),
    ('3 小檗碱治疗T2DM的药理机制',      1, 50008, '5'),
    ('  3.1 AMPK信号通路激活',           2, 50009, '5'),
    ('  3.2 肠道菌群调节',               2, 50010, '6'),
    ('  3.3 抗炎与抗氧化作用',           2, 50011, '7'),
    ('4 小檗碱治疗T2DM的临床研究进展',  1, 50012, '9'),
    ('  4.1 单药治疗研究',               2, 50013, '9'),
    ('  4.2 联合二甲双胍治疗',           2, 50014, '10'),
    ('  4.3 对并发症的改善作用',         2, 50015, '11'),
    ('5 小檗碱新剂型研究进展',           1, 50016, '13'),
    ('  5.1 纳米制剂',                   2, 50017, '13'),
    ('  5.2 脂质体与固体分散体',         2, 50018, '14'),
    ('6 小檗碱的安全性评价',             1, 50019, '16'),
    ('  6.1 急性与慢性毒性',             2, 50020, '16'),
    ('  6.2 药物相互作用',               2, 50021, '17'),
    ('7 总结与展望',                     1, 50022, '19'),
    ('参考文献',                         1, 50023, '21'),
    ('致  谢',                           1, 50024, '24'),
]

_toc_ppr_xml = (
    '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:pStyle w:val="{style}"/>'
    '<w:keepNext w:val="0"/><w:keepLines w:val="0"/>'
    '<w:pageBreakBefore w:val="0"/><w:widowControl w:val="0"/>'
    '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{tab_pos}"/></w:tabs>'
    '<w:kinsoku/><w:wordWrap/><w:overflowPunct/>'
    '<w:topLinePunct w:val="0"/><w:autoSpaceDE/><w:autoSpaceDN/>'
    '<w:bidi w:val="0"/><w:adjustRightInd/><w:snapToGrid/>'
    '<w:spacing w:line="300" w:lineRule="auto"/>'
    '<w:textAlignment w:val="auto"/>'
    '</w:pPr>'
)


def make_toc_entry(text, level, toc_id, page_num, is_first=False, toc1_style_id='10', toc2_style_id='11', tab_pos='8777'):
    """Create a TOC entry paragraph matching the example doc's field-based structure.
    toc1_style_id/toc2_style_id must be dynamically obtained from analyze_styles().
    tab_pos must be dynamically extracted from toc 1 style's tab definition."""
    style = toc1_style_id if level == 1 else toc2_style_id
    p = ET.Element(f'{wns}p')

    # pPr with dot leader tab (dynamic pos)
    ppr = ET.fromstring(_toc_ppr_xml.format(style=style, tab_pos=tab_pos))
    p.append(ppr)

    # First entry: TOC field begin
    if is_first:
        r_toc_begin = ET.SubElement(p, f'{wns}r')
        ET.SubElement(r_toc_begin, f'{wns}fldChar').set(f'{wns}fldCharType', 'begin')
        r_toc_instr = ET.SubElement(p, f'{wns}r')
        instr = ET.SubElement(r_toc_instr, f'{wns}instrText')
        instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instr.text = 'TOC \\o "1-3" \\h \\z \\u '
        r_toc_sep = ET.SubElement(p, f'{wns}r')
        ET.SubElement(r_toc_sep, f'{wns}fldChar').set(f'{wns}fldCharType', 'separate')

    # HYPERLINK field begin
    r_hl_begin = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r_hl_begin, f'{wns}fldChar').set(f'{wns}fldCharType', 'begin')
    r_hl_instr = ET.SubElement(p, f'{wns}r')
    hl_instr = ET.SubElement(r_hl_instr, f'{wns}instrText')
    hl_instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    hl_instr.text = f' HYPERLINK \\l _Toc{toc_id} '
    r_hl_sep = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r_hl_sep, f'{wns}fldChar').set(f'{wns}fldCharType', 'separate')

    # Text run with formatting
    r_text = ET.SubElement(p, f'{wns}r')
    rpr = ET.SubElement(r_text, f'{wns}rPr')
    rf = ET.SubElement(rpr, f'{wns}rFonts')
    rf.set(f'{wns}hint', 'eastAsia')
    lang = ET.SubElement(rpr, f'{wns}lang')
    lang.set(f'{wns}val', 'en-US')
    lang.set(f'{wns}eastAsia', 'zh-CN')
    t = ET.SubElement(r_text, f'{wns}t')
    t.text = text.strip() if text.startswith('  ') else text

    # Tab
    ET.SubElement(ET.SubElement(p, f'{wns}r'), f'{wns}tab')

    # PAGEREF field
    r_pr_begin = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r_pr_begin, f'{wns}fldChar').set(f'{wns}fldCharType', 'begin')
    r_pr_instr = ET.SubElement(p, f'{wns}r')
    pr_instr = ET.SubElement(r_pr_instr, f'{wns}instrText')
    pr_instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    pr_instr.text = f' PAGEREF _Toc{toc_id} \\h '
    r_pr_sep = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r_pr_sep, f'{wns}fldChar').set(f'{wns}fldCharType', 'separate')
    r_pr_num = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r_pr_num, f'{wns}t').text = page_num
    r_pr_end = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r_pr_end, f'{wns}fldChar').set(f'{wns}fldCharType', 'end')

    # HYPERLINK field end
    r_hl_end = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r_hl_end, f'{wns}fldChar').set(f'{wns}fldCharType', 'end')

    return p


def make_toc_end():
    """Create the TOC field end paragraph."""
    p = ET.Element(f'{wns}p')
    r = ET.SubElement(p, f'{wns}r')
    ET.SubElement(r, f'{wns}fldChar').set(f'{wns}fldCharType', 'end')
    return p


# =====================================================================
# Step 1: Copy example as base
# =====================================================================
src_dir = 'example_unpacked'
dst_dir = 'merged_unpacked'
if os.path.exists(dst_dir):
    shutil.rmtree(dst_dir)
shutil.copytree(src_dir, dst_dir)

# Step 0: Dynamically resolve style IDs (NEVER hardcode!)
style_map = analyze_styles(src_dir)
H1_ID = style_map.get('heading 1', '2')
HEADER_ID = style_map.get('header', '9')
FOOTER_ID = style_map.get('footer', '8')
TOC1_ID = style_map.get('toc 1', '10')
TOC2_ID = style_map.get('toc 2', '11')
TOC_TAB_POS = get_toc_tab_pos(src_dir, 'toc 1')
print(f"\nResolved: heading1={H1_ID}, header={HEADER_ID}, footer={FOOTER_ID}, toc1={TOC1_ID}, toc2={TOC2_ID}, tab_pos={TOC_TAB_POS}\n")

# =====================================================================
# Step 2: Parse documents
# =====================================================================
ex_tree = ET.parse(f'{src_dir}/word/document.xml')
ex_root = ex_tree.getroot()
ex_body = ex_root.find('.//w:body', ns)
ex_paras = ex_body.findall('w:p', ns)

rv_tree = ET.parse('review_unpacked/word/document.xml')
rv_root = rv_tree.getroot()
rv_body = rv_root.find('.//w:body', ns)
rv_paras = rv_body.findall('w:p', ns)

ex_all_children = list(ex_body)
ex_sectpr = ex_body.find('w:sectPr', ns)

# =====================================================================
# Step 3: Build front matter (example paras 0-35)
# =====================================================================
ex_keep_until = ex_paras[36]
ex_para_to_child = {}
new_body_children = []
for child in ex_all_children:
    if child is ex_keep_until:
        break
    if child.tag == f'{wns}p':
        idx = ex_paras.index(child) if child in ex_paras else -1
        if idx >= 0:
            ex_para_to_child[idx] = child
    new_body_children.append(child)

# Clear cover page info fields (user fills these)
cover_info_fields = {
    13: '学        院：                                   ',
    14: '专        业：                                   ',
    15: '班        级：                                   ',
    16: '姓        名：                                   ',
    17: '学        号：                                   ',
    18: '指导教师（职称）：                                   ',
    19: '起 止 日 期 ：                                   ',
}
for idx, new_text in cover_info_fields.items():
    if idx in ex_para_to_child:
        p = ex_para_to_child[idx]
        texts = p.findall('.//w:t', ns)
        if texts:
            texts[0].text = new_text
            for t in texts[1:]:
                t.text = ''

# Set actual titles on cover page (paras 9, 10 — example's title positions)
cover_titles = {
    9:  ('小檗碱治疗2型糖尿病的研究进展综述', None),
    10: ('Research Progress on Berberine in the Treatment of Type 2 Diabetes Mellitus', 'Times New Roman'),
}
for idx, (title_text, ascii_font) in cover_titles.items():
    if idx in ex_para_to_child:
        p = ex_para_to_child[idx]
        for r in p.findall('w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                rpr = ET.SubElement(r, f'{wns}rPr')
                r.remove(rpr)
                r.insert(0, rpr)
            # Add bold
            if rpr.find('w:b', ns) is None:
                ET.SubElement(rpr, f'{wns}b')
            # Set font size (三号 = sz=32)
            sz_el = rpr.find('w:sz', ns)
            if sz_el is None:
                sz_el = ET.SubElement(rpr, f'{wns}sz')
            sz_el.set(f'{wns}val', '32')
            szCs = rpr.find('w:szCs', ns)
            if szCs is None:
                szCs = ET.SubElement(rpr, f'{wns}szCs')
            szCs.set(f'{wns}val', '32')
            # Set ASCII font if needed
            if ascii_font:
                rfonts = rpr.find('w:rFonts', ns)
                if rfonts is None:
                    rfonts = ET.SubElement(rpr, f'{wns}rFonts')
                rfonts.set(f'{wns}ascii', ascii_font)
                rfonts.set(f'{wns}hAnsi', ascii_font)
        # Set text
        texts = p.findall('.//w:t', ns)
        if texts:
            texts[0].text = title_text
            for t in texts[1:]:
                t.text = ''

# Fix para 35 section break: replace its properties with 摘要 section properties
# This prevents double section break (para 35 from example + para 36 from our insertion)
# which would create an unwanted blank page between declarations and 摘要
if 35 in ex_para_to_child:
    p35 = ex_para_to_child[35]
    sp35 = p35.find('w:pPr/w:sectPr', ns)
    if sp35 is not None:
        # Remove existing header/footer refs and pgNumType
        for tag_name in ['headerReference', 'footerReference', 'pgNumType']:
            for elem in list(sp35.findall(f'w:{tag_name}', ns)):
                sp35.remove(elem)
        # Add footer rId31 (摘要 section footer)
        fr = ET.SubElement(sp35, f'{wns}footerReference')
        fr.set(f'{rns}id', 'rId31')
        fr.set(f'{rns}type', 'default')
        # Set upperRoman start=1
        pn = ET.SubElement(sp35, f'{wns}pgNumType')
        pn.set(f'{wns}fmt', 'upperRoman')
        pn.set(f'{wns}start', '1')

# =====================================================================
# Step 4: Define section breaks
# =====================================================================
# Section end definitions.
# Each entry: (next_start, header_title, hdr_rid, ftr_rid, pgfmt, pgstart)
# When we encounter a paragraph with orig_i == next_start, the section that
# ENDS at the previous paragraph gets this sectPr.  The header/footer IDs
# belong to the ENDING section (matching the example doc's pattern exactly).
sect_end_defs = [
    # ---- front-matter transitions ----
    # Cover section ends → before 摘要
    (50,  '',   'rId30', 'rId31', 'upperRoman', '1'),
    # 摘要 section ends → before Abstract
    (54,  '',  'rId32', 'rId33', 'upperRoman', '1'),
    # Abstract section ends → before 目录
    (58,  '',   'rId34', 'rId35', 'upperRoman', None),
    # ---- body chapters (decimal) ----
    # 目录 section ends → before 引言
    (83,  '',   'rId36', 'rId37', 'decimal', '1'),
    # 引言 section ends → before 小檗碱概述
    (87,  '引言',   'rId38', 'rId39', None, None),
    # 小檗碱概述 ends → before 药理机制
    (94,  '小檗碱概述',   'rId40', 'rId41', None, None),
    # 药理机制 ends → before 临床研究
    (104, '小檗碱治疗T2DM的药理机制',   'rId42', 'rId43', None, None),
    # 临床研究 ends → before 新剂型
    (114, '小檗碱治疗T2DM的临床研究进展',   'rId44', 'rId45', None, None),
    # 新剂型 ends → before 安全性
    (120, '小檗碱新剂型研究进展',   'rId46', 'rId47', None, None),
    # 安全性 ends → before 总结
    (126, '小檗碱的安全性评价',   'rId48', 'rId49', None, None),
    # 总结 ends → before 参考文献
    (130, '总结与展望',   'rId50', 'rId51', None, None),
    # 参考文献 ends → before 致谢
    (215, '参考文献', 'rId52', 'rId53', None, None),
    # Note: 致谢 is the LAST section → defined by body-level sectPr (rId54/rId55)
]

# Build lookup: next_start → sect_end_def (for O(1) lookup in the loop)
sect_end_map = {}
for sd in sect_end_defs:
    sect_end_map[sd[0]] = sd
sect_end_seen = set()

# =====================================================================
# Step 5: Process review paragraphs (paras 50+)
# =====================================================================
skip_indices = set()
for i in range(41, 50):     # template instruction block
    skip_indices.add(i)
for i in range(59, 82):     # old manual TOC entries (replaced with field-based)
    skip_indices.add(i)
for i in range(172, 215):   # duplicate formal 参考文献 section
    skip_indices.add(i)
skip_indices.add(218)        # incomplete trailing para

review_content = []  # list of (orig_i, paragraph_element)

for i, p in enumerate(rv_paras):
    if i < 50:
        continue
    if i in skip_indices:
        continue

    p_copy = copy.deepcopy(p)
    text = get_text(p_copy)
    style = get_style(p_copy)

    # ---- Strip ALL existing section breaks from review content ----
    for sp in list(p_copy.findall('.//w:sectPr', ns)):
        parent = sp.getparent() if hasattr(sp, 'getparent') else None
        # ElementTree doesn't have getparent, so find parent manually
        pass
    # Find sectPr in pPr
    ppr = p_copy.find('w:pPr', ns)
    if ppr is not None:
        sp = ppr.find('w:sectPr', ns)
        if sp is not None:
            ppr.remove(sp)

    # ---- Apply section break to the LAST para of the ending section ----
    if i in sect_end_map and i not in sect_end_seen:
        sect_end_seen.add(i)
        _, hdr_title, hdr_rid, ftr_rid, pgfmt, pgstart = sect_end_map[i]
        # Skip 摘要 (para 50): its break is already on front matter para 35
        if i != 50 and review_content:
            # Find the last paragraph in review_content to apply sectPr to.
            # Prefer a blank paragraph (separator), fall back to last content para.
            target_idx = len(review_content) - 1
            while target_idx >= 0:
                ti, tp = review_content[target_idx]
                if ti != -1:  # skip synthetic TOC entries
                    break
                target_idx -= 1
            if target_idx >= 0:
                _, target_para = review_content[target_idx]
                set_sect_break(target_para, hdr_rid, ftr_rid, pgfmt, pgstart)

    # ---- Text / style fixes ----

    # 摘要 → 摘  要
    if i == 50 and style == H1_ID:
        for t in p_copy.findall('.//w:t', ns):
            if t.text and '摘要' in t.text:
                t.text = t.text.replace('摘要', '摘  要')

    # Abstract → give it style [2] (heading 1)
    if text.strip() == 'Abstract' and style == '' and get_jc(p_copy) == 'center':
        ppr2 = p_copy.find('w:pPr', ns)
        if ppr2 is None:
            ppr2 = ET.SubElement(p_copy, f'{wns}pPr')
            p_copy.remove(ppr2)
            p_copy.insert(0, ppr2)
        pstyle = ppr2.find('w:pStyle', ns)
        if pstyle is None:
            pstyle = ET.SubElement(ppr2, f'{wns}pStyle')
        pstyle.set(f'{wns}val', H1_ID)
        jc = ppr2.find('w:jc', ns)
        if jc is not None:
            ppr2.remove(jc)
        for r in p_copy.findall('w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is not None:
                for tag in ['b', 'sz', 'szCs']:
                    elem = rpr.find(f'w:{tag}', ns)
                    if elem is not None:
                        rpr.remove(elem)

    # 目录 → 目  录
    if text.strip() == '目录' and get_jc(p_copy) == 'center':
        for t in p_copy.findall('.//w:t', ns):
            if t.text and '目录' in t.text:
                t.text = t.text.replace('目录', '目  录')
        for r in p_copy.findall('w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                rpr = ET.SubElement(r, f'{wns}rPr')
                r.remove(rpr)
                r.insert(0, rpr)
            rfonts = rpr.find('w:rFonts', ns)
            if rfonts is None:
                rfonts = ET.SubElement(rpr, f'{wns}rFonts')
            rfonts.set(f'{wns}ascii', '黑体')
            rfonts.set(f'{wns}hAnsi', '黑体')
            rfonts.set(f'{wns}eastAsia', '黑体')
            rfonts.set(f'{wns}cs', '黑体')

    # 致谢 → 致  谢
    if text.strip() == '致谢' and style == H1_ID:
        for t in p_copy.findall('.//w:t', ns):
            if t.text and '致谢' in t.text:
                t.text = t.text.replace('致谢', '致  谢')

    # 参考文献 inline heading → style [2]
    if i == 130 and text.strip() == '参考文献' and style == '':
        ppr2 = p_copy.find('w:pPr', ns)
        if ppr2 is None:
            ppr2 = ET.SubElement(p_copy, f'{wns}pPr')
            p_copy.remove(ppr2)
            p_copy.insert(0, ppr2)
        pstyle = ppr2.find('w:pStyle', ns)
        if pstyle is None:
            pstyle = ET.SubElement(ppr2, f'{wns}pStyle')
        pstyle.set(f'{wns}val', H1_ID)

    # Reference entries (paras 131-170) → 五号 (sz=21), no indent
    if 131 <= i <= 170:
        for r in p_copy.findall('w:r', ns):
            rpr = r.find('w:rPr', ns)
            if rpr is None:
                rpr = ET.SubElement(r, f'{wns}rPr')
                r.remove(rpr)
                r.insert(0, rpr)
            sz = rpr.find('w:sz', ns)
            if sz is None:
                sz = ET.SubElement(rpr, f'{wns}sz')
            sz.set(f'{wns}val', '21')
            szCs = rpr.find('w:szCs', ns)
            if szCs is None:
                szCs = ET.SubElement(rpr, f'{wns}szCs')
            szCs.set(f'{wns}val', '21')
            rfonts = rpr.find('w:rFonts', ns)
            if rfonts is None:
                rfonts = ET.SubElement(rpr, f'{wns}rFonts')
            rfonts.set(f'{wns}cs', 'Times New Roman')
        ppr3 = p_copy.find('w:pPr', ns)
        if ppr3 is not None:
            ind = ppr3.find('w:ind', ns)
            if ind is not None:
                ind.set(f'{wns}firstLine', '0')
                ind.set(f'{wns}firstLineChars', '0')
                ind.set(f'{wns}left', '0')
                ind.set(f'{wns}leftChars', '0')

    # Style [2] headings: remove jc=center (style definition handles it)
    if style == H1_ID:
        ppr4 = p_copy.find('w:pPr', ns)
        if ppr4 is not None:
            jc = ppr4.find('w:jc', ns)
            if jc is not None:
                ppr4.remove(jc)

    review_content.append((i, p_copy))

    # After the 目录 heading (para 58), insert field-based TOC entries
    if i == 58:
        for toc_idx, (text, level, toc_id, page_num) in enumerate(toc_entries):
            toc_p = make_toc_entry(text, level, toc_id, page_num, is_first=(toc_idx == 0),
                                    toc1_style_id=TOC1_ID, toc2_style_id=TOC2_ID, tab_pos=TOC_TAB_POS)
            review_content.append((-1, toc_p))
        # Add TOC field end paragraph
        review_content.append((-1, make_toc_end()))

# =====================================================================
# Step 6: Assemble full body
# =====================================================================
# Front matter + review content
for _, p in review_content:
    new_body_children.append(p)

# Final body sectPr → defines 致谢 section (last section)
final_sp = make_sectpr('rId54', 'rId55', 'decimal', None)
new_body_children.append(final_sp)

# =====================================================================
# Step 7: Write header/footer XML files
# =====================================================================
# Header map: rId → (filename, title)
header_map = {
    'rId30': ('header7.xml',  ''),           # cover section (blank)
    'rId32': ('header8.xml',  ''),           # 摘要 section (blank)
    'rId34': ('header9.xml',  ''),           # Abstract section (blank)
    'rId36': ('header10.xml', ''),           # 目录 section (blank)
    'rId38': ('header11.xml', '引言'),
    'rId40': ('header12.xml', '小檗碱概述'),
    'rId42': ('header13.xml', '小檗碱治疗T2DM的药理机制'),
    'rId44': ('header14.xml', '小檗碱治疗T2DM的临床研究进展'),
    'rId46': ('header15.xml', '小檗碱新剂型研究进展'),
    'rId48': ('header16.xml', '小檗碱的安全性评价'),
    'rId50': ('header17.xml', '总结与展望'),
    'rId52': ('header18.xml', '参考文献'),
    'rId54': ('header19.xml', '致  谢'),     # final body = 致谢
}

for rid, (fname, title) in header_map.items():
    with open(f'{dst_dir}/word/{fname}', 'w', encoding='utf-8') as f:
        f.write(make_header_xml(title, header_style_id=HEADER_ID))

# Footer map: rId → filename (all identical PAGE field)
footer_files = {
    'rId31': 'footer7.xml',
    'rId33': 'footer8.xml',
    'rId35': 'footer9.xml',
    'rId37': 'footer10.xml',
    'rId39': 'footer11.xml',
    'rId41': 'footer12.xml',
    'rId43': 'footer13.xml',
    'rId45': 'footer14.xml',
    'rId47': 'footer15.xml',
    'rId49': 'footer16.xml',
    'rId51': 'footer17.xml',
    'rId53': 'footer18.xml',
    'rId55': 'footer19.xml',
}

for rid, fname in footer_files.items():
    with open(f'{dst_dir}/word/{fname}', 'w', encoding='utf-8') as f:
        f.write(make_footer_xml(footer_style_id=FOOTER_ID))

print(f"Written {len(header_map)} headers, {len(footer_files)} footers")

# =====================================================================
# Step 8: Update relationships file
# =====================================================================
rels_path = f'{dst_dir}/word/_rels/document.xml.rels'
rels_tree = ET.parse(rels_path)
rels_root = rels_tree.getroot()
rns_uri = 'http://schemas.openxmlformats.org/package/2006/relationships'

new_rels = []
# Headers
for rid, (fname, _) in header_map.items():
    new_rels.append((rid, 'header', fname))
# Footers
for rid, fname in footer_files.items():
    new_rels.append((rid, 'footer', fname))

# Get existing rIds
existing_rids = set()
for rel in rels_root:
    existing_rids.add(rel.get('Id', ''))

for rid, rtype, target in new_rels:
    if rid in existing_rids:
        # Update existing
        for rel in rels_root:
            if rel.get('Id') == rid:
                rel.set('Type', f'http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rtype}')
                rel.set('Target', target)
                break
    else:
        rel = ET.SubElement(rels_root, f'{{{rns_uri}}}Relationship')
        rel.set('Id', rid)
        rel.set('Type', f'http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rtype}')
        rel.set('Target', target)

ET.register_namespace('', rns_uri)
rels_tree.write(rels_path, xml_declaration=True, encoding='UTF-8')
print(f"Rels updated: {len(new_rels)} new/updated relationships")

# =====================================================================
# Step 9: Update Content_Types.xml
# =====================================================================
ct_path = f'{dst_dir}/[Content_Types].xml'
ct_tree = ET.parse(ct_path)
ct_root = ct_tree.getroot()
ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'

existing_parts = set()
for override in ct_root.findall(f'{{{ct_ns}}}Override'):
    existing_parts.add(override.get('PartName', ''))

for i in range(7, 20):
    for prefix, ctype in [
        ('header', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'),
        ('footer', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'),
    ]:
        part_name = f'/word/{prefix}{i}.xml'
        if part_name not in existing_parts:
            ov = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
            ov.set('PartName', part_name)
            ov.set('ContentType', ctype)

ET.register_namespace('', ct_ns)
ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8')
print("Content_Types updated")

# =====================================================================
# Step 10: Rebuild body and write document.xml
# =====================================================================
for child in list(ex_body):
    ex_body.remove(child)
for child in new_body_children:
    ex_body.append(child)

# Serialize with namespace fix
xml_str = io.StringIO()
ex_tree.write(xml_str, xml_declaration=False, encoding='unicode')
xml_content = xml_str.getvalue()

root_tag_start = xml_content.find('<w:document')
root_tag_end = xml_content.find('>', root_tag_start)
root_tag = xml_content[root_tag_start:root_tag_end + 1]

ns_additions = []
if 'xmlns:wp14=' not in root_tag:
    ns_additions.append('xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"')
if 'xmlns:w15=' not in root_tag:
    ns_additions.append('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')
if 'xmlns:w10=' not in root_tag:
    ns_additions.append('xmlns:w10="urn:schemas-microsoft-com:office:word"')
if 'xmlns:wpsCustomData=' not in root_tag:
    ns_additions.append('xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData"')
if 'mc:Ignorable=' not in root_tag:
    ns_additions.append('mc:Ignorable="w14 w15 wp14"')

if ns_additions:
    insert_pos = root_tag_end
    new_root_tag = root_tag[:insert_pos] + ' ' + ' '.join(ns_additions) + root_tag[insert_pos:]
    xml_content = xml_content[:root_tag_start] + new_root_tag + xml_content[root_tag_end + 1:]

with open(f'{dst_dir}/word/document.xml', 'w', encoding='utf-8') as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    f.write(xml_content)

print("\n=== Merge v3 complete! ===")

# Print structure summary
result_paras = ex_body.findall('w:p', ns)
print(f"Total paragraphs: {len(result_paras)}")
sect_count = 0
for i, p in enumerate(result_paras):
    text = get_text(p)[:80]
    style = get_style(p)
    sp = p.find('.//w:sectPr', ns)
    if sp is not None:
        sect_count += 1
        hdr_refs = sp.findall('w:headerReference', ns)
        ftr_refs = sp.findall('w:footerReference', ns)
        pgnum = sp.find('w:pgNumType', ns)
        pg_info = ''
        if pgnum is not None:
            pg_info = f' pgnum={pgnum.get(f"{wns}fmt", "?")}/{pgnum.get(f"{wns}start", "?")}'
        hdr_info = ','.join(h.get(f'{rns}id', '?') for h in hdr_refs)
        ftr_info = ','.join(f2.get(f'{rns}id', '?') for f2 in ftr_refs)
        print(f"  [{sect_count}] para {i}: SECTBREAK hdr=[{hdr_info}] ftr=[{ftr_info}]{pg_info}")
    if style == H1_ID or (text.strip() and text.strip() in ['摘  要', 'Abstract', '目  录', '参考文献', '致  谢']):
        print(f"  para {i}: [{style}] {text}")

# Also check final sectPr
final = ex_body.find('w:sectPr', ns)
if final is not None:
    sect_count += 1
    hdr_refs = final.findall('w:headerReference', ns)
    hdr_info = ','.join(h.get(f'{rns}id', '?') for h in hdr_refs)
    print(f"  [{sect_count}] FINAL body sectPr hdr=[{hdr_info}]")
print(f"Total sections: {sect_count}")
