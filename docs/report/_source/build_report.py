# -*- coding: utf-8 -*-
"""Build the reference-format report DOCX."""
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE

NAVY = RGBColor(0x17, 0x36, 0x5D)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x59, 0x59, 0x59)
MIDGRAY = RGBColor(0x7F, 0x7F, 0x7F)
LINK = RGBColor(0x1F, 0x4E, 0x79)

BODY_FONT = "宋体"
HEAD_FONT = "黑体"
LATIN = "Times New Roman"
CONTENT_W = 15.4  # cm


def set_run_font(run, size=12, bold=False, color=None, east=BODY_FONT, latin=LATIN,
                 italic=False, superscript=False, spacing=None):
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if superscript:
        run.font.superscript = True
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), east)
    rFonts.set(qn("w:cs"), latin)
    if spacing is not None:
        drop_children(rPr, "w:spacing")
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(spacing))
        rPr.append(sp)


def set_para_spacing(p, before=0, after=0, line=1.5, first_indent=None,
                     left_indent=None, keep_next=False, keep_together=False):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line
    if first_indent is not None:
        pf.first_line_indent = Cm(first_indent)
    if left_indent is not None:
        pf.left_indent = Cm(left_indent)
    pf.keep_with_next = keep_next
    pf.keep_together = keep_together
    pf.widow_control = True




def drop_children(parent, tag):
    for el in parent.findall(qn(tag)):
        parent.remove(el)

CITE_RE = re.compile(r"\[\[(\d+(?:\s*[-,]\s*\d+)*)\]\]")


def add_rich_text(p, text, size=12, bold=False, color=None, east=BODY_FONT,
                  latin=LATIN, parse_cite=True):
    pos = 0
    if parse_cite:
        for m in CITE_RE.finditer(text):
            if m.start() > pos:
                r = p.add_run(text[pos:m.start()])
                set_run_font(r, size=size, bold=bold, color=color, east=east, latin=latin)
            r = p.add_run("[" + m.group(1).replace(" ", "") + "]")
            set_run_font(r, size=size, bold=False, color=color, east=east, latin=latin,
                         superscript=True)
            pos = m.end()
    if pos < len(text):
        r = p.add_run(text[pos:])
        set_run_font(r, size=size, bold=bold, color=color, east=east, latin=latin)


def add_body(doc, text, size=12, indent=True, before=0, after=6, line=1.5, color=None,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align
    set_para_spacing(p, before=before, after=after, line=line,
                     first_indent=(0.74 if indent else 0), keep_together=False)
    add_rich_text(p, text, size=size, color=color)
    return p


def add_bullet(doc, text, size=12, marker="•", indent=0.62, hang=0.42):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.left_indent = Cm(indent)
    pf.first_line_indent = Cm(-hang)
    set_para_spacing(p, before=0, after=4, line=1.4)
    r = p.add_run(marker + " ")
    set_run_font(r, size=size, color=NAVY)
    add_rich_text(p, text, size=size)
    return p


def add_label_para(doc, label, text, size=12, before=2, after=4):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.left_indent = Cm(0.62)
    set_para_spacing(p, before=before, after=after, line=1.45)
    r = p.add_run(label)
    set_run_font(r, size=size, bold=True, color=NAVY, east=HEAD_FONT)
    add_rich_text(p, text, size=size)
    return p


def add_h(doc, text, level=1, new_page=False):
    p = doc.add_paragraph(style=f"Heading {level}")
    if new_page:
        p.paragraph_format.page_break_before = True
    run = p.add_run(text)
    if level == 1:
        set_run_font(run, size=16, bold=False, color=DARK, east=HEAD_FONT)
        set_para_spacing(p, before=0 if new_page else 18, after=12, line=1.35, keep_next=True)
    elif level == 2:
        set_run_font(run, size=14, bold=False, color=DARK, east=HEAD_FONT)
        set_para_spacing(p, before=14, after=8, line=1.35, keep_next=True)
    else:
        set_run_font(run, size=12, bold=True, color=DARK, east=HEAD_FONT)
        set_para_spacing(p, before=10, after=5, line=1.35, keep_next=True)
    return p


def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    drop_children(tcPr, "w:shd")
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def cell_margins(cell, top=70, left=110, bottom=70, right=110):
    tcPr = cell._tc.get_or_add_tcPr()
    drop_children(tcPr, "w:tcMar")
    mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def table_borders(table, color="C9D2DE", size=4):
    tblPr = table._tbl.tblPr
    drop_children(tblPr, "w:tblBorders")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def set_cell(cell, text, size=10, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT,
             color=None, parse_cite=False, east=BODY_FONT, valign=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = valign
    cell_margins(cell)
    p = cell.paragraphs[0]
    p.alignment = align
    set_para_spacing(p, before=1, after=1, line=1.15, first_indent=0)
    add_rich_text(p, text, size=size, bold=bold, color=color, east=east, parse_cite=parse_cite)


def add_table(doc, headers, rows, widths, font_size=10, header_fill="E6EDF5",
              first_col_center=False, zebra=True, caption=None, note=None):
    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(cp, before=6, after=4, line=1.2, keep_next=True)
        r = cp.add_run(caption)
        set_run_font(r, size=10.5, bold=True, color=DARK, east=HEAD_FONT)
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"
    table_borders(table)
    tblPr = table._tbl.tblPr
    drop_children(tblPr, "w:tblLayout")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    drop_children(tblPr, "w:tblW")
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(sum(widths) * 567)))
    tblPr.append(tblW)

    hdr = table.rows[0]
    trPr = hdr._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    trPr.append(th)
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.width = Cm(widths[i])
        shade(c, header_fill)
        set_cell(c, h, size=font_size, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, east=HEAD_FONT)

    for ridx, row in enumerate(rows):
        tr = table.add_row()
        trPr = tr._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        trPr.append(cs)
        for i, val in enumerate(row):
            c = tr.cells[i]
            c.width = Cm(widths[i])
            if zebra and ridx % 2 == 1:
                shade(c, "F7F9FC")
            al = WD_ALIGN_PARAGRAPH.CENTER if (first_col_center and i == 0) else WD_ALIGN_PARAGRAPH.LEFT
            set_cell(c, str(val), size=font_size, align=al, parse_cite=True)

    for i, w in enumerate(widths):
        table.columns[i].width = Cm(w)
    if note:
        np_ = doc.add_paragraph()
        np_.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_para_spacing(np_, before=3, after=8, line=1.2)
        r = np_.add_run(note)
        set_run_font(r, size=9, color=GRAY)
    else:
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.line_spacing = Pt(1)
        _r = sp.add_run(" ")
        _r.font.size = Pt(1)
    return table


def add_callout(doc, title, lines, fill="F2F6FA", width=CONTENT_W, title_color=NAVY):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    drop_children(tblPr, "w:tblBorders")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        if edge == "left":
            el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "18"); el.set(qn("w:color"), "17365D")
        else:
            el.set(qn("w:val"), "nil")
        el.set(qn("w:space"), "0")
        borders.append(el)
    tblPr.append(borders)
    drop_children(tblPr, "w:tblW")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:type"), "dxa"); tblW.set(qn("w:w"), str(int(width * 567)))
    tblPr.append(tblW)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    cs = OxmlElement("w:cantSplit")
    trPr.append(cs)
    cell = t.rows[0].cells[0]
    cell.width = Cm(width)
    shade(cell, fill)
    cell_margins(cell, top=170, left=240, bottom=150, right=200)
    p0 = cell.paragraphs[0]
    set_para_spacing(p0, before=0, after=4, line=1.25)
    r = p0.add_run(title)
    set_run_font(r, size=11, bold=True, color=title_color, east=HEAD_FONT)
    for line in lines:
        p = cell.add_paragraph()
        set_para_spacing(p, before=0, after=2, line=1.35)
        add_rich_text(p, line, size=10.5, color=DARK)
    sp = doc.add_paragraph()
    set_para_spacing(sp, before=0, after=0, line=0.8)
    return t


def add_caption_below(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(p, before=2, after=10, line=1.2)
    r = p.add_run(text)
    set_run_font(r, size=9, color=GRAY)


def add_field(paragraph, instr, placeholder="1", size=10.5, color=GRAY):
    run = paragraph.add_run()
    set_run_font(run, size=size, color=color)
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText"); instr_el.set(qn("xml:space"), "preserve"); instr_el.text = f" {instr} "
    fld_sep = OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t"); txt.text = placeholder
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin); run._r.append(instr_el); run._r.append(fld_sep)
    run._r.append(txt); run._r.append(fld_end)


def set_numbering(section, fmt="decimal", start=1):
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sectPr.append(pg)
    pg.set(qn("w:fmt"), fmt)
    pg.set(qn("w:start"), str(start))


def add_header_footer(section, header_text=None, page_fmt=None, prefix="第 ", suffix=" 页"):
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    if header_text:
        hp = section.header.paragraphs[0]
        hp.text = ""
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(hp, before=0, after=0, line=1.0)
        r = hp.add_run(header_text)
        set_run_font(r, size=9, color=MIDGRAY)
        pPr = hp._p.get_or_add_pPr()
        drop_children(pPr, "w:pBdr")
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:space"), "1"); bottom.set(qn("w:color"), "BFC9D6")
        pbdr.append(bottom); pPr.append(pbdr)
    if page_fmt:
        fp = section.footer.paragraphs[0]
        fp.text = ""
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(fp, before=0, after=0, line=1.0)
        r = fp.add_run(prefix)
        set_run_font(r, size=10.5, color=GRAY)
        add_field(fp, "PAGE")
        r = fp.add_run(suffix)
        set_run_font(r, size=10.5, color=GRAY)


def add_hyperlink(paragraph, url, text=None, size=10.5):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), LATIN); rFonts.set(qn("w:hAnsi"), LATIN)
    rFonts.set(qn("w:eastAsia"), BODY_FONT)
    rPr.append(rFonts)
    color = OxmlElement("w:color"); color.set(qn("w:val"), "1F4E79"); rPr.append(color)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2))); rPr.append(sz)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement("w:t"); t.text = text or url; t.set(qn("xml:space"), "preserve")
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_ref_entry(doc, num, text, url=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.95)
    pf.first_line_indent = Cm(-0.95)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_para_spacing(p, before=0, after=3, line=1.18)
    pf.keep_together = True
    r = p.add_run(f"[{num}] ")
    set_run_font(r, size=10, bold=False, color=DARK)
    if url and url in text:
        pre, _, post = text.partition(url)
        if pre:
            r = p.add_run(pre)
            set_run_font(r, size=10)
        add_hyperlink(p, url, url + ("" if post == "" else ""), size=10)
        if post:
            r = p.add_run(post)
            set_run_font(r, size=10)
    else:
        r = p.add_run(text)
        set_run_font(r, size=10)
    return p

# ============================== DOCUMENT SETUP ==============================
doc = Document()
doc.core_properties.title = "参考文献著录与正文引用规范"
doc.core_properties.subject = "顺序编码制操作指南与个人文档智能整理课题文献清单"
doc.core_properties.author = "毕业设计（论文）写作规范专题报告"
doc.core_properties.keywords = "GB/T 7714; 顺序编码制; 参考文献; 毕业设计"

sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.54)
sec.bottom_margin = Cm(2.54)
sec.left_margin = Cm(3.0)
sec.right_margin = Cm(2.6)
sec.header_distance = Cm(1.5)
sec.footer_distance = Cm(1.6)

# Normal style
normal = doc.styles["Normal"]
normal.font.name = LATIN
normal.font.size = Pt(12)
normal.font.color.rgb = DARK
rPr = normal.element.get_or_add_rPr()
rFonts = rPr.find(qn("w:rFonts"))
if rFonts is None:
    rFonts = OxmlElement("w:rFonts"); rPr.insert(0, rFonts)
rFonts.set(qn("w:ascii"), LATIN)
rFonts.set(qn("w:hAnsi"), LATIN)
rFonts.set(qn("w:eastAsia"), BODY_FONT)
rFonts.set(qn("w:cs"), LATIN)

# Heading styles: make them black and Chinese
for lvl, sz in ((1, 16), (2, 14), (3, 12)):
    st = doc.styles[f"Heading {lvl}"]
    st.font.name = LATIN
    st.font.size = Pt(sz)
    st.font.bold = (lvl == 3)
    st.font.color.rgb = DARK
    rp = st.element.get_or_add_rPr()
    rf = rp.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rp.insert(0, rf)
    rf.set(qn("w:ascii"), LATIN)
    rf.set(qn("w:hAnsi"), LATIN)
    rf.set(qn("w:eastAsia"), HEAD_FONT)
    rf.set(qn("w:cs"), LATIN)

# Title style cleanup
tstyle = doc.styles["Title"]
tstyle.font.name = LATIN
tstyle.font.size = Pt(28)
tstyle.font.bold = True
tstyle.font.color.rgb = DARK
tstyle.paragraph_format.space_before = Pt(0)
tstyle.paragraph_format.space_after = Pt(6)
tstyle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
tpPr = tstyle.element.get_or_add_pPr()
drop_children(tpPr, "w:pBdr")
drop_children(tpPr, "w:shd")

# Compact TOC styles so the contents fit on one page
def ensure_toc_style(doc, idx, size, before, after, indent):
    styles_el = doc.styles.element
    sid = f"TOC{idx}"
    for el in styles_el.findall(qn("w:style")):
        if el.get(qn("w:styleId")) == sid:
            styles_el.remove(el)
    st = OxmlElement("w:style")
    st.set(qn("w:type"), "paragraph")
    st.set(qn("w:styleId"), sid)
    name = OxmlElement("w:name"); name.set(qn("w:val"), f"toc {idx}"); st.append(name)
    based = OxmlElement("w:basedOn"); based.set(qn("w:val"), "Normal"); st.append(based)
    nxt = OxmlElement("w:next"); nxt.set(qn("w:val"), "Normal"); st.append(nxt)
    ui = OxmlElement("w:uiPriority"); ui.set(qn("w:val"), "39"); st.append(ui)
    st.append(OxmlElement("w:qFormat"))
    pPr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(int(before * 20)))
    spacing.set(qn("w:after"), str(int(after * 20)))
    spacing.set(qn("w:line"), str(int(size * 1.08 * 20)))
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    if indent:
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(int(indent * 567)))
        pPr.append(ind)
    st.append(pPr)
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), LATIN)
    rFonts.set(qn("w:hAnsi"), LATIN)
    rFonts.set(qn("w:eastAsia"), BODY_FONT)
    rPr.append(rFonts)
    color = OxmlElement("w:color"); color.set(qn("w:val"), "1A1A1A"); rPr.append(color)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2))); rPr.append(sz)
    szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), str(int(size * 2))); rPr.append(szcs)
    st.append(rPr)
    styles_el.append(st)

for _idx, _size, _before, _after, _indent in ((1, 11, 3, 1, 0.0), (2, 10.5, 0, 0, 0.75), (3, 10, 0, 0, 1.5)):
    ensure_toc_style(doc, _idx, _size, _before, _after, _indent)

# ============================== COVER ==============================
bar = doc.add_table(rows=1, cols=1)
bar.autofit = False
btblPr = bar._tbl.tblPr
borders = OxmlElement("w:tblBorders")
for edge in ("top", "left", "bottom", "right"):
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "nil")
    borders.append(el)
btblPr.append(borders)
drop_children(btblPr, "w:tblW")
bw = OxmlElement("w:tblW"); bw.set(qn("w:type"), "dxa"); bw.set(qn("w:w"), str(int(15.4 * 567)))
btblPr.append(bw)
bcell = bar.rows[0].cells[0]
bcell.width = Cm(15.4)
shade(bcell, "17365D")
bcell.paragraphs[0].text = ""
set_para_spacing(bcell.paragraphs[0], before=0, after=0, line=0.35)
trPr = bar.rows[0]._tr.get_or_add_trPr()
trH = OxmlElement("w:trHeight"); trH.set(qn("w:val"), "170"); trH.set(qn("w:hRule"), "exact")
trPr.append(trH)

for _ in range(2):
    sp = doc.add_paragraph()
    set_para_spacing(sp, before=0, after=0, line=1.0)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(p, before=0, after=10, line=1.0)
r = p.add_run("毕 业 设 计 （ 论 文 ） 写 作 规 范")
set_run_font(r, size=12, bold=True, color=NAVY, east=HEAD_FONT, spacing=30)

sp = doc.add_paragraph()
set_para_spacing(sp, before=0, after=0, line=1.6)

tp = doc.add_paragraph(style="Title")
set_para_spacing(tp, before=0, after=8, line=1.25)
tr = tp.add_run("参考文献著录与正文引用规范")
set_run_font(tr, size=28, bold=True, color=DARK, east=HEAD_FONT)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(p, before=0, after=0, line=1.3)
r = p.add_run("顺序编码制操作指南与课题文献清单")
set_run_font(r, size=15, bold=False, color=GRAY, east=BODY_FONT)

for _ in range(3):
    sp = doc.add_paragraph()
    set_para_spacing(sp, before=0, after=0, line=1.2)

# Info block
info = doc.add_table(rows=6, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
info.autofit = False
itblPr = info._tbl.tblPr
ib = OxmlElement("w:tblBorders")
for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "nil")
    ib.append(el)
itblPr.append(ib)
drop_children(itblPr, "w:tblW")
iw = OxmlElement("w:tblW"); iw.set(qn("w:type"), "dxa"); iw.set(qn("w:w"), str(int(14.0 * 567)))
itblPr.append(iw)
info_rows = [
    ("课题名称", "个人文档智能整理与语义检索系统的设计与实现"),
    ("系部", ""),
    ("专业", ""),
    ("学号", ""),
    ("学生", ""),
    ("指导教师", ""),
]
for i, (lab, val) in enumerate(info_rows):
    c0, c1 = info.rows[i].cells
    c0.width = Cm(4.0); c1.width = Cm(10.0)
    for c in (c0, c1):
        cell_margins(c, top=90, left=80, bottom=90, right=80)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_para_spacing(p0, before=0, after=0, line=1.2)
    r = p0.add_run(lab + "：")
    set_run_font(r, size=12, color=DARK, east=HEAD_FONT)
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_para_spacing(p1, before=0, after=0, line=1.2)
    r = p1.add_run(val if val else " ")
    set_run_font(r, size=12, color=DARK)
    # bottom border under value
    tcPr = c1._tc.get_or_add_tcPr()
    bd = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "0"); bottom.set(qn("w:color"), "A6A6A6")
    bd.append(bottom); tcPr.append(bd)

for _ in range(4):
    sp = doc.add_paragraph()
    set_para_spacing(sp, before=0, after=0, line=1.2)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(p, before=0, after=0, line=1.2)
r = p.add_run("2026 年 11 月")
set_run_font(r, size=12, color=GRAY, east=BODY_FONT, spacing=20)

# ============================== TOC SECTION ==============================
sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
sec2.page_width = Cm(21.0); sec2.page_height = Cm(29.7)
sec2.top_margin = Cm(2.54); sec2.bottom_margin = Cm(2.54)
sec2.left_margin = Cm(3.0); sec2.right_margin = Cm(2.6)
set_numbering(sec2, "lowerRoman", 1)
add_header_footer(sec2, "参考文献著录与正文引用规范", page_fmt=True, prefix="", suffix="")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(p, before=6, after=16, line=1.2, keep_next=True)
r = p.add_run("目  录")
set_run_font(r, size=18, bold=True, color=DARK, east=HEAD_FONT, spacing=40)

tocp = doc.add_paragraph()
set_para_spacing(tocp, before=0, after=0, line=1.0)
run = tocp.add_run()
fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = ' TOC \\o "1-3" \\h \\z \\u '
fld_sep = OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
txt = OxmlElement("w:t"); txt.text = "目录将在文档打开后自动更新。"
fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_sep)
run._r.append(txt); run._r.append(fld_end)


# ============================== BODY SECTION ==============================
sec3 = doc.add_section(WD_SECTION.NEW_PAGE)
sec3.page_width = Cm(21.0); sec3.page_height = Cm(29.7)
sec3.top_margin = Cm(2.54); sec3.bottom_margin = Cm(2.54)
sec3.left_margin = Cm(3.0); sec3.right_margin = Cm(2.6)
set_numbering(sec3, "decimal", 1)
add_header_footer(sec3, "个人文档智能整理与语义检索系统 · 写作规范专题", page_fmt=True, prefix="第 ", suffix=" 页")

add_h(doc, "编制说明", 1, new_page=False)
add_body(doc, "本报告面向“个人文档智能整理与语义检索系统的设计与实现”课题，解决三个具体问题：正文中的引用怎样标注，文后的参考文献怎样著录，本课题应该引用哪些文献。")
add_body(doc, "报告依据 GB/T 7714—2015《信息与文献 参考文献著录规则》和学校本科生毕业设计（论文）撰写规范编写。第一章说明顺序编码制的总体要求；第二章给出正文标注的操作方法；第三章按文献类型给出著录模板；第四章提供 18 篇经过筛选的文献；第五章对照常见错误；第六章是提交前的检查清单；第七章结合任务书进度给出文献调研安排。附录中的示例段落和文献表可以直接复制到论文初稿，再按实际引用顺序调整编号。")
add_body(doc, "使用时先读第二章确定标注方式，再从第四章挑选与各章节对应的文献，最后按第六章逐项自查。报告中的编号只服务于这份清单；论文正式定稿时，必须按正文首次引用的先后顺序重新编号。")

add_callout(doc, "三条铁律", [
    "1. 正文引用在前，文后编号在后。编号由首次出现顺序决定，不按作者名气或出版时间排列。",
    "2. 正文标注与文后条目一一对应。多一条、少一条都会被格式审查发现。",
    "3. 格式统一到每一个标点和空格。参考文献表内使用英文半角标点，标点后空一格。",
])

add_body(doc, "本报告的清单包含中文文献 7 篇、外文文献 11 篇，覆盖专著、期刊、会议和电子文献四种类型。所有条目都按 GB/T 7714—2015 的格式写成了可直接使用的形式。")

# ============================== CHAPTER 1 ==============================
add_h(doc, "第一章  规范总览", 1, new_page=True)
add_h(doc, "1.1  什么是顺序编码制", 2)
add_body(doc, "顺序编码制是一种以正文引用顺序组织参考文献的著录体系。正文第一次引用某文献时，在引用处标注方括号序号；同一文献再次被引用时沿用原序号；全文引用结束后，按序号从小到大排列文后参考文献表。它与“著者—出版年制”的根本区别在于：编号由引用顺序决定，而不是由作者姓氏决定。")
add_body(doc, "这种制度有三个直接后果。第一，写作过程中新增文献可能引起后续序号整体后移，因此初稿阶段应使用文献管理软件或 Word 交叉引用，不要手工写死编号。第二，文后列表的顺序不是“重要程度”，也不是“发表时间”，而是“第一次被引用的先后”。第三，正文没有引用的文献不能出现在文后列表中，否则会被视为凑数。")

add_h(doc, "1.2  三项硬性指标", 2)
add_body(doc, "学校对参考文献提出三项可以直接检查的指标。三项都满足，才具备进入格式审查的基本条件。")
add_table(doc,
    ["指标", "要求", "检查方法"],
    [
        ["文献数量", "不少于 15 篇", "统计文后参考文献表的条目总数"],
        ["外文文献", "不少于 3 篇", "统计非中文文献的条目数，含译著和英文会议论文"],
        ["正文标注", "引用处 100% 标注", "逐段核对正文序号与文后条目是否一一对应"],
    ],
    widths=[4.0, 5.6, 5.8],
    font_size=10.5,
    first_col_center=True,
    caption="表 1-1  参考文献的三项硬性指标",
)
add_body(doc, "三项指标中，最容易失分的是第三项。数量可以通过补充文献达标，但“正文标注与文后条目对应”必须逐条核对。建议初稿完成后打印两份：一份只看正文标注，一份只看文后列表，用铅笔逐条勾连。凡是找不到对应条目的序号，或找不到正文引用的条目，都应当场处理。")

add_h(doc, "1.3  著录总原则", 2)
add_body(doc, "无论文献类型如何变化，参考文献表都遵循以下五条总原则。")
add_label_para(doc, "（1）标点。", "参考文献表内部使用英文半角标点，包括逗号、冒号、分号、括号和句点，标点后空一格。中文题名内部的冒号、顿号属于题名原文，保留原样。")
add_label_para(doc, "（2）作者。", "3 位及以内全部列出；超过 3 位只列前 3 位，其后加“，等”（中文）或“，et al.”（外文）。作者之间用逗号分隔。")
add_label_para(doc, "（3）类型标识。", "每条文献在题名后标注文献类型，如 [J]、[M]、[D]、[C]、[P]、[S]、[EB/OL]。标识放在题名之后、下一个著录项目之前。")
add_label_para(doc, "（4）一致性。", "同一篇论文中，期刊名、出版社名、作者姓名的写法必须前后一致。不要一处写“清华大学出版”，另一处写“清华大学出版社”。")
add_label_para(doc, "（5）完整性。", "期刊要有年、卷（期）、起止页码；专著要有版次（第一版不著录）、出版地、出版者、出版年；会议文献要有论文集名、出版地、出版者、出版年、页码。")

add_callout(doc, "关于卷号与页码的两个细节", [
    "有卷号时必须著录卷号；只有期号没有卷号的刊物，可以著录为“年（期）”。",
    "页码之间用半角短横线“-”，不用波浪号“~”；起止页码写全，不写“第 247 页”。",
])

add_h(doc, "1.4  三个常见理解误区", 2)
add_body(doc, "在正式写作前，有三个误解需要先澄清。它们不影响单条文献的格式，却会影响整份文献表的结构。")
add_bullet(doc, "误解一：文献越多越好。数量只是门槛，评审更看文献是否支撑了论文的具体判断。15 篇质量合格、引用位置明确的文献，胜过 30 篇只在文后出现、正文从未引用的条目。")
add_bullet(doc, "误解二：外文文献可以用译著凑数。译著可以计入外文文献，但必须按译著格式著录原作者和译者；如果只是把中文论文的英文摘要当作外文文献，不符合要求。")
add_bullet(doc, "误解三：参考文献表可以最后一次性补。编号依赖正文出现顺序，最后补写会导致大量序号错位。正确做法是边写边引、边引边记，定稿前只做格式统一和页码复核。")

# ============================== CHAPTER 2 ==============================
add_h(doc, "第二章  正文引用标注", 1, new_page=True)
add_h(doc, "2.1  标注位置", 2)
add_body(doc, "规则只有一句话：序号置于引文结束处最后一个字的右上角，通常位于句末标点之前。学校给出的示例是“……效率可提高25%[[14]]。”，序号在句号之前，并且是上标。")
add_body(doc, "下面三句用于对照。第一句正确；第二句把序号放到了句号之后；第三句没有使用上标。")
add_body(doc, "正例：相关研究表明，稠密向量检索能显著提升开放域问答的召回率[[13]]。", indent=True)
add_body(doc, "反例一：相关研究表明，稠密向量检索能显著提升开放域问答的召回率。[[13]]", indent=True)
add_body(doc, "反例二：相关研究表明，稠密向量检索能显著提升开放域问答的召回率[13]。", indent=True)
add_body(doc, "判断时有三个要点。第一，“最后一个字”指被引用内容的结尾，不是整段结尾；如果只引用半句，序号紧跟被引半句。第二，序号与句号的先后顺序按学校示例执行，全文统一，不要一部分在句号前、一部分在句号后。第三，上标只包括方括号和数字，不要把前面的文字一起设为上标。")

add_h(doc, "2.2  连续与不连续引用", 2)
add_body(doc, "同一处需要引用多篇文献时，把序号放进同一组方括号，序号之间用逗号分隔，连续序号用短横线连接。")
add_table(doc,
    ["引用情形", "规范写法", "说明"],
    [
        ["多篇不连续文献", "[3,5,7]", "序号从小到大排列，用逗号分隔"],
        ["多篇连续文献", "[3-5]", "表示第 3、4、5 篇，短横线前后不加空格"],
        ["连续与不连续混合", "[3,5-7]", "先写单个序号，再写连续区间"],
        ["同一文献不同页码", "[3]15-17", "页码写在方括号外，必要时才标注"],
    ],
    widths=[4.6, 4.2, 6.6],
    font_size=10.5,
    caption="表 2-1  多篇文献的标注写法",
)
add_body(doc, "不要写成“[3][5][7]”，也不要在方括号内使用中文顿号“[3、5、7]”。这两种写法在格式审查中都会被要求返工。")

add_h(doc, "2.3  特殊位置", 2)
add_label_para(doc, "（1）图表。", "图表标题或图注中引用他人数据时，序号标在说明文字末尾；图表下方的“资料来源”可以单独标注。")
add_label_para(doc, "（2）公式。", "引用他人公式时，序号标在公式编号之后或公式说明末尾，全文统一。")
add_label_para(doc, "（3）直接引语。", "整句引用时，序号标在引号之后、句号之前；引文较长另起段落时，标在引文末句的句号之前。")
add_label_para(doc, "（4）多句引用。", "同一观点连续引用多句时，可以在最后一句统一标注；如果中间插入自己的分析，应在各自引用处分别标注。")

add_h(doc, "2.4  与文后条目的对应", 2)
add_bullet(doc, "一个序号只对应一条文献；一条文献在正文中多次出现，始终使用同一序号。")
add_bullet(doc, "删除正文某处引用后，如果该文献不再被引用，应同时删除文后条目并重排后续序号。")
add_bullet(doc, "新增文献插入到中间时，后续序号顺延；建议用交叉引用或文献管理软件自动更新。")
add_bullet(doc, "文后列表中的每一条，都必须能在正文中找到至少一处标注。")
add_bullet(doc, "正文标注的序号必须连续，不能出现 [1][2][4] 这样的跳号。")

add_h(doc, "2.5  Word 中的操作步骤", 2)
add_body(doc, "下面的步骤适用于 Word 2016 及以上版本。WPS 的操作位置基本一致，但“交叉引用”的入口名称可能略有不同。")
add_label_para(doc, "步骤 1  设置上标。", "选中“[13]”，按 Ctrl+Shift+=；或打开“开始”→“字体”，勾选“上标”。")
add_label_para(doc, "步骤 2  建立自动编号。", "在文后参考文献表中选中所有条目，使用“开始”→“编号”→“定义新编号格式”，编号样式选“1, 2, 3, …”，格式填“[1]”。不要手工输入序号。")
add_label_para(doc, "步骤 3  插入交叉引用。", "把光标放在正文引用处，选择“引用”→“交叉引用”→“引用类型：编号项”→“引用内容：段落编号”，再选择对应条目。删除或移动条目后，正文编号可以随域更新。")
add_label_para(doc, "步骤 4  更新域。", "全选（Ctrl+A）后按 F9；打印前再执行一次“更新整个目录”。")
add_label_para(doc, "步骤 5  生成目录。", "使用“引用”→“目录”→“自动目录”。正文标题必须应用“标题 1/2/3”样式，不要用加粗大字模拟标题。")

add_callout(doc, "不要把尾注当成参考文献表", [
    "除非学校明确允许，不要用“尾注”代替文后参考文献表。尾注的编号位置、分隔符和列表格式与顺序编码制并不完全一致，答辩前修改成本很高。",
])

add_h(doc, "2.6  标注速查", 2)
add_body(doc, "把最常见的四种场景压缩成一张表，写作时可以随时对照。")
add_table(doc,
    ["场景", "写法", "示例"],
    [
        ["句末引用单篇", "序号上标，放在句号前", "……召回率[[13]]。"],
        ["同时引用多篇", "合并方括号，逗号分隔", "……性能提升[[3,5-7]]。"],
        ["同一文献多次引用", "沿用首次编号", "第一次[[14]]，第二次仍用[[14]]。"],
        ["图表或数据来源", "标在说明文字末尾", "数据来源：文献[[9]]。"],
    ],
    widths=[4.4, 5.2, 5.8],
    font_size=10.5,
    caption="表 2-2  正文标注速查",
)

# ============================== CHAPTER 3 ==============================
add_h(doc, "第三章  各类文献的著录格式", 1, new_page=True)
add_body(doc, "本章按文献类型给出模板。每条模板都可以直接套用，只需把示例中的作者、题名、年份替换成实际文献。所有模板均使用英文半角标点，标点后空一格。")

add_h(doc, "3.1  期刊论文 [J]", 2)
add_label_para(doc, "著录格式：", "主要责任者. 题名[J]. 刊名, 出版年, 卷(期): 起止页码.")
add_label_para(doc, "示例：", "[7] 刘知远, 孙茂松, 林衍凯, 等. 知识表示学习研究进展[J]. 计算机研究与发展, 2016, 53(2): 247-261.")
add_bullet(doc, "卷号与期号同时著录，卷号在前，期号加圆括号；没有卷号时才只写期号。")
add_bullet(doc, "页码写起止页，用半角短横线连接，不写“第 247 页”或“247~261”。")
add_bullet(doc, "期刊名写全称，不用缩写；外文期刊按刊物惯例保留全称。")
add_bullet(doc, "有 DOI 时可在页码后加“. DOI: 10.xxxx/xxxx.”；没有 DOI 不影响基本著录。")

add_h(doc, "3.2  专著 [M]", 2)
add_label_para(doc, "著录格式：", "主要责任者. 书名[M]. 版次. 出版地: 出版者, 出版年: 页码.")
add_label_para(doc, "示例：", "[1] 李航. 统计学习方法[M]. 2版. 北京: 清华大学出版社, 2019.")
add_bullet(doc, "第一版不著录版次；第二版及以上写“2版”“3版”，不写“第二版”。")
add_bullet(doc, "引用整本书时可以省略页码；引用具体章节时补上起止页码。")
add_bullet(doc, "译著先写原作者，再写译者，格式为“作者. 书名[M]. 译者, 译. 出版地: 出版者, 出版年.”。")
add_bullet(doc, "出版地写城市名，不写省份；有多地时写第一出版地。")

add_h(doc, "3.3  学位论文 [D]", 2)
add_label_para(doc, "著录格式：", "作者. 题名[D]. 保存地: 保存单位, 年份.")
add_label_para(doc, "格式示例：", "张某某. 面向个人文档的语义检索方法研究[D]. 南京: 南京大学, 2023.")
add_bullet(doc, "保存地是学位授予单位所在城市；保存单位写学校全称，不写学院。")
add_bullet(doc, "学位论文一般不写页码；如果引用某一页的具体观点，可在年份后补页码。")
add_bullet(doc, "上面的姓名与题名仅用于演示格式，不是真实文献；正式引用时必须替换为知网或学校图书馆检索到的条目。")

add_h(doc, "3.4  会议论文 [C]", 2)
add_label_para(doc, "著录格式：", "作者. 题名[C]//论文集编者. 论文集名. 出版地: 出版者, 出版年: 起止页码.")
add_label_para(doc, "示例：", "[10] VASWANI A, SHAZEER N, PARMAR N, et al. Attention is all you need[C]//Advances in Neural Information Processing Systems 30. Red Hook: Curran Associates, 2017: 5998-6008.")
add_bullet(doc, "“//”用于连接论文题名与论文集信息，前后不空格。")
add_bullet(doc, "论文集名写全称，包括会议名称和届次（如有）；出版地、出版者、出版年不能省略。")
add_bullet(doc, "没有正式出版的会议论文，按电子文献或报告处理，不要硬套 [C]。")

add_h(doc, "3.5  专利 [P]", 2)
add_label_para(doc, "著录格式：", "专利申请者. 专利题名: 国别, 专利号[P]. 公告日期.")
add_label_para(doc, "格式示例：", "张三. 一种文档语义检索方法: 中国, CN123456789A[P]. 2024-01-01.")
add_bullet(doc, "专利号写公开号或授权号，全文统一；公告日期写“YYYY-MM-DD”。")
add_bullet(doc, "国别写“中国”“美国”等中文名称，或按学校要求写国家代码。")
add_bullet(doc, "上述专利号仅用于演示格式，不是真实专利。")

add_h(doc, "3.6  电子文献 [EB/OL]", 2)
add_label_para(doc, "著录格式：", "主要责任者. 题名[EB/OL]. (更新日期)[引用日期]. 获取和访问路径.")
add_label_para(doc, "示例：", "[15] GAO Y, XIONG Y, GAO X, et al. Retrieval-augmented generation for large language models: A survey[EB/OL]. (2023-12-18)[2026-11-23]. https://arxiv.org/abs/2312.10997.")
add_bullet(doc, "引用日期必须写，格式为“[YYYY-MM-DD]”，写你实际访问页面的日期。")
add_bullet(doc, "更新日期可以省略；省略时保留引用日期，如“[2026-11-23]”。")
add_bullet(doc, "URL 写完整路径，末尾不加句号；如果需要句号，放在 URL 之后并空一格。")
add_bullet(doc, "技术文档要写清版本或页面名称，避免只写“官网”两个字。")

add_h(doc, "3.7  标准 [S]", 2)
add_label_para(doc, "著录格式：", "标准编号 标准名称[S]. 出版地: 出版者, 出版年.")
add_label_para(doc, "示例：", "GB/T 7714—2015 信息与文献 参考文献著录规则[S]. 北京: 中国标准出版社, 2015.")
add_bullet(doc, "标准编号中的年份用破折号“—”，这是标准编号本身的写法，不属于标点错误。")
add_bullet(doc, "标准名称写全称，不随意增删“中华人民共和国”等前缀。")

add_h(doc, "3.8  外文文献的姓名与大小写", 2)
add_bullet(doc, "姓前名后，姓全大写，名用缩写字母，缩写字母后不加点，如“MANNING C D”。")
add_bullet(doc, "作者超过 3 位，第 3 位后加“, et al.”；不要写成“etc.”或“等”。")
add_bullet(doc, "题名一般采用句首字母大写（sentence case），专有名词保留大写；刊名和论文集名保留原刊名的大小写。")
add_bullet(doc, "同一外文作者在全文中只使用一种译名或原文写法，不要混用。")

add_h(doc, "3.9  类型速查表", 2)
add_table(doc,
    ["文献类型", "标识", "必须著录的项目", "本报告示例"],
    [
        ["期刊论文", "[J]", "作者. 题名. 刊名, 年, 卷(期): 起止页码.", "[7]"],
        ["专著", "[M]", "作者. 书名. 版次. 出版地: 出版者, 年: 页码.", "[1]"],
        ["学位论文", "[D]", "作者. 题名. 保存地: 保存单位, 年.", "3.3 节示例"],
        ["会议论文", "[C]", "作者. 题名//论文集名. 出版地: 出版者, 年: 页码.", "[10]"],
        ["专利", "[P]", "申请者. 专利题名: 国别, 专利号. 公告日期.", "3.5 节示例"],
        ["电子文献", "[EB/OL]", "作者. 题名. (更新日期)[引用日期]. URL.", "[15]"],
        ["标准", "[S]", "标准编号 标准名称. 出版地: 出版者, 年.", "3.7 节示例"],
    ],
    widths=[3.0, 1.9, 7.4, 3.1],
    font_size=10,
    first_col_center=True,
    caption="表 3-1  常见文献类型的著录项目速查",
    note="注：表中的类型标识在正式著录时应紧跟题名，如“题名[J].”，不要单独成列。",
)

# ============================== CHAPTER 4 ==============================
add_h(doc, "第四章  本课题参考文献清单", 1, new_page=True)
add_body(doc, "本章给出 18 篇与课题直接相关的文献，按“理论方法—检索技术—检索增强生成—工程实现”四组排列。编号只服务于这份清单；正式论文中必须按正文首次引用顺序重新编号。")

add_h(doc, "4.1  编制原则", 2)
add_bullet(doc, "相关性：每条文献都能对应论文中的具体论述，不为凑数而列。")
add_bullet(doc, "权威性：优先经典专著、领域内主流期刊和会议、官方技术文档。")
add_bullet(doc, "时效性：基础理论可用经典文献，技术方案尽量包含近 3 年文献。")
add_bullet(doc, "类型均衡：兼顾专著、期刊、会议和电子文献，避免全部是网页。")
add_bullet(doc, "可核验：题名、作者、年份、页码以数据库页面为准，不凭记忆著录。")

add_h(doc, "4.2  文献清单", 2)
p = doc.add_paragraph()
set_para_spacing(p, before=4, after=4, line=1.3, keep_next=True)
r = p.add_run("（一）理论与方法基础")
set_run_font(r, size=11, bold=True, color=NAVY, east=HEAD_FONT)

refs = [
 (1, "李航. 统计学习方法[M]. 2版. 北京: 清华大学出版社, 2019.", None),
 (2, "宗成庆. 统计自然语言处理[M]. 2版. 北京: 清华大学出版社, 2013.", None),
 (3, "周志华. 机器学习[M]. 北京: 清华大学出版社, 2016.", None),
 (4, "邱锡鹏. 神经网络与深度学习[M]. 北京: 机械工业出版社, 2020.", None),
 (5, "吴军. 数学之美[M]. 3版. 北京: 人民邮电出版社, 2020.", None),
 (6, "王昊奋, 漆桂林, 陈华钧. 知识图谱：方法、实践与应用[M]. 北京: 电子工业出版社, 2019.", None),
 (7, "刘知远, 孙茂松, 林衍凯, 等. 知识表示学习研究进展[J]. 计算机研究与发展, 2016, 53(2): 247-261.", None),
]
for num, txt, url in refs:
    add_ref_entry(doc, num, txt, url)

p = doc.add_paragraph()
set_para_spacing(p, before=10, after=4, line=1.3, keep_next=True)
r = p.add_run("（二）信息检索与向量表示")
set_run_font(r, size=11, bold=True, color=NAVY, east=HEAD_FONT)

refs = [
 (8, "MANNING C D, RAGHAVAN P, SCHÜTZE H. Introduction to information retrieval[M]. Cambridge: Cambridge University Press, 2008.", None),
 (9, "ROBERTSON S, ZARAGOZA H. The probabilistic relevance framework: BM25 and beyond[J]. Foundations and Trends in Information Retrieval, 2009, 3(4): 333-389.", None),
 (10, "VASWANI A, SHAZEER N, PARMAR N, et al. Attention is all you need[C]//Advances in Neural Information Processing Systems 30. Red Hook: Curran Associates, 2017: 5998-6008.", None),
 (11, "DEVLIN J, CHANG M W, LEE K, et al. BERT: Pre-training of deep bidirectional transformers for language understanding[C]//Proceedings of the 2019 Conference of the North American Chapter of the Association for Computational Linguistics: Human Language Technologies. Stroudsburg: Association for Computational Linguistics, 2019: 4171-4186.", None),
 (12, "REIMERS N, GUREVYCH I. Sentence-BERT: Sentence embeddings using siamese BERT-networks[C]//Proceedings of the 2019 Conference on Empirical Methods in Natural Language Processing and the 9th International Joint Conference on Natural Language Processing. Stroudsburg: Association for Computational Linguistics, 2019: 3982-3992.", None),
 (13, "KARPUKHIN V, OGUZ B, MIN S, et al. Dense passage retrieval for open-domain question answering[C]//Proceedings of the 2020 Conference on Empirical Methods in Natural Language Processing. Stroudsburg: Association for Computational Linguistics, 2020: 6769-6781.", None),
]
for num, txt, url in refs:
    add_ref_entry(doc, num, txt, url)

p = doc.add_paragraph()
set_para_spacing(p, before=10, after=4, line=1.3, keep_next=True)
r = p.add_run("（三）检索增强生成")
set_run_font(r, size=11, bold=True, color=NAVY, east=HEAD_FONT)

refs = [
 (14, "LEWIS P, PEREZ E, PIKTUS A, et al. Retrieval-augmented generation for knowledge-intensive NLP tasks[C]//Advances in Neural Information Processing Systems 33. Red Hook: Curran Associates, 2020: 9459-9474.", None),
 (15, "GAO Y, XIONG Y, GAO X, et al. Retrieval-augmented generation for large language models: A survey[EB/OL]. (2023-12-18)[2026-11-23]. https://arxiv.org/abs/2312.10997.", "https://arxiv.org/abs/2312.10997"),
]
for num, txt, url in refs:
    add_ref_entry(doc, num, txt, url)

p = doc.add_paragraph()
set_para_spacing(p, before=10, after=4, line=1.3, keep_next=True)
r = p.add_run("（四）工程实现文档")
set_run_font(r, size=11, bold=True, color=NAVY, east=HEAD_FONT)

refs = [
 (16, "ELASTIC. Elasticsearch guide[EB/OL]. [2026-11-23]. https://www.elastic.co/guide/en/elasticsearch/reference/current/index.html.", "https://www.elastic.co/guide/en/elasticsearch/reference/current/index.html"),
 (17, "OPENAI. OpenAI API reference[EB/OL]. [2026-11-23]. https://platform.openai.com/docs/api-reference.", "https://platform.openai.com/docs/api-reference"),
 (18, "APACHE SOFTWARE FOUNDATION. Apache PDFBox[EB/OL]. [2026-11-23]. https://pdfbox.apache.org/.", "https://pdfbox.apache.org/"),
]
for num, txt, url in refs:
    add_ref_entry(doc, num, txt, url)

add_h(doc, "4.3  结构统计", 2)
add_body(doc, "清单共 18 篇，中文 7 篇、外文 11 篇，外文文献数量超过学校要求的 3 篇下限。类型上兼顾专著、期刊、会议和电子文献，时间跨度从 2008 年到 2023 年，另有 3 份持续更新的官方文档。")
add_table(doc,
    ["文献类型", "中文", "外文", "合计"],
    [
        ["专著 [M]", "6", "1", "7"],
        ["期刊 [J]", "1", "1", "2"],
        ["会议 [C]", "0", "5", "5"],
        ["电子文献 [EB/OL]", "0", "4", "4"],
        ["合计", "7", "11", "18"],
    ],
    widths=[5.0, 3.2, 3.2, 4.0],
    font_size=10.5,
    first_col_center=False,
    caption="表 4-1  参考文献的语种与类型分布",
)

add_h(doc, "4.4  章节引用映射", 2)
add_body(doc, "下表把文献分配到论文章节，目的是让每条文献都有明确的引用位置。映射是建议，不是强制；正文没有实际引用时，不要为了凑数硬塞。")
add_table(doc,
    ["论文章节", "建议引用文献", "引用用途"],
    [
        ["第 1 章  绪论", "[1,2,6-8,14,15]", "研究背景、知识组织、检索增强生成的提出"],
        ["第 2 章  相关技术", "[3-5,9-13]", "机器学习、深度学习、BM25、Transformer、BERT、句向量、稠密检索"],
        ["第 3 章  需求分析", "[1,8,16]", "用户场景、检索需求、全文检索能力"],
        ["第 4 章  系统设计", "[6,8,13,16,17]", "数据模型、索引设计、向量检索、接口设计"],
        ["第 5 章  系统实现", "[12,16-18]", "句向量、Elasticsearch、模型 API、PDF 解析"],
        ["第 6 章  测试与评价", "[14,15,17]", "问答效果评价、RAG 评价指标"],
    ],
    widths=[4.2, 5.4, 5.8],
    font_size=10,
    caption="表 4-2  文献与论文章节的建议映射",
)

add_callout(doc, "使用清单前必须知道的两件事", [
    "第一，清单中的编号不是论文最终编号。定稿时按正文首次引用顺序重新排列，正文序号与文后条目同步更新。",
    "第二，不同数据库对页码的著录偶有差异。提交前请在中国知网、ACL Anthology 或会议论文集中复核起止页码，以期刊或论文集原文为准。",
])

# ============================== CHAPTER 5 ==============================
add_h(doc, "第五章  常见错误对照", 1, new_page=True)
add_body(doc, "下表按“标注—著录—一致性”三类整理。第 1—2 行查正文标注，第 3—10 行查单条著录，第 11—15 行查全文一致性；第 13、14 行最易忽略，却决定正文与文后能否对应。")
add_table(doc,
    ["序号", "错误示例", "规范写法", "问题说明"],
    [
        ["1", "……效率可提高25%[14]。", "……效率可提高25%[[14]]。", "序号未设为上标"],
        ["2", "……召回率。[13]", "……召回率[[13]]。", "序号放在句号之后"],
        ["3", "计算机研究与发展，2016，53（2）：247-261.", "计算机研究与发展, 2016, 53(2): 247-261.", "使用全角标点，且未空格"],
        ["4", "247~261", "247-261", "页码连接符错误"],
        ["5", "刘知远, 孙茂松, 林衍凯, 谢若冰.", "刘知远, 孙茂松, 林衍凯, 等.", "超过三位作者未用“等”"],
        ["6", "Patrick Lewis, Ethan Perez, ...", "LEWIS P, PEREZ E, PIKTUS A, et al.", "英文姓名顺序与缩写错误"],
        ["7", "2016(2): 247-261", "2016, 53(2): 247-261", "缺少卷号"],
        ["8", "……[C]. 2017: 5998-6008.", "……[C]//Advances in Neural Information Processing Systems 30. Red Hook: Curran Associates, 2017: 5998-6008.", "会议论文集信息缺失"],
        ["9", "……[EB/OL]. https://example.com.", "……[EB/OL]. [2026-11-23]. https://example.com.", "电子文献缺引用日期"],
        ["10", "《统计学习方法》[M].", "统计学习方法[M].", "题名不应加书名号"],
        ["11", "[3][5][7]", "[3,5,7]", "多篇引用未合并方括号"],
        ["12", "[1][2][4]", "[1][2][3]", "序号跳号，未按出现顺序连续编号"],
        ["13", "正文出现 [19]，文后仅 18 条", "补齐 [19] 或删除该引用", "正文与文后不一一对应"],
        ["14", "[16] 只在文后出现", "在正文补引或删除条目", "文后条目未被正文引用"],
        ["15", "转引自[8]", "查原文并引用原始文献", "转引未溯源"],
    ],
    widths=[1.2, 4.5, 5.9, 3.8],
    font_size=10,
    first_col_center=True,
    caption="表 5-1  参考文献常见错误对照",
)

# ============================== CHAPTER 6 ==============================
add_h(doc, "第六章  提交前自查清单", 1, new_page=True)
add_body(doc, "这份清单按“先数量、再对应、后格式”的顺序排列。建议打印出来，每完成一项就在“结果”栏勾选；全部为“是”之后再提交。")

check_items = [
    "参考文献总数不少于 15 篇",
    "外文文献不少于 3 篇",
    "正文标注与文后条目一一对应",
    "编号按正文首次引用顺序排列",
    "同一文献重复引用时沿用同一编号",
    "序号连续，没有跳号或重号",
    "序号位于引文最后一个字的右上角",
    "多篇引用使用 [3,5-7] 形式，未拆成多个方括号",
    "作者不超过三位时全部列出，超过三位用“等”或“et al.”",
    "英文作者姓前名后、姓全大写、名缩写",
    "期刊论文有年、卷（期）、起止页码",
    "专著有版次、出版地、出版者、出版年",
    "会议论文有论文集名、出版地、出版者、出版年、页码",
    "学位论文有保存地、保存单位、年份",
    "电子文献有引用日期和完整 URL",
    "参考文献表全部使用半角标点，标点后空一格",
    "题名不加书名号，类型标识紧跟题名",
    "页码使用半角短横线“-”，不用“~”",
    "全文的期刊名、出版社名、作者名写法一致",
]
rows = [[str(i + 1), item, "☐ 是    ☐ 否", ""] for i, item in enumerate(check_items)]
add_table(doc,
    ["序号", "检查项", "结果", "备注"],
    rows,
    widths=[1.1, 10.0, 2.6, 1.7],
    font_size=10,
    first_col_center=True,
    caption="表 6-1  参考文献提交前自查清单",
)

add_h(doc, "6.1  三步核对法", 2)
add_body(doc, "第一步，查数量。统计文后条目总数和外文条目数，确认达到“15 篇、3 篇外文”的下限。不要用同一篇文献的不同版本重复计数。")
add_body(doc, "第二步，查对应。从正文第一个序号开始，按顺序在文后列表中找到对应条目，用铅笔在两条上做同一个记号；再反向从文后第一条开始，在正文中找到至少一处引用。两个方向都走通，才算一一对应。")
add_body(doc, "第三步，查格式。重点看四类问题：标点是否半角、作者是否符合三位规则、期刊是否有卷期页码、电子文献是否有引用日期。最后检查全文的字体、字号和行距是否统一。")

add_callout(doc, "打印前的最后五分钟", [
    "按 Ctrl+A 后按 F9，更新所有域和目录；确认目录页码与正文一致。",
    "随机抽取 3 条文献，到知网或会议论文集中核对题名、年份和页码。",
    "搜索全文中的“，”“（”“）”和“~”，确认参考文献表内没有混入全角标点。",
])

add_h(doc, "6.2  四类高频返工问题", 2)
add_body(doc, "根据格式审查的常见退回意见，下面四类问题出现频率最高。提交前用十分钟专门检查这四项，可以避免大部分返工。")
add_bullet(doc, "正文有序号、文后无条目。通常是删改正文时漏删了引用，或合并段落时丢了编号。")
add_bullet(doc, "上标位置不统一。同一篇论文中一部分在句号前、一部分在句号后，会被要求全文重排。")
add_bullet(doc, "英文作者格式错误。姓与名顺序颠倒、姓未大写、缩写后多写一个点，都是常见问题。")
add_bullet(doc, "电子文献缺引用日期。只写 URL 不写访问日期，无法证明文献的可获取性。")

# ============================== CHAPTER 7 ==============================
add_h(doc, "第七章  写作安排与工具建议", 1, new_page=True)
add_body(doc, "任务书把文献调研安排在 2026 年 11 月 25 日至 12 月 20 日，随后进入需求分析、系统设计和编码阶段。文献工作不应在开题前一次性完成，而应随着系统推进持续补充。")

add_h(doc, "7.1  结合任务书进度", 2)
add_table(doc,
    ["阶段", "时间", "文献工作", "产出"],
    [
        ["文献调研", "2026-11-25—12-05", "精读 10 篇核心文献，完成文献卡片", "10 张文献卡片"],
        ["综述撰写", "2026-12-06—12-15", "补齐至 15 篇以上，撰写综述初稿", "文献综述初稿"],
        ["开题完善", "2026-12-16—12-20", "统一著录格式，核对卷期页码", "开题报告与文献表"],
        ["系统设计", "2027-03-02—04-20", "每引入一项新技术补 1—2 篇文献", "技术选型依据"],
        ["测试评价", "2027-04-21—05-11", "补充评价方法与指标文献", "测试章节引用"],
        ["论文定稿", "2027-05-12—06-08", "全量核对，按正文顺序重排编号", "定稿参考文献表"],
    ],
    widths=[3.0, 4.4, 5.4, 2.6],
    font_size=10,
    caption="表 7-1  文献工作与任务书进度的对应安排",
)

add_h(doc, "7.2  文献管理工具", 2)
add_label_para(doc, "Zotero。", "免费，支持 GB/T 7714—2015 样式，可以从知网、ACL Anthology、IEEE Xplore 等站点抓取条目。导出后仍需人工核对卷期页码，不能完全依赖自动生成。")
add_label_para(doc, "EndNote。", "学校如果提供正版授权可以使用；同样需要在导出后检查样式，尤其是中文文献的标点和作者格式。")
add_label_para(doc, "知网导出。", "在文献详情页选择“参考文献”格式导出，复制到 Word 后统一替换全角标点，再按学校模板调整。")
add_label_para(doc, "文献卡片。", "每篇文献记录七项：题名、作者、年份、来源、核心方法、可引用章节、原文页码。写论文时先翻卡片，再回原文核对，避免凭印象引用。")

add_h(doc, "7.3  避免“凑参考文献”", 2)
add_bullet(doc, "先写论述，再补引用。每一处引用都应服务于一个具体判断，而不是为了增加数量。")
add_bullet(doc, "优先引用原始文献。引用 BERT 原文，不引用二手综述中的转述；找不到原文时，才考虑转引并注明。")
add_bullet(doc, "综述类文献用于交代领域全貌，方法类文献用于支撑技术选择，二者不要互相替代。")
add_bullet(doc, "近 3 年文献用于说明技术现状，经典文献用于说明方法来源；时间结构本身也是评审关注点。")

add_h(doc, "7.4  与开题报告的衔接", 2)
add_body(doc, "开题报告的文献综述通常包括研究现状、技术路线和参考文献三部分。本章的清单可以直接支撑前两部分：绪论用 [1,2,6-8,14,15] 说明研究背景，相关技术用 [3-5,9-13] 说明方法基础，系统设计用 [16,17] 说明实现方案。开题阶段不必把 18 篇全部写进正文，但文献表应完整列出，正文引用控制在 10—12 处。")
add_bullet(doc, "开题报告重点引用近 3 年文献，说明选题的时效性；经典文献放在技术基础部分。")
add_bullet(doc, "论文正文再补充方法细节和对比文献，不必与开题报告完全一致，但核心文献应保留。")
add_bullet(doc, "答辩 PPT 只列 5—8 篇核心文献，完整清单放在论文和开题报告中。")

# ============================== APPENDIX A ==============================
add_h(doc, "附录 A  正文引用示例", 1, new_page=True)
add_body(doc, "下面是一段模拟的绪论文字，用于演示上标位置。文中的编号对应附录 B 的文献表，句子内容仅用于说明写法。")

add_body(doc, "个人文档管理的研究重点，正在从文件存储转向内容理解。早期系统主要解决目录、标签和全文关键词匹配问题，其理论基础可以追溯到经典信息检索模型[[8]]。BM25 通过词频、逆文档频率和文档长度归一化对候选文档排序，在关键词明确、术语稳定的场景中仍然具有较低的计算成本和稳定的召回表现[[9]]。但个人文档往往存在主题交叉、表述口语化和同义词频繁等特点，仅依赖词项匹配容易漏检。以 BERT 为代表的预训练语言模型通过上下文编码获得语义表示[[11]]，Sentence-BERT 进一步把句子映射为固定维度向量，使大规模语义相似度计算成为可能[[12]]。在此基础上，稠密段落检索把查询和文档分别编码，并通过向量内积完成排序[[13]]。")

add_body(doc, "检索增强生成（RAG）把外部文档检索与大语言模型生成结合起来：系统先检索与问题相关的片段，再把片段作为上下文交给模型生成答案[[14]]。这一范式对个人文档系统尤其合适，因为用户的资料不断更新，而模型参数无法及时反映这些变化。相关综述进一步整理了 RAG 的检索器、生成器和增强策略[[15]]。")

add_body(doc, "本文的研究范围限定在单用户、本地部署的个人文档场景，重点解决 PDF、Word、TXT 等格式的解析、索引与自然语言查询问题。与面向开放域的大规模问答系统相比，本系统更强调数据可控、引用可追溯和结果可解释。")

add_h(doc, "A.1  标注说明", 2)
add_bullet(doc, "[8][9][11][12][13][14][15] 都是上标，位于句末标点之前。")
add_bullet(doc, "同一文献再次出现时沿用原序号，不重新编号。")
add_bullet(doc, "连续引用多篇文献时写成 [12-13]，不写成 [12][13]。")
add_bullet(doc, "如果只引用半句，序号紧跟被引半句，而不是整段末尾。")

# ============================== APPENDIX B ==============================
add_h(doc, "附录 B  参考文献表", 1, new_page=True)
add_body(doc, "以下条目可直接复制；复制后请按正文首次引用顺序重新编号。")

all_refs = [
 (1, "李航. 统计学习方法[M]. 2版. 北京: 清华大学出版社, 2019.", None),
 (2, "宗成庆. 统计自然语言处理[M]. 2版. 北京: 清华大学出版社, 2013.", None),
 (3, "周志华. 机器学习[M]. 北京: 清华大学出版社, 2016.", None),
 (4, "邱锡鹏. 神经网络与深度学习[M]. 北京: 机械工业出版社, 2020.", None),
 (5, "吴军. 数学之美[M]. 3版. 北京: 人民邮电出版社, 2020.", None),
 (6, "王昊奋, 漆桂林, 陈华钧. 知识图谱：方法、实践与应用[M]. 北京: 电子工业出版社, 2019.", None),
 (7, "刘知远, 孙茂松, 林衍凯, 等. 知识表示学习研究进展[J]. 计算机研究与发展, 2016, 53(2): 247-261.", None),
 (8, "MANNING C D, RAGHAVAN P, SCHÜTZE H. Introduction to information retrieval[M]. Cambridge: Cambridge University Press, 2008.", None),
 (9, "ROBERTSON S, ZARAGOZA H. The probabilistic relevance framework: BM25 and beyond[J]. Foundations and Trends in Information Retrieval, 2009, 3(4): 333-389.", None),
 (10, "VASWANI A, SHAZEER N, PARMAR N, et al. Attention is all you need[C]//Advances in Neural Information Processing Systems 30. Red Hook: Curran Associates, 2017: 5998-6008.", None),
 (11, "DEVLIN J, CHANG M W, LEE K, et al. BERT: Pre-training of deep bidirectional transformers for language understanding[C]//Proceedings of the 2019 Conference of the North American Chapter of the Association for Computational Linguistics: Human Language Technologies. Stroudsburg: Association for Computational Linguistics, 2019: 4171-4186.", None),
 (12, "REIMERS N, GUREVYCH I. Sentence-BERT: Sentence embeddings using siamese BERT-networks[C]//Proceedings of the 2019 Conference on Empirical Methods in Natural Language Processing and the 9th International Joint Conference on Natural Language Processing. Stroudsburg: Association for Computational Linguistics, 2019: 3982-3992.", None),
 (13, "KARPUKHIN V, OGUZ B, MIN S, et al. Dense passage retrieval for open-domain question answering[C]//Proceedings of the 2020 Conference on Empirical Methods in Natural Language Processing. Stroudsburg: Association for Computational Linguistics, 2020: 6769-6781.", None),
 (14, "LEWIS P, PEREZ E, PIKTUS A, et al. Retrieval-augmented generation for knowledge-intensive NLP tasks[C]//Advances in Neural Information Processing Systems 33. Red Hook: Curran Associates, 2020: 9459-9474.", None),
 (15, "GAO Y, XIONG Y, GAO X, et al. Retrieval-augmented generation for large language models: A survey[EB/OL]. (2023-12-18)[2026-11-23]. https://arxiv.org/abs/2312.10997.", "https://arxiv.org/abs/2312.10997"),
 (16, "ELASTIC. Elasticsearch guide[EB/OL]. [2026-11-23]. https://www.elastic.co/guide/en/elasticsearch/reference/current/index.html.", "https://www.elastic.co/guide/en/elasticsearch/reference/current/index.html"),
 (17, "OPENAI. OpenAI API reference[EB/OL]. [2026-11-23]. https://platform.openai.com/docs/api-reference.", "https://platform.openai.com/docs/api-reference"),
 (18, "APACHE SOFTWARE FOUNDATION. Apache PDFBox[EB/OL]. [2026-11-23]. https://pdfbox.apache.org/.", "https://pdfbox.apache.org/"),
]
for num, txt, url in all_refs:
    add_ref_entry(doc, num, txt, url)

p = doc.add_paragraph()
set_para_spacing(p, before=20, after=0, line=1.3)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("编制日期：2026 年 11 月 23 日")
set_run_font(r, size=9.5, color=MIDGRAY)

out = r"C:\Users\LEGION\Desktop\毕设\docs\report\参考文献著录与正文引用规范.docx"
doc.save(out)
print("saved", out)
