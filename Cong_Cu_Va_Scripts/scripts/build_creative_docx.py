import json, os, sys, re
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import lxml.etree as etree
import latex2mathml.converter

sys.stdout.reconfigure(encoding='utf-8')

# Load XSLT
xsl_path = r'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL'
if not os.path.exists(xsl_path):
    xsl_path = r'C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL'

xslt = etree.parse(xsl_path)
transform = etree.XSLT(xslt)

def latex_to_omml(latex_str):
    try:
        mathml = latex2mathml.converter.convert(latex_str)
        # Escape & not in entities
        mathml = re.sub(r'&(?!(#[0-9]+|#x[0-9a-fA-F]+|[a-zA-Z]+);)', '&amp;', mathml)
        tree = etree.fromstring(mathml)
        omml = transform(tree)
        return omml.getroot()
    except Exception as e:
        return None

def add_inline_math(paragraph, text, font_name="Times New Roman", font_size=Pt(11), color_rgb=None, bold=False):
    parts = text.split('$')
    for i, part in enumerate(parts):
        if not part:
            continue
        if i % 2 == 0:
            run = paragraph.add_run(part)
            run.font.name = font_name
            run.font.size = font_size
            run.bold = bold
            if color_rgb:
                run.font.color.rgb = color_rgb
        else:
            omml_root = latex_to_omml(part)
            if omml_root is not None:
                paragraph._p.append(omml_root)
            else:
                run = paragraph.add_run(part)
                run.font.name = font_name
                run.font.size = font_size
                run.bold = bold
                run.italic = True
                if color_rgb:
                    run.font.color.rgb = color_rgb

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_left_border(cell, border_color="059669", sz="18"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
    tcPr.append(borders)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def build_creative_document(all_items, chuyen_de_title, topic_code, output_docx_path, img_dir=None):
    doc = Document()

    # Standard A4 Margins
    for sec in doc.sections:
        sec.top_margin = Inches(0.79)
        sec.bottom_margin = Inches(0.79)
        sec.left_margin = Inches(1.18)
        sec.right_margin = Inches(0.79)

    # 1. Header
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(0)
    p_inst.paragraph_format.space_after = Pt(2)
    r1 = p_inst.add_run("BAN BIÊN SOẠN & PHÁT TRIỂN ĐỀ THI HỌC SINH GIỎI TOÁN 11\n")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(11)
    r1.bold = True
    r2 = p_inst.add_run("BỘ ĐỀ SÁNG TẠO 2026 — CHƯƠNG TRÌNH GDPT 2018 MỚI NHẤT")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(11)
    r2.bold = True

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.paragraph_format.space_before = Pt(0)
    p_div.paragraph_format.space_after = Pt(10)
    r_div = p_div.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    r_div.font.name = "Times New Roman"
    r_div.font.size = Pt(10)
    r_div.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)

    # 2. Document Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(chuyen_de_title)
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(15.5)
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A) # Deep Blue

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(10)
    r_sub = p_sub.add_run("Tuyển tập 20 bài toán sáng tạo đỉnh cao kèm Hướng dẫn tư duy & Lời giải mẫu mực\n(100% Công thức chuẩn Word Equation Alt + =, Không lỗi font)")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(10.5)
    r_sub.italic = True
    r_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # Table Info Banner
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = False
    col_widths = [Inches(2.0), Inches(2.3), Inches(2.0)]
    for row in table.rows:
        for idx_w, width in enumerate(col_widths):
            row.cells[idx_w].width = width

    cells = table.rows[0].cells
    info_texts = [
        ("Quy mô tài liệu", f"{len(all_items)} bài toán sáng tạo mới"),
        ("Cấu trúc sư phạm", "Đề bài • Gợi ý • Lời giải chi tiết"),
        ("Định dạng công thức", "100% Native Word Equation")
    ]
    for i, (k, v) in enumerate(info_texts):
        cell = cells[i]
        set_cell_background(cell, "F8FAFC")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        rk = p.add_run(f"{k}\n")
        rk.font.name = "Times New Roman"
        rk.font.size = Pt(10)
        rk.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        rv = p.add_run(v)
        rv.font.name = "Times New Roman"
        rv.font.size = Pt(10.5)
        rv.bold = True
        rv.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(6)
    p_space.paragraph_format.space_after = Pt(6)

    # 3. Iterate through all items
    for item in all_items:
        idx = item.get('index', 1)
        title = item.get('title', '')
        prompt = item.get('prompt', '')
        subs = item.get('sub_questions', [])
        guide_lines = item.get('guide_lines', [])
        solution_lines = item.get('solution_lines', [])

        # Item Header
        p_hdr = doc.add_paragraph()
        p_hdr.paragraph_format.space_before = Pt(12)
        p_hdr.paragraph_format.space_after = Pt(2)
        p_hdr.paragraph_format.keep_with_next = True

        r_badge = p_hdr.add_run(f"Bài {idx}. ")
        r_badge.bold = True
        r_badge.font.name = "Times New Roman"
        r_badge.font.size = Pt(12)
        r_badge.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

        src_text = f"[Sáng tạo HSG 11 — Mã: {topic_code}_{idx:02d}]"
        if title:
            src_text += f" — {title}"
        r_src = p_hdr.add_run(src_text)
        r_src.italic = True
        r_src.font.name = "Times New Roman"
        r_src.font.size = Pt(10)
        r_src.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

        # Prompt
        if prompt:
            p_prompt = doc.add_paragraph()
            p_prompt.paragraph_format.space_before = Pt(2)
            p_prompt.paragraph_format.space_after = Pt(2)
            p_prompt.paragraph_format.line_spacing = 1.15
            add_inline_math(p_prompt, prompt, font_size=Pt(11))

        # Subs
        if subs:
            for s in subs:
                p_sub = doc.add_paragraph()
                p_sub.paragraph_format.left_indent = Inches(0.25)
                p_sub.paragraph_format.space_before = Pt(1)
                p_sub.paragraph_format.space_after = Pt(2)
                p_sub.paragraph_format.line_spacing = 1.15
                add_inline_math(p_sub, s, font_size=Pt(11))

        # Images if any
        if img_dir:
            img_path = os.path.join(img_dir, f"hinh_{idx}.png")
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(4)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Inches(3.6))
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                rcap = p_cap.add_run(f"Hình {idx}")
                rcap.font.name = "Times New Roman"
                rcap.font.size = Pt(9.5)
                rcap.italic = True
                rcap.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

        # 1. Hướng dẫn giải & Gợi ý tư duy (Amber Box)
        if guide_lines:
            tbl_g = doc.add_table(rows=1, cols=1)
            tbl_g.alignment = WD_ALIGN_PARAGRAPH.CENTER
            tbl_g.autofit = False
            cg = tbl_g.cell(0, 0)
            cg.width = Inches(6.5)
            set_cell_background(cg, "FFFBEB") # Amber 50
            set_cell_left_border(cg, "D97706", sz="18") # Amber 600
            set_cell_margins(cg, top=70, bottom=70, left=140, right=140)

            p_g_hdr = cg.paragraphs[0]
            p_g_hdr.paragraph_format.space_before = Pt(1)
            p_g_hdr.paragraph_format.space_after = Pt(2)
            r_g_title = p_g_hdr.add_run("► Hướng dẫn giải & Gợi ý tư duy:")
            r_g_title.bold = True
            r_g_title.font.name = "Times New Roman"
            r_g_title.font.size = Pt(11)
            r_g_title.font.color.rgb = RGBColor(0xB4, 0x53, 0x09)

            for line in guide_lines:
                if not line.strip():
                    continue
                pl = cg.add_paragraph()
                pl.paragraph_format.space_before = Pt(1)
                pl.paragraph_format.space_after = Pt(1)
                pl.paragraph_format.line_spacing = 1.15
                add_inline_math(pl, line, font_size=Pt(10.5))

        # Space between boxes
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(2)
        p_sp.paragraph_format.space_after = Pt(2)

        # 2. Lời giải chi tiết (Emerald Box)
        if solution_lines:
            tbl_s = doc.add_table(rows=1, cols=1)
            tbl_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
            tbl_s.autofit = False
            cs = tbl_s.cell(0, 0)
            cs.width = Inches(6.5)
            set_cell_background(cs, "F0FDF4") # Emerald 50
            set_cell_left_border(cs, "059669", sz="18") # Emerald 600
            set_cell_margins(cs, top=70, bottom=70, left=140, right=140)

            p_s_hdr = cs.paragraphs[0]
            p_s_hdr.paragraph_format.space_before = Pt(1)
            p_s_hdr.paragraph_format.space_after = Pt(2)
            r_s_title = p_s_hdr.add_run("► Lời giải chi tiết:")
            r_s_title.bold = True
            r_s_title.font.name = "Times New Roman"
            r_s_title.font.size = Pt(11.5)
            r_s_title.font.color.rgb = RGBColor(0x04, 0x78, 0x57)

            for line in solution_lines:
                if not line.strip():
                    continue
                ps = cs.add_paragraph()
                ps.paragraph_format.space_before = Pt(1)
                ps.paragraph_format.space_after = Pt(1)
                ps.paragraph_format.line_spacing = 1.15
                add_inline_math(ps, line, font_size=Pt(10.5))

        # Divider
        p_sep = doc.add_paragraph()
        p_sep.paragraph_format.space_before = Pt(4)
        p_sep.paragraph_format.space_after = Pt(8)
        r_sep = p_sep.add_run("· · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · · ·")
        r_sep.font.size = Pt(8)
        r_sep.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)

    doc.save(output_docx_path)
    print(f"Creative Document saved successfully: {output_docx_path} ({os.path.getsize(output_docx_path)} bytes)")

if __name__ == '__main__':
    print("build_creative_docx ready.")
