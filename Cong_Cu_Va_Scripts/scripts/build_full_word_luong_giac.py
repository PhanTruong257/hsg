import os
import re
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import latex2mathml.converter
from lxml import etree

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = r'c:\Users\truongpv\Downloads\Nguyệt'
CHUYEN_DE_DIR = os.path.join(ROOT_DIR, 'Chuyen_De_HSG_Toan_11')
TU_LUAN_DIR = os.path.join(ROOT_DIR, 'downloads_hsg_toan11_tu_luan')

# Load XSLT for OMML
xslt_path = r'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL'
xslt_doc = etree.parse(xslt_path)
transform = etree.XSLT(xslt_doc)

def latex_to_word_omml(latex_code):
    try:
        mathml = latex2mathml.converter.convert(latex_code)
        tree = etree.fromstring(mathml)
        omml = transform(tree)
        return omml.getroot()
    except Exception as e:
        return None

from generate_master_chuyen_de import CURATED_LUONG_GIAC, extract_all_questions, find_pdf_info

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout_box(doc, text, title=None, fill_color="F0F4F8"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, fill_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="1A365D"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if title:
        run_t = p.add_run(f"{title}\n")
        run_t.bold = True
        run_t.font.name = "Times New Roman"
        run_t.font.size = Pt(11)
        run_t.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    run_txt = p.add_run(text)
    run_txt.font.name = "Times New Roman"
    run_txt.font.size = Pt(11)

# ==============================================================================
# PUA & SYMBOL TRANSLATION MAP (Resolves all font encoding / □ issues)
# ==============================================================================
SYMBOL_MAP = {
    0x2B: '+', 0x2D: '-', 0x3D: '=', 0x28: '(', 0x29: ')', 0x3C: '<', 0x3E: '>',
    0x2F: '/', 0x3A: ':', 0x3B: ';', 0x5B: '[', 0x5D: ']', 0x7B: '{', 0x7D: '}',
    0x2C: ',', 0x2E: '.', 0x2A: '*', 0x24: '$', 0x25: '%', 0x22: '"',
    0x70: 'π', 0x61: 'α', 0x62: 'β', 0x67: 'γ', 0x64: 'δ', 0x65: 'ε',
    0x71: 'θ', 0x6C: 'λ', 0x6D: 'μ', 0x73: 'σ', 0x74: 'τ', 0x66: 'φ', 0x77: 'ω',
    0x44: 'Δ', 0x47: 'Γ', 0x4C: 'Λ', 0x53: 'Σ', 0x57: 'Ω', 0x50: 'Π',
    0xA1: '', 0xA2: "'", 0xA3: '≤', 0xA5: '∞', 0xAE: '→', 0xB0: '°',
    0xB1: '±', 0xB2: "''", 0xB3: '≥', 0xB4: '×', 0xB7: '·', 0xB9: '≠',
    0xBA: '≡', 0xBB: '≈', 0xBC: '…', 0xC6: '∅', 0xC7: '∩', 0xC8: '∪',
    0xCE: '∈', 0xCF: '∉', 0xD0: '∠', 0xD6: '√', 0xD7: '·', 0xDB: '⇔',
    0xDE: '⇒', 0xE5: '∑', 0xF2: '∫'
}

BRACKET_PUA = {
    0xF0E6: '(', 0xF0E7: '', 0xF0E8: ')', 0xF0F6: ')', 0xF0F7: '', 0xF0F8: ')',
    0xF0E9: '[', 0xF0EA: '', 0xF0EB: ']', 0xF0F9: ']', 0xF0FA: '', 0xF0FB: ']',
    0xF0EC: '{', 0xF0ED: '', 0xF0EE: '}', 0xF0EF: '', 0xF0FC: '}', 0xF0FD: '', 0xF0FE: '}',
    0xF8F1: '{', 0xF8F2: '', 0xF8F3: '}', 0xF8F4: '',
    0xF8EB: '[', 0xF8EC: '', 0xF8ED: ']', 0xF8F6: ']', 0xF8F7: '', 0xF8F8: ']'
}

def clean_pua(text):
    res = []
    for ch in text:
        code = ord(ch)
        if 0xF000 <= code <= 0xF0FF:
            offset = code - 0xF000
            if offset in SYMBOL_MAP:
                res.append(SYMBOL_MAP[offset])
            elif code in BRACKET_PUA:
                res.append(BRACKET_PUA[code])
            else:
                res.append('')
        elif code in BRACKET_PUA:
            res.append(BRACKET_PUA[code])
        elif 0xF000 <= code <= 0xF8FF:
            res.append('')
        else:
            res.append(ch)
    return ''.join(res)

def clean_and_sanitize_text(text):
    t = clean_pua(text)
    t = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', t)
    t = t.replace('□', '').replace('', '∞')
    for ch in ['', '', '', '', '', '', '', '', '', '', '', '']:
        t = t.replace(ch, '')
    return t

def sanitize_xml(text):
    return clean_and_sanitize_text(text)

def build_full_document():
    print("Extracting all questions from 130 exam PDFs...")
    collections = extract_all_questions()
    all_lg_items = collections['LUONG_GIAC']['items']
    print(f"Total Luong Giac items extracted: {len(all_lg_items)}")

    doc = Document()

    # Standard A4 Margins
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(4)

    # Header / Title Block
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst1 = p_inst.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO — TỔ TOÁN THPT\n")
    r_inst1.bold = True
    r_inst1.font.size = Pt(11)
    r_inst1.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    r_inst2 = p_inst.add_run("NGÂN HÀNG ĐỀ THI HỌC SINH GIỎI TOÁN LỚP 11 TOÀN QUỐC")
    r_inst2.bold = True
    r_inst2.font.size = Pt(12)
    r_inst2.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    r_div.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("CHUYÊN ĐỀ 1: PHƯƠNG TRÌNH & HỆ THỨC LƯỢNG GIÁC\nHSG TOÁN 11")
    r_title.bold = True
    r_title.font.size = Pt(18)
    r_title.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)
    p_title.paragraph_format.space_after = Pt(6)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"Tuyển tập toàn bộ {len(all_lg_items)} bài toán Tự luận Lượng giác trích từ 130 đề thi HSG cấp Tỉnh / Olympic (2012 – 2026)\n(Soạn thảo đầy đủ từng bài — Chuẩn Microsoft Word .docx)")
    r_sub.italic = True
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p_sub.paragraph_format.space_after = Pt(16)

    add_callout_box(
        doc,
        f"• Tài liệu bao gồm TOÀN BỘ {len(all_lg_items)} câu hỏi Lượng giác trích xuất từ 130 đề thi HSG cấp Tỉnh/TP và Olympic (2012 – 2026).\n"
        "• PHẦN I: Tuyển chọn 21 bài toán trọng tâm đã được chuẩn hóa công thức toán bằng Microsoft Word Equation (OMML). Thầy cô có thể nhấp chuột sửa công thức trực tiếp bằng [Alt + =].\n"
        f"• PHẦN II: Soạn thảo chi tiết toàn bộ {len(all_lg_items)} đề bài cụ thể, phân nhóm khoa học theo từng năm học từ 2026 về 2012.",
        "📌 HƯỚNG DẪN DÀNH CHO GIÁO VIÊN BỒI DƯỠNG HSG",
        "EBF8FF"
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==========================================
    # PHẦN 1: 21 BÀI TOÁN TUYỂN CHỌN EQUATION EDITOR
    # ==========================================
    p_sec1 = doc.add_paragraph()
    r_sec1 = p_sec1.add_run("PHẦN I: TUYỂN CHỌN 21 BÀI TOÁN LƯỢNG GIÁC MẪU (CÔNG THỨC EQUATION CHUẨN)")
    r_sec1.bold = True
    r_sec1.font.size = Pt(13.5)
    r_sec1.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    p_sec1.paragraph_format.space_before = Pt(14)
    p_sec1.paragraph_format.space_after = Pt(8)

    q_curated_idx = 0
    for sec in CURATED_LUONG_GIAC:
        p_cat = doc.add_paragraph()
        r_cat = p_cat.add_run(f"❖ {sec['category']}")
        r_cat.bold = True
        r_cat.font.size = Pt(12.5)
        r_cat.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
        p_cat.paragraph_format.space_before = Pt(8)
        p_cat.paragraph_format.space_after = Pt(4)

        for q in sec['questions']:
            q_curated_idx += 1
            p_q = doc.add_paragraph()
            r_qnum = p_q.add_run(f"Bài mẫu {q_curated_idx}. ")
            r_qnum.bold = True
            r_qnum.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)
            
            r_src = p_q.add_run(f"({q['title']})\n")
            r_src.italic = True
            r_src.font.size = Pt(11)
            r_src.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            
            latex_lines = q['latex'].split(r'\\')
            for line in latex_lines:
                line = line.strip()
                if not line:
                    continue
                m_txt = re.match(r'\\text\{([^}]+)\}\s*(.*)', line)
                p_math = doc.add_paragraph()
                p_math.paragraph_format.left_indent = Inches(0.3)
                p_math.paragraph_format.space_before = Pt(2)
                p_math.paragraph_format.space_after = Pt(4)
                
                if m_txt:
                    txt_prefix = m_txt.group(1).strip()
                    math_part = m_txt.group(2).strip()
                    r_pre = p_math.add_run(txt_prefix + " ")
                    r_pre.font.name = "Times New Roman"
                    r_pre.font.size = Pt(12)
                    
                    if math_part:
                        omml = latex_to_word_omml(math_part)
                        if omml is not None:
                            p_math._p.append(omml)
                        else:
                            p_math.add_run(math_part)
                else:
                    omml = latex_to_word_omml(line)
                    if omml is not None:
                        p_math._p.append(omml)
                    else:
                        p_math.add_run(line)

            p_q.paragraph_format.space_after = Pt(4)

    # Page break before Section 2
    doc.add_page_break()

    # ==========================================
    # PHẦN 2: TOÀN BỘ NGÂN HÀNG 101 BÀI TOÁN LƯỢNG GIÁC
    # ==========================================
    p_sec2 = doc.add_paragraph()
    r_sec2 = p_sec2.add_run(f"PHẦN II: TOÀN BỘ NGÂN HÀNG {len(all_lg_items)} BÀI TOÁN LƯỢNG GIÁC TỰ LUẬN")
    r_sec2.bold = True
    r_sec2.font.size = Pt(14)
    r_sec2.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    p_sec2.paragraph_format.space_before = Pt(10)
    p_sec2.paragraph_format.space_after = Pt(4)

    p_sec2_note = doc.add_paragraph()
    r_s2_n = p_sec2_note.add_run(f"Dưới đây là nội dung toàn bộ {len(all_lg_items)} câu hỏi tự luận Lượng giác được phân nhóm chi tiết theo từng năm học. Mỗi bài đều có thông tin trường/tỉnh, điểm số và đề bài cụ thể để giáo viên sao chép hoặc in ấn ra đề.")
    r_s2_n.font.size = Pt(11)
    r_s2_n.italic = True
    p_sec2_note.paragraph_format.space_after = Pt(14)

    # Group questions by School Year
    year_order = [
        "Năm học 2025 - 2026",
        "Năm học 2024 - 2025",
        "Năm học 2023 - 2024",
        "Năm học 2022 - 2023",
        "Năm học 2021 - 2022",
        "Năm học 2020 - 2021",
        "Năm học 2019 - 2020",
        "Năm học 2018 - 2019",
        "Năm học 2017 - 2018",
        "Năm học 2016 - 2017",
        "Năm học 2012 - 2013",
        "Tuyển tập tổng hợp"
    ]

    grouped_by_year = {y: [] for y in year_order}
    grouped_by_year["Khác"] = []

    for item in all_lg_items:
        src = item['source']
        m_yr = re.search(r'\((Năm học.*?)\)', src)
        assigned = False
        if m_yr:
            yr_str = m_yr.group(1)
            for yk in year_order:
                if yk in yr_str:
                    grouped_by_year[yk].append(item)
                    assigned = True
                    break
        if not assigned:
            grouped_by_year["Khác"].append(item)

    global_item_count = 0
    for yk in year_order + ["Khác"]:
        group_items = grouped_by_year.get(yk, [])
        if not group_items:
            continue

        # Year Section Heading
        p_yr = doc.add_paragraph()
        r_yr = p_yr.add_run(f"📁 {yk.upper()} ({len(group_items)} BÀI TOÁN)")
        r_yr.bold = True
        r_yr.font.size = Pt(13)
        r_yr.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        p_yr.paragraph_format.space_before = Pt(14)
        p_yr.paragraph_format.space_after = Pt(6)

        # Decorative line
        p_line = doc.add_paragraph()
        r_l = p_line.add_run("────────────────────────────────────────────────────────")
        r_l.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE0)
        p_line.paragraph_format.space_after = Pt(6)

        for item in group_items:
            global_item_count += 1
            src = item['source']
            body = item['content']

            # Find matching PDF info
            pdf_info = find_pdf_info(src)
            rel_pdf = pdf_info['rel_pdf'] if pdf_info else None
            clean_title = re.sub(r'\(Năm học.*?\)', '', src).strip()

            # Question Card Paragraph
            p_item_header = doc.add_paragraph()
            p_item_header.paragraph_format.space_before = Pt(8)
            p_item_header.paragraph_format.space_after = Pt(2)
            p_item_header.paragraph_format.keep_with_next = True

            r_inum = p_item_header.add_run(f"Bài {global_item_count}. ")
            r_inum.bold = True
            r_inum.font.name = "Times New Roman"
            r_inum.font.size = Pt(11.5)
            r_inum.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

            r_isrc = p_item_header.add_run(f"[{clean_title}]")
            r_isrc.bold = True
            r_isrc.font.name = "Times New Roman"
            r_isrc.font.size = Pt(11)
            r_isrc.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

            # Question Content Box / Paragraph
            p_body = doc.add_paragraph()
            p_body.paragraph_format.left_indent = Inches(0.2)
            p_body.paragraph_format.space_before = Pt(2)
            p_body.paragraph_format.space_after = Pt(8)

            lines = [l.strip() for l in body.split('\n') if l.strip()]
            cleaned_text = "\n".join(lines)
            cleaned_text = sanitize_xml(cleaned_text)

            r_b = p_body.add_run(cleaned_text)
            r_b.font.name = "Times New Roman"
            r_b.font.size = Pt(11.5)

    # Save to Word file
    out_docx_path = os.path.join(CHUYEN_DE_DIR, "01_Luong_Giac_HSG_11.docx")
    doc.save(out_docx_path)
    print(f"\n==========================================")
    print(f"SUCCESS: Fully compiled Word document saved to:")
    print(f"{out_docx_path}")
    print(f"- Total curated equations: {q_curated_idx}")
    print(f"- Total questions written out in full: {global_item_count}")
    print(f"==========================================")

if __name__ == '__main__':
    build_full_document()
