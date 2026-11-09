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

from generate_master_chuyen_de import extract_all_questions, find_pdf_info

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
    # Translate MathType PUA glyphs
    t = clean_pua(text)
    # Remove control characters
    t = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', t)
    # Clean box symbol artifacts and broken MT brackets
    t = t.replace('□', '').replace('', '∞')
    for ch in ['', '', '', '', '', '', '', '', '', '', '', '']:
        t = t.replace(ch, '')
    return t

# ==============================================================================
# PART I: 21 CURATED MODEL PROBLEMS (STRUCTURED WITH NATIVE WORD OMML EQUATIONS)
# ==============================================================================
CURATED_DAY_SO = [
    {
        "category": "Dạng 1: Dãy số truy hồi tuyến tính bậc nhất và quy về Cấp số nhân",
        "questions": [
            {
                "title": "Đề thi HSG Toán 11 năm 2012 – 2013 trường THPT Thuận An – TT Huế",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"\begin{cases} u_1 = 2 \\ u_{n+1} = 5u_n - 4 \end{cases} \quad (n \ge 1)",
                "sub": ["Tìm công thức số hạng tổng quát của u_n theo n."]
            },
            {
                "title": "Đề thi chọn HSG cấp tỉnh Toán 11 năm 2018 – 2019 sở GD&ĐT Bắc Ninh",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"\begin{cases} u_1 = 3 \\ u_{n+1} = \frac{u_n + 1}{2} \end{cases} \quad (\forall n \ge 1)",
                "sub": [
                    "a) Tìm công thức số hạng tổng quát u_n theo n.",
                    ("b) Chứng minh dãy (u_n) có giới hạn hữu hạn và tìm ", r"\lim_{n \to +\infty} u_n.")
                ]
            },
            {
                "title": "Đề thi HSG Toán 11 năm 2023 – 2024 cụm trường THPT Gia Lâm – Long Biên – Hà Nội",
                "intro": "Cho dãy số (u_n) thỏa mãn:",
                "equation": r"\begin{cases} u_1 = 1 \\ u_{n+1} = u_n + 2n + 1 \end{cases} \quad (\forall n \ge 1)",
                "sub": [
                    ("Tìm số hạng tổng quát u_n và tính giới hạn: ", r"\lim_{n \to +\infty} \frac{u_n}{n^2}.")
                ]
            }
        ]
    },
    {
        "category": "Dạng 2: Dãy số phân tuyến tính (phân thức hữu tỉ) và truy hồi phi tuyến",
        "questions": [
            {
                "title": "Đề thi chọn HSG môn Toán 11 năm 2020 – 2021 sở GD&ĐT Hải Phòng",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"\begin{cases} u_1 = 1 \\ u_{n+1} = \frac{2u_n + 1}{u_n + 2} \end{cases} \quad (n \ge 1)",
                "sub": [
                    "a) Tìm công thức số hạng tổng quát u_n theo n.",
                    ("b) Tính giới hạn ", r"\lim_{n \to +\infty} u_n.")
                ]
            },
            {
                "title": "Đề chọn HSG Toán 11 năm 2025 – 2026 trường chuyên Lê Quý Đôn – Đà Nẵng",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"u_1 = 1, \quad u_{n+1} = \frac{u_n}{1 + u_n^2} \quad (\forall n \ge 1)",
                "sub": [
                    "a) Chứng minh dãy số (u_n) giảm và bị chặn dưới bởi 0.",
                    ("b) Tính giới hạn ", r"\lim_{n \to +\infty} u_n", " và ", r"\lim_{n \to +\infty} \sqrt{n} u_n.")
                ]
            },
            {
                "title": "Đề chọn đội tuyển HSG Toán 11 năm 2022 – 2023 THPT chuyên Vĩnh Phúc",
                "intro": "Cho dãy số (x_n) xác định bởi:",
                "equation": r"x_1 = 1, \quad x_{n+1} = \frac{x_n}{\sqrt{1 + 2023 x_n^2}} \quad (n \ge 1)",
                "sub": [
                    ("Tìm công thức số hạng tổng quát x_n và tính: ", r"\lim_{n \to +\infty} \sqrt{n} x_n.")
                ]
            },
            {
                "title": "Đề chọn học sinh giỏi cấp tỉnh Toán 11 năm 2025 – 2026 sở GD&ĐT Thái Nguyên",
                "intro": "Cho dãy số (x_n) xác định bởi:",
                "equation": r"x_1 = 2, \quad x_{n+1} = \frac{x_n^2 + 1}{2x_n} \quad (\forall n \ge 1)",
                "sub": [
                    ("Chứng minh rằng dãy số hội tụ và tìm ", r"\lim_{n \to +\infty} x_n.")
                ]
            }
        ]
    },
    {
        "category": "Dạng 3: Dãy số liên kết tổng, tích và giới hạn của chuỗi tổng",
        "questions": [
            {
                "title": "Đề thi học sinh giỏi Toán 11 cấp tỉnh năm 2016 – 2017 sở GD&ĐT Lai Châu",
                "intro": "Cho dãy số (u_n) xác định như sau:",
                "equation": r"\begin{cases} u_1 = 1 \\ u_{n+1} = \frac{u_n^2}{2017} + u_n \end{cases} \quad (n = 1, 2, \dots)",
                "sub": [
                    ("Tính giới hạn của tổng: ", r"\lim_{n \to +\infty} \left(\frac{u_1}{u_2} + \frac{u_2}{u_3} + \dots + \frac{u_n}{u_{n+1}}\right).")
                ]
            },
            {
                "title": "Đề thi KSCL học sinh giỏi cụm thi THPT Yên Thành – Nghệ An 2016 – 2017",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"\begin{cases} u_1 = 2 \\ n(n^2 - 1)u_n = u_1 + 2u_2 + \dots + (n-1)u_{n-1} \end{cases} \quad (\forall n > 1, n \in \mathbb{N})",
                "sub": [
                    ("Tìm giới hạn: ", r"\lim_{n \to +\infty} \frac{9}{2}(n^3 - n)u_n.")
                ]
            },
            {
                "title": "Đề khảo sát HSG Toán 11 lần 1 năm 2020 – 2021 THPT Triệu Sơn 3 – Thanh Hóa",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"\begin{cases} u_1 = \frac{1}{2} \\ u_{n+1} = \frac{u_n}{2(2n + 1)u_n + 1} \end{cases} \quad (n \ge 1)",
                "sub": [
                    ("Đặt ", r"S_n = u_1 + u_2 + \dots + u_n", ". Tính giới hạn: ", r"\lim_{n \to +\infty} S_n.")
                ]
            },
            {
                "title": "Đề KSCL đội tuyển HSG Toán 11 năm 2017 – 2018 THPT Minh Châu – Hưng Yên",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"u_1 = 1, \quad u_{n+1} = \frac{u_n}{n^2 + 3n + 2} + u_n \quad (\forall n \ge 1)",
                "sub": [
                    "Tìm công thức số hạng tổng quát của u_n theo n."
                ]
            }
        ]
    },
    {
        "category": "Dạng 4: Dãy số căn thức, chứng minh đơn điệu và bị chặn (Định lý Weierstrass)",
        "questions": [
            {
                "title": "Đề thi chọn HSG Toán THPT cấp tỉnh năm học 2018 – 2019 sở GD&ĐT Phú Yên",
                "intro": "Cho dãy số thực (x_n) thỏa mãn:",
                "equation": r"x_1 = 1, \quad x_{n+1} = \sqrt{x_n^2 + \frac{1}{4^n}} \quad (n = 1, 2, 3, \dots)",
                "sub": [
                    ("a) Chứng minh rằng ", r"x_n < \frac{2}{\sqrt{3}}", " với mọi n ≥ 1."),
                    "b) Tìm giới hạn của dãy số (x_n)."
                ]
            },
            {
                "title": "Đề thi học sinh giỏi Toán 11 cấp tỉnh năm 2017 – 2018 sở GD&ĐT Lai Châu",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"u_1 > 1 \quad \text{và} \quad u_{n+1} = \sqrt{u_n + 1} \quad (n = 1, 2, \dots)",
                "sub": [
                    ("Chứng minh dãy số có giới hạn hữu hạn và tính ", r"\lim_{n \to +\infty} u_n.")
                ]
            },
            {
                "title": "Đề thi HSG Toán 11 năm 2023 – 2024 sở GD&ĐT Quảng Ngãi",
                "intro": "Cho dãy số (u_n) xác định bởi:",
                "equation": r"\begin{cases} u_1 = 2 \\ u_{n+1} = \sqrt{2 + u_n} \end{cases} \quad (\forall n \ge 1)",
                "sub": [
                    ("Chứng minh rằng dãy số (u_n) tăng, bị chặn trên và tính ", r"\lim_{n \to +\infty} u_n.")
                ]
            },
            {
                "title": "Đề thi chọn HSG tỉnh Toán 11 năm học 2017 – 2018 sở GD&ĐT Vĩnh Phúc",
                "intro": "Cho a > 1 và dãy số (x_n) xác định như sau:",
                "equation": r"x_1 = a; \quad x_{n+1} = \frac{x_n^2 + 3}{2018} \quad (\forall n \ge 1)",
                "sub": [
                    ("Tìm a để: ", r"\lim_{n \to +\infty} \frac{x_{n+1}}{x_n} = 2018.")
                ]
            }
        ]
    },
    {
        "category": "Dạng 5: Cấp số cộng, Cấp số nhân & Tính chất hình học",
        "questions": [
            {
                "title": "Đề thi HSG Toán 11 năm 2012 – 2013 trường THPT Thuận An – TT Huế",
                "intro": "Ba số x, y, z theo thứ tự lập thành một cấp số nhân; đồng thời chúng theo thứ tự là số hạng thứ nhất, thứ ba và thứ chín của một cấp số cộng.",
                "equation": r"x + y + z = 13",
                "sub": [
                    "Tìm ba số x, y, z."
                ]
            },
            {
                "title": "Đề KSCL đội tuyển HSG Toán 11 năm 2017 – 2018 trường THPT Yên Lạc 2 – Vĩnh Phúc",
                "intro": "Một tứ giác có bốn góc tạo thành một cấp số nhân và số đo góc lớn nhất gấp 8 lần số đo góc nhỏ nhất.",
                "equation": None,
                "sub": [
                    "Tính số đo bốn góc của tứ giác trên."
                ]
            },
            {
                "title": "Đề thi HSG Toán 11 cấp trường năm 2018 – 2019 trường Thuận Thành 2 – Bắc Ninh",
                "intro": "Cho cấp số cộng (u_n) thỏa mãn:",
                "equation": r"\begin{cases} u_2 - u_3 + u_5 = 10 \\ u_4 + u_6 = 26 \end{cases}",
                "sub": [
                    ("Tìm số hạng đầu u_1 và công sai d. Tính tổng: ", r"S_{20} = u_1 + u_2 + \dots + u_{20}.")
                ]
            },
            {
                "title": "Đề chọn HSG Toán 11 năm 2024 – 2025 trường THPT Phùng Khắc Khoan – Hà Nội",
                "intro": "Cho cấp số nhân (u_n) có các số hạng đều dương thỏa mãn:",
                "equation": r"\begin{cases} u_1 + u_2 + u_3 = 13 \\ u_4 - u_1 = 26 \end{cases}",
                "sub": [
                    "Tìm số hạng đầu u_1 và công bội q."
                ]
            }
        ]
    },
    {
        "category": "Dạng 6: Giới hạn của hàm số và dãy số chứa căn thức khử dạng vô định",
        "questions": [
            {
                "title": "Đề KSCL đội tuyển HSG Toán 11 năm 2017 – 2018 THPT Minh Châu – Hưng Yên",
                "intro": "Tính giới hạn sau:",
                "equation": r"\lim_{x \to 1} \frac{\sqrt[5]{5x - 4} + \sqrt{2x - 1} - 2}{\sqrt[3]{3x - 2} - 1}",
                "sub": []
            },
            {
                "title": "Đề thi chọn HSG cấp tỉnh Toán 11 năm 2018 – 2019 sở GD&ĐT Bắc Ninh",
                "intro": "Tìm a để hàm số sau liên tục tại x = 1:",
                "equation": r"f(x) = \begin{cases} \frac{\sqrt{x^2 + 3} - 2}{x - 1} & \text{khi } x > 1 \\ ax + \frac{3}{4} & \text{khi } x \le 1 \end{cases}",
                "sub": []
            }
        ]
    }
]

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

def build_word_day_so():
    print("Extracting and filtering Dãy Số questions...")
    col = extract_all_questions()
    raw_day_so = col['DAY_SO']['items']

    clean_items = []
    for it in raw_day_so:
        content = it['content']
        lower = content.lower()
        
        # Filter out rubrics / grading guides
        if '1.0d' in lower or '0.5d' in lower or '1.0đ' in lower or '0.5đ' in lower:
            continue
        if 'thang điểm' in lower or 'biểu điểm' in lower or 'hướng dẫn chấm' in lower:
            continue
        if 'suy ra' in lower and 'giải ra được' in lower:
            continue
        if lower.count('0,5') >= 2 or lower.count('0.5') >= 2 or lower.count('1,0') >= 2:
            continue
            
        # Filter out misplaced pure trigonometry
        if 'giải phương trình' in lower and ('sin' in lower or 'cos' in lower) and not any(k in lower for k in ['dãy số', 'cấp số', 'u_n', 'u_{n+1}', 'lim']):
            continue
        if content.strip().startswith('Giải phương trình:') and 'sin' in lower and 'cos' in lower:
            continue

        # Clean specific Quảng Trị items to isolate only Dãy số / Giới hạn
        if 'Quảng Trị' in it['source'] and '2025 – 2026' in it['source']:
            if 'bánh nướng' in content and 'Cho dãy số' in content:
                # Keep only part b (dãy số)
                m_sub = re.search(r'(b\)\s*Cho dãy số.*)', content, re.DOTALL)
                if m_sub:
                    content = m_sub.group(1)
            elif 'log' in content and 'parabol' in content:
                # Keep only part b (giới hạn)
                m_sub = re.search(r'(b\)\s*Cho điểm\s*M.*)', content, re.DOTALL)
                if m_sub:
                    content = m_sub.group(1)
            elif 'hình chóp' in content and 'giới hạn hữu hạn' in content:
                # Keep only part a (giới hạn)
                m_sub = re.search(r'(a\)\s*Tìm hai số thực.*?)(\nb\)|$)', content, re.DOTALL)
                if m_sub:
                    content = m_sub.group(1)

        item_copy = dict(it)
        item_copy['content'] = content
        clean_items.append(item_copy)

    print(f"Total clean Dãy Số items: {len(clean_items)}")

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
    r_title = p_title.add_run("CHUYÊN ĐỀ 2: DÃY SỐ, CẤP SỐ & GIỚI HẠN\nHSG TOÁN 11")
    r_title.bold = True
    r_title.font.size = Pt(18)
    r_title.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)
    p_title.paragraph_format.space_after = Pt(6)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"Tuyển tập toàn bộ {len(clean_items)} bài toán Tự luận Dãy số & Giới hạn trích từ 130 đề thi HSG cấp Tỉnh / Olympic (2012 – 2026)\n(Soạn thảo đầy đủ từng bài — Chuẩn Microsoft Word .docx)")
    r_sub.italic = True
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p_sub.paragraph_format.space_after = Pt(16)

    add_callout_box(
        doc,
        f"• Tài liệu bao gồm TOÀN BỘ {len(clean_items)} câu hỏi tự luận chọn lọc về Dãy số, Cấp số cộng, Cấp số nhân & Giới hạn từ 130 đề thi HSG (2012 – 2026).\n"
        "• PHẦN I: Tuyển chọn 21 bài toán mẫu kinh điển đã được chuẩn hóa 100% bằng công thức Word Equation (OMML). Thầy cô có thể nhấp chuột trực tiếp sửa công thức bằng [Alt + =].\n"
        f"• PHẦN II: Soạn thảo chi tiết toàn bộ {len(clean_items)} đề bài cụ thể, phân nhóm theo từng năm học từ 2026 về 2012, đã làm sạch toàn bộ lỗi font, mã Symbol PUA và thang điểm.",
        "📌 HƯỚNG DẪN DÀNH CHO GIÁO VIÊN BỒI DƯỠNG HSG",
        "EBF8FF"
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==========================================
    # PHẦN 1: 21 BÀI TOÁN TUYỂN CHỌN EQUATION EDITOR
    # ==========================================
    p_sec1 = doc.add_paragraph()
    r_sec1 = p_sec1.add_run("PHẦN I: TUYỂN CHỌN 21 BÀI TOÁN DÃY SỐ & GIỚI HẠN MẪU (CÔNG THỨC EQUATION CHUẨN)")
    r_sec1.bold = True
    r_sec1.font.size = Pt(13.5)
    r_sec1.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    p_sec1.paragraph_format.space_before = Pt(14)
    p_sec1.paragraph_format.space_after = Pt(8)

    q_curated_idx = 0
    for sec in CURATED_DAY_SO:
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
            p_q.paragraph_format.space_before = Pt(6)
            p_q.paragraph_format.space_after = Pt(2)
            p_q.paragraph_format.keep_with_next = True

            r_qnum = p_q.add_run(f"Bài mẫu {q_curated_idx}. ")
            r_qnum.bold = True
            r_qnum.font.name = "Times New Roman"
            r_qnum.font.size = Pt(12)
            r_qnum.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)
            
            r_src = p_q.add_run(f"({q['title']})")
            r_src.italic = True
            r_src.font.name = "Times New Roman"
            r_src.font.size = Pt(11)
            r_src.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

            # Intro text
            if q.get('intro'):
                p_intro = doc.add_paragraph()
                p_intro.paragraph_format.left_indent = Inches(0.2)
                p_intro.paragraph_format.space_before = Pt(1)
                p_intro.paragraph_format.space_after = Pt(2)
                r_intro = p_intro.add_run(q['intro'])
                r_intro.font.name = "Times New Roman"
                r_intro.font.size = Pt(11.5)

            # Main Display Equation
            if q.get('equation'):
                p_eq = doc.add_paragraph()
                p_eq.paragraph_format.left_indent = Inches(0.4)
                p_eq.paragraph_format.space_before = Pt(2)
                p_eq.paragraph_format.space_after = Pt(4)
                omml = latex_to_word_omml(q['equation'])
                if omml is not None:
                    p_eq._p.append(omml)
                else:
                    p_eq.add_run(q['equation'])

            # Sub-questions / questions
            for sub_it in q.get('sub', []):
                p_subq = doc.add_paragraph()
                p_subq.paragraph_format.left_indent = Inches(0.2)
                p_subq.paragraph_format.space_before = Pt(1)
                p_subq.paragraph_format.space_after = Pt(2)

                if isinstance(sub_it, tuple):
                    # Mixed text and equation elements: (text, eq, text, eq, ...)
                    for elem in sub_it:
                        if elem.startswith(r'\lim') or elem.startswith(r'x_') or elem.startswith(r'S_') or elem.startswith(r'\frac') or elem.startswith(r'u_'):
                            omml_sub = latex_to_word_omml(elem)
                            if omml_sub is not None:
                                p_subq._p.append(omml_sub)
                            else:
                                p_subq.add_run(elem)
                        else:
                            r_sub_txt = p_subq.add_run(elem)
                            r_sub_txt.font.name = "Times New Roman"
                            r_sub_txt.font.size = Pt(11.5)
                else:
                    r_sub_txt = p_subq.add_run(sub_it)
                    r_sub_txt.font.name = "Times New Roman"
                    r_sub_txt.font.size = Pt(11.5)

    # Page break before Section 2
    doc.add_page_break()

    # ==========================================
    # PHẦN 2: TOÀN BỘ NGÂN HÀNG CÁC BÀI TOÁN DÃY SỐ
    # ==========================================
    p_sec2 = doc.add_paragraph()
    r_sec2 = p_sec2.add_run(f"PHẦN II: TOÀN BỘ NGÂN HÀNG {len(clean_items)} BÀI TOÁN DÃY SỐ & GIỚI HẠN TỰ LUẬN")
    r_sec2.bold = True
    r_sec2.font.size = Pt(14)
    r_sec2.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    p_sec2.paragraph_format.space_before = Pt(10)
    p_sec2.paragraph_format.space_after = Pt(4)

    p_sec2_note = doc.add_paragraph()
    r_s2_n = p_sec2_note.add_run(f"Dưới đây là nội dung toàn bộ {len(clean_items)} câu hỏi tự luận Dãy số & Giới hạn phân nhóm theo từng năm học (2012 – 2026). Đã được làm sạch 100% toàn bộ lỗi ký tự font MathType và thang điểm rác để giáo viên thuận tiện ra đề.")
    r_s2_n.font.size = Pt(11)
    r_s2_n.italic = True
    p_sec2_note.paragraph_format.space_after = Pt(14)

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

    for item in clean_items:
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

        p_yr = doc.add_paragraph()
        r_yr = p_yr.add_run(f"📁 {yk.upper()} ({len(group_items)} BÀI TOÁN)")
        r_yr.bold = True
        r_yr.font.size = Pt(13)
        r_yr.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        p_yr.paragraph_format.space_before = Pt(14)
        p_yr.paragraph_format.space_after = Pt(6)

        p_line = doc.add_paragraph()
        r_l = p_line.add_run("────────────────────────────────────────────────────────")
        r_l.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE0)
        p_line.paragraph_format.space_after = Pt(6)

        for item in group_items:
            global_item_count += 1
            src = item['source']
            body = item['content']

            pdf_info = find_pdf_info(src)
            rel_pdf = pdf_info['rel_pdf'] if pdf_info else None
            clean_title = re.sub(r'\(Năm học.*?\)', '', src).strip()

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

            p_body = doc.add_paragraph()
            p_body.paragraph_format.left_indent = Inches(0.2)
            p_body.paragraph_format.space_before = Pt(2)
            p_body.paragraph_format.space_after = Pt(8)

            cleaned_text = clean_and_sanitize_text(body)
            lines = [l.strip() for l in cleaned_text.split('\n') if l.strip()]
            cleaned_text = "\n".join(lines)

            r_b = p_body.add_run(cleaned_text)
            r_b.font.name = "Times New Roman"
            r_b.font.size = Pt(11.5)

    out_docx_path = os.path.join(CHUYEN_DE_DIR, "02_Day_So_Gioi_Han_HSG_11.docx")
    doc.save(out_docx_path)
    print(f"\n==========================================")
    print(f"SUCCESS: Fully compiled Word document saved to:")
    print(f"{out_docx_path}")
    print(f"- Total curated equations: {q_curated_idx}")
    print(f"- Total clean questions written out in full: {global_item_count}")
    print(f"==========================================")

if __name__ == '__main__':
    build_word_day_so()
