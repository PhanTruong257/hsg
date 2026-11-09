import os
import re
import sys
from pypdf import PdfReader

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = r'c:\Users\truongpv\Downloads\Nguyệt'
TU_LUAN_DIR = os.path.join(ROOT_DIR, 'downloads_hsg_toan11_tu_luan')
OUT_DIR = os.path.join(ROOT_DIR, 'Chuyen_De_HSG_Toan_11')
PREVIEW_DIR = os.path.join(ROOT_DIR, 'de_thi_previews')

os.makedirs(OUT_DIR, exist_ok=True)

# Index PDF files and preview images
pdf_files = {}
for root, dirs, files in os.walk(TU_LUAN_DIR):
    for f in files:
        if f.endswith('.pdf'):
            full_path = os.path.join(root, f)
            rel_pdf = os.path.relpath(full_path, OUT_DIR).replace('\\', '/')
            base = f.replace('.pdf', '')
            clean_title = base.replace('_', ' ')
            
            preview_name = f"{base}_trang1.png"
            preview_path = os.path.join(PREVIEW_DIR, preview_name)
            rel_preview = f"../de_thi_previews/{preview_name}" if os.path.exists(preview_path) else None
            
            pdf_files[clean_title] = {
                'filename': f,
                'rel_pdf': rel_pdf,
                'rel_preview': rel_preview,
                'clean_title': clean_title
            }

def find_pdf_info(source_str):
    title_clean = re.sub(r'\(Năm học.*?\)', '', source_str).replace('🏷️', '').strip()
    for k, v in pdf_files.items():
        if title_clean in k or k in title_clean:
            return v
    words = [w for w in title_clean.split() if len(w) > 3]
    for k, v in pdf_files.items():
        if sum(1 for w in words if w in k) >= len(words) - 2:
            return v
    return None

def classify_question(text):
    tl = text.lower()
    
    # 1. Xác suất & Tổ hợp
    if any(k in tl for k in ['xác suất', 'chọn ngẫu nhiên', 'biến cố', 'hoán vị', 'chỉnh hợp', 'tổ hợp', 'nhị thức newton', 'xúc xắc', 'chữ số phân biệt', 'sắp xếp']):
        return 'XAC_SUAT'
        
    # 2. Dãy số & Giới hạn
    if any(k in tl for k in ['dãy số', 'cấp số cộng', 'cấp số nhân', 'giới hạn', 'tính lim', 'u_n', 'u_{n+1}']) or re.search(r'\blim\b', tl):
        return 'DAY_SO'
        
    # 3. Hình học không gian
    if any(k in tl for k in ['hình chóp', 'lăng trụ', 'tứ diện', 'khoảng cách', 'góc giữa đường thẳng', 'góc giữa hai mặt phẳng', 'thiết diện']) and not any(k in tl for k in ['giải phương trình', 'nghiệm']):
        return 'HINH_KHONG_GIAN'

    # 4. Lượng giác
    has_trig = False
    if any(k in tl for k in ['lượng giác', 'nghiệm trong đoạn', 'nghiệm thuộc khoảng']):
        has_trig = True
    if ('sin' in text and 'cos' in text) or ('sin' in text and 'tan' in text) or ('cos' in text and 'tan' in text):
        has_trig = True
    if re.search(r'\b(sin|cos|tan|cot)\s*[\(\d2345x]', text, re.IGNORECASE):
        has_trig = True
    if has_trig:
        return 'LUONG_GIAC'

    return 'KHAC'

CURATED_LUONG_GIAC = [
    {
        "category": "Dạng 1: Phương trình biến đổi tích thành tổng, quy về tích & hạ bậc",
        "questions": [
            {
                "title": "Đề thi HSG Toán 11 năm 2012 – 2013 trường THPT Thuận An – TT Huế",
                "latex": r"""\text{Giải các phương trình sau:}\\
1) \quad 4\sin x\cos x + \tan\left(\frac{10\pi}{3}\right)(2\sin^2 x - 1) = 2\sin 3x\\
2) \quad 1 + \sin\frac{x}{2}\sin x - \cos\frac{x}{2}\sin^2 x = 2\cos^2\left(\frac{\pi}{4} - \frac{x}{2}\right)""",
                "source_kw": "Thuận An"
            },
            {
                "title": "Đề thi HSG cấp tỉnh Toán 11 năm 2016 – 2017 sở GD&ĐT Lai Châu",
                "latex": r"""\text{Giải phương trình: } \sin 2x + 2\cos 2x = 1 + \sin x - 4\cos x""",
                "source_kw": "Lai Châu"
            },
            {
                "title": "Đề thi chọn HSG cấp tỉnh Toán 11 năm 2018 – 2019 sở GD&ĐT Bắc Ninh",
                "latex": r"""\text{Giải phương trình: } \sin^3 x + \cos^3 x = \sin x + \cos x + \sqrt{2}\sin 2x""",
                "source_kw": "Bắc Ninh"
            },
            {
                "title": "Đề giao lưu HSG Toán 11 cụm Gia Bình – Lương Tài – Bắc Ninh 2019 – 2020",
                "latex": r"""\text{Giải phương trình: } \sin^4 x + \cos^4 x + \cos\left(x - \frac{\pi}{4}\right)\sin\left(3x - \frac{\pi}{4}\right) - \frac{3}{2} = 0""",
                "source_kw": "Gia Bình"
            },
            {
                "title": "Đề chọn HSG Toán 11 năm 2024 – 2025 THPT Phùng Khắc Khoan – Hà Nội",
                "latex": r"""\text{Giải phương trình: } \sin^2 x + \sin^2 2x + \sin^2 3x = \frac{3}{2}""",
                "source_kw": "Phùng Khắc Khoan"
            },
            {
                "title": "Đề chọn đội tuyển HSG Toán 11 năm 2022 – 2023 THPT chuyên Vĩnh Phúc",
                "latex": r"""\text{Giải phương trình: } \cos^3 x\cos 3x - \sin^3 x\sin 3x = \frac{2 + 3\sqrt{2}}{8}""",
                "source_kw": "chuyên Vĩnh Phúc"
            }
        ]
    },
    {
        "category": "Dạng 2: Phương trình bậc nhất đối với sin và cos ($a\sin x + b\cos x = c$)",
        "questions": [
            {
                "title": "Đề KSCL đội tuyển HSG Toán 11 năm 2016 – 2017 trường Lê Lợi – Thanh Hóa",
                "latex": r"""\text{Giải phương trình: } \sin 2x + \sqrt{3}\cos 2x + (2 + \sqrt{3})\sin x - \cos x = 1 + \sqrt{3}""",
                "source_kw": "Lê Lợi"
            },
            {
                "title": "Đề thi KSCL học sinh giỏi cụm thi THPT Yên Thành – Nghệ An 2016 – 2017",
                "latex": r"""\text{Giải phương trình: } 2(\sin x + 3)\cos^4\left(\frac{x}{2}\right) - \sin x(1 + \cos x) - 3\cos x - 1 = 0""",
                "source_kw": "Yên Thành"
            },
            {
                "title": "Đề KSCL đội tuyển HSG Toán 11 năm 2017 – 2018 THPT Minh Châu – Hưng Yên",
                "latex": r"""\text{Giải phương trình: } 2\sqrt{3}\cos^2 x - \sin 2x = 2\cos\left(x - \frac{\pi}{4}\right)""",
                "source_kw": "Minh Châu"
            },
            {
                "title": "Đề thi HSG Toán 11 năm 2023 – 2024 cụm trường THPT Gia Lâm – Long Biên – Hà Nội",
                "latex": r"""\text{Giải phương trình: } \sqrt{3}\sin 2x + \cos 2x - 2\cos x = \sqrt{3}""",
                "source_kw": "GL - LB"
            },
            {
                "title": "Đề thi HSG cấp tỉnh môn Toán 11 năm 2018 – 2019 sở GD&ĐT Quảng Trị",
                "latex": r"""\text{Giải phương trình: } 2\sin 2x - \cos 2x = 7\sin x + 2\cos x - 4""",
                "source_kw": "Quảng Trị"
            }
        ]
    },
    {
        "category": "Dạng 3: Phương trình có điều kiện xác định, chứa ẩn dưới mẫu & tang, cotang",
        "questions": [
            {
                "title": "Đề KSCL đội tuyển HSG Toán 11 năm 2017 – 2018 THPT Minh Châu – Hưng Yên",
                "latex": r"""\text{Giải phương trình: } \frac{2\sin^2 x - 1}{2(1 - \cos x)(1 + \cot x)} = \frac{\sin x - \cos x}{\sin x + \cos x}""",
                "source_kw": "Minh Châu"
            },
            {
                "title": "Đề thi chọn HSG Toán 11 năm 2021 – 2022 sở GD&ĐT Quảng Bình",
                "latex": r"""\text{Giải phương trình: } 3\tan^2 x + 2\sqrt{3}\tan x - 1 = \frac{2}{\cos^2 x}""",
                "source_kw": "Quảng Bình"
            },
            {
                "title": "Đề khảo sát HSG Toán 11 năm 2019 – 2020 THPT Hậu Lộc 2 – Thanh Hóa",
                "latex": r"""\text{Giải phương trình: } \frac{\sin^4 2x + \cos^4 2x}{\tan\left(\frac{\pi}{6} - x\right)\tan\left(\frac{\pi}{3} + x\right)} = \frac{1}{6}\cos^4 4x - \frac{1}{24}""",
                "source_kw": "Hậu Lộc"
            },
            {
                "title": "Đề chọn HSG Toán 11 năm 2025 – 2026 chuyên Lê Quý Đôn – Đà Nẵng",
                "latex": r"""\text{Giải phương trình: } \frac{\sin 3x - \sin x}{\sqrt{1 - \cos 2x}} = \cos 2x + \sin 2x \quad \text{với } x \in (0; 2\pi)""",
                "source_kw": "Lê Quý Đôn"
            }
        ]
    },
    {
        "category": "Dạng 4: Phương trình đối xứng theo sin x ± cos x và phương trình chứa tham số m",
        "questions": [
            {
                "title": "Đề thi HSG Toán 11 cấp trường năm 2018 – 2019 THPT Lưu Hoàng – Hà Nội",
                "latex": r"""\text{Cho phương trình: } (2\sin x - 1)(2\cos 2x + 2\sin x + m) = 3 - 4\cos^2 x \quad (m \text{ là tham số}).\\
\text{Tìm tất cả các giá trị của } m \text{ để phương trình có đúng 2 nghiệm thuộc } [0; \pi].""",
                "source_kw": "Lưu Hoàng"
            },
            {
                "title": "Đề KSCL đội tuyển HSG Toán 11 năm 2017 – 2018 THPT Yên Lạc 2 – Vĩnh Phúc",
                "latex": r"""\text{Tìm } m \text{ để hàm số: } y = \frac{\cos x}{3\sin 5x - 4\cos 5x - 2m + 3} \text{ có tập xác định là } \mathbb{R}.""",
                "source_kw": "Yên Lạc"
            },
            {
                "title": "Đề thi HSG Toán 11 năm 2023 – 2024 sở GD&ĐT Quảng Ngãi",
                "latex": r"""\text{Giải phương trình: } \sin x + \cos x = \sqrt{2}(2 - \sin 2x)""",
                "source_kw": "Quảng Ngãi"
            },
            {
                "title": "Đề thi HSG Toán 11 năm 2020 – 2021 sở GD&ĐT Hà Nam",
                "latex": r"""\text{Giải phương trình: } 2\sin^3\left(x + \frac{\pi}{4}\right) = \sqrt{2}\sin x""",
                "source_kw": "Hà Nam"
            }
        ]
    },
    {
        "category": "Dạng 5: Bất đẳng thức & Hệ thức lượng giác trong tam giác (HSG Cấp Tỉnh)",
        "questions": [
            {
                "title": "Đề thi HSG Toán 11 cấp trường năm 2018 – 2019 THPT Lưu Hoàng – Hà Nội",
                "latex": r"""\text{Cho } A, B, C \text{ là ba góc của tam giác } ABC.\\
\text{a) Chứng minh tam giác } ABC \text{ vuông nếu: } \frac{\cos A + \cos B}{\sin A + \sin B} = \frac{\sin C}{\cos C}.\\
\text{b) Chứng minh rằng trong mọi tam giác } ABC \text{ ta luôn có: }\\
\sin A + \sin B + \sin C \le \frac{3\sqrt{3}}{2}""",
                "source_kw": "Lưu Hoàng"
            },
            {
                "title": "Đề chọn HSG Toán 11 năm 2025 – 2026 sở GD&ĐT Thái Nguyên",
                "latex": r"""\text{Giải phương trình: } \cos 3x + \sqrt{2 - \cos^2 3x} = 2(1 + \sin^2 2x)""",
                "source_kw": "Thái Nguyên"
            }
        ]
    }
]

CSS = """
<style>
  :root {
    --primary: #1a365d;
    --primary-light: #2b6cb0;
    --bg: #f7fafc;
    --card-bg: #ffffff;
    --text: #2d3748;
    --border: #e2e8f0;
  }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    background: var(--bg);
    color: var(--text);
    margin: 0;
    padding: 20px 40px;
    line-height: 1.6;
  }
  .top-banner {
    background: linear-gradient(135deg, #1a365d 0%, #2b6cb0 100%);
    color: white;
    padding: 26px 32px;
    border-radius: 12px;
    margin-bottom: 25px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.12);
  }
  .top-banner h1 { margin: 0 0 8px 0; font-size: 22px; }
  .top-banner p { margin: 0; opacity: 0.92; font-size: 14.5px; }
  .tabs { display: flex; gap: 12px; margin-bottom: 25px; }
  .tab-btn {
    padding: 10px 22px;
    font-size: 15px;
    font-weight: 600;
    border: none;
    border-radius: 8px;
    cursor: pointer;
    background: white;
    color: #4a5568;
    box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    transition: all 0.2s;
  }
  .tab-btn.active {
    background: #2b6cb0;
    color: white;
    box-shadow: 0 3px 8px rgba(43,108,176,0.3);
  }
  .section-title {
    font-size: 19px;
    font-weight: 700;
    color: #1a365d;
    margin: 25px 0 15px 0;
    padding-bottom: 8px;
    border-bottom: 2px solid #cbd5e0;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .card {
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 20px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.04);
  }
  .card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 12px;
    border-bottom: 1px dashed var(--border);
    padding-bottom: 8px;
    flex-wrap: wrap;
    gap: 10px;
  }
  .q-badge {
    background: #ebf8ff;
    color: #2b6cb0;
    padding: 4px 10px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 13.5px;
  }
  .q-source {
    font-weight: 600;
    color: #2d3748;
    font-size: 13.5px;
  }
  .actions {
    display: flex;
    gap: 8px;
    margin-left: auto;
  }
  .btn-action {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 6px 12px;
    border-radius: 6px;
    font-size: 13px;
    font-weight: 600;
    text-decoration: none;
    cursor: pointer;
    border: none;
  }
  .btn-pdf { background: #e2e8f0; color: #2d3748; }
  .btn-pdf:hover { background: #cbd5e0; }
  .btn-preview { background: #3182ce; color: white; }
  .btn-preview:hover { background: #2b6cb0; }
  .btn-copy { background: #edf2f7; color: #4a5568; }
  .btn-copy:hover { background: #e2e8f0; }
  .math-box {
    background: #f8fafc;
    border-left: 4px solid #3182ce;
    padding: 16px 20px;
    border-radius: 0 8px 8px 0;
    margin: 12px 0;
    font-size: 15.5px;
  }
  .raw-box {
    background: #fdfdfd;
    border-left: 3px solid #3182ce;
    padding: 12px 16px;
    font-family: 'Consolas', monospace;
    font-size: 13.5px;
    white-space: pre-wrap;
    color: #333;
    margin-top: 10px;
  }
  details.preview-collapse {
    margin-top: 14px;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    overflow: hidden;
  }
  details.preview-collapse summary {
    background: #edf2f7;
    padding: 10px 16px;
    font-weight: 600;
    font-size: 13.5px;
    cursor: pointer;
    user-select: none;
    color: #2b6cb0;
  }
  details.preview-collapse summary:hover { background: #e2e8f0; }
  details.preview-collapse .preview-content {
    padding: 15px;
    background: #fff;
    text-align: center;
  }
  details.preview-collapse img {
    max-width: 100%;
    height: auto;
    border: 1px solid #cbd5e0;
    border-radius: 6px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
  }
  .modal {
    display: none;
    position: fixed;
    z-index: 9999;
    left: 0; top: 0;
    width: 100%; height: 100%;
    background: rgba(0,0,0,0.75);
    overflow: auto;
  }
  .modal-content {
    background: white;
    margin: 3% auto;
    padding: 20px;
    border-radius: 10px;
    width: 80%;
    max-width: 1000px;
    position: relative;
  }
  .close-modal {
    position: absolute;
    right: 20px; top: 15px;
    font-size: 28px;
    font-weight: bold;
    color: #a0aec0;
    cursor: pointer;
  }
  .close-modal:hover { color: #e53e3e; }
  .modal img { width: 100%; height: auto; border-radius: 6px; }
  .floating-print {
    position: fixed;
    bottom: 25px;
    right: 25px;
    padding: 12px 20px;
    background: #38a169;
    color: white;
    font-weight: bold;
    border-radius: 30px;
    box-shadow: 0 4px 12px rgba(56,161,105,0.4);
    border: none;
    cursor: pointer;
    font-size: 14px;
    z-index: 100;
  }
  @media print {
    .floating-print, .actions, .tabs, .top-banner p, details.preview-collapse { display: none !important; }
    body { padding: 0; background: white; }
    .card { box-shadow: none; border: 1px solid #ccc; page-break-inside: avoid; }
  }
</style>
<script>
  window.MathJax = {
    tex: {
      inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
      displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']]
    }
  };
</script>
<script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
"""

JS = """
<script>
function openModal(imgSrc, title) {
  document.getElementById('modalImg').src = imgSrc;
  document.getElementById('modalTitle').innerText = title;
  document.getElementById('examModal').style.display = 'block';
}
function closeModal() {
  document.getElementById('examModal').style.display = 'none';
}
window.onclick = function(event) {
  var modal = document.getElementById('examModal');
  if (event.target == modal) { modal.style.display = 'none'; }
}
function copyText(elemId, btn) {
  var text = document.getElementById(elemId).innerText;
  navigator.clipboard.writeText(text).then(function() {
    var oldText = btn.innerHTML;
    btn.innerHTML = '✅ Đã chép!';
    setTimeout(function() { btn.innerHTML = oldText; }, 2000);
  });
}
function switchTab(tabId) {
  document.getElementById('tab-curated').style.display = tabId === 'curated' ? 'block' : 'none';
  document.getElementById('tab-all').style.display = tabId === 'all' ? 'block' : 'none';
  document.getElementById('btn-curated').classList.toggle('active', tabId === 'curated');
  document.getElementById('btn-all').classList.toggle('active', tabId === 'all');
}
</script>
"""

MODAL_HTML = """
<div id="examModal" class="modal">
  <div class="modal-content">
    <span class="close-modal" onclick="closeModal()">&times;</span>
    <h3 id="modalTitle" style="margin-top:0; color:#1a365d;">Trang Đề Thi Gốc</h3>
    <img id="modalImg" src="" alt="Đề thi gốc sắc nét">
  </div>
</div>
"""

def extract_all_questions():
    collections = {
        'LUONG_GIAC': {
            'title': 'CHUYÊN ĐỀ 1: PHƯƠNG TRÌNH & HỆ THỨC LƯỢNG GIÁC HSG TOÁN 11',
            'desc': 'Tuyển tập câu hỏi tự luận Lượng giác chọn lọc từ 130 đề thi Học Sinh Giỏi Toán 11 Tỉnh/TP & Olympic (2012 – 2026). Đã tích hợp MathJax công thức chuẩn và Trình xem trực tiếp Đề gốc sắc nét 100%.',
            'file_name': '01_Luong_Giac_HSG_11.html',
            'items': []
        },
        'DAY_SO': {
            'title': 'CHUYÊN ĐỀ 2: DÃY SỐ, CẤP SỐ & GIỚI HẠN HSG TOÁN 11',
            'desc': 'Tuyển tập câu hỏi Tự luận Dãy số, Cấp số cộng, Cấp số nhân & Giới hạn từ các kỳ thi HSG Toán 11 (2012 – 2026).',
            'file_name': '02_Day_So_Gioi_Han_HSG_11.html',
            'items': []
        },
        'XAC_SUAT': {
            'title': 'CHUYÊN ĐỀ 3: TỔ HỢP, XÁC SUẤT & NHỊ THỨC NEWTON HSG TOÁN 11',
            'desc': 'Tuyển tập câu hỏi Tự luận Tổ hợp, Xác suất, Nhị thức Newton từ các kỳ thi HSG Toán 11 (2012 – 2026).',
            'file_name': '03_Xac_Suat_To_Hop_HSG_11.html',
            'items': []
        },
        'HINH_KHONG_GIAN': {
            'title': 'CHUYÊN ĐỀ 4: HÌNH HỌC KHÔNG GIAN (GÓC & KHOẢNG CÁCH) HSG TOÁN 11',
            'desc': 'Tuyển tập câu hỏi Tự luận Hình học không gian (quan hệ song song, vuông góc, khoảng cách, góc, thiết diện) từ các kỳ thi HSG Toán 11 (2012 – 2026).',
            'file_name': '04_Hinh_Hoc_Khong_Gian_HSG_11.html',
            'items': []
        }
    }

    count_files = 0
    for root, dirs, files in os.walk(TU_LUAN_DIR):
        for f in files:
            if not f.endswith('.pdf'):
                continue
            count_files += 1
            pdf_path = os.path.join(root, f)
            source_name = f.replace('.pdf', '').replace('_', ' ')
            year_name = os.path.basename(root).replace('Nam_hoc_', 'Năm học ').replace('_', ' - ')

            try:
                reader = PdfReader(pdf_path)
                max_pages = min(2, len(reader.pages))
                exam_text = '\n'.join([reader.pages[i].extract_text() or '' for i in range(max_pages)])

                pattern = r'(?:(?:\n|^)(?:Bài|Câu)\s+(?:[0-9IVX]+|I|II|III|IV|V|VI|VII)[\.\:\s][\s\S]*?)(?=(?:\n(?:Bài|Câu)\s+(?:[0-9IVX]+|I|II|III|IV|V|VI|VII)[\.\:\s])|(?:\n\s*HẾT)|(?:\n\s*---)|$)'
                questions = re.findall(pattern, exam_text)

                for q in questions:
                    q_clean = q.strip()
                    lower = q_clean.lower()
                    if len(q_clean) < 35 or 'hướng dẫn chấm' in lower or 'bảng đáp án' in lower or 'biểu điểm' in lower or 'thang điểm' in lower:
                        continue
                    if lower.count('0,5') >= 2 or lower.count('0.5') >= 2:
                        continue
                    cat = classify_question(q_clean)
                    if cat in collections:
                        collections[cat]['items'].append({
                            'source': f"{source_name} ({year_name})",
                            'content': q_clean
                        })
            except Exception as e:
                pass

    print(f"Scanned {count_files} PDFs.")
    for k, v in collections.items():
        print(f"- {k}: {len(v['items'])} questions")

    return collections

def build_all_html():
    collections = extract_all_questions()

    for cat_key, col in collections.items():
        file_path = os.path.join(OUT_DIR, col['file_name'])
        is_luong_giac = (cat_key == 'LUONG_GIAC')

        cleaned_cards = []
        for idx, item in enumerate(col['items'], 1):
            src = item['source']
            body_clean = item['content']

            pdf_info = find_pdf_info(src)
            rel_pdf = pdf_info['rel_pdf'] if pdf_info else None
            rel_preview = pdf_info['rel_preview'] if pdf_info else None
            clean_title = pdf_info['clean_title'] if pdf_info else src
            clean_title_esc = clean_title.replace("'", "\\'").replace('"', '&quot;')

            actions_html = []
            if rel_preview:
                actions_html.append(f'<button class="btn-action btn-preview" onclick="openModal(\'{rel_preview}\', \'{clean_title_esc}\')">👁️ Phóng to đề gốc</button>')
            if rel_pdf:
                actions_html.append(f'<a class="btn-action btn-pdf" href="{rel_pdf}" target="_blank">📄 Mở PDF gốc</a>')

            preview_collapse = ""
            if rel_preview:
                preview_collapse = f"""
                <details class="preview-collapse">
                  <summary>🔍 Nhấp xem nhanh trang đề thi gốc sắc nét (Chống lỗi font MathType)</summary>
                  <div class="preview-content">
                    <img src="{rel_preview}" alt="Trang đề gốc" loading="lazy">
                  </div>
                </details>
                """

            cleaned_cards.append(f"""
            <div class="card">
              <div class="card-header">
                <span class="q-badge">Câu {idx}</span>
                <span class="q-source">🏷️ {src}</span>
                <div class="actions">
                  {''.join(actions_html)}
                </div>
              </div>
              <div class="raw-box">{body_clean}</div>
              {preview_collapse}
            </div>
            """)

        if is_luong_giac:
            # Add Curated LaTeX Tab
            curated_sections_html = []
            curated_id = 0
            for sec in CURATED_LUONG_GIAC:
                sec_cards = []
                for item in sec["questions"]:
                    curated_id += 1
                    cid = f"curated_q_{curated_id}"
                    pdf_info = None
                    for k, v in pdf_files.items():
                        if item["source_kw"].lower() in k.lower():
                            pdf_info = v
                            break
                    rel_pdf = pdf_info['rel_pdf'] if pdf_info else None
                    rel_preview = pdf_info['rel_preview'] if pdf_info else None
                    title_esc = item['title'].replace("'", "\\'")

                    actions_html = [
                        f'<button class="btn-action btn-copy" onclick="copyText(\'{cid}\', this)">📋 Chép đề</button>'
                    ]
                    if rel_preview:
                        actions_html.append(f'<button class="btn-action btn-preview" onclick="openModal(\'{rel_preview}\', \'{title_esc}\')">👁️ Đề gốc</button>')
                    if rel_pdf:
                        actions_html.append(f'<a class="btn-action btn-pdf" href="{rel_pdf}" target="_blank">📄 PDF</a>')

                    preview_collapse = ""
                    if rel_preview:
                        preview_collapse = f"""
                        <details class="preview-collapse">
                          <summary>🔍 Nhấp xem nhanh trang đề thi gốc sắc nét</summary>
                          <div class="preview-content">
                            <img src="{rel_preview}" alt="Trang đề gốc" loading="lazy">
                          </div>
                        </details>
                        """

                    sec_cards.append(f"""
                    <div class="card">
                      <div class="card-header">
                        <span class="q-badge">Bài {curated_id}</span>
                        <span class="q-source">🏷️ {item['title']}</span>
                        <div class="actions">
                          {''.join(actions_html)}
                        </div>
                      </div>
                      <div class="math-box" id="{cid}">
                        $${item['latex']}$$
                      </div>
                      {preview_collapse}
                    </div>
                    """)

                curated_sections_html.append(f"""
                <div class="section-title">{sec['category']} ({len(sec['questions'])} câu)</div>
                {''.join(sec_cards)}
                """)

            body_content = f"""
            <div class="tabs">
              <button id="btn-curated" class="tab-btn active" onclick="switchTab('curated')">⭐ Đề Tuyển Chọn Chuẩn Hóa LaTeX ({curated_id} bài mẫu đỉnh cao)</button>
              <button id="btn-all" class="tab-btn" onclick="switchTab('all')">📚 Toàn Bộ Ngân Hàng Câu Hỏi ({len(cleaned_cards)} câu + Đề Gốc Sắc Nét)</button>
            </div>

            <div id="tab-curated">
              {''.join(curated_sections_html)}
            </div>

            <div id="tab-all" style="display:none;">
              <div class="section-title">Ngân hàng toàn bộ {len(cleaned_cards)} câu hỏi Lượng giác (Đã làm sạch thang điểm)</div>
              <p style="color:#718096; margin-bottom:20px; font-size:14px;">Mỗi câu hỏi dưới đây đều có tính năng <b>"Xem nhanh trang đề thi gốc"</b> hoặc <b>"Phóng to đề gốc"</b> để xem công thức nguyên bản sắc nét 100% từ đề của Sở/Bộ mà không bao giờ bị xô lệch.</p>
              {''.join(cleaned_cards)}
            </div>
            """
        else:
            body_content = f"""
            <div class="cards-container">
              {''.join(cleaned_cards)}
            </div>
            """

        full_html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{col['title']}</title>
{CSS}
</head>
<body>
<button class="floating-print" onclick="window.print()">🖨️ In Tuyển Tập Đề</button>

<div class="top-banner">
  <h1>📚 {col['title']}</h1>
  <p>{col['desc']}</p>
  <p style="margin-top: 8px; font-size: 13.5px; opacity: 0.85;">💡 <i>Mỗi câu hỏi đều tích hợp <b>"Xem nhanh trang đề thi gốc"</b> hoặc <b>"Phóng to đề gốc"</b> để xem hình vẽ, sơ đồ và công thức nguyên bản sắc nét 100%.</i></p>
</div>

{body_content}

{MODAL_HTML}
{JS}
</body>
</html>
"""
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        print(f"Generated {col['file_name']}: {len(cleaned_cards)} questions.")

if __name__ == '__main__':
    build_all_html()
