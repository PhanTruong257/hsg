---
name: math-exam-word-equation
description: >-
  Standard operating procedure and toolkit for digitizing mathematical competition exams (HSG) 
  from PDF/MathType documents into clean JSON question banks and Microsoft Word (.docx) documents 
  featuring 100% native Word Equations (Alt + =, m:oMath), zero scrambled formulas, zero PUA/font corruption, 
  and professional typography.
---

# Math Exam Word Equation Digitization Skill

## 1. Bối cảnh & Nguyên nhân gốc rễ (Root Cause Analysis)
Khi trích xuất văn bản từ các tệp đề thi toán THPT/HSG dạng PDF (thường được xuất từ Microsoft Word + MathType):
- **Phân rã tọa độ 2D:** MathType trong PDF lưu công thức dưới dạng các đối tượng đồ họa rời rạc. Tử số, mẫu số, số mũ, chỉ số dưới được đặt ở các tọa độ $(x, y)$ riêng biệt; dấu gạch ngang phân số và dấu căn là các nét vẽ vector.
- **Lỗi quét dòng tuyến tính:** Các thư viện đọc text thông thường (`pypdf`, `pdfplumber`, `pdftotext`) gom text theo thứ tự quét $y$, dẫn đến phân số bị xé thành các dòng dọc, số mũ rơi xuống ngang hàng (`4sin 3 cos 2...`), và hệ phương trình bị vỡ.
- **Ký tự Private Use Area (PUA):** MathType dùng vùng font PUA riêng (`0xE000 - 0xF8FF`), khiến các ký hiệu toán đặc biệt biến thành ô vuông `□` hoặc rác font.

## 2. Quy trình chuẩn 4 bước (Standard Operating Procedure - SOP)

```
[1. PDF gốc] ──(PyMuPDF crop)──> [Ảnh nét cao từng câu]
                                        │
                               (Vision Inspection)
                                        ↓
[2. AI Subagents] ─────────────> [LaTeX chuẩn hóa 100%]
                                        │
                                 (Structured Data)
                                        ↓
[3. JSON Database] ────────────> [Ngân hàng câu hỏi sạch]
                                        │
                            (latex2mathml + MML2OMML.XSL)
                                        ↓
[4. Python-Docx Engine] ───────> [Native Word Equation (Alt + =, m:oMath)]
```

### Bước 1: Trích xuất thị giác nguyên bản (Vision Cropping)
- Tuyệt đối không dùng text extract cho các khối có công thức phức tạp.
- Sử dụng thư viện `fitz` (`PyMuPDF`) định vị tọa độ và cắt ảnh khối đề bài:
```python
import fitz # PyMuPDF
doc = fitz.open(pdf_path)
page = doc[page_index]
# Crop rect với zoom 2x (150-300 DPI) để ảnh rõ nét
pix = page.get_pixmap(clip=rect, dpi=200)
pix.save("crop.png")
```

### Bước 2: Phiên âm toán học sang LaTeX (Mathematical Transcription)
- Xem ảnh crop thực tế để đối chiếu chính xác:
  - Phân số 2 tầng, căn bậc hai lồng nhau, mũ lũy thừa.
  - Các hàm lượng giác $\sin, \cos, \tan, \cot$, giới hạn $\lim$, tích phân $\int$, v.v.
  - Hệ phương trình `\begin{cases} ... \end{cases}`.
  - Loại bỏ các câu hỏi lạc đề (nếu trang PDF in chung câu thuộc chuyên đề khác).
- Chuẩn hóa toàn bộ công thức sang mã nguồn LaTeX hợp lệ:
  - Công thức trong dòng (inline): `$formula$`
  - Công thức độc lập (display): `$$formula$$` hoặc trường `latex` riêng.

### Bước 3: Đóng gói Cơ sở dữ liệu JSON (`database_<chuyen_de>.json`)
Mỗi câu hỏi phải tuân thủ nghiêm ngặt cấu trúc:
```json
{
  "id": "HSG11_MA_001",
  "index": 1,
  "year": "Năm học 2023 - 2024",
  "source": "Đề thi HSG Toán 11 Tỉnh Nghệ An",
  "category": "Phương trình bậc hai đối với một hàm số lượng giác",
  "prompt": "Giải phương trình sau:",
  "latex": "2\\cos^2 x - 3\\sin x = 0",
  "sub_questions": [
    "a) Giải phương trình khi $m = 1$.",
    "b) Tìm $m$ để phương trình có nghiệm trên đoạn $[0; \\pi]$."
  ]
}
```

### Bước 4: Biên dịch Native Word Equation (`Alt + =`, `m:oMath`)
Sử dụng cầu nối:
$$\text{LaTeX} \xrightarrow{\text{latex2mathml}} \text{MathML XML} \xrightarrow{\text{MML2OMML.XSL}} \text{Word OpenXML } \langle\text{m:oMath}\rangle$$

- File XSLT của Microsoft Office nằm tại:
  `C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL` (hoặc bản 64-bit).
- Đoạn mã Python chuẩn:
```python
import docx
import latex2mathml.converter
import lxml.etree as etree

xsl_path = r'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL'
xslt = etree.parse(xsl_path)
transform = etree.XSLT(xslt)

def latex_to_omml(latex_str):
    mathml = latex2mathml.converter.convert(latex_str)
    # Bắt buộc: Escape các dấu & độc lập (từ aligned, cases, matrix) trước khi parse XML
    mathml = re.sub(r'&(?!(#[0-9]+|#x[0-9a-fA-F]+|[a-zA-Z]+);)', '&amp;', mathml)
    tree = etree.fromstring(mathml)
    return transform(tree).getroot()

def add_inline_math(paragraph, text):
    parts = text.split('$')
    for i, part in enumerate(parts):
        if not part:
            continue
        if i % 2 == 0:
            paragraph.add_run(part)
        else:
            omml = latex_to_omml(part)
            paragraph._p.append(omml)
```

## 3. Tiêu chuẩn Thể thức & Trình bày Tài liệu (Word Styling Guide)
1. **Khổ giấy & Căn lề chuẩn Bộ GD&ĐT:**
   - Khổ giấy: A4.
   - Căn lề: Top 2.0 cm, Bottom 2.0 cm, Left 3.0 cm, Right 2.0 cm.
2. **Font chữ & Khoảng cách dòng:**
   - Font: `Times New Roman` cho toàn bộ văn bản.
   - Cỡ chữ: 12pt (tiêu đề bài 12pt bold màu Navy `#1E3A8A`, phân dạng 10.5pt bold Slate/Teal).
   - Giãn dòng: Multiple 1.15.
   - Giãn đoạn: Space after 2-4pt cho câu hỏi, 6pt kết thúc bài toán.
3. **Quy tắc làm sạch dữ liệu:**
   - Tuyệt đối loại bỏ các dòng `📄 Nguồn gốc file đề: ...pdf`. Thay bằng thông tin ngắn gọn `(Năm học | Đơn vị ra đề)` đặt cùng dòng tiêu đề bài.
   - Toàn bộ công thức toán độc lập được căn giữa (`WD_ALIGN_PARAGRAPH.CENTER`).
   - Các ý con $a), b)$ hoặc $A, B, C, D$ thụt đầu dòng `left_indent = Inches(0.25)`.

## 4. Bộ tiêu chí Kiểm định Chất lượng (QA Verification Checklist)
Mỗi tài liệu xuất ra phải vượt qua script kiểm định tự động với các tiêu chí:
1. `Số bài toán`: Đạt 100% số lượng (không bị cắt ngắn thành bài mẫu hay để link).
2. `Thẻ m:oMath`: Đạt $\ge$ 100% số lượng công thức trong đề.
3. `Ký tự rác PUA (0xE000-0xF8FF)`: Phải bằng **0**.
4. `Ký tự ô vuông lỗi (□, \ufffd)`: Phải bằng **0**.
5. `Rò rỉ mã lệnh LaTeX thô trong text`: Phải bằng **0**.
6. `Dòng rác nguồn gốc file`: Phải bằng **0**.
7. `Ký tự thừa / lỗi nhân đôi regex (\r\r, \r)`: Phải bằng **0**.

---

## 5. Bảng tổng hợp các bẫy lỗi kỹ thuật & cách phòng tránh triệt để (Pitfalls & Solutions)

Trong quá trình số hóa thực tế từ 130 đề thi HSG Toán 11, các bẫy kỹ thuật sau đã được đúc rút để không bao giờ tái phạm:

### ⚠️ Bẫy 1: Lỗi `\r\r` do thay thế chuỗi thô (Regex / String Replacement Overreach)
- **Hiện tượng:** Trong Word công thức hiển thị các ký tự thừa kỳ lạ như $\tan(\frac{10\pi}{3}\backslash\text{r}\backslash\text{r})$.
- **Nguyên nhân gốc:** Khi một bài toán bị thiếu ký tự (ví dụ gõ nhầm `ight)` thay vì `\right)`), việc dùng lệnh replace ngây thơ `.replace('ight)', r'\right)')` hoặc regex `(?<!\\)ight` quét trên toàn bộ tài liệu sẽ biến các công thức chuẩn vốn có `\right)` thành `\r` + `\right)` $\to$ `\r\r\right)`. Khi qua bộ chuyển đổi Equation của Word, `\r` bị coi là macro/ký tự điều khiển và in thẳng ra màn hình.
- **Quy tắc phòng ngừa:**
  - **TUYỆT ĐỐI KHÔNG** dùng lệnh replace chuỗi thô đối với các từ khóa cấu trúc LaTeX (`\left`, `\right`, `\begin`, `\end`) trên toàn bộ văn bản.
  - Sửa lỗi cú pháp trực tiếp tại phần tử dữ liệu gốc (Database/JSON) thay vì vá lỗi ở tầng converter trung gian.

### ⚠️ Bẫy 2: Lỗi sập XML `XMLSyntaxError: xmlParseEntityRef: no name` do dấu `&` trong môi trường căn dòng
- **Hiện tượng:** Toàn bộ công thức hệ phương trình (`\begin{cases}`) hoặc biến đổi tương đương nhiều dòng (`\begin{aligned}`) bị chuyển đổi thất bại và rớt xuống dạng text trần chứa đầy mã LaTeX.
- **Nguyên nhân gốc:** Trong LaTeX, dấu `&` được dùng để gióng hàng (ví dụ: `3x &= 2x`). Khi thư viện `latex2mathml` sinh mã XML MathML, nó tạo thẻ `<mi>&</mi>`. Bộ phân tích cú pháp `lxml.etree` của Python sẽ sập vì trong chuẩn XML, dấu `&` bắt buộc phải là thực thể `&amp;`. Khi parser sập, hàm bắt exception và in mã LaTeX thô ra Word.
- **Đoạn mã bắt buộc trong engine chuyển đổi:**
  ```python
  def latex_to_omml(latex_str):
      try:
          mathml = latex2mathml.converter.convert(latex_str)
          # Bắt buộc: Escape các dấu & độc lập không phải entity XML trước khi parse
          mathml = re.sub(r'&(?!(#[0-9]+|#x[0-9a-fA-F]+|[a-zA-Z]+);)', '&amp;', mathml)
          tree = etree.fromstring(mathml)
          omml = transform(tree)
          return omml.getroot()
      except Exception as e:
          return None
  ```

### ⚠️ Bẫy 3: Rò rỉ mã LaTeX trần trong Thẻ Phân Dạng (Category Badge)
- **Hiện tượng:** Tiêu đề bài toán hiển thị: `[Dạng 2: Phương trình bậc nhất đối với sin và cos ($a\sin x + b\cos x = c$)]` dưới dạng chữ thường có cả dấu `$`.
- **Nguyên nhân gốc:** Người lập trình dùng hàm thêm text thường `p.add_run(f"[{cat}]")` mà quên rằng tên phân dạng cũng có thể chứa công thức toán học đặt trong cặp `$ ... $`.
- **Quy tắc phòng ngừa:** Mọi thành phần hiển thị (Tiêu đề, Badge phân dạng, Prompt, Sub-questions, Lời giải) đều bắt buộc đi qua hàm `add_inline_math()`.

### ⚠️ Bẫy 4: Bất đồng bộ dữ liệu khi hợp nhất đa tiến trình (Multi-Agent Data Inconsistency)
- **Hiện tượng:** Khi chạy nhiều subagents giải song song (Part 1, Part 2, Part 3), lỗi đã sửa ở file tổng hợp `solutions_all.json` nhưng khi chạy lại script build Word thì lỗi cũ lại xuất hiện.
- **Nguyên nhân gốc:** Script build tự động nạp lại từ các file thành phần con (`solutions_part1.json`) chưa được cập nhật và ghi đè lên file tổng hợp.
- **Quy tắc phòng ngừa:** Tuân thủ nguyên tắc *Single Source of Truth*: Sửa lỗi từ file tầng thấp nhất (part files) trước khi chạy hàm merge, hoặc chỉ đọc duy nhất từ Database trung tâm `database_<chuyen_de>.json`.

### ⚠️ Bẫy 5: Lỗi bảng mã Windows PowerShell CP1252 (`UnicodeEncodeError`)
- **Hiện tượng:** Python script bị crash ngay khi in log tiếng Việt ra màn hình: `UnicodeEncodeError: 'charmap' codec can't encode character '\u1ec7'`.
- **Đoạn mã bắt buộc đầu mọi script Python:**
  ```python
  import sys
  sys.stdout.reconfigure(encoding='utf-8')
  sys.stderr.reconfigure(encoding='utf-8')
  ```

### ⚠️ Bẫy 6: Cân bằng cặp dấu `$` trên từng dòng lời giải
- **Hiện tượng:** Một dòng lời giải bị lẻ dấu `$` (ví dụ 1 hoặc 3 dấu `$`), khiến toàn bộ đoạn văn bản phía sau bị hoán đổi giữa chế độ văn bản thường và công thức toán.
- **Quy tắc kiểm tra tự động trước khi lưu:**
  ```python
  assert line.count('$') % 2 == 0, f"Lỗi lẻ dấu $ tại dòng: {line}"
  ```

### ⚠️ Bẫy 7: Lỗi chia cho 0 khi chèn hình vào Word (`ZeroDivisionError: division by zero` do DPI = 0)
- **Hiện tượng:** Khi gọi `paragraph.add_run().add_picture(img_path, width=...)`, `python-docx` ném ngoại lệ sập chương trình: `ZeroDivisionError: division by zero` tại `return Inches(self.px_width / self.horz_dpi)`.
- **Nguyên nhân gốc:** Ảnh trích xuất từ file PDF (bằng `pymupdf` / `fitz`) hoặc ảnh đồ họa thô thường không chứa metadata mật độ điểm ảnh (DPI), hoặc giá trị DPI trong chunk pHYs / JFIF bị gán bằng 0. Hàm tính toán tỷ lệ co giãn ảnh `scaled_dimensions` của `python-docx` lấy `self.px_width / self.horz_dpi`, dẫn đến chia cho 0.
- **Quy tắc phòng ngừa:** Luôn chuẩn hóa mọi hình ảnh qua Pillow với DPI chuẩn 300 trước khi nhúng vào tài liệu Word:
  ```python
  from PIL import Image
  im = Image.open(src_path).convert('RGB')
  im.save(dst_png_path, 'PNG', dpi=(300, 300))
  ```

### ⚠️ Bẫy 8: Lỗi `docx.image.exceptions.UnrecognizedImageError` khi chèn ảnh PDF trích xuất
- **Hiện tượng:** `python-docx` ném ngoại lệ `UnrecognizedImageError` dù file ảnh mở và xem bình thường trong Windows Photos.
- **Nguyên nhân gốc:** Bộ giải mã ảnh nội tại của `python-docx` (`DocxImage.from_file`) yêu cầu cấu trúc header magic bytes chuẩn (PNG, standard JFIF/EXIF JPEG). Các ảnh nhúng trong PDF đôi khi sử dụng hệ màu CMYK, raw stream JPX/JBIG2, hoặc thiếu chunk đánh dấu định dạng ảnh chuẩn Word.
- **Quy tắc phòng ngừa:**
  - Tuyệt đối không nhúng trực tiếp file trích xuất thô từ PDF (`.jpeg`, `.bin`, `.tmp`).
  - Toàn bộ hình vẽ minh họa bắt buộc phải được chuyển đổi (convert) sang định dạng `.png` 24-bit RGB tiêu chuẩn trước khi gọi `add_picture()`.

### ⚠️ Bẫy 9: Hình vẽ hình học không gian thiếu chi tiết hoặc xung đột giữa các ý a, b, c (Sub-question Geometry Conflict)
- **Hiện tượng:** 
  1. Hình vẽ chỉ có khung đáy và đỉnh chóp chung chung ($S, A, B, C, D$), thiếu hoàn toàn các điểm phụ ($M, N, P, H, K, I$), đường cao, hình chiếu, thiết diện được chứng minh trong lời giải câu a, b, c.
  2. Câu a và câu b có hai thiết diện khác nhau hoặc hai hình thể khác nhau (ví dụ: câu 1 là hình chóp, câu 2 là tứ diện đều hoặc lăng trụ), nhưng người soạn gom chung vào một hình vẽ gây xung đột và sai lệch hình học.
- **Quy tắc giải quyết chuẩn mực:**
  1. **Đầy đủ chi tiết:** Mọi đối tượng hình học xuất hiện trong lời giải (chân đường vuông góc $H$, đường cao $SH$, hình chiếu $K$, trung điểm $M, N$, thiết diện đa giác tô mờ alpha) đều phải hiển thị trên hình vẽ.
  2. **Tách hình khi có xung đột (Conflict Separation):** Khi các ý a, b, c sử dụng các cấu hình hình học đối lập hoặc các khối đa diện khác nhau, bắt buộc phải vẽ các hình độc lập:
     - `hinh_bai_{idx:02d}_a.png`: Minh họa câu a).
     - `hinh_bai_{idx:02d}_b.png`: Minh họa câu b).
     - `hinh_bai_{idx:02d}_c.png`: Minh họa câu c) (nếu có).
  3. **Trình bày song song (Side-by-side Layout):** Trong Word, đặt 2 hình vẽ cạnh nhau bằng bảng 2 cột ẩn viền (mỗi ô rộng 3.2 inches, ảnh rộng 3.0 inches), căn giữa kèm chú thích chuẩn mực (`Hình {idx}a. Hình vẽ minh họa câu a)`, `Hình {idx}b. Hình vẽ minh họa câu b)`), giúp tiết kiệm diện tích trang và tối ưu trải nghiệm đọc của giáo viên/học sinh.



