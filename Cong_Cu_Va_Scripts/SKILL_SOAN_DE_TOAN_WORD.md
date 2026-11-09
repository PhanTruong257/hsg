# HƯỚNG DẪN QUY CHUẨN KỸ THUẬT (SKILL): SOẠN THẢO & SỐ HÓA ĐỀ THI TOÁN SANG WORD NATIVE EQUATION
*(Chống 100% lỗi vỡ công thức MathType, lỗi font ô vuông □, và rác văn bản từ PDF)*

---

## 1. NGUYÊN NHÂN GỐC RỄ CÁC LỖI ĐỊNH DẠNG TỪ FILE ĐỀ PDF
Các đề thi học sinh giỏi (HSG) và đề thi THPT môn Toán thường được soạn bằng Microsoft Word + MathType, sau đó xuất sang PDF:
1. **Lỗi phân rã công thức 2D:** MathType khi sang PDF sẽ biến công thức thành các tọa độ hình học rời rạc. Tử số nằm ở hàng trên, mẫu số nằm ở hàng dưới, dấu gạch ngang phân số chỉ là 1 nét vẽ vector, dấu căn là hình vector ghép. Khi phần mềm trích xuất text PDF quét qua, nó gom chữ theo trục dọc, khiến phân số bị gãy nát thành nhiều dòng, số mũ rơi xuống ngang cơ số (`4sin 3 cos 2...` thay vì $4\sin^3 x \cos^2 x$).
2. **Lỗi font ký tự Private Use Area (PUA):** MathType dùng vùng font PUA riêng (`0xE000 - 0xF8FF`). Nếu copy thẳng sang Word thông thường, các ký hiệu toán học (dấu vector, góc, tích phân, ký hiệu tập số $\mathbb{R}$) sẽ bị biến thành ô vuông `□` hoặc ký tự rác.
3. **Lỗi rò rỉ mã lệnh hoặc link file đề:** Khi trích xuất tự động không kiểm soát, các dòng metadata như `📄 Nguồn gốc file đề: ...pdf` hoặc link Google Docs/Drive bị chèn bừa bãi làm loãng nội dung đề thi.

---

## 2. QUY TRÌNH 4 BƯỚC CHUẨN HÓA (STANDARD OPERATING PROCEDURE - SOP)

Mọi quá trình số hóa từ PDF đề thi sang tài liệu Word phải tuân thủ nghiêm ngặt 4 bước:

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
- Tuyệt đối không dùng công cụ đọc text tuyến tính đối với các khối công thức phức tạp.
- Sử dụng thư viện `PyMuPDF` (`fitz`) để cắt ảnh vùng chứa bài toán với độ phân giải cao (DPI 150 - 200).
- Mục đích: Giữ nguyên vẹn 100% hình ảnh mắt người nhìn thấy, bảo toàn cấu trúc phân số 2 tầng, ma trận, hệ phương trình và hình vẽ minh họa.

### Bước 2: Phiên âm toán học sang LaTeX (Mathematical Transcription)
- Đối chiếu trực tiếp từ ảnh cắt để dịch công thức sang mã nguồn LaTeX hợp lệ:
  - Phân số: `\frac{tử}{mẫu}`
  - Căn bậc hai: `\sqrt{...}`
  - Lượng giác: `\sin, \cos, \tan, \cot`, góc $\frac{\pi}{4}$
  - Hệ phương trình: `\begin{cases} ... \end{cases}`
  - Công thức trong dòng (inline) đặt trong cặp `$ ... $` (ví dụ: `$x \in [0; \pi]$`).
  - Lọc bỏ các câu hỏi lạc đề (nếu trang PDF in chung câu thuộc chuyên đề khác).

### Bước 3: Đóng gói vào Database có cấu trúc (`database_<chuyen_de>.json`)
Mỗi câu hỏi phải được lưu dưới dạng JSON chuẩn để quản lý ngân hàng câu hỏi lâu dài:
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
    "b) Tìm $m$ để phương trình có đúng 2 nghiệm thuộc $[0; \\pi]$."
  ]
}
```

### Bước 4: Biên dịch Native Microsoft Word Equation (`Alt + =`, `m:oMath`)
- Sử dụng cầu nối chuyển đổi chuẩn của Microsoft:
  $$\text{LaTeX} \xrightarrow{\text{latex2mathml}} \text{MathML XML} \xrightarrow{\text{MML2OMML.XSL}} \text{OpenXML } \langle\text{m:oMath}\rangle$$
- File XSLT có sẵn trong máy tính Windows tại:
  `C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL` (hoặc `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL`).
- Đưa trực tiếp thẻ `<m:oMath>` vào cấu trúc đoạn văn của `python-docx`.
- Kết quả: Word nhận diện là **công thức gốc của Word (Equation)**, cho phép bấm `Alt + =` chỉnh sửa số liệu, đổi font Cambria Math tùy ý.

---

## 3. QUY CHUẨN TRÌNH BÀY VĂN BẢN (TYPOGRAPHY & STYLING)
1. **Khổ giấy & Căn lề chuẩn văn bản Bộ GD&ĐT:**
   - Khổ giấy: A4.
   - Căn lề: Trên 2.0 cm, Dưới 2.0 cm, Trái 3.0 cm, Phải 2.0 cm.
2. **Font chữ & Cỡ chữ:**
   - Font: `Times New Roman` cho toàn bộ nội dung.
   - Cỡ chữ nội dung: 12pt, giãn dòng 1.15pt, space after 2-4pt.
   - Tiêu đề câu hỏi: **Bài [số].** màu Deep Navy (`#1E3A8A`), cỡ 12pt, in đậm.
   - Thẻ phân dạng: `[Dạng bài...]` màu Dark Teal (`#0F766E`), cỡ 10.5pt, in đậm.
   - Thẻ nguồn đề: `(Năm học | Đơn vị ra đề)` màu xám Slate (`#64748B`), cỡ 10pt, in nghiêng.
3. **Nguyên tắc làm sạch:**
   - Tuyệt đối KHÔNG có dòng `📄 Nguồn gốc file đề: ...pdf`.
   - KHÔNG để link rút gọn thay cho đề bài; phải soạn ra đầy đủ 100% các câu hỏi.
   - Các ý con $a), b)$ hoặc $A, B, C, D$ được thụt lề đầu dòng $0.25\text{ inch}$ ($0.63\text{ cm}$).

---

## 4. CHECKLIST KIỂM ĐỊNH TỰ ĐỘNG TRƯỚC KHI BÀN GIAO
Trước khi kết luận hoàn thành bất kỳ file đề nào, bắt buộc chạy script kiểm tra:
- [x] **Số lượng câu hỏi:** Đủ 100% số lượng (không bị cắt ngắn thành bài mẫu hay để link).
- [x] **Số lượng công thức Native Word Equation (`m:oMath`):** Đạt $\ge 100\%$ số câu hỏi.
- [x] **Ký tự rác PUA (`0xE000 - 0xF8FF`):** Phải bằng **0**.
- [x] **Ký tự ô vuông lỗi (`□`, `\ufffd`):** Phải bằng **0**.
- [x] **Mã LaTeX thô rò rỉ trong text (`\frac`, `\sqrt`, `\begin{cases}`):** Phải bằng **0**.
- [x] **Dòng rác nguồn gốc file:** Phải bằng **0**.
- [x] **Ký tự thừa / lỗi nhân đôi regex (`\r\r`, `\r`):** Phải bằng **0**.

---

## 5. BẢNG TỔNG HỢP CÁC BẪY LỖI KỸ THUẬT & CÁCH PHÒNG TRÁNH TRIỆT ĐỂ (PITFALLS & SOLUTIONS)

Trong quá trình số hóa thực tế từ 130 đề thi HSG Toán 11, các bẫy kỹ thuật sau đã được phát hiện và phải tuân thủ nghiêm ngặt để không bao giờ tái phạm:

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



