import docx, re, os, sys
sys.stdout.reconfigure(encoding='utf-8')

docs = [
    ("Chuyên đề 1: Lượng giác", "06_Bo_De_Sang_Tao_HSG_11/01_Sang_Tao_Luong_Giac_HSG_11.docx", 20),
    ("Chuyên đề 2: Dãy số & Giới hạn", "06_Bo_De_Sang_Tao_HSG_11/02_Sang_Tao_Day_So_Gioi_Han_HSG_11.docx", 20),
    ("Chuyên đề 3: Tổ hợp & Xác suất", "06_Bo_De_Sang_Tao_HSG_11/03_Sang_Tao_Xac_Suat_To_Hop_HSG_11.docx", 20),
    ("Chuyên đề 4: Hình học không gian", "06_Bo_De_Sang_Tao_HSG_11/04_Sang_Tao_Hinh_Hoc_Khong_Gian_HSG_11.docx", 20)
]

slash_r = "\\r"
print(f"{'Chuyên đề':<35} {'Dung lượng':<12} {'Số bài':<8} {'OMML Math':<10} {'Callouts':<10} {'PUA':<5} {slash_r:<5} {'Lệch $':<8}")
print("-" * 100)

total_items = 0
total_omml = 0

for name, path, expected in docs:
    assert os.path.exists(path), f"Missing {path}"
    size = f"{os.path.getsize(path):,} B"
    doc = docx.Document(path)
    xml = doc._part.blob.decode('utf-8', errors='ignore')
    omml = len(re.findall(r'<m:oMath\b', xml))
    pua = len(re.findall(r'[\uE000-\uF8FF]', xml))
    r_err = len(re.findall(r'\\r\\r|\\r', xml))
    tables = len(doc.tables)

    # Check unbalanced $ in paragraphs & tables
    bad_dollar = 0
    for p in doc.paragraphs:
        if p.text.count('$') % 2 != 0:
            bad_dollar += 1
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    if p.text.count('$') % 2 != 0:
                        bad_dollar += 1

    total_items += expected
    total_omml += omml

    print(f"{name:<35} {size:<12} {expected:<8} {omml:<10} {tables:<10} {pua:<5} {r_err:<5} {bad_dollar:<8}")

print("-" * 100)
print(f"{'TỔNG CỘNG 4 CHUYÊN ĐỀ SÁNG TẠO':<35} {'':<12} {total_items:<8} {total_omml:<10}")
