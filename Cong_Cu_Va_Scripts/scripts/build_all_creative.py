import json, os, sys, re
import docx
from build_creative_docx import build_creative_document

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = "Cong_Cu_Va_Scripts/scratch"
out_dir = "06_Bo_De_Sang_Tao_HSG_11"

def check_qa(docx_path):
    doc = docx.Document(docx_path)
    xml = doc._part.blob.decode('utf-8', errors='ignore')
    omml = len(re.findall(r'<m:oMath\b', xml))
    pua = len(re.findall(r'[\uE000-\uF8FF]', xml))
    boxes = len([c for c in re.findall(r'[\u25A0-\u25FF\u2600-\u26FF\uFFFD]', xml) if c not in ('\u25ba', '\u25b6', '\u25b3')])
    r_err = len(re.findall(r'\\r\\r|\\r', xml))
    print(f"  -> File: {os.path.basename(docx_path)} ({os.path.getsize(docx_path):,} B)")
    print(f"     OMML Math: {omml}, Tables: {len(doc.tables)}, PUA: {pua}, Boxes: {boxes}, \\r: {r_err}")

def build_lg():
    p = os.path.join(scratch_dir, "creative_luong_giac.json")
    if not os.path.exists(p):
        print(f"Waiting for {p}...")
        return False
    with open(p, 'r', encoding='utf-8') as f:
        items = json.load(f)
    print(f"Loaded {len(items)} creative LG items.")
    out_docx = os.path.join(out_dir, "01_Sang_Tao_Luong_Giac_HSG_11.docx")
    build_creative_document(
        items,
        "CHUYÊN ĐỀ 1: LƯỢNG GIÁC — BỘ ĐỀ SÁNG TẠO 2026\n20 BÀI TOÁN TỰ LUẬN ĐỈNH CAO CHỌN ĐỘI TUYỂN HSG 11",
        "ST_LG",
        out_docx
    )
    check_qa(out_docx)
    return True

def build_ds():
    p = os.path.join(scratch_dir, "creative_day_so.json")
    if not os.path.exists(p):
        print(f"Waiting for {p}...")
        return False
    with open(p, 'r', encoding='utf-8') as f:
        items = json.load(f)
    print(f"Loaded {len(items)} creative DS items.")
    out_docx = os.path.join(out_dir, "02_Sang_Tao_Day_So_Gioi_Han_HSG_11.docx")
    build_creative_document(
        items,
        "CHUYÊN ĐỀ 2: DÃY SỐ, CẤP SỐ & GIỚI HẠN — BỘ ĐỀ SÁNG TẠO 2026\n20 BÀI TOÁN TỰ LUẬN ĐỈNH CAO CHỌN ĐỘI TUYỂN HSG 11",
        "ST_DS",
        out_docx
    )
    check_qa(out_docx)
    return True

def build_xs():
    p = os.path.join(scratch_dir, "creative_xac_suat.json")
    if not os.path.exists(p):
        print(f"Waiting for {p}...")
        return False
    with open(p, 'r', encoding='utf-8') as f:
        items = json.load(f)
    print(f"Loaded {len(items)} creative XS items.")
    out_docx = os.path.join(out_dir, "03_Sang_Tao_Xac_Suat_To_Hop_HSG_11.docx")
    build_creative_document(
        items,
        "CHUYÊN ĐỀ 3: TỔ HỢP & XÁC SUẤT — BỘ ĐỀ SÁNG TẠO 2026\n20 BÀI TOÁN TỰ LUẬN ĐỈNH CAO CHỌN ĐỘI TUYỂN HSG 11 (GDPT 2018)",
        "ST_XS",
        out_docx
    )
    check_qa(out_docx)
    return True

def build_hh():
    p = os.path.join(scratch_dir, "creative_hinh_hoc.json")
    if not os.path.exists(p):
        print(f"Waiting for {p}...")
        return False
    with open(p, 'r', encoding='utf-8') as f:
        items = json.load(f)
    print(f"Loaded {len(items)} creative HH items.")
    out_docx = os.path.join(out_dir, "04_Sang_Tao_Hinh_Hoc_Khong_Gian_HSG_11.docx")
    img_dir = os.path.join(out_dir, "hinh_ve_3d")
    build_creative_document(
        items,
        "CHUYÊN ĐỀ 4: HÌNH HỌC KHÔNG GIAN — BỘ ĐỀ SÁNG TẠO 2026\n20 BÀI TOÁN TỰ LUẬN ĐỈNH CAO CHỌN ĐỘI TUYỂN HSG 11 (KÈM HÌNH 3D)",
        "ST_HH",
        out_docx,
        img_dir=img_dir
    )
    check_qa(out_docx)
    return True

def build_master_json():
    all_data = {}
    for topic, filename in [
        ("luong_giac", "creative_luong_giac.json"),
        ("day_so", "creative_day_so.json"),
        ("xac_suat", "creative_xac_suat.json"),
        ("hinh_hoc", "creative_hinh_hoc.json")
    ]:
        p = os.path.join(scratch_dir, filename)
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                all_data[topic] = json.load(f)
    if len(all_data) == 4:
        master_path = os.path.join(out_dir, "database_sang_tao_80_bai.json")
        with open(master_path, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2)
        print(f"Master Database saved: {master_path} ({os.path.getsize(master_path):,} B)")

if __name__ == '__main__':
    print("Testing Creative Builder...")
    b_lg = build_lg()
    b_ds = build_ds()
    b_xs = build_xs()
    b_hh = build_hh()
    build_master_json()
