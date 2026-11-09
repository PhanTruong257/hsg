import os
import re
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(ROOT_DIR, 'downloads_hsg_toan11_tu_luan')

def classify_exam(filename):
    name = filename.replace('.pdf', '')
    
    school_year = 'Chưa xác định'
    year_folder = 'Tuyen_tap_tong_hop'
    
    # 1. Tìm khoảng năm học: vd 2025-2026, 2024-2025...
    m_range = re.search(r'(20\d{2})[_\s–-]+(20\d{2})', name)
    if m_range:
        y1, y2 = int(m_range.group(1)), int(m_range.group(2))
        if y2 - y1 == 1:
            school_year = f'Năm học {y1} - {y2}'
            year_folder = f'Nam_hoc_{y1}_{y2}'
        elif y2 - y1 > 1:
            school_year = f'Tuyển tập nhiều năm ({y1} - {y2})'
            year_folder = 'Tuyen_tap_tong_hop'
            
    # 2. Nếu chỉ có 1 năm đơn lẻ (vd: Olympic 30/4 năm 2026, Trại hè Hùng Vương 2024...)
    if school_year == 'Chưa xác định':
        m_single = re.search(r'năm[_\s]+(20\d{2})', name, re.IGNORECASE) or re.search(r'(20\d{2})', name)
        if m_single:
            yr = int(m_single.group(1))
            # Olympic diễn ra vào tháng 4-8, thuộc năm học (yr-1) - yr
            school_year = f'Năm học {yr-1} - {yr} (Kỳ thi năm {yr})'
            year_folder = f'Nam_hoc_{yr-1}_{yr}'

    # 3. Phân loại cấp độ
    level = 'Cấp Trường'
    if any(k in name for k in ['Olympic', 'Trại_hè', 'DH&ĐB_Bắc_Bộ', 'Lê_Hồng_Phong']):
        level = 'Olympic / Chuyên đỉnh cao'
    elif any(k in name for k in ['cấp_tỉnh', 'sở_GD&ĐT', 'tỉnh', 'chuyên']):
        level = 'Cấp Tỉnh / Sở GD&ĐT'
    elif any(k in name for k in ['cụm', 'liên_cụm']):
        level = 'Cụm trường'

    # 4. Xác định tỉnh thành
    province = 'Toàn quốc / Chưa rõ'
    provinces = [
        'Hà_Nội', 'TP_HCM', 'Đà_Nẵng', 'Hải_Phòng', 'Bắc_Ninh', 'Quảng_Trị', 
        'Quảng_Bình', 'Quảng_Nam', 'Quảng_Ngãi', 'Bình_Định', 'Hà_Tĩnh', 
        'Hà_Nam', 'Nghệ_An', 'Thanh_Hóa', 'Thái_Nguyên', 'Tuyên_Quang', 
        'Bạc_Liêu', 'An_Giang', 'Phú_Yên', 'Vĩnh_Phúc', 'Hưng_Yên', 'Bắc_Giang', 
        'BR_VT', 'Nam_Định', 'Lâm_Đồng', 'Cần_Thơ', 'Hải_Dương', 'Bến_Tre', 'Gia_Lai'
    ]
    for p in provinces:
        if p in name:
            province = p.replace('_', ' ')
            break

    display_title = name.replace('_', ' ')
    
    return {
        'filename': filename,
        'display_title': display_title,
        'school_year': school_year,
        'year_folder': year_folder,
        'level': level,
        'province': province
    }

def run_organize():
    # Quét tất cả file PDF (bao gồm cả file ở root và file trong subdirs)
    all_pdf_paths = []
    for root, dirs, files in os.walk(SOURCE_DIR):
        for f in files:
            if f.endswith('.pdf'):
                all_pdf_paths.append(os.path.join(root, f))
                
    print(f'🔍 Tìm thấy {len(all_pdf_paths)} file PDF trong thư mục.')
    
    classified_items = []
    seen_files = set()
    
    for pdf_path in all_pdf_paths:
        fname = os.path.basename(pdf_path)
        if fname in seen_files:
            continue
        seen_files.add(fname)
        
        info = classify_exam(fname)
        classified_items.append(info)
        
        target_dir = os.path.join(SOURCE_DIR, info['year_folder'])
        os.makedirs(target_dir, exist_ok=True)
        
        dest_path = os.path.join(target_dir, fname)
        if os.path.abspath(pdf_path) != os.path.abspath(dest_path):
            shutil.move(pdf_path, dest_path)
            
    # Dọn dẹp các file thừa ở root nếu có
    for item in os.listdir(SOURCE_DIR):
        item_path = os.path.join(SOURCE_DIR, item)
        if os.path.isfile(item_path) and item.endswith('.pdf'):
            os.remove(item_path)

    # Sắp xếp danh mục theo năm mới nhất -> cũ nhất
    classified_items.sort(key=lambda x: (x['school_year'], x['level']), reverse=True)
    
    # Xuất Markdown Index
    md_path = os.path.join(ROOT_DIR, 'DANH_MUC_DE_THI_HSG_THEO_NAM.md')
    with open(md_path, 'w', encoding='utf-8') as md:
        md.write('# 📚 TOÀN TẬP ĐỀ THI HSG TOÁN 11 (TỰ LUẬN) TOANMATH (2015 – 2026)\n\n')
        md.write(f'> Tổng số lượng: **{len(classified_items)} đề thi Tự luận** đã phân loại theo từng năm học.\n\n')
        
        current_year = ''
        count = 1
        for item in classified_items:
            if item['school_year'] != current_year:
                current_year = item['school_year']
                folder_name = item['year_folder']
                md.write(f'\n## 📅 {current_year} (Thư mục: `downloads_hsg_toan11_tu_luan/{folder_name}/`)\n\n')
                md.write('| STT | Tên Đề Thi | Cấp Độ Kỳ Thi | Tỉnh / Thành Phố | File PDF |\n')
                md.write('| :---: | :--- | :---: | :---: | :--- |\n')
                count = 1
            
            rel_link = f"downloads_hsg_toan11_tu_luan/{item['year_folder']}/{item['filename']}"
            md.write(f"| {count} | {item['display_title']} | **{item['level']}** | {item['province']} | [📄 Mở đề]({rel_link}) |\n")
            count += 1
            
    print(f'✅ Đã tổ chức hoàn tất {len(classified_items)} đề thi vào các thư mục năm tương ứng!')
    print(f'📄 Đã cập nhật file mục lục: {md_path}')

if __name__ == '__main__':
    run_organize()
