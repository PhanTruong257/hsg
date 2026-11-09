import os
import sys
import time
import pymupdf

sys.stdout.reconfigure(encoding='utf-8')

out_dir = r'c:\Users\truongpv\Downloads\Nguyệt\de_thi_previews'
os.makedirs(out_dir, exist_ok=True)

folder = r'c:\Users\truongpv\Downloads\Nguyệt\downloads_hsg_toan11_tu_luan'
t0 = time.time()
count = 0
for root, dirs, files in os.walk(folder):
    for f in files:
        if f.endswith('.pdf'):
            pdf_path = os.path.join(root, f)
            base_name = f.replace('.pdf', '')
            # Chỉ render trang 1 (hoặc thêm trang 2 nếu đề dài 2 trang)
            out_img = os.path.join(out_dir, f'{base_name}_trang1.png')
            if not os.path.exists(out_img):
                try:
                    doc = pymupdf.open(pdf_path)
                    if len(doc) > 0:
                        page = doc[0]
                        pix = page.get_pixmap(dpi=150)
                        pix.save(out_img)
                        count += 1
                except Exception as e:
                    print(f'Error rendering {f}: {e}')

print(f'Rendered {count} exam preview images in {time.time()-t0:.2f}s!')
