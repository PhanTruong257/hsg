/**
 * KỊCH BẢN PHÂN LOẠI FILE ĐÃ CÀO SẴN TRÊN MÁY
 * Tác vụ:
 * 1. Đọc 20 file PDF đã tải trong thư mục "downloads_hsg_toan11"
 * 2. Phân loại theo nội dung đề thi: TỰ LUẬN hay TRẮC NGHIỆM
 * 3. Chuyển các đề TỰ LUẬN sang thư mục "downloads_hsg_toan11_tu_luan"
 * 4. Chuyển các đề TRẮC NGHIỆM sang thư mục "downloads_hsg_toan11_trac_nghiem"
 */

const fs = require('fs');
const path = require('path');

const SOURCE_DIR = path.join(__dirname, 'downloads_hsg_toan11');
const TU_LUAN_DIR = path.join(__dirname, 'downloads_hsg_toan11_tu_luan');
const TRAC_NGHIEM_DIR = path.join(__dirname, 'downloads_hsg_toan11_trac_nghiem');

if (!fs.existsSync(TU_LUAN_DIR)) fs.mkdirSync(TU_LUAN_DIR, { recursive: true });
if (!fs.existsSync(TRAC_NGHIEM_DIR)) fs.mkdirSync(TRAC_NGHIEM_DIR, { recursive: true });

// Danh sách các đề trắc nghiệm đã được kiểm chứng từ nội dung bài viết TOANMATH
const TRAC_NGHIEM_SIGNATURES = [
    'Bắc_Ninh',        // Cụm trường Bắc Ninh (20 TN + 8 đúng sai + 6 ngắn)
    'Thuận_Thành_1',    // Thuận Thành 1 (Đề trắc nghiệm mã đề)
    'Hà_Tĩnh',          // Sở Hà Tĩnh (Trắc nghiệm phương án A B C D)
    'Kiến_Thụy',        // Kiến Thụy Hải Phòng (Mã đề 0004 - 0005)
    'cụm_trường_THPT_số_10', // Cụm 10 Hà Nội (Trắc nghiệm ABCD)
    'Hùng_Vương_–_Đồng_Nai'  // Chọn HSG THPT Hùng Vương (Trắc nghiệm ABCD)
];

async function filterExisting() {
    if (!fs.existsSync(SOURCE_DIR)) {
        console.log('Thư mục downloads_hsg_toan11 chưa tồn tại.');
        return;
    }

    const files = fs.readdirSync(SOURCE_DIR).filter(f => f.endsWith('.pdf'));
    console.log(`🔍 Tìm thấy ${files.length} file PDF đã tải trong thư mục.`);

    let tuLuanCount = 0;
    let tracNghiemCount = 0;

    for (const file of files) {
        const srcPath = path.join(SOURCE_DIR, file);

        const isTracNghiem = TRAC_NGHIEM_SIGNATURES.some(sig => file.includes(sig));

        if (isTracNghiem) {
            tracNghiemCount++;
            const destPath = path.join(TRAC_NGHIEM_DIR, file);
            fs.copyFileSync(srcPath, destPath);
            console.log(`🚫 [TRẮC NGHIỆM] Đã phân loại vào folder Trắc Nghiệm: ${file}`);
        } else {
            tuLuanCount++;
            const destPath = path.join(TU_LUAN_DIR, file);
            fs.copyFileSync(srcPath, destPath);
            console.log(`✅ [TỰ LUẬN] Đã chuyển vào folder Tự Luận: ${file}`);
        }
    }

    console.log('\n=============================================');
    console.log(`🎉 PHÂN LOẠI XONG CÁC FILE ĐÃ TẢI!`);
    console.log(`📁 Thư mục Tự Luận (${tuLuanCount} file): ${TU_LUAN_DIR}`);
    console.log(`📁 Thư mục Trắc Nghiệm (${tracNghiemCount} file): ${TRAC_NGHIEM_DIR}`);
    console.log('=============================================');
}

filterExisting();
