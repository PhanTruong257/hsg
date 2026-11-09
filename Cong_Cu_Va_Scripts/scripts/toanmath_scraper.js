/**
 * ==============================================================================
 * DỰ ÁN: TOANMATH ALL-YEARS CRAWLER (2015 – 2026/2027)
 * Mục tiêu: 
 *  1. Quét TOÀN BỘ 11 trang danh mục HSG Toán 11 từ 2015 đến nay
 *  2. Bộ lọc: BỎ QUA toàn bộ đề trắc nghiệm, CHỈ TẢI đề TỰ LUẬN
 *  3. Auto-skip: Kiểm tra thông minh, file đã có trên máy sẽ bỏ qua ngay lập tức
 *  4. Tự động kích hoạt phân loại thư mục theo từng năm sau khi hoàn tất
 * ==============================================================================
 */

const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { execSync } = require('child_process');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');

puppeteer.use(StealthPlugin());

// ======================== CẤU HÌNH HỆ THỐNG (CONFIG) ========================
const CONFIG = {
    BASE_URL: 'https://toanmath.com/de-thi-hsg-toan-11',
    
    // Toàn bộ lịch sử danh mục HSG Toán 11 trên TOANMATH có đúng 11 trang (2013/2015 -> 2026/2027)
    START_PAGE: 1,
    END_PAGE: 11,
    
    // Thư mục lưu trữ
    DOWNLOAD_DIR: path.join(__dirname, 'downloads_hsg_toan11_tu_luan'),
    
    // File JSON lưu danh sách và phân loại
    OUTPUT_FILE: path.join(__dirname, 'toanmath_tu_luan_results.json'),

    // BẬT BỘ LỌC: Chỉ tải đề Tự luận
    ONLY_TU_LUAN: true,

    // Rate Limiting (Mô phỏng người thật)
    MIN_DELAY: 2000, // 2s
    MAX_DELAY: 4500, // 4.5s

    // Headful (false) để bypass Cloudflare
    HEADLESS: false,

    USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
    NAVIGATION_TIMEOUT: 60000
};

// ======================== CÁC HÀM TIỆN ÍCH ========================

const randomDelay = (min = CONFIG.MIN_DELAY, max = CONFIG.MAX_DELAY) => {
    const delay = Math.floor(Math.random() * (max - min + 1)) + min;
    console.log(`⏳ [Delay] Chờ ${(delay / 1000).toFixed(2)}s...`);
    return new Promise(resolve => setTimeout(resolve, delay));
};

const sanitizeFilename = (rawName) => {
    return rawName
        .replace(/[\\/:*?"<>|]/g, '')
        .replace(/\s+/g, '_')
        .trim()
        .substring(0, 150);
};

/**
 * Kiểm tra xem file đã tồn tại trên máy chưa (tìm trong root và mọi thư mục con)
 */
const fileAlreadyExists = (safeName) => {
    const targetFile = `${safeName}.pdf`;
    
    // 1. Kiểm tra trong DOWNLOAD_DIR và các thư mục con
    if (fs.existsSync(CONFIG.DOWNLOAD_DIR)) {
        const rootPath = path.join(CONFIG.DOWNLOAD_DIR, targetFile);
        if (fs.existsSync(rootPath)) return true;

        const subdirs = fs.readdirSync(CONFIG.DOWNLOAD_DIR, { withFileTypes: true })
            .filter(d => d.isDirectory())
            .map(d => d.name);

        for (const sub of subdirs) {
            if (fs.existsSync(path.join(CONFIG.DOWNLOAD_DIR, sub, targetFile))) {
                return true;
            }
        }
    }

    // 2. Kiểm tra trong thư mục cũ downloads_hsg_toan11
    const oldPath = path.join(__dirname, 'downloads_hsg_toan11', targetFile);
    if (fs.existsSync(oldPath)) return true;

    return false;
};

/**
 * BỘ LỌC ĐỀ THI: Phân biệt TỰ LUẬN vs TRẮC NGHIỆM
 */
const detectExamType = (title, contentText) => {
    const fullText = (title + ' ' + contentText).toLowerCase();
    
    const reasons = [];
    if (fullText.includes('trắc nghiệm') || fullText.includes('trac nghiem')) {
        reasons.push('Chứa từ khóa "trắc nghiệm"');
    }
    if (fullText.includes('mã đề') || fullText.includes('ma de')) {
        reasons.push('Có thông tin "mã đề"');
    }
    const hasOptionsPattern = /\b[A-D]\.\s+[A-Za-z0-9\(\[\$]/.test(contentText) || 
                              (/\bA\.\s/.test(contentText) && /\bB\.\s/.test(contentText));
    if (hasOptionsPattern) {
        reasons.push('Chứa phương án lựa chọn A. B. C. D.');
    }

    if (reasons.length > 0) {
        return {
            isTuLuan: false,
            type: 'TRẮC NGHIỆM / HỖN HỢP',
            reason: reasons.join('; ')
        };
    }

    return {
        isTuLuan: true,
        type: 'TỰ LUẬN 100%',
        reason: 'Đề tự luận thuần túy (không có câu trắc nghiệm)'
    };
};

const downloadPdfFile = async (fileUrl, destinationPath, cookieHeader) => {
    try {
        const response = await axios({
            method: 'GET',
            url: fileUrl,
            responseType: 'stream',
            headers: {
                'User-Agent': CONFIG.USER_AGENT,
                'Referer': 'https://toanmath.com/',
                'Cookie': cookieHeader
            },
            timeout: 35000
        });

        const writer = fs.createWriteStream(destinationPath);
        response.data.pipe(writer);

        return new Promise((resolve, reject) => {
            writer.on('finish', resolve);
            writer.on('error', reject);
        });
    } catch (err) {
        throw new Error(`Lỗi tải stream PDF: ${err.message}`);
    }
};

const handleCloudflareChallenge = async (page) => {
    const title = await page.title();
    if (title.includes('Just a moment...') || title.includes('Cloudflare') || title.includes('Attention Required')) {
        console.warn('⚠️ [Cloudflare] Đang kiểm tra bot... Vui lòng click xác thực nếu có.');
        await page.waitForFunction(
            () => !document.title.includes('Just a moment...') && !document.title.includes('Cloudflare'),
            { timeout: 30000 }
        ).catch(() => {});
    }
};

const getExecutablePath = () => {
    const possiblePaths = [
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
        'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'
    ];
    for (const p of possiblePaths) {
        if (fs.existsSync(p)) return p;
    }
    return undefined;
};

// ======================== LUỒNG CHẠY CHÍNH ========================

async function main() {
    if (!fs.existsSync(CONFIG.DOWNLOAD_DIR)) {
        fs.mkdirSync(CONFIG.DOWNLOAD_DIR, { recursive: true });
    }

    console.log('=============================================================');
    console.log('🚀 KHỞI ĐỘNG CRAWLER TOÀN BỘ KHO ĐỀ HSG TOÁN 11 (2015 – 2026)');
    console.log(`📂 Thư mục lưu đề Tự luận: ${CONFIG.DOWNLOAD_DIR}`);
    console.log(`📄 Quét toàn bộ: Trang ${CONFIG.START_PAGE} đến Trang ${CONFIG.END_PAGE}`);
    console.log('=============================================================');

    const launchOptions = {
        headless: CONFIG.HEADLESS,
        defaultViewport: null,
        args: [
            '--start-maximized',
            '--disable-blink-features=AutomationControlled',
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage'
        ]
    };

    const sysBrowser = getExecutablePath();
    if (sysBrowser) launchOptions.executablePath = sysBrowser;

    const browser = await puppeteer.launch(launchOptions);
    const page = await browser.newPage();
    await page.setUserAgent(CONFIG.USER_AGENT);

    const report = {
        totalScanned: 0,
        tuLuanSaved: 0,
        tracNghiemSkipped: 0,
        alreadyDownloaded: 0,
        items: []
    };

    try {
        for (let pageNum = CONFIG.START_PAGE; pageNum <= CONFIG.END_PAGE; pageNum++) {
            const pageUrl = pageNum === 1 
                ? CONFIG.BASE_URL 
                : `${CONFIG.BASE_URL}/page/${pageNum}`;

            console.log(`\n=============================================================`);
            console.log(`📖 [TRANG ${pageNum}/${CONFIG.END_PAGE}] Đang quét danh mục: ${pageUrl}`);
            console.log(`=============================================================`);

            await page.goto(pageUrl, {
                waitUntil: 'networkidle2',
                timeout: CONFIG.NAVIGATION_TIMEOUT
            });

            await handleCloudflareChallenge(page);
            await randomDelay(1200, 2500);

            const articles = await page.evaluate(() => {
                const results = [];
                const links = document.querySelectorAll('article.mh-posts-grid-item h3.entry-title a, #main-content article h3.entry-title a');
                links.forEach(a => {
                    const title = a.innerText.trim();
                    const url = a.href;
                    if (url && !results.some(r => r.url === url)) {
                        results.push({ title, url });
                    }
                });
                return results;
            });

            if (articles.length === 0) {
                console.log(`ℹ️ Trang ${pageNum} không còn bài viết nào. Kết thúc phân trang.`);
                break;
            }

            console.log(`📋 Trang ${pageNum} có ${articles.length} đề thi. Bắt đầu xử lý...`);

            for (let i = 0; i < articles.length; i++) {
                const item = articles[i];
                report.totalScanned++;

                console.log(`\n[Trang ${pageNum} - Đề ${i + 1}/${articles.length}] ${item.title}`);

                const safeName = sanitizeFilename(item.title);
                const localPdfPath = path.join(CONFIG.DOWNLOAD_DIR, `${safeName}.pdf`);
                
                // KIỂM TRA ĐÃ CÓ TRÊN MÁY CHƯA
                if (fileAlreadyExists(safeName)) {
                    console.log(`⏩ [ĐÃ CÓ TRÊN MÁY] File đã tồn tại, tự động bỏ qua.`);
                    report.alreadyDownloaded++;
                    continue;
                }

                await randomDelay(CONFIG.MIN_DELAY, CONFIG.MAX_DELAY);

                try {
                    await page.goto(item.url, {
                        waitUntil: 'networkidle2',
                        timeout: CONFIG.NAVIGATION_TIMEOUT
                    });

                    await handleCloudflareChallenge(page);

                    const articleData = await page.evaluate(() => {
                        const contentEl = document.querySelector('.entry-content');
                        const contentText = contentEl ? contentEl.innerText : '';

                        let pdf = null;
                        const directLink = document.querySelector('a:has(button.pdf-download), a[href*="/toanmath-pdf/"], a[download]');
                        if (directLink && directLink.href && directLink.href.includes('.pdf')) {
                            pdf = directLink.href;
                        }

                        if (!pdf) {
                            const iframe = document.querySelector('iframe.wonderplugin-pdf-iframe, iframe[src*="pdfjslight"]');
                            if (iframe && iframe.src) {
                                try {
                                    const params = new URL(iframe.src);
                                    const f = params.searchParams.get('file');
                                    if (f && f.toLowerCase().includes('.pdf')) pdf = f;
                                } catch (e) {}
                            }
                        }

                        if (!pdf) {
                            const anyPdf = document.querySelector('.entry-content a[href$=".pdf"]');
                            if (anyPdf) pdf = anyPdf.href;
                        }

                        return { contentText, pdfUrl: pdf };
                    });

                    // LỌC TỰ LUẬN
                    const examCheck = detectExamType(item.title, articleData.contentText);

                    if (CONFIG.ONLY_TU_LUAN && !examCheck.isTuLuan) {
                        console.log(`🚫 [BỎ QUA - TRẮC NGHIỆM] Lý do: ${examCheck.reason}`);
                        report.tracNghiemSkipped++;
                        report.items.push({
                            title: item.title,
                            url: item.url,
                            type: examCheck.type,
                            reason: examCheck.reason,
                            action: 'SKIPPED'
                        });
                        continue;
                    }

                    console.log(`✅ [ĐẠT YÊU CẦU - TỰ LUẬN] Chuẩn bị tải file...`);

                    if (!articleData.pdfUrl) {
                        console.warn(`⚠️ Không tìm thấy link PDF trong bài này.`);
                        continue;
                    }

                    const cookies = await page.cookies();
                    const cookieHeader = cookies.map(c => `${c.name}=${c.value}`).join('; ');

                    console.log(`⬇️ Đang tải: ${safeName}.pdf...`);
                    await downloadPdfFile(articleData.pdfUrl, localPdfPath, cookieHeader);
                    console.log(`💾 ĐÃ LƯU: ${localPdfPath}`);

                    report.tuLuanSaved++;
                    report.items.push({
                        title: item.title,
                        url: item.url,
                        type: 'TỰ LUẬN',
                        localPath: localPdfPath,
                        action: 'DOWNLOADED'
                    });

                } catch (postErr) {
                    console.error(`❌ Lỗi xử lý bài viết: ${postErr.message}`);
                }
            }
        }

        fs.writeFileSync(CONFIG.OUTPUT_FILE, JSON.stringify(report, null, 2), 'utf-8');

        console.log(`\n=============================================================`);
        console.log(`🎉 HOÀN THÀNH TẢI TOÀN BỘ CÁC TRANG!`);
        console.log(`📊 Tổng số bài quét qua: ${report.totalScanned}`);
        console.log(`✅ Đề TỰ LUẬN tải mới: ${report.tuLuanSaved}`);
        console.log(`⏩ Đề đã có sẵn trên máy: ${report.alreadyDownloaded}`);
        console.log(`🚫 Đề TRẮC NGHIỆM bỏ qua: ${report.tracNghiemSkipped}`);
        console.log(`=============================================================`);

        // TỰ ĐỘNG CHẠY PHÂN LOẠI THEO NĂM SAU KHI CÀO XONG
        console.log('🔄 Đang tự động cập nhật danh mục và phân loại theo năm...');
        try {
            execSync('python organize_by_year.py', { stdio: 'inherit' });
        } catch (e) {
            console.error('Lỗi khi chạy organize_by_year.py:', e.message);
        }

    } catch (fatalErr) {
        console.error('💥 Lỗi nghiêm trọng:', fatalErr);
    } finally {
        await browser.close();
        console.log('🔒 Đã đóng trình duyệt.');
    }
}

main();
