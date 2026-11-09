# CHECKPOINT CONTINUITY PACKET (CCP) — MASTER THÚY ENGLISH STUDIO v2.0

- **Mã Checkpoint:** `CHECKP_MASTER_THUY_ENGLISH_STUDIO_v2_0_CCP`
- **Ngày giờ phát hành:** `2026-11-05T02:08:00+07:00`
- **Phiên bản hệ thống:** `v2.0 (Production & Cloud-Ready — Hybrid Zero-Server-Cost)`
- **Thư mục gốc Production (Local):** `D:\YOURDREAM.VN\master-thuy-english-studio`
- **Thư mục quản trị chính thức (Google Drive):** `G:\Drive của tôi\YOURDREAM_AI_WORKSPACE\ENGLISH_TEACHING_STUDIO`
- **Nhánh xuất bản GCS Bucket:** `gs://ai-studio-bucket-336999988003-asia-southeast1/master-thuy-studio/`
- **Chủ nhiệm chuyên môn:** `MASTER THÚY` (Zalo: `0836 686 874` — Không logo đồ họa, chuẩn nhận diện chữ in đậm)

---

## 1. BẢNG NGHIỆM THU 6 HẠNG MỤC NÂNG CẤP v2.0

| Hạng mục | Nội dung thực thi | Bằng chứng kỹ thuật (Verified) | Trạng thái 4 Cấp YOURDREAM |
| :--- | :--- | :--- | :--- |
| **1. Di trú khỏi vùng nháp `scratch`** | Chuyển toàn bộ dự án từ `C:\Users\Admin\.gemini\antigravity\scratch\master-thuy-english-studio` sang thư mục chuẩn `D:\YOURDREAM.VN\master-thuy-english-studio`. Quy hoạch kho lưu trữ nguyên bản `raw_63_archive/drive_37` (37 tệp) và `raw_63_archive/unsaved_26` (26 tệp). | Toàn bộ đường dẫn trong Python (`build_knowledge_maps_and_shards.py`, `generate_dual_worksheet.py`, `server_and_print_bridge.py`, `start_master_thuy_bridge.ps1`) trỏ tuyệt đối về `D:\YOURDREAM.VN\master-thuy-english-studio`. | **Cấp 2 — Đã kiểm chứng kỹ thuật** (Sẵn sàng Cấp 3/4) |
| **2. Tầng Dữ liệu Tiếng Anh (Đủ 63 tệp & Giữ định dạng)** | Quét và bóc tách toàn bộ **63 tệp gốc** (gồm cả bảng `.docx` `1x4`, bảng lồng nhau `2x2`, 3 đoạn văn sửa lỗi `Proofreading` và tệp `.pdf`) thành **20 Bộ đề chuẩn** với **953 câu hỏi nguyên tử** trong `db/master_thuy_studio.db` và `public/shards/items/*.json`. | - **953/953 câu** đạt 100% qua 4 cổng QA (`verify_questions.py`).<br>- **337 câu** bảo toàn 100% thẻ `<u>...</u>` và `<b>...</b>` từ Word run.<br>- **50 biển báo** trích xuất ảnh gốc (`public/assets/signs/sign_01.png`..`sign_50.png`) đóng khung `.sign-visual-box`. | **Cấp 2 — Đã kiểm chứng kỹ thuật 100%** |
| **3. Giao diện & Sư phạm (`teacher/` & `student/`)** | - Chuyển 100% đường dẫn sang tương đối (`./...`, `../...`) + tạo `public/favicon.ico` xóa sạch lỗi 404.<br>- Nút loa phát âm chuẩn **🇬🇧 Anh-Anh (`en-GB`) / 🇺🇸 Anh-Mỹ (`en-US`)** + tốc độ chậm `0.75x` bằng Web Speech API (0 đồng).<br>- Cơ chế **Làm bài 2 Pha** (Pha 1 khóa đáp án & giải thích; Pha 2 bấm Nộp bài mới mở đáp án kèm giải thích theo 7 nhóm bẫy `map2-traps-and-levels.json`).<br>- Bộ chấm tự luận chuẩn hóa viết tắt tương đương (`don't = do not`, `isn't = is not`, v.v.). | - Kiểm thử HTTP E2E đạt `200 OK` trên toàn bộ tài nguyên tĩnh (0 lỗi 404).<br>- Hàm `MT_HYBRID.isWrittenAnswerMatch()` xử lý đầy đủ 27 cặp viết tắt tiếng Anh phổ biến.<br>- Giao diện 100% ngôn ngữ sư phạm thân thiện, không thuật ngữ kỹ thuật. | **Cấp 2 — Đã kiểm chứng kỹ thuật** |
| **4. Kiến trúc Hybrid Zero-Server-Cost (GCS + Localhost)** | Tự động phát hiện môi trường trong `public/shared/js/hybrid-cloud-engine.js`:<br>- Khi chạy trên `localhost:8765`: gọi Python API.<br>- Khi chạy tĩnh trên GCS Bucket (`storage.googleapis.com`) lúc máy tính tắt:<br>  (a) `/api/generate-docx`: Trộn **4 mã đề 101–104** và xuất trực tiếp file `.docx` (OpenXML ZIP chuẩn CRC32) gồm Bản Học Sinh + Bản Giáo Viên (có Bảng đáp án đầu Trang 1 & Bảng từ vựng gấp đôi IPA, bảo toàn `<w:u>` và `<w:b>`) ngay trên trình duyệt.<br>  (b) `/api/submissions`: Chấm trực tiếp trên trình duyệt & đồng bộ `localStorage` + server.<br>  (c) `/api/print-job`: Đẩy lệnh in vào hàng đợi `00_PRINT_QUEUE` trên Google Drive; daemon `poll_drive_print_queue_loop()` trong `server_and_print_bridge.py` tự động quét, xuất Word và in khi máy tính bật. | Kiểm thử E2E (`verify_v2_e2e.py`):<br>- API `/api/generate-docx` sinh đủ 4 mã đề `101-104` (8 file `.docx`).<br>- Thả vé in `PJ_CLOUD_DRIVE_TEST_V2.json` vào `G:\Drive của tôi\YOURDREAM_AI_WORKSPACE\ENGLISH_TEACHING_STUDIO\00_PRINT_QUEUE`, daemon tự động nhận diện sau 5 giây, sinh file Word và chuyển vé sang `00_PRINT_QUEUE/PROCESSED/`. | **Cấp 2 — Đã kiểm chứng kỹ thuật** |
| **5. Đóng gói GCS Bucket & Kết nối `/control/`** | Đóng gói `dist-gcs/master-thuy-studio/` (156 tệp, 14.99 MB) sẵn sàng xuất bản lên `gs://ai-studio-bucket-336999988003-asia-southeast1/master-thuy-studio/`. Cập nhật `dist-gcs/control/index.html` thêm nút điều hướng trên thanh Header và thẻ sản phẩm Cấp 4 trỏ trực tiếp sang `../master-thuy-studio/teacher/index.html` và `../master-thuy-studio/student/index.html`. | Cả đường dẫn tương đối `../master-thuy-studio/teacher/index.html` và đường dẫn trực tiếp đều trả về `200 OK` khi kiểm thử từ `/control/index.html`. Kèm kịch bản `dist-gcs/deploy_to_gcs.ps1` và `deploy_to_gcs.sh`. | **Cấp 2 — Đã kiểm chứng kỹ thuật** |
| **6. Đồng bộ Drive, GitHub & Xuất Checkpoint** | Kiểm định `verify_questions.py` đạt **953/953 (100.0%)**, khởi tạo Git commit v2.0, đồng bộ mã nguồn + học liệu + các file `.docx` xuất ra + gói `dist-gcs` lên `G:\Drive của tôi\YOURDREAM_AI_WORKSPACE\ENGLISH_TEACHING_STUDIO` và lập tệp CCP này. | Đã đọc kiểm chứng ngược (verified readback) từ ổ `G:\Drive của tôi\YOURDREAM_AI_WORKSPACE\ENGLISH_TEACHING_STUDIO\CHECKPOINTS\MASTER_THUY_CHECKP_v2_0_20261105`. | **Cấp 2 — Đã kiểm chứng kỹ thuật** |
| **7. Bảo mật Đáp án 2 Pha Web Crypto & Trạm Nghiên Cứu 06:00 Sáng (`08_AI_LIBRARIAN`)** | - Tích hợp mô-đun `ZeroLeakAssessment` (Web Crypto API: `HMAC-SHA256` khóa một chiều Pha 1 chống F12 xem trộm đáp án + `AES-256-GCM` mở khóa lời giải Pha 2) vào `public/shared/js/hybrid-cloud-engine.js`, `public/student/index.html`, toàn bộ 25 tệp `public/shards/items/WS_*.json` (1.165 câu) và `student/v1.html` của Toán học YOURDREAM.<br>- Hợp nhất nhánh `feature/daily-06am-research-workflow` vào `main` và cập nhật `control/index.html` hiển thị trực tiếp **Trạm Nghiên Cứu 06:00 Sáng (`08_AI_LIBRARIAN`)**. | - Xác thực tự động bằng Node.js Web Crypto API đạt `1165/1165` câu hỏi vượt qua kiểm thử mã hóa/giải mã 2 chiều (`HMAC-SHA256` + `AES-256-GCM`).<br>- Kiểm định `verify_questions.py` đạt **953/953 (100.0%)**. | **Cấp 2 — Đã kiểm chứng kỹ thuật 100%** |

---

## 2. KIỂM KÊ DỮ LIỆU 20 BỘ ĐỀ CHUẨN (953 CÂU HỎI TỪ 63 TỆP GỐC)

| Mã Bộ Đề (`id`) | Khối | Tên Học Liệu Trên WebApp | Số câu | Đặc tả Định dạng & Nguồn Gốc |
| :--- | :---: | :--- | :---: | :--- |
| `WS_G06_U03_ADV` | Lớp 6 | Lớp 6 — Unit 3: My Friends (Nâng cao 50 câu) | 50 | Giữ nguyên `<u>...</u>` ngữ âm & từ vựng nâng cao |
| `WS_G06_ERROR_48` | Lớp 6 | Lớp 6 — Sửa lỗi sai & Đoạn văn Proofreading (Units 1–3) | 48 | 36 câu tìm lỗi gạch chân `<u>` + 3 bảng Proofreading (Câu 37–48) |
| `WS_G06_SIGNS_50` | Lớp 6 | Lớp 6 — 50 Câu Biển Báo Thực Tế (Signs Around Us) | 50 | 50 ảnh biển báo gốc (`sign_01.png`..`sign_50.png`) trong `.sign-visual-box` |
| `WS_G06_PS_30` | Lớp 6 | Lớp 6 — Chuyên đề Thì Hiện Tại Đơn (Present Simple) | 30 | Từ bộ sao lưu 26 tệp gốc (`Present_Simple_G6`) |
| `WS_G06_PC_30` | Lớp 6 | Lớp 6 — Chuyên đề Thì Hiện Tại Tiếp Diễn (Present Continuous) | 30 | Từ bộ sao lưu 26 tệp gốc (`Present_Continuous_G6`) |
| `WS_G06_THERE_30` | Lớp 6 | Lớp 6 — Cấu trúc There is / There are & Giới từ chỉ nơi chốn | 30 | Từ bộ sao lưu 26 tệp gốc (`There_is_There_are_G6`) |
| `WS_G06_MID1_50` | Lớp 6 | Lớp 6 — Đề Kiểm Tra Giữa Kì I (Mục tiêu 8–10 điểm) | 50 | Từ bộ sao lưu 26 tệp gốc (`G6_Midterm_8_10`) |
| `WS_G07_MID1_50` | Lớp 7 | Lớp 7 — Đề Ôn Tập Giữa Kì I (50 câu Trắc nghiệm) | 50 | Quét cả đoạn văn và bảng phương án `1x4` trong Word |
| `WS_G07_ZALO_50` | Lớp 7 | Lớp 7 — Đề Tổng Hợp Giữa Kì I (Bản Final Zalo) | 50 | Quét đầy đủ câu hỏi ngữ âm `<u>` + viết lại câu |
| `WS_G07_PS_30` | Lớp 7 | Lớp 7 — Chuyên đề Hiện Tại Đơn & Trạng từ tần suất | 30 | Từ bộ sao lưu 26 tệp gốc (`Present_Simple_G7`) |
| `WS_G07_TARGET_45` | Lớp 7 | Lớp 7 — Đề Luyện Thi Mục Tiêu 8–10 Điểm | 45 | Từ bộ sao lưu 26 tệp gốc (`G7_Target_8_10`) |
| `WS_G08_MID1_50` | Lớp 8 | Lớp 8 — Đề Kiểm Tra Giữa Kì I (50 câu chuẩn) | 50 | Bảo toàn `<u>` phát âm `/ʊ/` vs `/uː/`, `/ə/` vs `/ɪ/` |
| `WS_G08_MID1_80` | Lớp 8 | Lớp 8 — Đại Đề Ôn Tập Giữa Kì I (80 câu Toàn diện) | 80 | Đề lớn 80 câu (Trắc nghiệm + Tự luận chia động từ & viết câu) |
| `WS_G09_MID1_50` | Lớp 9 | Lớp 9 — Đề Kiểm Tra Giữa Kì I (50 câu Trọng tâm) | 50 | Phrasal verbs, Double comparatives, Wish clauses |
| `WS_G09_MID1_80` | Lớp 9 | Lớp 9 — Đại Đề Ôn Tập Giữa Kì I (Units 1–3, 80 câu) | 80 | Quét đầy đủ bảng phương án `1x4` + bài đọc hiểu |
| `WS_G09_WF_50` | Lớp 9 | Lớp 9 — Chuyên đề Cấu Tạo Từ (Word Formation 50 câu) | 50 | Từ bộ sao lưu 26 tệp gốc (`Word_Form_G9`) |
| `WS_TOP_GET_HAVE` | Chuyên đề | Chuyên đề Đặc Biệt: Cấu trúc GET & HAVE (40 câu) | 40 | Causative form (`have sb do` / `get sb to do` / `have sth done`) |
| `WS_G11_101` | Lớp 11 | Lớp 11 — Đề Kiểm Tra Chuẩn THPT (Mã đề 101) | 40 | Đề chuẩn THPT Quốc Gia (Cloze test, Reading, Arrangement) |
| `WS_G11_102` | Lớp 11 | Lớp 11 — Đề Kiểm Tra Chuẩn THPT (Mã đề 102) | 40 | Đề chuẩn THPT Quốc Gia (Hoán vị đối chiếu Mã 101) |
| `WS_G12_101` | Lớp 12 | Lớp 12 — Đề Luyện Thi Tốt Nghiệp THPT (60 câu) | 60 | Đề tổng hợp Lớp 12 từ bộ sao lưu 26 tệp gốc |
| **TỔNG CỘNG** | **Lớp 6–12** | **20 Bộ Đề Chuẩn (Bao phủ 63/63 tệp gốc)** | **953 câu** | **100.0% vượt qua 4 cổng QA tự động** |

---

## 3. HƯỚNG DẪN KHỞI ĐỘNG NHANH & PHỤC HỒI (RESTORE INSTRUCTIONS)

### 3.1. Chạy trên Máy tính Cô Thúy (Localhost + Tự động In từ Google Drive)
```powershell
powershell -ExecutionPolicy Bypass -File "D:\YOURDREAM.VN\master-thuy-english-studio\local-pc-bridge\start_master_thuy_bridge.ps1"
```
- **Trang chủ Studio:** `http://localhost:8765/index.html`
- **Phòng của Cô Thúy:** `http://localhost:8765/teacher/index.html`
- **Góc Học Sinh:** `http://localhost:8765/student/index.html`
- **Trang Điều Hành Control:** `http://localhost:8765/control/index.html`

### 3.2. Xuất bản lên Google Cloud Storage (Khi chạy trên Cloud Shell / máy có `gcloud`)
```bash
bash dist-gcs/deploy_to_gcs.sh
```
- Khi chạy trên GCS Bucket mà máy tính ở nhà tắt, `hybrid-cloud-engine.js` tự động chuyển sang chế độ **Cloud-Ready (0 đồng server)**: trộn và tải trực tiếp 4 mã đề `.docx` ngay trên trình duyệt và đẩy lệnh in vào `00_PRINT_QUEUE`.
