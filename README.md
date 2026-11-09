# 🇬🇧 YOURDREAM ENGLISH — MASTER THÚY ENGLISH STUDIO v2.0

> **Phiên bản:** `v2.0 (Production & Cloud-Ready — Hybrid Zero-Server-Cost & Web Crypto 2-Phase Assessment)`  
> **Chủ nhiệm chuyên môn:** `MASTER THÚY` (Thạc sĩ · Giảng viên Tiếng Anh · Zalo / ĐT: `0836 686 874`)  
> **Chuẩn chương trình:** SGK *Global Success* (Lớp 6 → Lớp 12)  
> **Kho lưu trữ nội bộ cố định:** `/shared/user/.gemini/yourdream_english/studio_v2_prod/`

---

## 1. Tổng Quan Kiến Trúc Hệ Thống

**MASTER THÚY — English Studio v2.0** là hệ sinh thái quản lý học liệu, trộn đề Word (`.docx`) 4 mã đề (`101–104`), giảng dạy sơ đồ tư duy ngữ pháp/từ vựng và kiểm tra đánh giá trực tuyến 2 pha dành cho môn Tiếng Anh phổ thông (Lớp 6–12):

1. **Ngân Hàng Học Liệu Chuẩn Hóa (63 Tệp Gốc → 20 Bộ Đề → 953 Câu Hỏi Nguyên Tử):**
   - Bóc tách đầy đủ **63 tệp gốc** (`.docx` bảng `1x4`, bảng lồng nhau `2x2`, đoạn văn sửa lỗi `Proofreading` và `.pdf`) thành **20 bộ đề tương tác** với **953 câu hỏi nguyên tử** trong SQLite (`db/master_thuy_studio.db`) và JSON Shards (`public/shards/items/*.json`).
   - Bảo toàn **100% định dạng gốc** `<u>...</u>` (gạch chân phát âm/trọng âm/tìm lỗi sai — 337 câu) và `<b>...</b>`, cùng **50 biển báo thực tế** (`public/assets/signs/sign_01.png`..`sign_50.png`) đóng khung `.sign-visual-box`.
   - Đạt **953/953 (100.0%)** qua 4 cổng kiểm định tự động (`engine-python/qa_gates/verify_questions.py`).

2. **Bản Đồ Tri Thức 4 Tầng (Map 1 → Map 4):**
   - **Map 1 (`map1-curriculum-g6-g12.json`):** Khung chương trình *Global Success* Lớp 6–12, khóa chặt danh mục ngữ pháp vượt cấp và tích hợp bảng từ vựng gấp đôi IPA.
   - **Map 2 (`map2-traps-and-levels.json`):** Phân loại vi kỹ năng & **7 nhóm bẫy phòng thi Tiếng Anh** kèm quy tắc phản xạ 5 giây.
   - **Map 3 (`map3-grammar-mindmaps.json`):** Sơ đồ tư duy ngữ pháp có nút **[ ? ]** tự che công thức (Active Recall).
   - **Map 4 (`map4-archetype-templates.json`):** Khuôn xuất bản phiếu bài tập A4 chuẩn nhận diện **MASTER THÚY** (tách biệt Bản Học Sinh & Bản Giáo Viên).

3. **Kiến Trúc Hybrid Zero-Server-Cost (`public/shared/js/hybrid-cloud-engine.js`):**
   - **Chế độ Local Bridge (`localhost:8765`):** Kết nối trực tiếp với Python API (`local-pc-bridge/server_and_print_bridge.py`) và cơ sở dữ liệu SQLite.
   - **Chế độ Cloud-Ready (GCS Bucket — 0 đồng máy chủ):**
     - Trộn **4 mã đề `101–104`** và đóng gói trực tiếp tệp `.docx` chuẩn OpenXML (Bản Học Sinh + Bản Giáo Viên có Bảng đáp án đối chiếu 4 mã đề ở Trang 1 & Bảng từ vựng gấp đôi IPA) ngay trên trình duyệt.
     - Đẩy lệnh in từ điện thoại vào hàng đợi `00_PRINT_QUEUE` trên Google Drive; daemon `poll_drive_print_queue_loop()` tự động quét, sinh file Word và in khi máy tính bật.
   - **Loa Phát Âm Chuẩn 🇬🇧 Anh-Anh (`en-GB`) / 🇺🇸 Anh-Mỹ (`en-US`):** Hỗ trợ tốc độ chuẩn `1.0x` và chậm `0.75x` qua Web Speech API.

4. **Bảo Mật Đáp Án 2 Pha (`ZeroLeakAssessment` — Web Crypto API):**
   - **Pha 1 (Đang làm bài):** Khóa một chiều bằng `HMAC-SHA256` chống xem trộm đáp án qua DevTools (`F12`), đồng thời hỗ trợ chấm tự luận chấp nhận 27 cặp viết tắt tương đương (`don't = do not`, `isn't = is not`, v.v.).
   - **Pha 2 (Sau khi nộp bài):** Giải mã lời giải chi tiết & cảnh báo bẫy bằng `AES-256-GCM`.

---

## 2. Cấu Trúc Thư Mục Dự Án

```text
yourdream-english/
├── CHECKP_MASTER_THUY_ENGLISH_STUDIO_v2_0_CCP.md   # Hồ sơ nghiệm thu & Checkpoint Continuity Packet v2.0
├── README.md                                       # Tài liệu tổng quan kiến trúc & vận hành dự án
├── db/
│   ├── master_thuy_studio.db                       # Cơ sở dữ liệu SQLite (953 câu hỏi, 20 bộ đề, học sinh, bài nộp)
│   └── schema.sql                                  # Lược đồ CSDL 8 bảng chuẩn Production
├── engine-python/
│   ├── build_knowledge_maps_and_shards.py          # Bộ bóc tách 63 tệp gốc -> SQLite & JSON Shards (Map 1-4)
│   ├── generate_dual_worksheet.py                  # Bộ trộn 4 mã đề (101-104) & xuất Word (.docx) Học sinh / Giáo viên
│   ├── build_gcs_bundle.py                         # Đóng gói bản tĩnh dist-gcs/ & kết nối cổng điều hành /control/
│   ├── sync_drive_and_git_v2.py                    # Đồng bộ Google Drive, Git & kiểm chứng SHA-256
│   └── qa_gates/
│       ├── verify_questions.py                     # Kiểm định tự động 4 cổng sư phạm & kỹ thuật (953/953 PASS)
│       └── verify_v2_e2e.py                        # Kiểm định End-to-End HTTP, API trộn đề & hàng đợi in 00_PRINT_QUEUE
├── local-pc-bridge/
│   ├── server_and_print_bridge.py                  # Máy chủ HTTP cổng 8765 + SQLite API + Daemon quét 00_PRINT_QUEUE
│   └── start_master_thuy_bridge.ps1                # Kịch bản PowerShell khởi động nhanh trên máy tính Cô Thúy
├── public/
│   ├── index.html                                  # Cổng chào chính Master Thúy English Studio v2.0
│   ├── teacher/index.html                          # Phòng làm việc của Cô Thúy (Tủ đề, 4 Bản đồ, Trộn đề, Sổ điểm)
│   ├── student/index.html                          # Góc học tập Học sinh (Làm bài 2 pha Web Crypto, Nghe UK/US)
│   ├── control/index.html                          # Cổng điều hành tổng hợp & Trạm Nghiên Cứu 06:00 Sáng
│   ├── shared/
│   │   ├── css/studio.css                          # Hệ thống giao diện sư phạm chuẩn Master Thúy
│   │   └── js/hybrid-cloud-engine.js               # Nhân Hybrid Cloud, OpenXML DOCX Builder & ZeroLeakAssessment
│   ├── shards/
│   │   ├── map1-curriculum-g6-g12.json             # Bản đồ 1: Khung chương trình Global Success 6-12
│   │   ├── map2-traps-and-levels.json              # Bản đồ 2: 7 nhóm bẫy phòng thi Tiếng Anh
│   │   ├── map3-grammar-mindmaps.json              # Bản đồ 3: Sơ đồ tư duy ngữ pháp [ ? ]
│   │   ├── map4-archetype-templates.json           # Bản đồ 4: Khuôn phiếu bài tập chuẩn A4
│   │   ├── worksheets-catalog.json                 # Danh mục 20 bộ đề chuẩn
│   │   ├── qa-verification-report.json             # Báo cáo kết quả kiểm định 4 cổng QA
│   │   └── items/*.json                            # Các gói câu hỏi phân mảnh theo từng bộ đề (WS_*.json)
│   ├── assets/signs/                               # 50 ảnh biển báo thực tế (sign_01.png .. sign_50.png)
│   ├── generated-docx/                             # Các tệp Word (.docx) xuất bản (Bản Học sinh & Giáo viên)
│   └── archive-files/                              # Kho lưu trữ 63 tệp học liệu gốc (.docx & .pdf)
└── dist-gcs/
    ├── deploy_to_gcs.sh                            # Kịch bản triển khai lên Google Cloud Storage (Bash)
    ├── deploy_to_gcs.ps1                           # Kịch bản triển khai lên Google Cloud Storage (PowerShell)
    ├── control/index.html                          # Trang điều hành đã tích hợp điều hướng English Studio v2.0
    └── master-thuy-studio/                         # Gói tĩnh sẵn sàng xuất bản lên GCS Bucket
```

---

## 3. Danh Mục 20 Bộ Đề Chuẩn (953 Câu Hỏi Từ 63 Tệp Gốc)

| Mã Bộ Đề (`id`) | Khối | Tên Học Liệu Trên WebApp | Số câu |
| :--- | :---: | :--- | :---: |
| `WS_G06_U03_ADV` | Lớp 6 | Lớp 6 — Unit 3: My Friends (Nâng cao 50 câu) | 50 |
| `WS_G06_ERROR_48` | Lớp 6 | Lớp 6 — Sửa lỗi sai & Đoạn văn Proofreading (Units 1–3) | 48 |
| `WS_G06_SIGNS_50` | Lớp 6 | Lớp 6 — 50 Câu Biển Báo Thực Tế (Signs Around Us) | 50 |
| `WS_G06_PS_30` | Lớp 6 | Lớp 6 — Chuyên đề Thì Hiện Tại Đơn (Present Simple) | 30 |
| `WS_G06_PC_30` | Lớp 6 | Lớp 6 — Chuyên đề Thì Hiện Tại Tiếp Diễn (Present Continuous) | 30 |
| `WS_G06_THERE_30` | Lớp 6 | Lớp 6 — Cấu trúc There is / There are & Giới từ chỉ nơi chốn | 30 |
| `WS_G06_MID1_50` | Lớp 6 | Lớp 6 — Đề Kiểm Tra Giữa Kì I (Mục tiêu 8–10 điểm) | 50 |
| `WS_G07_MID1_50` | Lớp 7 | Lớp 7 — Đề Ôn Tập Giữa Kì I (50 câu Trắc nghiệm) | 50 |
| `WS_G07_ZALO_50` | Lớp 7 | Lớp 7 — Đề Tổng Hợp Giữa Kì I (Bản Final Zalo) | 50 |
| `WS_G07_PS_30` | Lớp 7 | Lớp 7 — Chuyên đề Hiện Tại Đơn & Trạng từ tần suất | 30 |
| `WS_G07_TARGET_45` | Lớp 7 | Lớp 7 — Đề Luyện Thi Mục Tiêu 8–10 Điểm | 45 |
| `WS_G08_MID1_50` | Lớp 8 | Lớp 8 — Đề Kiểm Tra Giữa Kì I (50 câu chuẩn) | 50 |
| `WS_G08_MID1_80` | Lớp 8 | Lớp 8 — Đại Đề Ôn Tập Giữa Kì I (80 câu Toàn diện) | 80 |
| `WS_G09_MID1_50` | Lớp 9 | Lớp 9 — Đề Kiểm Tra Giữa Kì I (50 câu Trọng tâm) | 50 |
| `WS_G09_MID1_80` | Lớp 9 | Lớp 9 — Đại Đề Ôn Tập Giữa Kì I (Units 1–3, 80 câu) | 80 |
| `WS_G09_WF_50` | Lớp 9 | Lớp 9 — Chuyên đề Cấu Tạo Từ (Word Formation 50 câu) | 50 |
| `WS_TOP_GET_HAVE` | Chuyên đề | Chuyên đề Đặc Biệt: Cấu trúc GET & HAVE (40 câu) | 40 |
| `WS_G11_101` | Lớp 11 | Lớp 11 — Đề Kiểm Tra Chuẩn THPT (Mã đề 101) | 40 |
| `WS_G11_102` | Lớp 11 | Lớp 11 — Đề Kiểm Tra Chuẩn THPT (Mã đề 102) | 40 |
| `WS_G12_101` | Lớp 12 | Lớp 12 — Đề Luyện Thi Tốt Nghiệp THPT (60 câu) | 60 |

---

## 4. Hướng Dẫn Vận Hành & Kiểm Định

### 4.1. Chạy Kiểm Định Chất Lượng 4 Cổng (QA Gates)
```bash
python3 engine-python/qa_gates/verify_questions.py
```

### 4.2. Khởi Động Máy Chủ Local Bridge & Trạm In Tự Động (Cổng `8765`)
```bash
python3 local-pc-bridge/server_and_print_bridge.py
```
Hoặc trên Windows PowerShell:
```powershell
powershell -ExecutionPolicy Bypass -File "local-pc-bridge\start_master_thuy_bridge.ps1"
```
- **Trang chủ Studio:** `http://localhost:8765/index.html`
- **Phòng làm việc của Cô Thúy:** `http://localhost:8765/teacher/index.html`
- **Góc học tập Học sinh:** `http://localhost:8765/student/index.html`
- **Trang điều hành Control:** `http://localhost:8765/control/index.html`

### 4.3. Xuất Bản Lên Google Cloud Storage (GCS Bucket)
```bash
bash dist-gcs/deploy_to_gcs.sh
```
