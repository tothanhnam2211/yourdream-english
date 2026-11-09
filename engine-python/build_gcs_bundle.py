# -*- coding: utf-8 -*-
"""
Đóng gói Gói Tĩnh GCS Bucket (dist-gcs/) & Kết nối trang điều hành /control/index.html
sang ../master-thuy-studio/teacher/index.html
"""
import shutil
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
PUBLIC_DIR = BASE_DIR / "public"
DIST_GCS_DIR = BASE_DIR / "dist-gcs"
CONTROL_HTML = DIST_GCS_DIR / "control" / "index.html"
TARGET_STUDIO_DIR = DIST_GCS_DIR / "master-thuy-studio"


def patch_control_html():
    txt = CONTROL_HTML.read_text(encoding="utf-8").replace("\r\n", "\n")

    if "../master-thuy-studio/teacher/index.html" not in txt:
        # 1. Header navigation buttons
        anchor_header = '<a href="https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/student/app.html"'
        assert anchor_header in txt, "anchor_header not found"
        nav_buttons = """<a href="../master-thuy-studio/teacher/index.html" target="_blank" style="text-decoration:none; background:linear-gradient(135deg,#0f766e,#0d9488); color:#fff; padding:6px 14px; border-radius:20px; font-size:12px; font-weight:800; box-shadow:0 0 14px rgba(15,118,110,0.6); border:1px solid #5eead4;">
        👩‍🏫 MỞ PHÒNG CÔ THÚY (ENGLISH STUDIO v2.0)
      </a>
      <a href="../master-thuy-studio/student/index.html" target="_blank" style="text-decoration:none; background:linear-gradient(135deg,#d97706,#b45309); color:#fff; padding:6px 14px; border-radius:20px; font-size:12px; font-weight:800; box-shadow:0 0 12px rgba(217,119,6,0.4);">
        🎒 GÓC HỌC SINH TIẾNG ANH
      </a>
      """
        txt = txt.replace(anchor_header, nav_buttons + anchor_header, 1)

        # 2. Filter button
        anchor_filter = "onclick=\"filterCategory('toan')\""
        assert anchor_filter in txt, "anchor_filter not found"
        eng_filter_btn = "onclick=\"filterCategory('english')\">🇬🇧 Master Thúy — Tiếng Anh (953 câu / 63 tệp)</button>\n    <button class=\"filter-btn\" onclick=\"filterCategory('toan')\""
        txt = txt.replace(anchor_filter, eng_filter_btn, 1)

        # 3. Items array card
        old_items = "let items = ["
        assert old_items in txt, "old_items not found"
        new_items = """let items = [
      {
        id: 11,
        title: "⭐ MASTER THÚY — English Studio v2.0 (953 Câu & 63 Tệp Gốc)",
        desc: "Phòng làm việc Giáo viên & Góc Học sinh Tiếng Anh Lớp 6–12: 20 bộ đề (953 câu chuẩn QA 100%), xuất Word 4 mã đề 101-104 ngay trên trình duyệt, đọc phát âm UK/US và In từ xa qua Google Drive.",
        stage: 4,
        cat: "english",
        type: "english_studio",
        agent: "master-thuy-english-studio-v2",
        thumb: "../master-thuy-studio/assets/signs/sign_01.png",
        path: "../master-thuy-studio/teacher/index.html",
        proofText: `=== BÁO CÁO NGHIỆM THU MASTER THÚY ENGLISH STUDIO v2.0 ===
[DATA ENGINE] 63/63 tệp gốc (37 tệp Drive + 26 tệp ZIP) -> 20 Bộ đề chuẩn -> 953 câu hỏi nguyên tử.
[QA GATES] 953/953 câu đạt 100.0% qua 4 cổng kiểm định tự động (verify_questions.py).
[FORMATTING] Bảo toàn 100% gạch chân <u>...</u> và in đậm <b>...</b> (337 câu Ngữ âm & Sửa lỗi + 3 bảng Proofreading).
[SIGNS VISUALS] 50/50 biển báo trực quan (.sign-visual-box) trích xuất nguyên bản từ Word.
[HYBRID CLOUD] Xuất Word .docx 4 mã đề (101-104) + Bảng đáp án Trang 1 + Bảng từ vựng gấp đôi IPA ngay trên trình duyệt (0 đồng server).
[REMOTE PRINT] Hàng đợi 00_PRINT_QUEUE đồng bộ Google Drive tự động in khi bật máy tính.`,
        visuals: []
      },"""
        txt = txt.replace(old_items, new_items, 1)

        # 4. Tag label & direct launch button on card
        anchor_tag = "item.cat === 'video' ? '🎬 VIDEO' : '⚡ HỆ THỐNG'"
        if anchor_tag in txt:
            txt = txt.replace(
                anchor_tag,
                "item.cat === 'video' ? '🎬 VIDEO' : (item.cat === 'english' ? '🇬🇧 TIẾNG ANH' : '⚡ HỆ THỐNG')",
                1,
            )

        anchor_btn = 'onclick="event.stopPropagation(); openInspection(${item.id})"'
        if anchor_btn in txt:
            txt = txt.replace(
                f'<button class="btn-soi-bai" {anchor_btn}>',
                '${item.type === \'english_studio\' ? `<a href="../master-thuy-studio/teacher/index.html" target="_blank" class="btn-soi-bai" style="text-decoration:none; background:#0f766e; color:#fff; border-color:#14b8a6; margin-right:6px;" onclick="event.stopPropagation();">🚀 Mở Studio</a>` : \'\'}<button class="btn-soi-bai" '
                + anchor_btn
                + ">",
                1,
            )

        # 5. Modal inspection branch
        old_modal_branch = "else if(item.type === 'slides_gallery') {"
        new_modal_branch = """else if(item.type === 'english_studio') {
        tabs.push({ id: 'tab-eng-links', label: '🇬🇧 ĐIỀU HƯỚNG MASTER THÚY STUDIO v2.0', content: `
          <div style="background: var(--bg-card); border: 1px solid var(--border); border-radius: 16px; padding: 22px;">
            <h3 style="color: #2dd4bf; font-size: 17px; margin-bottom: 8px;">👩‍🏫 MASTER THÚY — ENGLISH STUDIO v2.0 (PRODUCTION & CLOUD-READY)</h3>
            <p style="font-size: 13.5px; color: #cbd5e1; margin-bottom: 18px; line-height: 1.6;">
              Hệ thống quản lý học liệu, trộn 4 mã đề Word (.docx) kèm Bảng từ vựng gấp đôi IPA, phát âm chuẩn Anh-Anh / Anh-Mỹ (0 đồng), làm bài trực tuyến 2 pha và In từ xa qua Google Drive cho Cô Thúy (Zalo: 0836 686 874).
            </p>
            <div class="drive-links-grid">
              <a href="../master-thuy-studio/teacher/index.html" target="_blank" class="drive-btn-card" style="border-color:#14b8a6;">
                <span class="drive-icon">👩‍🏫</span>
                <div>
                  <b style="color: #2dd4bf; display: block; font-size: 14px;">Phòng Làm Việc Của Cô Thúy (Teacher Studio)</b>
                  <span style="font-size: 12px; color: #94a3b8;">../master-thuy-studio/teacher/index.html — Duyệt 20 bộ đề (953 câu), Trộn 4 mã đề Word & In từ xa</span>
                </div>
              </a>
              <a href="../master-thuy-studio/student/index.html" target="_blank" class="drive-btn-card" style="border-color:#f59e0b;">
                <span class="drive-icon">🎒</span>
                <div>
                  <b style="color: #fbbf24; display: block; font-size: 14px;">Góc Học Tập Học Sinh (Student Corner)</b>
                  <span style="font-size: 12px; color: #94a3b8;">../master-thuy-studio/student/index.html — Làm bài 2 pha, Nghe phát âm UK/US & Giải thích 7 nhóm bẫy</span>
                </div>
              </a>
              <a href="../master-thuy-studio/index.html" target="_blank" class="drive-btn-card" style="border-color:#38bdf8;">
                <span class="drive-icon">🏠</span>
                <div>
                  <b style="color: #38bdf8; display: block; font-size: 14px;">Cổng Chào Master Thúy Studio</b>
                  <span style="font-size: 12px; color: #94a3b8;">../master-thuy-studio/index.html — Trang chủ lựa chọn vai trò Giáo viên / Học sinh</span>
                </div>
              </a>
            </div>
          </div>
        `});
      }
      else if(item.type === 'slides_gallery') {"""
        assert old_modal_branch in txt, "old_modal_branch not found"
        txt = txt.replace(old_modal_branch, new_modal_branch, 1)

        CONTROL_HTML.write_text(txt, encoding="utf-8")
        print("[OK] Patched dist-gcs/control/index.html with ../master-thuy-studio/teacher/index.html navigation.")
    else:
        print("[OK] dist-gcs/control/index.html already patched.")

    pub_control_dir = PUBLIC_DIR / "control"
    pub_control_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(CONTROL_HTML, pub_control_dir / "index.html")
    print("[OK] Synced control/index.html to public/control/index.html")


def build_master_thuy_studio_bundle():
    if TARGET_STUDIO_DIR.exists():
        shutil.rmtree(TARGET_STUDIO_DIR)
    shutil.copytree(
        PUBLIC_DIR,
        TARGET_STUDIO_DIR,
        ignore=shutil.ignore_patterns("control"),
    )
    all_files = [p for p in TARGET_STUDIO_DIR.rglob("*") if p.is_file()]
    total_bytes = sum(p.stat().st_size for p in all_files)
    print(f"[OK] Built dist-gcs/master-thuy-studio/: {len(all_files)} files ({total_bytes / (1024*1024):.2f} MB)")


if __name__ == "__main__":
    patch_control_html()
    build_master_thuy_studio_bundle()
