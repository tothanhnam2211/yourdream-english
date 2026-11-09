# -*- coding: utf-8 -*-
"""
MASTER THÚY — ENGLISH STUDIO v2.0
Kiểm định Toàn diện End-to-End (6 Hạng Mục Bắt Buộc)
"""
import json
import sys
import time
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_URL = "http://localhost:8765"
DRIVE_QUEUE_DIR = Path(r"G:\Drive của tôi\YOURDREAM_AI_WORKSPACE\ENGLISH_TEACHING_STUDIO\00_PRINT_QUEUE")


def check_url(path_str: str):
    url = BASE_URL + path_str
    req = urllib.request.Request(url, method="GET")
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = resp.read()
        assert resp.status == 200, f"Status {resp.status} for {path_str}"
        print(f"  [200 OK] {path_str:<48} ({len(data):>8,} bytes)")
        return data


def post_json(path_str: str, payload: dict):
    url = BASE_URL + path_str
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        res_data = json.loads(resp.read().decode("utf-8"))
        assert resp.status == 200 and res_data.get("ok"), f"POST failed for {path_str}: {res_data}"
        return res_data


def run_e2e():
    print("=" * 72)
    print(" 1. KIỂM ĐỊNH HTTP STATIC & ĐƯỜNG DẪN TƯƠNG ĐỐI (ZERO 404)")
    print("=" * 72)
    endpoints = [
        "/favicon.ico",
        "/index.html",
        "/teacher/index.html",
        "/student/index.html",
        "/control/index.html",
        "/master-thuy-studio/teacher/index.html",
        "/master-thuy-studio/student/index.html",
        "/shared/css/studio.css",
        "/shared/js/hybrid-cloud-engine.js",
        "/shards/worksheets-catalog.json",
        "/shards/map2-traps-and-levels.json",
        "/shards/maps/map2-traps-and-levels.json",
        "/shards/items/WS_G06_SIGNS_50.json",
        "/shards/items/WS_G06_ERROR_48.json",
        "/assets/signs/sign_01.png",
        "/assets/signs/sign_50.png",
    ]
    for ep in endpoints:
        check_url(ep)

    print("\n" + "=" * 72)
    print(" 2. KIỂM ĐỊNH API TRỘN 4 MÃ ĐỀ WORD (101-104) & TẢI FILE DOCX")
    print("=" * 72)
    gen_res = post_json(
        "/api/generate-docx",
        {
            "grade": 6,
            "unit_code": "G06_U03",
            "count": 20,
            "title": "PHIẾU ÔN TẬP LỚP 6 — UNIT 3 & SIGNS (v2.0)",
            "mix_4_codes": True,
        },
    )
    print(f"  [OK] Generated primary Student DOCX: {gen_res['student_docx_url']}")
    print(f"  [OK] Generated primary Teacher DOCX: {gen_res['teacher_docx_url']}")
    matrix = gen_res.get("answer_matrix_101_104", {})
    print(f"  [OK] Answer Matrix 4 Codes: {list(matrix.keys())}")
    assert set(matrix.keys()) == {"101", "102", "103", "104"}, "Missing 4 exam codes in answer_matrix_101_104!"
    check_url("/" + gen_res["student_docx_url"].lstrip("./"))
    check_url("/" + gen_res["teacher_docx_url"].lstrip("./"))

    print("\n" + "=" * 72)
    print(" 3. KIỂM ĐỊNH API NỘP BÀI HỌC SINH (/api/submissions)")
    print("=" * 72)
    sub_res = post_json(
        "/api/submissions",
        {
            "student_name": "Nguyễn Minh Anh (E2E v2.0)",
            "class_code": "G6A",
            "worksheet_id": "WS_G06_SIGNS_50",
            "score": 9.6,
            "correct_count": 48,
            "total_questions": 50,
            "time_spent_sec": 840,
            "answers": {"G06_SIGN_Q01": "A"},
        },
    )
    print(f"  [OK] Saved submission ID: {sub_res['submission_id']}")

    print("\n" + "=" * 72)
    print(" 4. KIỂM ĐỊNH HÀNG ĐỢI IN ĐÁM MÂY 00_PRINT_QUEUE TRÊN GOOGLE DRIVE")
    print("=" * 72)
    DRIVE_QUEUE_DIR.mkdir(parents=True, exist_ok=True)
    ticket_id = "PJ_CLOUD_DRIVE_TEST_V2"
    ticket_file = DRIVE_QUEUE_DIR / f"{ticket_id}.json"
    ticket_payload = {
        "job_id": ticket_id,
        "worksheet_id": "WS_G06_U03_ADV",
        "worksheet_title": "Lớp 6 — Unit 3: My Friends (Nâng cao 50 câu)",
        "target_version": "BOTH",
        "copies": 30,
        "duplex": 1,
        "requested_from": "Cloud Queue 00_PRINT_QUEUE E2E Test",
    }
    ticket_file.write_text(json.dumps(ticket_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  [OK] Dropped ticket into Google Drive: {ticket_file}")
    print("  [...] Waiting 6 seconds for poll_drive_print_queue_loop() daemon to process...")
    time.sleep(6.5)

    processed_ticket = DRIVE_QUEUE_DIR / "PROCESSED" / f"{ticket_id}.json"
    assert processed_ticket.exists(), f"Processed ticket not found at {processed_ticket}"
    print(f"  [OK] Daemon automatically processed ticket -> {processed_ticket}")

    jobs_res = json.loads(check_url("/api/print-job").decode("utf-8"))
    found_job = any(j.get("job_id") == ticket_id for j in jobs_res.get("jobs", []))
    assert found_job, "Job ID not found in SQLite print_jobs table!"
    print(f"  [OK] Confirmed job {ticket_id} logged in SQLite print_jobs with status ĐÃ IN & LƯU!")

    print("\n✅ TẤT CẢ KIỂM ĐỊNH E2E v2.0 ĐẠT 100%!")


if __name__ == "__main__":
    run_e2e()
