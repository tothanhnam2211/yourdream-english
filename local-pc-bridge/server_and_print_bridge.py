# -*- coding: utf-8 -*-
"""
MASTER THÚY — ENGLISH STUDIO v2.0 (PRODUCTION & CLOUD-READY)
Unified WebApp Server, SQLite API & Google Drive 00_PRINT_QUEUE Auto-Poller Bridge
Port: 8765
"""

import json
import os
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import base64
_WIN_ROOT = Path(r"D:\YOURDREAM.VN\master-thuy-english-studio")
PROJECT_ROOT = _WIN_ROOT if _WIN_ROOT.exists() else Path(__file__).resolve().parent.parent
PUBLIC_DIR = PROJECT_ROOT / "public"
DB_PATH = PROJECT_ROOT / "db" / "master_thuy_studio.db"
ENGINE_DIR = PROJECT_ROOT / "engine-python"
INBOX_LOCAL_DIR = PROJECT_ROOT / "00_LENH_DIEU_KHIEN_INBOX"
sys.path.insert(0, str(ENGINE_DIR))

from generate_dual_worksheet import generate_worksheet_from_db

DRIVE_STUDIO_ROOT = Path(r"G:\Drive của tôi\YOURDREAM_AI_WORKSPACE\ENGLISH_TEACHING_STUDIO")
DRIVE_PRINT_QUEUE_DIR = DRIVE_STUDIO_ROOT / "00_PRINT_QUEUE"
DRIVE_PRINT_PROCESSED_DIR = DRIVE_PRINT_QUEUE_DIR / "PROCESSED"


def ensure_drive_print_queue_dirs():
    try:
        if DRIVE_STUDIO_ROOT.exists():
            DRIVE_PRINT_QUEUE_DIR.mkdir(parents=True, exist_ok=True)
            DRIVE_PRINT_PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
            return True
    except Exception:
        pass
    local_fallback = PROJECT_ROOT / "00_PRINT_QUEUE"
    (local_fallback / "PROCESSED").mkdir(parents=True, exist_ok=True)
    return False


def process_single_print_ticket(ticket_data: dict, ticket_file_path: Path = None):
    job_id = str(ticket_data.get("job_id", f"PJ_{datetime.now().strftime('%Y%m%d_%H%M%S')}"))
    ws_id = str(ticket_data.get("worksheet_id", "WS_G06_U03_ADV"))
    ws_title = str(ticket_data.get("worksheet_title", "Phiếu bài tập Cô Thúy"))
    version = str(ticket_data.get("target_version", "BOTH"))
    copies = int(ticket_data.get("copies", 25))
    duplex = int(ticket_data.get("duplex", 1))
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Generate A4 Word (.docx) files for the print job
    grade_match = 6
    for g in (6, 7, 8, 9, 10, 11, 12):
        if f"G0{g}" in ws_id or f"G{g}" in ws_id or f"Lớp {g}" in ws_title:
            grade_match = g
            break

    gen_res = generate_worksheet_from_db(
        grade=grade_match,
        unit_code=f"G0{grade_match}_U01",
        count=15,
        title=ws_title,
        exam_code="101"
    )

    # Copy generated DOCX files + ticket to 00_PRINT_QUEUE/PROCESSED on Drive
    ensure_drive_print_queue_dirs()
    target_processed_dir = DRIVE_PRINT_PROCESSED_DIR if DRIVE_STUDIO_ROOT.exists() else (PROJECT_ROOT / "00_PRINT_QUEUE" / "PROCESSED")
    try:
        s_docx = Path(gen_res["student_docx_path"])
        t_docx = Path(gen_res["teacher_docx_path"])
        if s_docx.exists():
            shutil.copy2(s_docx, target_processed_dir / f"{job_id}_{s_docx.name}")
        if t_docx.exists():
            shutil.copy2(t_docx, target_processed_dir / f"{job_id}_{t_docx.name}")

        if os.environ.get("ENABLE_PHYSICAL_PRINTER") == "1" and s_docx.exists():
            subprocess.Popen(["powershell", "-Command", f'Start-Process -FilePath "{s_docx}" -Verb Print'])
    except Exception as e:
        print("Print queue copy warning:", e)

    if ticket_file_path and ticket_file_path.exists():
        try:
            shutil.move(str(ticket_file_path), str(target_processed_dir / ticket_file_path.name))
        except Exception:
            pass

    status_str = "ĐÃ IN & LƯU 00_PRINT_QUEUE/PROCESSED"
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """INSERT OR REPLACE INTO print_jobs
           (job_id, worksheet_id, worksheet_title, target_version, copies, paper_size, duplex, status, requested_from, created_at, completed_at)
           VALUES (?, ?, ?, ?, ?, 'A4', ?, ?, '00_PRINT_QUEUE (Google Drive)', ?, ?)""",
        (job_id, ws_id, ws_title, version, copies, duplex, status_str, now_str, now_str)
    )
    conn.commit()
    conn.close()
    return {"ok": True, "job_id": job_id, "status": status_str, "student_docx": gen_res["student_docx_url"]}


def poll_drive_print_queue_loop():
    """Background daemon that scans 00_PRINT_QUEUE on Google Drive every 5 seconds."""
    ensure_drive_print_queue_dirs()
    while True:
        try:
            queue_dir = DRIVE_PRINT_QUEUE_DIR if DRIVE_STUDIO_ROOT.exists() else (PROJECT_ROOT / "00_PRINT_QUEUE")
            if queue_dir.exists():
                for item in queue_dir.glob("*.json"):
                    if item.is_file():
                        try:
                            data = json.loads(item.read_text(encoding="utf-8"))
                            print(f"[00_PRINT_QUEUE] Detected cloud print ticket: {item.name}")
                            process_single_print_ticket(data, ticket_file_path=item)
                        except Exception as e:
                            print(f"[00_PRINT_QUEUE] Error processing {item.name}: {e}")
        except Exception:
            pass
        time.sleep(5)


class MasterThuyStudioHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(PUBLIC_DIR), **kwargs)

    def translate_path(self, path):
        if path.startswith("/master-thuy-studio/"):
            path = "/" + path[len("/master-thuy-studio/"):]
        elif path == "/master-thuy-studio":
            path = "/index.html"
        return super().translate_path(path)

    def _send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path == "/api/print-job":
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute("SELECT * FROM print_jobs ORDER BY created_at DESC LIMIT 30")
            jobs = [dict(r) for r in cur.fetchall()]
            conn.close()
            self._send_json({"ok": True, "jobs": jobs})
            return

        if self.path == "/api/students":
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute("SELECT * FROM students ORDER BY grade ASC, student_code ASC")
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self._send_json({"ok": True, "students": rows})
            return

        if self.path == "/api/inbox-request":
            INBOX_LOCAL_DIR.mkdir(parents=True, exist_ok=True)
            items = []
            for p in sorted(INBOX_LOCAL_DIR.glob("*.json"), reverse=True)[:30]:
                try:
                    items.append(json.loads(p.read_text(encoding="utf-8")))
                except Exception:
                    pass
            self._send_json({"ok": True, "requests": items})
            return

        if self.path == "/api/submissions":
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute("SELECT * FROM submissions ORDER BY submitted_at DESC LIMIT 50")
            rows = []
            for r in cur.fetchall():
                d = dict(r)
                d["wrong_traps"] = json.loads(d.get("wrong_traps_json") or "[]")
                rows.append(d)
            conn.close()
            self._send_json({"ok": True, "submissions": rows})
            return

        return super().do_GET()

    def do_POST(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
        data = json.loads(raw_body or "{}")

        if self.path == "/api/generate-docx":
            grade = int(data.get("grade", 6))
            unit_code = str(data.get("unit_code", f"G0{grade}_U01"))
            count = int(data.get("count", 15))
            title = str(data.get("title", f"Phiếu Luyện Tập Khối {grade}"))
            exam_code = str(data.get("exam_code", "101"))

            result = generate_worksheet_from_db(
                grade=grade,
                unit_code=unit_code,
                count=count,
                title=title,
                exam_code=exam_code
            )
            self._send_json({"ok": True, "mode": "LOCAL_PYTHON_BRIDGE", **result})
            return

        if self.path == "/api/print-job":
            # Also write ticket into 00_PRINT_QUEUE on Drive and process immediately
            ensure_drive_print_queue_dirs()
            queue_dir = DRIVE_PRINT_QUEUE_DIR if DRIVE_STUDIO_ROOT.exists() else (PROJECT_ROOT / "00_PRINT_QUEUE")
            job_id = str(data.get("job_id", f"PJ_{datetime.now().strftime('%Y%m%d_%H%M%S')}"))
            data["job_id"] = job_id
            ticket_path = queue_dir / f"{job_id}_TICKET.json"
            ticket_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

            res = process_single_print_ticket(data, ticket_file_path=ticket_path)
            self._send_json(res)
            return

        if self.path == "/api/students":
            student_code = str(data.get("student_code", "")).strip().upper()
            full_name = str(data.get("full_name", "Học sinh mới")).strip()
            grade = int(data.get("grade", 6))
            class_id = str(data.get("class_id", f"CLS_G0{grade}A"))
            pin_code = str(data.get("pin_code", "1234")).strip() or "1234"
            if not student_code:
                student_code = f"MT{grade}{int(time.time()) % 100:02d}"
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute(
                "INSERT OR IGNORE INTO classes (class_id, class_name, grade) VALUES (?, ?, ?)",
                (class_id, data.get("class_name") or f"Lớp {grade}A", grade)
            )
            cur.execute(
                "INSERT OR REPLACE INTO students (student_code, full_name, class_id, grade, pin_code, streak_days, xp_points) VALUES (?, ?, ?, ?, ?, 1, 100)",
                (student_code, full_name, class_id, grade, pin_code)
            )
            conn.commit()
            conn.close()
            self._send_json({"ok": True, "student_code": student_code, "full_name": full_name, "grade": grade})
            return

        if self.path == "/api/inbox-request":
            INBOX_LOCAL_DIR.mkdir(parents=True, exist_ok=True)
            ticket_id = str(data.get("ticket_id", f"INBOX_REQ_{int(time.time())}"))
            saved_files = []
            for idx, att in enumerate(data.get("attachments") or []):
                fname = Path(att.get("name") or f"photo_{idx+1}.jpg").name
                durl = str(att.get("data_url") or "")
                if "," in durl:
                    try:
                        raw_bytes = base64.b64decode(durl.split(",", 1)[1])
                        out_f = INBOX_LOCAL_DIR / f"{ticket_id}_{fname}"
                        out_f.write_bytes(raw_bytes)
                        saved_files.append(out_f.name)
                    except Exception:
                        pass
            clean_ticket = {k: v for k, v in data.items() if k != "attachments"}
            clean_ticket["saved_attachment_files"] = saved_files
            OUTBOX_LOCAL_DIR = PROJECT_ROOT / "00_BAO_CAO_KET_QUA_OUTBOX"
            OUTBOX_LOCAL_DIR.mkdir(parents=True, exist_ok=True)
            grade_num = int(data.get("grade") or 6) if str(data.get("grade") or "").isdigit() else 6
            gen_res = generate_worksheet_from_db(
                grade=grade_num,
                unit_code=f"G0{grade_num}_U01",
                count=15,
                title=str(data.get("request_prompt") or f"Phiếu Bài Tập Khối {grade_num}")[:80],
                exam_code="101"
            )
            try:
                s_docx = Path(gen_res["student_docx_path"])
                t_docx = Path(gen_res["teacher_docx_path"])
                if s_docx.exists():
                    shutil.copy2(s_docx, OUTBOX_LOCAL_DIR / f"{ticket_id}_{s_docx.name}")
                if t_docx.exists():
                    shutil.copy2(t_docx, OUTBOX_LOCAL_DIR / f"{ticket_id}_{t_docx.name}")
            except Exception:
                pass
            # Auto-publish OUTBOX worksheet into WebApp Catalog & Shards so Teacher & Students see it immediately
            new_ws_id = str(gen_res.get("ws_id") or f"WS_OUTBOX_G0{grade_num}_{int(time.time()) % 100000}")
            ws_title_clean = str(data.get("request_prompt") or f"Phiếu Bài Tập Khối {grade_num}")[:90]
            conn_cat = sqlite3.connect(DB_PATH)
            conn_cat.row_factory = sqlite3.Row
            cur_cat = conn_cat.cursor()
            cur_cat.execute(
                "SELECT * FROM question_bank WHERE grade=? ORDER BY RANDOM() LIMIT 15",
                (grade_num,)
            )
            q_rows = [dict(r) for r in cur_cat.fetchall()]
            shard_questions = []
            for q_r in q_rows:
                shard_questions.append({
                    "question_id": q_r.get("question_id"),
                    "grade": q_r.get("grade"),
                    "unit_id": q_r.get("unit_id"),
                    "skill_domain": q_r.get("skill_domain"),
                    "difficulty_level": q_r.get("difficulty_level", 2),
                    "trap_code": q_r.get("trap_code") or "TRAP_TENSE_SIGNAL",
                    "stem_text": q_r.get("stem_text"),
                    "passage_or_sign_svg": q_r.get("passage_or_sign_svg") or "",
                    "options": json.loads(q_r.get("options_json") or "[]"),
                    "correct_index": int(q_r.get("correct_index") or 0),
                    "answer_mode": q_r.get("answer_mode") or "MCQ",
                    "written_answers": json.loads(q_r.get("written_answers_json") or "[]"),
                    "explanation_vi": q_r.get("explanation_vi") or ""
                })
            shard_rel = f"shards/items/{new_ws_id}.json"
            shard_file = PUBLIC_DIR / shard_rel
            shard_file.parent.mkdir(parents=True, exist_ok=True)
            shard_file.write_text(json.dumps({
                "worksheet_id": new_ws_id,
                "title": ws_title_clean,
                "grade": grade_num,
                "question_count": len(shard_questions),
                "questions": shard_questions
            }, ensure_ascii=False, indent=2), encoding="utf-8")

            new_cat_entry = {
                "worksheet_id": new_ws_id,
                "title": ws_title_clean,
                "grade": grade_num,
                "unit_scope": f"Đề mới từ OUTBOX ({ticket_id})",
                "category": "OUTBOX_AUTO_PUBLISHED",
                "question_count": len(shard_questions) or 15,
                "duration_minutes": 45,
                "local_student_docx": gen_res.get("student_docx_url"),
                "local_teacher_docx": gen_res.get("teacher_docx_url"),
                "shard_json_path": f"/{shard_rel}",
                "approval_state": "✨ Mới đồng bộ từ OUTBOX",
                "source_group": "00_BAO_CAO_KET_QUA_OUTBOX",
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            cat_file = PUBLIC_DIR / "shards/worksheets-catalog.json"
            if cat_file.exists():
                try:
                    cat_json = json.loads(cat_file.read_text(encoding="utf-8"))
                    curated = cat_json.get("curated_sets") or []
                    curated = [w for w in curated if w.get("worksheet_id") != new_ws_id]
                    curated.insert(0, new_cat_entry)
                    cat_json["curated_sets"] = curated
                    cat_file.write_text(json.dumps(cat_json, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception:
                    pass
            conn_cat.close()

            clean_ticket["status"] = "PUBLISHED_FROM_OUTBOX_TO_WEBAPP"
            clean_ticket["outbox_result"] = {
                "worksheet_id": new_ws_id,
                "title": ws_title_clean,
                "grade": grade_num,
                "question_count": len(shard_questions) or 15,
                "student_docx_url": gen_res.get("student_docx_url"),
                "teacher_docx_url": gen_res.get("teacher_docx_url"),
                "catalog_entry": new_cat_entry,
                "questions": shard_questions
            }
            ticket_path = INBOX_LOCAL_DIR / f"{ticket_id}.json"
            ticket_path.write_text(json.dumps(clean_ticket, ensure_ascii=False, indent=2), encoding="utf-8")
            self._send_json({
                "ok": True,
                "ticket_id": ticket_id,
                "saved_path": str(OUTBOX_LOCAL_DIR),
                "saved_attachments": saved_files,
                "outbox_result": clean_ticket["outbox_result"]
            })
            return

        if self.path in ("/api/submit", "/api/submissions"):
            sub_id = str(data.get("submission_id", f"SUB_{int(datetime.now().timestamp())}"))
            ws_id = str(data.get("worksheet_id", "WS_G06_U03_ADV"))
            student_code = str(data.get("student_code", "MT601"))
            student_name = str(data.get("student_name", "Nguyễn Minh Anh"))
            class_id = str(data.get("class_id", "CLS_G06A"))
            exam_code = str(data.get("exam_code", "101"))
            auto_score = float(data.get("auto_score", 9.0))
            final_score = float(data.get("final_score", auto_score))
            teacher_approved = int(data.get("teacher_approved", 0))
            wrong_traps = json.dumps(data.get("wrong_traps", []), ensure_ascii=False)
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute(
                """INSERT OR REPLACE INTO submissions
                   (submission_id, assignment_id, worksheet_id, student_code, student_name, class_id, exam_code, answers_json, auto_score, final_score, teacher_approved, wrong_traps_json, submitted_at)
                   VALUES (?, 'ASSIGN_DEFAULT', ?, ?, ?, ?, ?, '{}', ?, ?, ?, ?, ?)""",
                (sub_id, ws_id, student_code, student_name, class_id, exam_code, auto_score, final_score, teacher_approved, wrong_traps, now_str)
            )
            conn.commit()
            conn.close()
            self._send_json({"ok": True, "submission_id": sub_id})
            return

        self._send_json({"ok": False, "error": "Unknown endpoint"}, status=404)


def run_server(port=8765):
    poller = threading.Thread(target=poll_drive_print_queue_loop, daemon=True)
    poller.start()

    server = ThreadingHTTPServer(("0.0.0.0", port), MasterThuyStudioHandler)
    print(f"======================================================================")
    print(f" MASTER THUY — ENGLISH STUDIO v2.0 & 00_PRINT_QUEUE BRIDGE IS RUNNING!")
    print(f" - Root Dir:             {PROJECT_ROOT}")
    print(f" - Drive Print Queue:    {DRIVE_PRINT_QUEUE_DIR}")
    print(f" - Trang chu:            http://localhost:{port}/index.html")
    print(f" - Phong cua Co Thuy:    http://localhost:{port}/teacher/index.html")
    print(f" - Goc Hoc Sinh Online:  http://localhost:{port}/student/index.html")
    print(f"======================================================================")
    server.serve_forever()


if __name__ == "__main__":
    run_server(8765)
