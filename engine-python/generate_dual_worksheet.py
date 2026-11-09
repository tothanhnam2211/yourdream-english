# -*- coding: utf-8 -*-
"""
MASTER THÚY — ENGLISH STUDIO v2.0
Dual-Output Worksheet & 4-Code Exam Mixer Engine (Map 4)
Preserves 100% <u>...</u> and <b>...</b> rich runs in generated Word (.docx) files!
"""

import json
import random
import re
import sqlite3
from pathlib import Path
from datetime import datetime
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

PROJECT_ROOT = Path(r"D:\YOURDREAM.VN\master-thuy-english-studio")
DB_PATH = PROJECT_ROOT / "db" / "master_thuy_studio.db"
OUTPUT_DOCX_DIR = PROJECT_ROOT / "public" / "generated-docx"


def set_Font(run, name="Times New Roman", size=11, bold=False, italic=False, underline=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_rich_html_runs(p, html_text: str, size=11, default_bold=False, default_italic=False):
    """Parses <u>...</u>, <b>...</b>, and <br/> tags and adds properly styled Word runs."""
    if not html_text:
        return
    text = html_text.replace("<br/>", "\n").replace("<br>", "\n")
    tokens = re.split(r"(</?[ub]>)", text, flags=re.I)
    is_b = default_bold
    is_u = False
    for tok in tokens:
        low = tok.lower()
        if low == "<b>":
            is_b = True
        elif low == "</b>":
            is_b = default_bold
        elif low == "<u>":
            is_u = True
        elif low == "</u>":
            is_u = False
        elif tok:
            clean_tok = re.sub(r"<[^>]+>", "", tok)
            if clean_tok:
                r = p.add_run(clean_tok)
                set_Font(r, size=size, bold=is_b, italic=default_italic, underline=is_u)


def add_master_thuy_header(doc, title_text, grade, exam_code="101", is_teacher=False):
    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    cell_tl = table.cell(0, 0)
    p_tl = cell_tl.paragraphs[0]
    r1 = p_tl.add_run("MASTER THÚY\n")
    set_Font(r1, size=12, bold=True)
    r2 = p_tl.add_run("Thạc sĩ · Giảng viên Tiếng Anh\nZalo / ĐT: 0836 686 874")
    set_Font(r2, size=10, italic=True)

    cell_tr = table.cell(0, 1)
    p_tr = cell_tr.paragraphs[0]
    p_tr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    role_badge = " [BẢN GIÁO VIÊN - ĐÁP ÁN]" if is_teacher else ""
    r_title = p_tr.add_run(f"{title_text.upper()}{role_badge}\n")
    set_Font(r_title, size=11, bold=True)
    r_code = p_tr.add_run(f"Khối {grade} · MÃ ĐỀ: {exam_code}")
    set_Font(r_code, size=11, bold=True)

    cell_bl = table.cell(1, 0)
    p_bl = cell_bl.paragraphs[0]
    if is_teacher:
        r_bl = p_bl.add_run("Tài liệu lưu hành nội bộ — Đã kiểm định 4 cổng sư phạm")
        set_Font(r_bl, size=10, italic=True)
    else:
        r_bl = p_bl.add_run("Họ và tên học sinh: ....................................................   Lớp: ..............")
        set_Font(r_bl, size=10)

    cell_br = table.cell(1, 1)
    p_br = cell_br.paragraphs[0]
    p_br.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if is_teacher:
        r_br = p_br.add_run(f"Ngày xuất bản: {datetime.now().strftime('%d/%m/%Y')}")
        set_Font(r_br, size=10)
    else:
        r_br = p_br.add_run("Điểm: ............... / 10   Nhận xét: ...............................")
        set_Font(r_br, size=10, bold=True)

    doc.add_paragraph()


def normalize_options_dict(raw_options):
    letters = ["A", "B", "C", "D"]
    if isinstance(raw_options, dict):
        return {k: str(v) for k, v in raw_options.items()}
    if isinstance(raw_options, list):
        out = {}
        for idx, val in enumerate(raw_options[:4]):
            s = str(val).strip()
            for let in letters:
                if s.startswith(f"{let}.") or s.startswith(f"{let})"):
                    s = s[2:].strip()
                    break
            out[letters[idx]] = s
        return out
    return {}


def mix_exam_codes(questions, codes=("101", "102", "103", "104"), seed_base=2026):
    mixed = {}
    letters = ["A", "B", "C", "D"]
    for idx, code in enumerate(codes):
        rng = random.Random(seed_base + idx * 101)
        q_copy = [dict(q) for q in questions]
        if idx > 0:
            rng.shuffle(q_copy)

        code_questions = []
        for new_num, q in enumerate(q_copy, start=1):
            opts = dict(q.get("options", {}))
            orig_ans = q.get("answer", "A")
            if len(opts) == 4 and idx > 0:
                items = list(opts.items())
                rng.shuffle(items)
                new_opts = {}
                new_ans = orig_ans
                for new_let, (old_let, txt) in zip(letters, items):
                    new_opts[new_let] = txt
                    if old_let == orig_ans:
                        new_ans = new_let
                opts = new_opts
                ans = new_ans
            else:
                ans = orig_ans

            code_questions.append({
                **q,
                "display_num": new_num,
                "options": opts,
                "answer": ans,
                "exam_code": code
            })
        mixed[code] = code_questions
    return mixed


def build_dual_docx(ws_id, title, grade, questions, vocab_foldable=None, exam_code="101"):
    OUTPUT_DOCX_DIR.mkdir(parents=True, exist_ok=True)
    mixed_variants = mix_exam_codes(questions, codes=("101", "102", "103", "104"))
    selected_questions = mixed_variants.get(str(exam_code), mixed_variants["101"])

    # 1. STUDENT DOCX
    student_doc = docx.Document()
    for section in student_doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    add_master_thuy_header(student_doc, title, grade, exam_code=exam_code, is_teacher=False)

    if vocab_foldable:
        p_voc_head = student_doc.add_paragraph()
        r_vh = p_voc_head.add_run("I. BẢNG TỪ VỰNG GẤP ĐÔI TỰ KIỂM TRA (Gấp mép giấy dọc đường kẻ đứt để tự dò nghĩa):")
        set_Font(r_vh, size=11, bold=True)
        v_table = student_doc.add_table(rows=len(vocab_foldable) + 1, cols=2)
        v_table.cell(0, 0).paragraphs[0].add_run("Từ vựng & Phiên âm IPA").bold = True
        v_table.cell(0, 1).paragraphs[0].add_run("┆  Nghĩa Tiếng Việt (Gấp giấy che cột này)").bold = True
        for row_idx, v in enumerate(vocab_foldable, start=1):
            r_l = v_table.cell(row_idx, 0).paragraphs[0].add_run(
                f"{row_idx}. {v.get('word', '')} {v.get('ipa', '')} ({v.get('pos', '')})"
            )
            set_Font(r_l, size=10, bold=True)
            r_r = v_table.cell(row_idx, 1).paragraphs[0].add_run(f"┆  {v.get('meaning_vi', '')}")
            set_Font(r_r, size=10)
        student_doc.add_paragraph()

    p_q_head = student_doc.add_paragraph()
    r_qh = p_q_head.add_run("PHẦN BÀI TẬP TRẮC NGHIỆM & VẬN DỤNG (Khoanh tròn chữ cái A, B, C hoặc D):")
    set_Font(r_qh, size=11, bold=True)

    for q in selected_questions:
        p = student_doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r_num = p.add_run(f"Question {q['display_num']}: ")
        set_Font(r_num, size=11, bold=True)
        add_rich_html_runs(p, q["stem"], size=11)

        opts = q.get("options", {})
        if opts:
            max_len = max((len(re.sub(r"<[^>]+>", "", str(v))) for v in opts.values()), default=0)
            p_opt = student_doc.add_paragraph()
            p_opt.paragraph_format.left_indent = Inches(0.25)
            p_opt.paragraph_format.space_after = Pt(4)
            sep = "      " if max_len <= 22 else "\n"
            for o_i, (k, v) in enumerate(opts.items()):
                if o_i > 0:
                    r_sep = p_opt.add_run(sep)
                    set_Font(r_sep, size=11)
                r_k = p_opt.add_run(f"{k}. ")
                set_Font(r_k, size=11, bold=True)
                add_rich_html_runs(p_opt, str(v), size=11)

    student_filename = f"{ws_id}_Code{exam_code}_STUDENT.docx"
    student_path = OUTPUT_DOCX_DIR / student_filename
    student_doc.save(str(student_path))

    # 2. TEACHER DOCX
    teacher_doc = docx.Document()
    for section in teacher_doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    add_master_thuy_header(teacher_doc, title, grade, exam_code="101-104", is_teacher=True)

    p_grid_title = teacher_doc.add_paragraph()
    r_gt = p_grid_title.add_run("BẢNG ĐÁP ÁN CHẤM NHANH TRANG 1 (ĐỐI CHIẾU 4 MÃ ĐỀ 101 · 102 · 103 · 104):")
    set_Font(r_gt, size=11, bold=True)

    total_q = len(selected_questions)
    grid_table = teacher_doc.add_table(rows=total_q + 1, cols=5)
    headers = ["Câu", "Mã 101", "Mã 102", "Mã 103", "Mã 104"]
    for c_idx, h_text in enumerate(headers):
        r_h = grid_table.cell(0, c_idx).paragraphs[0].add_run(h_text)
        set_Font(r_h, size=10, bold=True)

    for q_idx in range(total_q):
        r_c0 = grid_table.cell(q_idx + 1, 0).paragraphs[0].add_run(f"Câu {q_idx + 1}")
        set_Font(r_c0, size=10, bold=True)
        for c_i, code_str in enumerate(("101", "102", "103", "104"), start=1):
            ans_letter = mixed_variants[code_str][q_idx]["answer"]
            r_ans = grid_table.cell(q_idx + 1, c_i).paragraphs[0].add_run(ans_letter)
            set_Font(r_ans, size=10, bold=True)

    teacher_doc.add_paragraph()
    p_exp_title = teacher_doc.add_paragraph()
    r_et = p_exp_title.add_run("CHI TIẾT GIẢI THÍCH SƯ PHẠM & CẢNH BÁO BẪY (THEO MÃ GỐC 101):")
    set_Font(r_et, size=11, bold=True)

    for q in mixed_variants["101"]:
        p = teacher_doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        r_num = p.add_run(f"Câu {q['display_num']} [Đáp án: {q['answer']}] ({q.get('trap_id', 'CORE')}): ")
        set_Font(r_num, size=10.5, bold=True)
        add_rich_html_runs(p, f"{q['stem']}\n", size=10.5)
        add_rich_html_runs(p, f"   → Giải thích: {q.get('explanation_vi', '')}", size=10, default_italic=True)

    teacher_filename = f"{ws_id}_AllCodes_TEACHER.docx"
    teacher_path = OUTPUT_DOCX_DIR / teacher_filename
    teacher_doc.save(str(teacher_path))

    answer_matrix = {
        code_str: [
            {"num": item["display_num"], "answer": item["answer"], "q_id": item["q_id"]}
            for item in mixed_variants[code_str]
        ]
        for code_str in ("101", "102", "103", "104")
    }

    return {
        "ws_id": ws_id,
        "title": title,
        "grade": grade,
        "exam_code": exam_code,
        "question_count": len(selected_questions),
        "student_docx_url": f"../generated-docx/{student_filename}",
        "teacher_docx_url": f"../generated-docx/{teacher_filename}",
        "student_docx_path": str(student_path),
        "teacher_docx_path": str(teacher_path),
        "answer_matrix_101_104": answer_matrix,
        "questions_preview": selected_questions[:20],
    }


def generate_worksheet_from_db(grade=6, unit_code="G06_U03", count=15, title=None, exam_code="101"):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute(
        "SELECT * FROM question_bank WHERE grade = ? ORDER BY RANDOM() LIMIT ?",
        (int(grade), int(count))
    )
    rows = cur.fetchall()
    if not rows:
        cur.execute("SELECT * FROM question_bank ORDER BY RANDOM() LIMIT ?", (int(count),))
        rows = cur.fetchall()

    letters = ["A", "B", "C", "D"]
    questions = []
    for r in rows:
        raw_opts = json.loads(r["options_json"] or "[]")
        opts_dict = normalize_options_dict(raw_opts)
        c_idx = int(r["correct_index"] or 0)
        ans_letter = letters[c_idx] if 0 <= c_idx < 4 else "A"
        questions.append({
            "q_id": r["question_id"],
            "grade": r["grade"],
            "unit_code": r["unit_id"],
            "trap_id": r["trap_code"] or "TRAP_TENSE_SIGNAL",
            "level": r["difficulty_level"],
            "q_type": r["skill_domain"],
            "stem": r["stem_text"],
            "options": opts_dict,
            "answer": ans_letter,
            "signal_words": json.loads(r["signal_keywords_json"] or "[]"),
            "explanation_vi": r["explanation_vi"],
        })

    norm_unit = unit_code.replace("-", "_").replace("G6", "G06").replace("G7", "G07").replace("G8", "G08").replace("G9", "G09")
    cur.execute("SELECT core_vocab_json FROM map1_curriculum_units WHERE unit_id = ? OR grade = ? LIMIT 1", (norm_unit, int(grade)))
    u_row = cur.fetchone()
    vocab_foldable = json.loads(u_row["core_vocab_json"]) if u_row and u_row["core_vocab_json"] else []
    conn.close()

    ws_id = f"WS_CUSTOM_G{grade}_{datetime.now().strftime('%H%M%S')}"
    ws_title = title or f"Phiếu Luyện Tập Tiếng Anh Lớp {grade} ({unit_code})"
    return build_dual_docx(ws_id, ws_title, grade, questions, vocab_foldable=vocab_foldable, exam_code=exam_code)


if __name__ == "__main__":
    res = generate_worksheet_from_db(grade=6, unit_code="G06_U03", count=15, title="Phiếu Ôn Tập Lớp 6 — Unit 3 My Friends")
    print("Generated Dual DOCX v2.0 successfully:")
    print(" - Student:", res["student_docx_path"])
    print(" - Teacher:", res["teacher_docx_path"])
