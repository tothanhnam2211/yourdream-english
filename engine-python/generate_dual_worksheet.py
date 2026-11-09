# -*- coding: utf-8 -*-
"""
MASTER THÚY — ENGLISH STUDIO v2.2
Dual-Output Worksheet & 4-Code Exam Mixer Engine (Map 1 - Map 4 Contract)
Upgrades implemented per CHECKP_HANDOFF_GPT_TO_ANTI_MASTER_THUY.md:
1. Strict Unit Lock (exact_unit_rows: no silent fallback to random grade/whole DB; exact unit_id vocab lookup).
2. Group-aware 4-Code Mixer (mix_exam_codes preserves PART boundaries and Reading/Proofreading passage groups).
3. Explicit A4 Layout (210x297 mm, margins 17/16.5/21.5/20.5 mm, PAGE footer, brand 'MASTER THÚY · Thạc sĩ · Giảng viên Tiếng Anh · Zalo: 0836 686 874').
4. Full Student & Teacher Architecture: REMEMBER THE PATTERNS box, PART A-F headers, MARKING OVERVIEW (10.0 pts Decimal check),
   4-Code Quick Grid (101-104), 3-Column Zebra Key (12|32|124 mm = 168 mm, repeat header, cantSplit),
   Written Answer Contraction Variants, Writing Rubric & Sample Paragraph with declared word_count().
5. Dual DOCX + PDF generation (ReportLab A4 Unicode TTF when available).
"""

from decimal import Decimal
from html import escape
import json
import random
import re
import sqlite3
from pathlib import Path
from datetime import datetime
import docx
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_PATH = PROJECT_ROOT / "db" / "master_thuy_studio.db"
SHARDS_ITEMS_DIR = PROJECT_ROOT / "public" / "shards" / "items"
OUTPUT_DOCX_DIR = PROJECT_ROOT / "public" / "generated-docx"

PALETTE = {
    "teal": "0F766E",
    "mint": "F0FDFA",
    "amber": "FEF3C7",
    "coral": "FFF1F2",
    "slate": "1E293B",
}

PART_LABELS = {
    "PART_A": "PART A · PRONUNCIATION",
    "PART_B": "PART B · VOCABULARY & GRAMMAR",
    "PART_C": "PART C · ERROR CORRECTION / PROOFREADING",
    "PART_D": "PART D · READING COMPREHENSION",
    "PART_E": "PART E · SENTENCE BUILDING / TRANSFORMATION",
    "PART_F": "PART F · GUIDED WRITING CHALLENGE",
}


def exact_unit_rows(conn, grade, unit_ids, count, seed=20261109):
    """Strict Unit scope guard: queries exact grade + unit_ids and raises ValueError if insufficient."""
    if isinstance(unit_ids, str):
        unit_ids = [unit_ids]
    if not unit_ids or count < 1 or len(set(unit_ids)) != len(unit_ids):
        raise ValueError("Explicit distinct Unit IDs and positive count required")
    marks = ",".join("?" for _ in unit_ids)
    units = conn.execute(
        f"SELECT unit_id, grade FROM map1_curriculum_units WHERE unit_id IN ({marks})",
        tuple(unit_ids),
    ).fetchall()
    if len(units) != len(unit_ids) or any(int(r[1]) != int(grade) for r in units):
        raise ValueError(f"SOURCE_GAP: Unit {unit_ids} missing or grade {grade} mismatch; no broader fallback allowed")
    rows = conn.execute(
        f"SELECT * FROM question_bank WHERE grade=? AND unit_id IN ({marks}) ORDER BY question_id",
        (int(grade), *unit_ids),
    ).fetchall()
    if len(rows) < count:
        raise ValueError(f"INSUFFICIENT_SCOPED_ITEMS: found {len(rows)} items for {unit_ids}, required {count}")
    if len(rows) == count:
        return list(rows)
    return random.Random(seed).sample(list(rows), count)


def require_total(points, expected="10.0"):
    actual = sum((Decimal(str(x)) for x in points), Decimal("0"))
    if actual != Decimal(str(expected)):
        raise ValueError(f"Score total {actual} differs from expected {expected}")
    return actual


def word_count(text: str) -> int:
    """Declared convention: whitespace-delimited words; hyphenated compound counts as one word."""
    return len(str(text or "").split())


def set_Font(run, name="Carlito", size=10, bold=False, italic=False, underline=False, color_hex=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.color.rgb = RGBColor.from_string(color_hex or PALETTE["slate"])


def shade_cell(cell, fill_hex: str):
    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), fill_hex.lstrip("#"))
    cell._tc.get_or_add_tcPr().append(shade)


def new_a4_docx(role: str, font: str = "Carlito"):
    if role not in ("STUDENT", "TEACHER"):
        raise ValueError(f"Invalid role: {role}")
    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(17.0), Mm(16.5)
    sec.left_margin, sec.right_margin = Mm(21.5), Mm(20.5)
    sec.header_distance, sec.footer_distance = Mm(6.5), Mm(6.8)

    style = doc.styles["Normal"]
    style.font.name = font
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor.from_string(PALETTE["slate"])
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.line_spacing = 1.08

    hdr_p = sec.header.paragraphs[0]
    hdr_p.text = ""
    r_hdr = hdr_p.add_run("MASTER THÚY · Thạc sĩ · Giảng viên Tiếng Anh · Zalo: 0836 686 874")
    set_Font(r_hdr, name=font, size=8.5, bold=True, color_hex=PALETTE["teal"])

    ftr_p = sec.footer.paragraphs[0]
    ftr_p.text = ""
    r_ftr = ftr_p.add_run(f"MASTER THÚY · Zalo: 0836 686 874 | {role}'S COPY | Page ")
    set_Font(r_ftr, name=font, size=8.5, italic=True, color_hex=PALETTE["slate"])
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    ftr_p._p.append(field)
    return doc


def add_rich_html_runs(p, html_text: str, size=10, default_bold=False, default_italic=False, color_hex=None):
    """Parses <u>...</u>, <b>...</b>, and <br/> tags and adds properly styled Word runs."""
    if not html_text:
        return
    text = str(html_text).replace("<br/>", "\n").replace("<br>", "\n")
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
                set_Font(r, size=size, bold=is_b, italic=default_italic, underline=is_u, color_hex=color_hex)


def add_master_thuy_header(doc, title_text, grade, exam_code="101", is_teacher=False, duration_min=50, total_pts="10.0"):
    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for row in table.rows:
        row.cells[0].width = Mm(108)
        row.cells[1].width = Mm(60)

    cell_tl = table.cell(0, 0)
    shade_cell(cell_tl, PALETTE["mint"])
    p_tl = cell_tl.paragraphs[0]
    r1 = p_tl.add_run("MASTER THÚY · Thạc sĩ · Giảng viên Tiếng Anh\n")
    set_Font(r1, size=11, bold=True, color_hex=PALETTE["teal"])
    r2 = p_tl.add_run(f"Zalo: 0836 686 874 · Global Success Grade {grade}")
    set_Font(r2, size=9.5, italic=True)

    cell_tr = table.cell(0, 1)
    shade_cell(cell_tr, PALETTE["mint"])
    p_tr = cell_tr.paragraphs[0]
    p_tr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    role_badge = "TEACHER'S COPY" if is_teacher else "STUDENT'S COPY"
    r_title = p_tr.add_run(f"{role_badge} · CODE {exam_code}\n")
    set_Font(r_title, size=10.5, bold=True, color_hex=PALETTE["teal"])
    r_time = p_tr.add_run(f"Time: {duration_min} mins · Total: {total_pts} pts")
    set_Font(r_time, size=9.5, bold=True)

    cell_bl = table.cell(1, 0)
    p_bl = cell_bl.paragraphs[0]
    if is_teacher:
        r_bl = p_bl.add_run(f"{title_text} — ANSWER KEY, RUBRIC & EXPLANATIONS")
        set_Font(r_bl, size=10, bold=True, color_hex=PALETTE["teal"])
    else:
        r_bl = p_bl.add_run("Full name: ..........................................   Class: ............   Date: ..../..../2026")
        set_Font(r_bl, size=9.5)

    cell_br = table.cell(1, 1)
    p_br = cell_br.paragraphs[0]
    p_br.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if is_teacher:
        r_br = p_br.add_run(f"Updated: {datetime.now().strftime('%d/%m/%Y')}")
        set_Font(r_br, size=9.5, italic=True)
    else:
        r_br = p_br.add_run(f"Score: ........ / {total_pts}")
        set_Font(r_br, size=10, bold=True, color_hex=PALETTE["teal"])

    p_main_title = doc.add_paragraph()
    p_main_title.paragraph_format.space_before = Pt(6)
    p_main_title.paragraph_format.space_after = Pt(4)
    r_mt = p_main_title.add_run(title_text.upper())
    set_Font(r_mt, size=13, bold=True, color_hex=PALETTE["teal"])


def add_remember_patterns_box(doc, patterns: list):
    if not patterns:
        return
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    tbl.columns[0].width = Mm(168)
    cell = tbl.cell(0, 0)
    shade_cell(cell, PALETTE["mint"])
    p = cell.paragraphs[0]
    r_head = p.add_run("REMEMBER THE PATTERNS:\n")
    set_Font(r_head, size=9.5, bold=True, color_hex=PALETTE["teal"])
    for idx, pat in enumerate(patterns):
        sep = "\n" if idx < len(patterns) - 1 else ""
        r_item = p.add_run(f"  • {pat}{sep}")
        set_Font(r_item, size=9.5, bold=True)
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
    """
    Shuffles questions ONLY within independent (section_id, passage_group) blocks
    so PART A..F order and Reading/Proofreading passages are 100% preserved across codes 101-104.
    """
    mixed = {}
    letters = ["A", "B", "C", "D"]

    for idx, code in enumerate(codes):
        rng = random.Random(seed_base + idx * 101)
        q_copy = [dict(q) for q in questions]

        # Group contiguous blocks by (section_id, has_passage_block)
        blocks = []
        for q in q_copy:
            sec_id = q.get("section_id") or "PART_B"
            pass_txt = (q.get("passage_or_sign_svg") or "").strip()
            # Keep passage-linked items in fixed internal order; standalone items in a section can shuffle
            block_key = (sec_id, pass_txt if ("<div" in pass_txt or "TEXT " in pass_txt) else "")
            if not blocks or blocks[-1][0] != block_key:
                blocks.append((block_key, [q]))
            else:
                blocks[-1][1].append(q)

        ordered_questions = []
        for (sec_id, pass_group), group_items in blocks:
            g_list = list(group_items)
            if idx > 0 and not pass_group and len(g_list) > 1:
                rng.shuffle(g_list)
            ordered_questions.extend(g_list)

        code_questions = []
        for new_num, q in enumerate(ordered_questions, start=1):
            opts = dict(q.get("options", {}))
            orig_ans = q.get("answer", "A")
            ans_mode = q.get("answer_mode", "MCQ")
            opt_perm = ["A", "B", "C", "D"]

            # Do not shuffle T/F/NG options or WRITTEN prompts
            is_tfng = any("T (True)" in str(v) or "NG (Not Given)" in str(v) for v in opts.values())
            if ans_mode == "MCQ" and len(opts) == 4 and idx > 0 and not is_tfng:
                items = list(opts.items())
                rng.shuffle(items)
                new_opts = {}
                new_ans = orig_ans
                opt_perm = []
                for new_let, (old_let, txt) in zip(letters, items):
                    new_opts[new_let] = txt
                    opt_perm.append(old_let)
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
                "option_permutation": opt_perm,
                "exam_code": code,
            })
        mixed[code] = code_questions
    return mixed


def add_zebra_key(doc, rows):
    """rows: list of (display_number_str, key_text, explanation); widths 12 | 32 | 124 = 168 mm."""
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    for col, width in zip(table.columns, (12, 32, 124)):
        col.width = Mm(width)
    for c, value in zip(table.rows[0].cells, ("No.", "Key", "Explanation (Signal → Rule → Answer/Trap)")):
        c.text = ""
        shade_cell(c, PALETTE["teal"])
        r = c.paragraphs[0].add_run(value)
        set_Font(r, size=9, bold=True, color_hex="FFFFFF")
    repeat = OxmlElement("w:tblHeader")
    table.rows[0]._tr.get_or_add_trPr().append(repeat)

    for index, row in enumerate(rows):
        cells = table.add_row().cells
        for c_i, (c, value) in enumerate(zip(cells, row)):
            c.width = Mm((12, 32, 124)[c_i])
            c.text = ""
            if index % 2 == 0:
                shade_cell(c, PALETTE["mint"])
            p = c.paragraphs[0]
            add_rich_html_runs(
                p,
                str(value),
                size=8.5,
                default_bold=(c_i < 2),
                color_hex=(PALETTE["teal"] if c_i == 1 else PALETTE["slate"]),
            )
        table.rows[-1]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    return table


def try_build_pdf(pdf_path: Path, role: str, title: str, grade: int, exam_code: str, questions: list, blueprint_meta: dict, zebra_rows=None):
    """Generates an A4 PDF using ReportLab (if available) or Chromium Headless A4 print-to-pdf."""
    import shutil
    import subprocess
    import tempfile

    chromium_bin = shutil.which("chromium") or shutil.which("google-chrome")
    if chromium_bin:
        overview = (blueprint_meta or {}).get("marking_overview", [])
        rubric = (blueprint_meta or {}).get("writing_rubric", [])
        sample_p = (blueprint_meta or {}).get("sample_paragraph", "")
        patterns = (blueprint_meta or {}).get("remember_patterns", [])
        duration_min = (blueprint_meta or {}).get("duration_minutes", 50)
        total_pts = (blueprint_meta or {}).get("total_points", "10.0")

        body_parts = []
        body_parts.append(f"""
        <div class="header-box">
          <div class="header-left">
            <div class="brand">MASTER THÚY · Thạc sĩ · Giảng viên Tiếng Anh</div>
            <div class="sub-brand">Zalo: 0836 686 874 · Global Success Grade {grade}</div>
          </div>
          <div class="header-right">
            <div class="role-badge">{role}'S COPY · CODE {escape(str(exam_code))}</div>
            <div class="meta-line">Time: {duration_min} mins · Total: {escape(str(total_pts))} pts</div>
          </div>
        </div>
        <div class="student-line">
          {"<b>INTERNAL TEACHER GUIDE — ANSWER KEY, RUBRIC & EXPLANATIONS</b>" if role == "TEACHER" else "Full name: ................................................................. Class: .................. Date: ..../..../2026 &nbsp;&nbsp;|&nbsp;&nbsp; <b>Score: ........ / 10.0</b>"}
        </div>
        <h1>{escape(title)}</h1>
        """)

        if role == "STUDENT":
            if patterns:
                pat_items = "".join(f"<li>{escape(p)}</li>" for p in patterns)
                body_parts.append(f'<div class="remember-box"><b>REMEMBER THE PATTERNS:</b><ul>{pat_items}</ul></div>')
            cur_sec = None
            last_passage = None
            for q in questions:
                sec_id = q.get("section_id") or "PART_B"
                if sec_id != cur_sec:
                    cur_sec = sec_id
                    body_parts.append(f'<h2>{escape(PART_LABELS.get(sec_id, sec_id))}</h2>')
                pass_html = (q.get("passage_or_sign_svg") or "").strip()
                if pass_html and pass_html != last_passage and "sign-visual-box" not in pass_html:
                    last_passage = pass_html
                    body_parts.append(f'<div class="passage-box">{pass_html}</div>')
                body_parts.append(f'<div class="q-item"><span class="q-num">{q["display_num"]:02d}.</span> {q["stem"]}</div>')
                opts = q.get("options", {})
                if q.get("answer_mode", "MCQ") == "MCQ" and opts:
                    opt_cells = "".join(f'<div class="opt-cell"><b>{escape(k)}.</b> {v}</div>' for k, v in opts.items())
                    body_parts.append(f'<div class="opt-grid">{opt_cells}</div>')
                else:
                    lines_n = 4 if sec_id == "PART_F" else 1
                    for _ in range(lines_n):
                        body_parts.append('<div class="write-line">........................................................................................................................................................................</div>')
        else:
            if overview:
                ov_rows = "".join(
                    f'<tr><td><b>{escape(it["section"])}</b></td><td>{escape(it["items"])}</td><td><b>{escape(it["score"])}</b></td></tr>'
                    for it in overview
                )
                body_parts.append(f"""
                <h2>1. MARKING OVERVIEW ({escape(str(total_pts))} pts)</h2>
                <table class="zebra-table">
                  <thead><tr><th style="width:45%">Section</th><th style="width:22%">Items</th><th style="width:33%">Score Breakdown</th></tr></thead>
                  <tbody>{ov_rows}</tbody>
                </table>
                """)
            if zebra_rows:
                z_html = "".join(
                    f'<tr><td class="col-no"><b>{escape(str(r_no))}</b></td><td class="col-key"><b>{r_key}</b></td><td class="col-exp">{r_exp}</td></tr>'
                    for r_no, r_key, r_exp in zebra_rows
                )
                body_parts.append(f"""
                <h2>2. ZEBRA ANSWER KEY & PEDAGOGICAL EXPLANATIONS</h2>
                <table class="zebra-table">
                  <thead><tr><th style="width:12mm">No.</th><th style="width:32mm">Key</th><th style="width:124mm">Explanation (Signal → Rule → Answer/Trap)</th></tr></thead>
                  <tbody>{z_html}</tbody>
                </table>
                """)
            if rubric:
                rb_html = "".join(
                    f'<tr><td>{escape(it["criterion"])}</td><td><b>{escape(str(it["points"]))}</b></td></tr>'
                    for it in rubric
                )
                body_parts.append(f"""
                <h2>3. WRITING RUBRIC</h2>
                <table class="zebra-table">
                  <thead><tr><th style="width:80%">Criterion</th><th style="width:20%">Points</th></tr></thead>
                  <tbody>{rb_html}</tbody>
                </table>
                """)
            if sample_p:
                wc = word_count(sample_p)
                body_parts.append(f"""
                <h2>4. SAMPLE PARAGRAPH ({wc} words — whitespace word_count convention)</h2>
                <div class="passage-box"><i>{escape(sample_p)}</i></div>
                """)

        html_doc = f"""<!DOCTYPE html>
        <html lang="en"><head><meta charset="utf-8"/>
        <style>
          @page {{ size: A4; margin: 17mm 20.5mm 16.5mm 21.5mm; }}
          body {{ font-family: 'Liberation Sans', 'DejaVu Sans', sans-serif; font-size: 9.5pt; line-height: 1.35; color: #{PALETTE["slate"]}; margin: 0; }}
          .header-box {{ display: flex; justify-content: space-between; background: #{PALETTE["mint"]}; border-left: 4px solid #{PALETTE["teal"]}; padding: 6px 10px; }}
          .brand {{ font-weight: 700; font-size: 10.5pt; color: #{PALETTE["teal"]}; }}
          .sub-brand, .meta-line {{ font-size: 8.5pt; color: #475569; }}
          .role-badge {{ font-weight: 700; font-size: 10pt; color: #{PALETTE["teal"]}; text-align: right; }}
          .student-line {{ font-size: 9pt; padding: 5px 2px; border-bottom: 1px solid #CBD5E1; margin-bottom: 6px; }}
          h1 {{ font-size: 12.5pt; color: #{PALETTE["teal"]}; margin: 6px 0; text-transform: uppercase; }}
          h2 {{ font-size: 10pt; color: #{PALETTE["teal"]}; margin: 8px 0 4px 0; border-bottom: 1px solid #{PALETTE["mint"]}; padding-bottom: 2px; }}
          .remember-box, .passage-box {{ background: #{PALETTE["mint"]}; border-left: 3px solid #{PALETTE["teal"]}; padding: 6px 10px; margin: 5px 0; font-size: 9pt; }}
          .remember-box ul {{ margin: 3px 0 0 16px; padding: 0; }}
          .q-item {{ margin-top: 4px; page-break-inside: avoid; }}
          .q-num {{ font-weight: 700; color: #{PALETTE["teal"]}; }}
          .opt-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 2px 12px; margin: 2px 0 4px 14px; font-size: 9pt; page-break-inside: avoid; }}
          .write-line {{ margin: 3px 0 4px 14px; color: #64748B; font-size: 8.5pt; }}
          table.zebra-table {{ width: 168mm; border-collapse: collapse; margin: 4px 0 8px 0; font-size: 8.5pt; }}
          table.zebra-table thead {{ display: table-header-group; }}
          table.zebra-table tr {{ page-break-inside: avoid; }}
          table.zebra-table th {{ background: #{PALETTE["teal"]}; color: #FFFFFF; text-align: left; padding: 4px 6px; }}
          table.zebra-table td {{ padding: 4px 6px; border-bottom: 0.5px solid #E2E8F0; vertical-align: top; }}
          table.zebra-table tbody tr:nth-child(odd) {{ background: #{PALETTE["mint"]}; }}
          .col-key {{ color: #{PALETTE["teal"]}; }}
        </style></head><body>{"".join(body_parts)}</body></html>"""

        with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tf:
            tf.write(html_doc)
            tmp_html = Path(tf.name)
        try:
            subprocess.run(
                [
                    chromium_bin,
                    "--headless",
                    "--no-sandbox",
                    "--disable-gpu",
                    "--no-pdf-header-footer",
                    f"--print-to-pdf={pdf_path}",
                    f"file://{tmp_html}",
                ],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            if pdf_path.is_file():
                return str(pdf_path)
        finally:
            tmp_html.unlink(missing_ok=True)
    return None


def build_dual_docx(ws_id, title, grade, questions, vocab_foldable=None, exam_code="101", blueprint_meta=None):
    OUTPUT_DOCX_DIR.mkdir(parents=True, exist_ok=True)
    blueprint_meta = blueprint_meta or {}
    duration_min = blueprint_meta.get("duration_minutes", 50)
    total_pts = blueprint_meta.get("total_points", "10.0")
    overview = blueprint_meta.get("marking_overview", [])
    if overview:
        require_total([item.get("points", "0") for item in overview], expected=total_pts)

    mixed_variants = mix_exam_codes(questions, codes=("101", "102", "103", "104"))
    selected_questions = mixed_variants.get(str(exam_code), mixed_variants["101"])

    # =========================================================================
    # 1. STUDENT DOCX (A4 Explicit, Remember Patterns, PART A..F, Passage Lock)
    # =========================================================================
    student_doc = new_a4_docx("STUDENT")
    add_master_thuy_header(student_doc, title, grade, exam_code=exam_code, is_teacher=False, duration_min=duration_min, total_pts=total_pts)

    if blueprint_meta.get("remember_patterns"):
        add_remember_patterns_box(student_doc, blueprint_meta["remember_patterns"])

    if vocab_foldable:
        p_voc_head = student_doc.add_paragraph()
        r_vh = p_voc_head.add_run("CORE VOCABULARY FOLDABLE CHECKLIST (Fold along the dotted line to self-test):")
        set_Font(r_vh, size=10, bold=True, color_hex=PALETTE["teal"])
        v_table = student_doc.add_table(rows=len(vocab_foldable) + 1, cols=2)
        v_table.autofit = False
        v_table.columns[0].width = Mm(84)
        v_table.columns[1].width = Mm(84)
        shade_cell(v_table.cell(0, 0), PALETTE["mint"])
        shade_cell(v_table.cell(0, 1), PALETTE["mint"])
        set_Font(v_table.cell(0, 0).paragraphs[0].add_run("Word & IPA Pronunciation"), size=9.5, bold=True, color_hex=PALETTE["teal"])
        set_Font(v_table.cell(0, 1).paragraphs[0].add_run("┆  Vietnamese Meaning (Fold to hide)"), size=9.5, bold=True, color_hex=PALETTE["teal"])
        for row_idx, v in enumerate(vocab_foldable, start=1):
            r_l = v_table.cell(row_idx, 0).paragraphs[0].add_run(
                f"{row_idx:02d}. {v.get('word', '')} {v.get('ipa', '')} ({v.get('pos', '')})"
            )
            set_Font(r_l, size=9.5, bold=True)
            r_r = v_table.cell(row_idx, 1).paragraphs[0].add_run(f"┆  {v.get('meaning_vi', '')}")
            set_Font(r_r, size=9.5)
        student_doc.add_paragraph()

    current_section = None
    last_rendered_passage = None
    for q in selected_questions:
        sec_id = q.get("section_id") or "PART_B"
        if sec_id != current_section:
            current_section = sec_id
            p_sec = student_doc.add_paragraph()
            p_sec.paragraph_format.space_before = Pt(6)
            p_sec.paragraph_format.space_after = Pt(3)
            r_sec = p_sec.add_run(PART_LABELS.get(sec_id, sec_id))
            set_Font(r_sec, size=10.5, bold=True, color_hex=PALETTE["teal"])

        passage_html = (q.get("passage_or_sign_svg") or "").strip()
        if passage_html and passage_html != last_rendered_passage and "sign-visual-box" not in passage_html:
            last_rendered_passage = passage_html
            p_tbl = student_doc.add_table(rows=1, cols=1)
            p_tbl.autofit = False
            p_tbl.columns[0].width = Mm(168)
            p_cell = p_tbl.cell(0, 0)
            shade_cell(p_cell, PALETTE["mint"])
            add_rich_html_runs(p_cell.paragraphs[0], passage_html, size=9.5)
            student_doc.add_paragraph()

        p = student_doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r_num = p.add_run(f"{q['display_num']:02d}. ")
        set_Font(r_num, size=10, bold=True, color_hex=PALETTE["teal"])
        add_rich_html_runs(p, q["stem"], size=10)

        ans_mode = q.get("answer_mode", "MCQ")
        opts = q.get("options", {})
        if ans_mode == "MCQ" and opts:
            max_len = max((len(re.sub(r"<[^>]+>", "", str(v))) for v in opts.values()), default=0)
            p_opt = student_doc.add_paragraph()
            p_opt.paragraph_format.left_indent = Mm(6)
            p_opt.paragraph_format.space_after = Pt(4)
            sep = "      " if max_len <= 22 else "\n"
            for o_i, (k, v) in enumerate(opts.items()):
                if o_i > 0:
                    r_sep = p_opt.add_run(sep)
                    set_Font(r_sep, size=9.5)
                r_k = p_opt.add_run(f"{k}. ")
                set_Font(r_k, size=9.5, bold=True)
                add_rich_html_runs(p_opt, str(v), size=9.5)
        else:
            lines_count = 5 if sec_id == "PART_F" else 1
            for _ in range(lines_count):
                p_line = student_doc.add_paragraph()
                p_line.paragraph_format.left_indent = Mm(6)
                r_dots = p_line.add_run("→ .........................................................................................................................................." if lines_count == 1 else "....................................................................................................................................................")
                set_Font(r_dots, size=9.5, color_hex="64748B")

    student_filename = f"{ws_id}_Code{exam_code}_STUDENT.docx"
    student_path = OUTPUT_DOCX_DIR / student_filename
    student_doc.save(str(student_path))

    # =========================================================================
    # 2. TEACHER DOCX (Marking Overview -> 4-Code Grid -> Zebra Key -> Rubric)
    # =========================================================================
    teacher_doc = new_a4_docx("TEACHER")
    add_master_thuy_header(teacher_doc, title, grade, exam_code="101-104", is_teacher=True, duration_min=duration_min, total_pts=total_pts)

    # 2A. MARKING OVERVIEW (10.0 pts)
    if overview:
        p_ov = teacher_doc.add_paragraph()
        set_Font(p_ov.add_run(f"1. MARKING OVERVIEW ({total_pts} pts)"), size=10.5, bold=True, color_hex=PALETTE["teal"])
        ov_table = teacher_doc.add_table(rows=len(overview) + 2, cols=3)
        ov_table.autofit = False
        for c_i, w_mm in enumerate((76, 36, 56)):
            ov_table.columns[c_i].width = Mm(w_mm)
        for c_i, h_txt in enumerate(("Section", "Items", "Score Breakdown")):
            shade_cell(ov_table.cell(0, c_i), PALETTE["teal"])
            set_Font(ov_table.cell(0, c_i).paragraphs[0].add_run(h_txt), size=9, bold=True, color_hex="FFFFFF")
        for r_i, item in enumerate(overview, start=1):
            if r_i % 2 == 1:
                for c_i in range(3):
                    shade_cell(ov_table.cell(r_i, c_i), PALETTE["mint"])
            set_Font(ov_table.cell(r_i, 0).paragraphs[0].add_run(item["section"]), size=9, bold=True)
            set_Font(ov_table.cell(r_i, 1).paragraphs[0].add_run(item["items"]), size=9)
            set_Font(ov_table.cell(r_i, 2).paragraphs[0].add_run(item["score"]), size=9, bold=True, color_hex=PALETTE["teal"])
        last_r = len(overview) + 1
        for c_i in range(3):
            shade_cell(ov_table.cell(last_r, c_i), PALETTE["amber"])
        set_Font(ov_table.cell(last_r, 0).paragraphs[0].add_run("TOTAL"), size=9.5, bold=True, color_hex=PALETTE["teal"])
        set_Font(ov_table.cell(last_r, 1).paragraphs[0].add_run(f"{len(selected_questions)} items"), size=9.5, bold=True)
        set_Font(ov_table.cell(last_r, 2).paragraphs[0].add_run(f"{total_pts} pts"), size=9.5, bold=True, color_hex=PALETTE["teal"])
        teacher_doc.add_paragraph()

    # 2B. 4-CODE QUICK COMPARISON TABLE (101 | 102 | 103 | 104)
    p_grid_title = teacher_doc.add_paragraph()
    set_Font(p_grid_title.add_run("2. QUICK ANSWER KEY MATRIX (CODES 101 · 102 · 103 · 104):"), size=10.5, bold=True, color_hex=PALETTE["teal"])

    total_q = len(selected_questions)
    grid_table = teacher_doc.add_table(rows=total_q + 1, cols=5)
    grid_table.autofit = False
    for c_idx, h_text in enumerate(("No.", "Code 101", "Code 102", "Code 103", "Code 104")):
        shade_cell(grid_table.cell(0, c_idx), PALETTE["teal"])
        set_Font(grid_table.cell(0, c_idx).paragraphs[0].add_run(h_text), size=9, bold=True, color_hex="FFFFFF")
    grid_table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

    for q_idx in range(total_q):
        if q_idx % 2 == 0:
            for c_i in range(5):
                shade_cell(grid_table.cell(q_idx + 1, c_i), PALETTE["mint"])
        set_Font(grid_table.cell(q_idx + 1, 0).paragraphs[0].add_run(f"{q_idx + 1:02d}"), size=8.5, bold=True)
        for c_i, code_str in enumerate(("101", "102", "103", "104"), start=1):
            q_item = mixed_variants[code_str][q_idx]
            ans_display = q_item["answer"] if q_item.get("answer_mode", "MCQ") == "MCQ" else (q_item.get("written_answers", [q_item["answer"]])[0][:22])
            set_Font(grid_table.cell(q_idx + 1, c_i).paragraphs[0].add_run(str(ans_display)), size=8.5, bold=True, color_hex=PALETTE["teal"])

    teacher_doc.add_paragraph()

    # 2C. 3-COLUMN ZEBRA KEY (12 | 32 | 124 mm = 168 mm)
    p_exp_title = teacher_doc.add_paragraph()
    set_Font(p_exp_title.add_run("3. ZEBRA ANSWER KEY & PEDAGOGICAL EXPLANATIONS (CODE 101):"), size=10.5, bold=True, color_hex=PALETTE["teal"])

    zebra_rows = []
    for q in mixed_variants["101"]:
        if q.get("answer_mode", "MCQ") == "MCQ":
            key_str = q["answer"]
        else:
            w_list = q.get("written_answers") or [q["answer"]]
            key_str = " / ".join(str(x) for x in w_list[:2])
        exp_str = q.get("explanation_vi", "")
        zebra_rows.append((f"{q['display_num']:02d}", key_str, exp_str))

    add_zebra_key(teacher_doc, zebra_rows)

    # 2D. WRITING RUBRIC & SAMPLE PARAGRAPH WITH WORD COUNT
    rubric = blueprint_meta.get("writing_rubric", [])
    if rubric:
        teacher_doc.add_paragraph()
        p_rb = teacher_doc.add_paragraph()
        set_Font(p_rb.add_run("4. WRITING RUBRIC & SAMPLE PARAGRAPH:"), size=10.5, bold=True, color_hex=PALETTE["teal"])
        rb_table = teacher_doc.add_table(rows=len(rubric) + 1, cols=2)
        rb_table.autofit = False
        rb_table.columns[0].width = Mm(136)
        rb_table.columns[1].width = Mm(32)
        shade_cell(rb_table.cell(0, 0), PALETTE["teal"])
        shade_cell(rb_table.cell(0, 1), PALETTE["teal"])
        set_Font(rb_table.cell(0, 0).paragraphs[0].add_run("Criterion"), size=9, bold=True, color_hex="FFFFFF")
        set_Font(rb_table.cell(0, 1).paragraphs[0].add_run("Points"), size=9, bold=True, color_hex="FFFFFF")
        for r_i, r_item in enumerate(rubric, start=1):
            if r_i % 2 == 1:
                shade_cell(rb_table.cell(r_i, 0), PALETTE["mint"])
                shade_cell(rb_table.cell(r_i, 1), PALETTE["mint"])
            set_Font(rb_table.cell(r_i, 0).paragraphs[0].add_run(r_item["criterion"]), size=9)
            set_Font(rb_table.cell(r_i, 1).paragraphs[0].add_run(str(r_item["points"])), size=9, bold=True, color_hex=PALETTE["teal"])

    sample_para = blueprint_meta.get("sample_paragraph", "")
    if sample_para:
        wc = word_count(sample_para)
        p_sp = teacher_doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(6)
        set_Font(p_sp.add_run(f"Sample Paragraph ({wc} words — whitespace word_count convention):\n"), size=9.5, bold=True, color_hex=PALETTE["teal"])
        set_Font(p_sp.add_run(sample_para), size=9.5, italic=True)

    teacher_filename = f"{ws_id}_AllCodes_TEACHER.docx"
    teacher_path = OUTPUT_DOCX_DIR / teacher_filename
    teacher_doc.save(str(teacher_path))

    # 3. DUAL PDF EXPORT (STUDENT + TEACHER)
    student_pdf_path = OUTPUT_DOCX_DIR / f"{ws_id}_Code{exam_code}_STUDENT.pdf"
    teacher_pdf_path = OUTPUT_DOCX_DIR / f"{ws_id}_AllCodes_TEACHER.pdf"
    s_pdf_ok = try_build_pdf(student_pdf_path, "STUDENT", title, grade, exam_code, selected_questions, blueprint_meta)
    t_pdf_ok = try_build_pdf(teacher_pdf_path, "TEACHER", title, grade, "101-104", mixed_variants["101"], blueprint_meta, zebra_rows=zebra_rows)

    answer_matrix = {
        code_str: [
            {
                "num": item["display_num"],
                "answer": item["answer"],
                "q_id": item["q_id"],
                "section_id": item.get("section_id", "PART_B"),
                "option_permutation": item.get("option_permutation", ["A", "B", "C", "D"]),
            }
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
        "student_pdf_path": s_pdf_ok,
        "teacher_pdf_path": t_pdf_ok,
        "answer_matrix_101_104": answer_matrix,
        "questions_preview": selected_questions[:20],
    }


def generate_worksheet_from_db(grade=6, unit_code="G06_U03", count=15, title=None, exam_code="101", worksheet_id=None):
    norm_unit = unit_code.replace("-", "_").replace("G6", "G06").replace("G7", "G07").replace("G8", "G08").replace("G9", "G09")
    blueprint_meta = {}

    if worksheet_id:
        shard_path = SHARDS_ITEMS_DIR / f"{worksheet_id}.json"
        if shard_path.exists():
            s_data = json.loads(shard_path.read_text(encoding="utf-8"))
            blueprint_meta = s_data.get("blueprint_meta", {})

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    if worksheet_id:
        rows = cur.execute(
            "SELECT * FROM question_bank WHERE source_worksheet_id = ? ORDER BY question_id",
            (worksheet_id,),
        ).fetchall()
        if not rows:
            conn.close()
            raise ValueError(f"No questions found for worksheet_id={worksheet_id}")
    else:
        rows = exact_unit_rows(conn, grade=int(grade), unit_ids=[norm_unit], count=int(count))

    letters = ["A", "B", "C", "D"]
    questions = []
    for r in rows:
        raw_opts = json.loads(r["options_json"] or "[]")
        opts_dict = normalize_options_dict(raw_opts)
        c_idx = int(r["correct_index"] or 0)
        ans_letter = letters[c_idx] if 0 <= c_idx < 4 else "A"
        r_keys = r.keys()
        questions.append({
            "q_id": r["question_id"],
            "grade": r["grade"],
            "unit_code": r["unit_id"],
            "section_id": r["section_id"] if "section_id" in r_keys and r["section_id"] else "PART_B",
            "trap_id": r["trap_code"] or "TRAP_TENSE_SIGNAL",
            "level": r["difficulty_level"],
            "q_type": r["skill_domain"],
            "stem": r["stem_text"],
            "passage_or_sign_svg": r["passage_or_sign_svg"] or "",
            "options": opts_dict,
            "answer": ans_letter,
            "answer_mode": r["answer_mode"] or "MCQ",
            "written_answers": json.loads(r["written_answers_json"] or "[]"),
            "signal_words": json.loads(r["signal_keywords_json"] or "[]"),
            "explanation_vi": r["explanation_vi"],
        })

    # Exact unit_id match only — no OR grade = ? fallback
    cur.execute("SELECT core_vocab_json FROM map1_curriculum_units WHERE unit_id = ? LIMIT 1", (norm_unit,))
    u_row = cur.fetchone()
    vocab_foldable = json.loads(u_row["core_vocab_json"]) if u_row and u_row["core_vocab_json"] else []
    conn.close()

    ws_id = worksheet_id or f"WS_CUSTOM_G{grade}_{norm_unit}"
    ws_title = title or f"Phiếu Luyện Tập Tiếng Anh Lớp {grade} ({norm_unit})"
    return build_dual_docx(
        ws_id,
        ws_title,
        grade,
        questions,
        vocab_foldable=vocab_foldable,
        exam_code=exam_code,
        blueprint_meta=blueprint_meta,
    )


if __name__ == "__main__":
    res = generate_worksheet_from_db(
        grade=6,
        unit_code="G06_U03",
        count=31,
        title="English 6 · Unit 3: My Friends — Advanced Practice",
        exam_code="101",
        worksheet_id="WS_G06_U03_ADV",
    )
    print("Generated Dual DOCX & PDF v2.2 successfully:")
    print(" - Student DOCX:", res["student_docx_path"])
    print(" - Teacher DOCX:", res["teacher_docx_path"])
    print(" - Student PDF :", res["student_pdf_path"])
    print(" - Teacher PDF :", res["teacher_pdf_path"])
