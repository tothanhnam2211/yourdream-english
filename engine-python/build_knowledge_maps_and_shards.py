# -*- coding: utf-8 -*-
"""
MASTER THÚY — ENGLISH STUDIO v2.0 (PRODUCTION & CLOUD-READY)
Full 63-File Extractor, Rich Formatting Preserver (<u>...</u> & <b>...</b>),
Visual Sign Frame Builder, 4-Maps Compiler & Sharded JSON/SQLite Engine.
"""

import csv
import json
import re
import shutil
import sqlite3
import zipfile
from datetime import datetime
from pathlib import Path
import docx
from docx.text.paragraph import Paragraph
from docx.table import Table
import pymupdf

PROJECT_ROOT = Path(r"D:\YOURDREAM.VN\master-thuy-english-studio")
RAW_ARCHIVE_DIR = PROJECT_ROOT / "raw_63_archive"
DRIVE_37_DIR = RAW_ARCHIVE_DIR / "drive_37"
UNSAVED_26_DIR = RAW_ARCHIVE_DIR / "unsaved_26" / "FILES_PENDING_DRIVE"

DB_PATH = PROJECT_ROOT / "db" / "master_thuy_studio.db"
SCHEMA_PATH = PROJECT_ROOT / "db" / "schema.sql"
SHARDS_DIR = PROJECT_ROOT / "public" / "shards"
MAPS_DIR = SHARDS_DIR / "maps"
ITEMS_DIR = SHARDS_DIR / "items"
SIGNS_ASSET_DIR = PROJECT_ROOT / "public" / "assets" / "signs"
ARCHIVE_PUBLIC_DIR = PROJECT_ROOT / "public" / "archive-files"


# ============================================================================
# 1. RICH RUN FORMATTER (PRESERVES 100% <u>...</u> AND <b>...</b> FROM WORD)
# ============================================================================
def paragraph_to_rich_html(para: Paragraph) -> str:
    """Extracts text from a docx Paragraph while preserving <u>...</u> and <b>...</b> runs."""
    if not para.runs:
        return para.text.strip()
    parts = []
    for run in para.runs:
        txt = run.text
        if not txt:
            continue
        is_u = bool(run.underline) or ("<w:u " in run._r.xml and 'w:val="none"' not in run._r.xml)
        is_b = bool(run.bold) or ("<w:b/>" in run._r.xml or "<w:b " in run._r.xml)
        if is_u and txt.strip():
            # Keep leading/trailing spaces outside <u>
            lead = len(txt) - len(txt.lstrip())
            trail = len(txt) - len(txt.rstrip())
            core = txt.strip()
            txt = f"{txt[:lead]}<u>{core}</u>{txt[len(txt)-trail:] if trail else ''}"
        if is_b and txt.strip():
            lead = len(txt) - len(txt.lstrip())
            trail = len(txt) - len(txt.rstrip())
            core = txt.strip()
            txt = f"{txt[:lead]}<b>{core}</b>{txt[len(txt)-trail:] if trail else ''}"
        parts.append(txt)
    merged = "".join(parts).strip()
    # Merge adjacent tags like </u><u> or </b><b>
    merged = merged.replace("</u><u>", "").replace("</b><b>", "")
    return merged


def cell_to_rich_html(cell) -> str:
    lines = []
    for p in cell.paragraphs:
        s = paragraph_to_rich_html(p)
        if s:
            lines.append(s)
    return " ".join(lines).strip()


def iter_block_items(doc):
    """Yields each paragraph and table child within document body in true document order."""
    parent_elm = doc.element.body
    for child in parent_elm.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def enhance_phonetic_underline(stem: str, options: list) -> list:
    """
    Ensures Pronunciation options always have clear <u>...</u> around the target sound
    even if the original PDF or run only specified it in the stem (e.g. 'final -ed', 'final -s/-es', 'letter f').
    """
    low_stem = stem.lower()
    if any("<u>" in opt for opt in options):
        return options

    is_ed = "-ed" in low_stem or "ending -ed" in low_stem or all(
        re.sub(r"[^a-zA-Z]", "", o).lower().endswith("ed") for o in options if o
    )
    is_ses = "-s" in low_stem or "-es" in low_stem or all(
        re.sub(r"[^a-zA-Z]", "", o).lower().endswith("s") for o in options if o
    )

    new_opts = []
    for opt in options:
        if is_ed and re.search(r"ed\b", opt, re.I):
            new_opts.append(re.sub(r"(ed)\b", r"<u>\1</u>", opt, count=1, flags=re.I))
        elif is_ses and re.search(r"(?:es|s)\b", opt, re.I):
            new_opts.append(re.sub(r"(es|s)\b", r"<u>\1</u>", opt, count=1, flags=re.I))
        elif "letter f" in low_stem and re.search(r"f+", opt, re.I):
            new_opts.append(re.sub(r"(f+)", r"<u>\1</u>", opt, count=1, flags=re.I))
        elif "underlined" in low_stem:
            # Highlight common vowel/consonant cluster in short words
            w = opt.strip()
            if all("ch" in x.lower() for x in options):
                new_opts.append(re.sub(r"(ch)", r"<u>\1</u>", w, count=1, flags=re.I))
            elif all("th" in x.lower() for x in options):
                new_opts.append(re.sub(r"(th)", r"<u>\1</u>", w, count=1, flags=re.I))
            elif all("oo" in x.lower() for x in options):
                new_opts.append(re.sub(r"(oo)", r"<u>\1</u>", w, count=1, flags=re.I))
            else:
                new_opts.append(re.sub(r"([aeiouy]+)", r"<u>\1</u>", w, count=1, flags=re.I))
        else:
            new_opts.append(opt)
    return new_opts


def strip_opt_prefix(s: str) -> str:
    return re.sub(r"^(?:<b>)?[A-D][\.\)]\s*(?:</b>)?\s*", "", s.strip()).strip()


def detect_skill_and_trap(stem: str, opts: list):
    low = (stem + " " + " ".join(opts)).lower()
    if "sign" in low or "biển báo" in low or "notice" in low:
        return "SIGNS", "TRAP_SIGN_DISTRACTOR", ["must", "must not", "cannot", "only"]
    if "pronounced" in low or "underlined" in low or "stress" in low or "<u>" in " ".join(opts):
        if "ed</u>" in low or "-ed" in low:
            return "PHONETICS", "TRAP_PRON_ED", ["-ed", "/t/", "/d/", "/ɪd/"]
        return "PHONETICS", "TRAP_PRON_SES", ["-s", "-es", "stress"]
    if "have" in low or "get" in low or "got" in low or "had" in low:
        if any(w in low for w in ["repaired", "cut", "cleaned", "painted", "fixed", "checked", "tested", "delivered"]):
            return "GRAMMAR", "TRAP_CAUSATIVE_PASSIVE", ["have", "get", "by"]
    if any(w in low for w in ["every day", "usually", "always", "often", "now", "at the moment", "look!", "listen!", "yesterday", "last ", "ago", "since"]):
        return "GRAMMAR", "TRAP_TENSE_SIGNAL", ["every day", "usually", "now", "at the moment", "look!", "listen!", "yesterday", "last"]
    if "error" in low or "lỗi" in low or "<u>" in stem:
        return "ERROR_FIX", "TRAP_SV_AGREE", ["every day", "always", "there is", "there are"]
    if any(w in low for w in ["unfriendly", "careful", "careless", "helpful", "helpless", "pollution", "tradition", "diffic"]):
        return "WORD_FORM", "TRAP_WORD_FORM_NEG", ["un-", "im-", "-ful", "-less", "-tion"]
    return "GRAMMAR", "TRAP_SV_AGREE", ["usually", "every", "now"]


def build_contraction_variants(ans_str: str) -> list:
    """Generates equivalent full and contracted English answer strings for fair written grading."""
    base = re.sub(r"\s+", " ", ans_str.strip().rstrip(".")).strip()
    variants = {base, base.lower()}
    # Split slash alternatives like "do not work / don't work" or "There is not (isn't)"
    cleaned_paren = re.sub(r"\(([^)]+)\)", "", base).strip()
    if cleaned_paren:
        variants.add(cleaned_paren)
    for part in re.split(r"\s*/\s*|\s+or\s+", base):
        p = re.sub(r"\(([^)]+)\)", "", part).strip()
        if p:
            variants.add(p)

    contraction_pairs = [
        ("don't", "do not"), ("doesn't", "does not"), ("didn't", "did not"),
        ("isn't", "is not"), ("aren't", "are not"), ("wasn't", "was not"), ("weren't", "were not"),
        ("haven't", "have not"), ("hasn't", "has not"), ("hadn't", "had not"),
        ("won't", "will not"), ("wouldn't", "would not"), ("can't", "cannot"), ("can't", "can not"),
        ("couldn't", "could not"), ("shouldn't", "should not"), ("mustn't", "must not"),
        ("i'm", "i am"), ("you're", "you are"), ("he's", "he is"), ("she's", "she is"),
        ("it's", "it is"), ("we're", "we are"), ("they're", "they are"),
        ("there's", "there is"), ("let's", "let us")
    ]
    expanded = set(variants)
    for v in list(variants):
        v_norm = v.replace("’", "'")
        expanded.add(v_norm)
        for short_f, long_f in contraction_pairs:
            if re.search(rf"\b{re.escape(short_f)}\b", v_norm, re.I):
                expanded.add(re.sub(rf"\b{re.escape(short_f)}\b", long_f, v_norm, flags=re.I))
            if re.search(rf"\b{re.escape(long_f)}\b", v_norm, re.I):
                expanded.add(re.sub(rf"\b{re.escape(long_f)}\b", short_f, v_norm, flags=re.I))
    return sorted(x for x in expanded if x)


def make_4_mcq_from_written(correct_str: str, stem: str = "") -> tuple:
    """Builds 4 distinct options [A, B, C, D] and correct_index=0 (shuffled deterministically) from a written answer."""
    variants = build_contraction_variants(correct_str)
    primary = variants[0] if variants else correct_str.strip()
    low = primary.lower()

    distractors = []
    if low in ("there is", "there's"):
        distractors = ["There are", "There be", "It have"]
    elif low == "there are":
        distractors = ["There is", "There be", "They is"]
    elif low.startswith("is ") and low.endswith("ing"):
        v_ing = primary.split(" ", 1)[1]
        distractors = [f"are {v_ing}", v_ing, f"be {v_ing}"]
    elif low.startswith("are ") and low.endswith("ing"):
        v_ing = primary.split(" ", 1)[1]
        distractors = [f"is {v_ing}", v_ing, f"am {v_ing}"]
    elif low.startswith("am ") and low.endswith("ing"):
        v_ing = primary.split(" ", 1)[1]
        distractors = [f"is {v_ing}", f"are {v_ing}", v_ing]
    elif low.endswith("es") and len(low) > 3:
        root = primary[:-2]
        distractors = [root, f"is {root}ing", f"to {root}"]
    elif low.endswith("s") and not low.endswith("ss") and len(low) > 2:
        root = primary[:-1]
        distractors = [root, f"is {root}ing", f"to {root}"]
    else:
        distractors = [
            f"{primary}s" if not low.endswith("s") else primary[:-1],
            f"is {primary}ing" if " " not in primary else f"not {primary}",
            f"to {primary}"
        ]

    # Ensure 4 distinct options
    pool = [primary]
    for d in distractors:
        if d.lower() not in [x.lower() for x in pool]:
            pool.append(d)
    while len(pool) < 4:
        pool.append(f"{primary} (form {len(pool)})")

    # Deterministic rotation based on stem length
    shift = len(stem) % 4
    rotated = pool[shift:] + pool[:shift]
    c_idx = rotated.index(primary)
    return rotated, c_idx


# ============================================================================
# 2. SPECIALIZED PARSER A: 50 SIGNS AROUND US (WITH VISUAL SIGN CARDS & PNGS)
# ============================================================================
def parse_50_signs_worksheet() -> list:
    student_docx = DRIVE_37_DIR / "MASTER_THUY_G6_50_Signs_STUDENT.docx"
    teacher_docx = DRIVE_37_DIR / "MASTER_THUY_G6_50_Signs_TEACHER.docx"
    SIGNS_ASSET_DIR.mkdir(parents=True, exist_ok=True)

    # Extract 50 PNG sign graphics
    with zipfile.ZipFile(student_docx, "r") as zf:
        for i in range(1, 51):
            src_name = f"word/media/image{i}.png"
            if src_name in zf.namelist():
                target_png = SIGNS_ASSET_DIR / f"sign_{i:02d}.png"
                target_png.write_bytes(zf.read(src_name))

    # Parse Teacher Answer Key & Explanations from 3 tables in TEACHER.docx
    t_doc = docx.Document(str(teacher_docx))
    key_map = {}
    letters = ["A", "B", "C", "D"]
    for tbl in t_doc.tables:
        for row in tbl.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) >= 4 and cells[0].isdigit():
                q_num = int(cells[0])
                ans_let = cells[1].strip().upper()
                meaning_en = cells[2].strip()
                exp_vi = cells[3].strip()
                key_map[q_num] = {
                    "correct_index": letters.index(ans_let) if ans_let in letters else 0,
                    "meaning_en": meaning_en,
                    "explanation_vi": f"Biển báo #{q_num:02d}: {exp_vi} (Đáp án chuẩn: {meaning_en})"
                }

    # Parse Student 50 tables (T1..T50)
    s_doc = docx.Document(str(student_docx))
    questions = []
    for q_num in range(1, 51):
        if q_num >= len(s_doc.tables):
            break
        tbl = s_doc.tables[q_num]
        c1 = tbl.cell(0, 1)
        stem_raw = c1.paragraphs[0].text.strip() if c1.paragraphs else f"{q_num:02d}. What does the sign mean?"
        stem_clean = re.sub(r"^\d+[\.\)]\s*", "", stem_raw).strip() or "What does the sign mean?"

        opts = []
        if c1.tables:
            for r in c1.tables[0].rows:
                for cell in r.cells:
                    txt = cell_to_rich_html(cell)
                    if txt:
                        opts.append(strip_opt_prefix(txt))
        while len(opts) < 4:
            opts.append(f"Option {len(opts)+1}")

        k_info = key_map.get(q_num, {"correct_index": 0, "meaning_en": opts[0], "explanation_vi": "Quan sát kỹ biển báo."})
        part_badge = "PART A · EASY" if q_num <= 15 else ("PART B · INTERMEDIATE" if q_num <= 35 else "PART C · CHALLENGING")

        sign_visual_html = (
            f'<div class="sign-visual-box">'
            f'<div class="sign-visual-header">⚠️ BIỂN BÁO THỰC TẾ #{q_num:02d} ({part_badge})</div>'
            f'<div class="sign-visual-body">'
            f'<img src="../assets/signs/sign_{q_num:02d}.png" alt="Sign {q_num:02d}" class="sign-visual-img" '
            f'onerror="this.src=\'./assets/signs/sign_{q_num:02d}.png\'" />'
            f'</div></div>'
        )

        questions.append({
            "question_id": f"Q_WS_G06_SIGNS_50_{q_num:03d}",
            "grade": 6,
            "unit_id": "G06_U02",
            "skill_domain": "SIGNS",
            "difficulty_level": 1 if q_num <= 15 else (2 if q_num <= 35 else 3),
            "trap_code": "TRAP_SIGN_DISTRACTOR",
            "stem_text": f"<b>[Sign #{q_num:02d}]</b> {stem_clean}",
            "passage_or_sign_svg": sign_visual_html,
            "options": opts[:4],
            "correct_index": k_info["correct_index"],
            "answer_mode": "MCQ",
            "written_answers": [opts[k_info["correct_index"]]],
            "signal_keywords": ["must", "must not", "cannot", "should", "only", "no "],
            "explanation_vi": k_info["explanation_vi"],
            "source_worksheet_id": "WS_G06_SIGNS_50"
        })
    return questions


# ============================================================================
# 3. SPECIALIZED PARSER B: 48 ERROR CORRECTION & PROOFREADING (<u> & <b>)
# ============================================================================
def parse_48_error_correction_worksheet() -> list:
    student_docx = DRIVE_37_DIR / "MASTER_THUY_G6_Units1-3_ErrorCorrection_STUDENT.docx"
    teacher_docx = DRIVE_37_DIR / "MASTER_THUY_G6_Units1-3_ErrorCorrection_TEACHER.docx"

    t_doc = docx.Document(str(teacher_docx))
    key_map = {}
    # Tables T1, T2, T3 have 4 cols: ['#', 'Lỗi sai', 'Sửa đúng', 'Giải thích ngắn']
    for ti in [1, 2, 3]:
        if ti < len(t_doc.tables):
            for row in t_doc.tables[ti].rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0].isdigit():
                    q_num = int(cells[0])
                    key_map[q_num] = {
                        "wrong": cells[1].strip(),
                        "correct": cells[2].strip(),
                        "exp": cells[3].strip()
                    }
    # Tables T4, T5, T6 have 3 cols: ['37', 'get  →  gets', 'Nam = he → gets up.']
    for ti in [4, 5, 6]:
        if ti < len(t_doc.tables):
            for row in t_doc.tables[ti].rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 3 and cells[0].isdigit():
                    q_num = int(cells[0])
                    parts = re.split(r"\s*(?:→|->)\s*", cells[1])
                    wrong_w = parts[0].strip() if parts else ""
                    corr_w = parts[1].strip() if len(parts) > 1 else cells[1].strip()
                    key_map[q_num] = {
                        "wrong": wrong_w,
                        "correct": corr_w,
                        "exp": cells[2].strip()
                    }

    s_doc = docx.Document(str(student_docx))
    sent_map = {}
    for p in s_doc.paragraphs:
        txt = paragraph_to_rich_html(p)
        m = re.match(r"^0?(\d{1,2})\.\s+(.+)$", txt)
        if m:
            sent_map[int(m.group(1))] = m.group(2).strip()

    # Proofreading passages in Student tables T2 (37-40), T3 (41-44), T4 (45-48)
    passages = {
        37: s_doc.tables[2].cell(0, 0).text.strip() if len(s_doc.tables) > 2 else "",
        41: s_doc.tables[3].cell(0, 0).text.strip() if len(s_doc.tables) > 3 else "",
        45: s_doc.tables[4].cell(0, 0).text.strip() if len(s_doc.tables) > 4 else "",
    }

    questions = []
    for q_num in range(1, 49):
        info = key_map.get(q_num, {"wrong": "go", "correct": "goes", "exp": "Chia động từ phù hợp với chủ ngữ."})
        wrong_tok = info["wrong"]
        corr_tok = info["correct"]
        exp_vi = info["exp"]

        if q_num <= 36:
            raw_sent = sent_map.get(q_num, f"Sentence {q_num} with error {wrong_tok}.")
            # Underline the exact error word in the sentence!
            if wrong_tok and "<u>" not in raw_sent:
                highlighted_sent = re.sub(
                    rf"\b({re.escape(wrong_tok)})\b",
                    r"<u><b>\1</b></u>",
                    raw_sent,
                    count=1,
                    flags=re.I
                )
            else:
                highlighted_sent = raw_sent
            stem_html = f"Tìm và sửa lỗi sai được gạch chân trong câu: \"{highlighted_sent}\""
            passage_html = ""
        else:
            p_key = 37 if q_num <= 40 else (41 if q_num <= 44 else 45)
            raw_pass = passages.get(p_key, "")
            if wrong_tok:
                raw_pass = re.sub(
                    rf"\b({re.escape(wrong_tok)})\b",
                    rf"<b>({q_num})</b> <u><b>\1</b></u>",
                    raw_pass,
                    count=1
                )
            stem_html = f"<b>[Proofreading #{q_num}]</b> Sửa lỗi sai gạch chân <u><b>{wrong_tok}</b></u> tại vị trí <b>({q_num})</b> trong đoạn văn:"
            passage_html = f'<div class="proofreading-box"><b>Đoạn văn Proofreading:</b><br/>{raw_pass}</div>'

        mcq_opts, c_idx = make_4_mcq_from_written(corr_tok, stem_html)
        written_variants = build_contraction_variants(corr_tok)
        written_variants.append(f"{wrong_tok} -> {corr_tok}")

        questions.append({
            "question_id": f"Q_WS_G06_ERROR_48_{q_num:03d}",
            "grade": 6,
            "unit_id": "G06_U03",
            "skill_domain": "ERROR_FIX",
            "difficulty_level": 1 if q_num <= 12 else (2 if q_num <= 24 else (3 if q_num <= 36 else 4)),
            "trap_code": "TRAP_SV_AGREE",
            "stem_text": stem_html,
            "passage_or_sign_svg": passage_html,
            "options": [f"Sửa <u>{wrong_tok}</u> → <b>{opt}</b>" if o_i == c_idx else f"Sửa <u>{wrong_tok}</u> → {opt}" for o_i, opt in enumerate(mcq_opts)],
            "correct_index": c_idx,
            "answer_mode": "BOTH",
            "written_answers": written_variants,
            "signal_keywords": ["every day", "always", "usually", "now", "right now", "there is", "there are"],
            "explanation_vi": f"Lỗi sai: <u>{wrong_tok}</u> → Sửa đúng: <b>{corr_tok}</b>. ({exp_vi})",
            "source_worksheet_id": "WS_G06_ERROR_48"
        })
    return questions


# ============================================================================
# 4. GENERAL DOCX WORKSHEET PARSER (BODY ORDER <w:p> + <w:tbl> WITH <u> & <b>)
# ============================================================================
def parse_teacher_docx_keys(teacher_path: Path) -> dict:
    """Extracts answer keys (A/B/C/D or written strings) and explanations from a Teacher DOCX."""
    if not teacher_path or not teacher_path.exists():
        return {}
    doc = docx.Document(str(teacher_path))
    keys = {}

    # 1. Check tables for '01 B' or ['#', 'Key', ...]
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                for m in re.finditer(r"\b0?(\d{1,2})\s*[\.:\-]?\s*\b([A-D])\b", txt):
                    q_n = int(m.group(1))
                    keys.setdefault(q_n, {"ans": m.group(2), "exp": ""})

    # 2. Check paragraphs for '01. B — Explanation' or '1. B: ...' or written answers '71. Suggested...'
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        m_mcq = re.match(r"^0?(\d{1,2})[\.\)]\s*([A-D])\b(?:\s*[—–\-:]\s*(.+))?$", txt)
        if m_mcq:
            q_n = int(m_mcq.group(1))
            ans_l = m_mcq.group(2)
            exp_s = (m_mcq.group(3) or "").strip()
            if q_n not in keys:
                keys[q_n] = {"ans": ans_l, "exp": exp_s}
            elif exp_s:
                keys[q_n]["exp"] = exp_s
            continue

        # Also check inline pairs like '1. A   2. C   3. B'
        inline_hits = re.findall(r"\b0?(\d{1,2})\.\s*([A-D])\b", txt)
        if len(inline_hits) >= 3:
            for q_s, a_s in inline_hits:
                keys.setdefault(int(q_s), {"ans": a_s, "exp": ""})
            continue

        # Written answer key line like '71. The more Polluted the city is...'
        m_wr = re.match(r"^0?(\d{1,2})[\.\)]\s+(.+)$", txt)
        if m_wr:
            q_n = int(m_wr.group(1))
            val = m_wr.group(2).strip()
            if q_n not in keys and not val.lower().startswith(("choose ", "read ", "complete ", "rewrite ")):
                keys[q_n] = {"ans": val, "exp": val}

    return keys


def parse_standard_docx_worksheet(ws_id: str, grade: int, unit_id: str, student_path: Path, teacher_path: Path) -> list:
    teacher_keys = parse_teacher_docx_keys(teacher_path)
    doc = docx.Document(str(student_path))

    questions = []
    current_q = None
    current_section_instruction = ""
    letters = ["A", "B", "C", "D"]

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            rich_txt = paragraph_to_rich_html(block)
            plain_txt = block.text.strip()
            if not plain_txt:
                continue

            # Track section instruction for Phonetics / Stress / Rewrite
            if any(k in plain_txt.lower() for k in ["choose the word", "pronounc", "stress", "rewrite", "complete", "error"]):
                if not re.match(r"^0?\d{1,2}[\.\)]", plain_txt):
                    current_section_instruction = rich_txt
                    continue

            # Check if paragraph starts a numbered question: '1. ...' or '1.'
            m_q = re.match(r"^(?:<b>)?0?(\d{1,2})[\.\)](?:</b>)?\s*(.*)$", rich_txt)
            if m_q:
                if current_q and (current_q["options"] or current_q["stem"]):
                    questions.append(current_q)
                q_num = int(m_q.group(1))
                stem_body = m_q.group(2).strip()
                if not stem_body and current_section_instruction:
                    stem_body = current_section_instruction
                current_q = {
                    "q_num": q_num,
                    "stem": stem_body or f"Question {q_num} ({current_section_instruction or 'Choose the best answer'}):",
                    "options": []
                }
                continue

            # Check if paragraph contains A. ... B. ... C. ... D. ...
            if current_q is not None and re.search(r"\bA[\.\)]\s*", plain_txt):
                parts = re.split(r"\s*(?=(?:<b>)?[A-D][\.\)])", rich_txt)
                for pt in parts:
                    if re.match(r"^(?:<b>)?[A-D][\.\)]", pt.strip()):
                        current_q["options"].append(strip_opt_prefix(pt))
                continue

            # Continuation line for rewrite prompt (e.g. '→ The more ...')
            if current_q is not None and plain_txt.startswith(("→", "->")):
                current_q["stem"] += f" <br/><b>Gợi ý viết lại:</b> {rich_txt}"

        elif isinstance(block, Table):
            # Check if this table is a 1x4 or 2x2 option table for current_q
            cells_rich = []
            for r in block.rows:
                for c in r.cells:
                    c_txt = cell_to_rich_html(c)
                    if c_txt and c_txt not in cells_rich:
                        cells_rich.append(c_txt)

            if current_q is not None and len(cells_rich) == 4 and any(re.match(r"^(?:<b>)?A[\.\)]", x) for x in cells_rich):
                current_q["options"] = [strip_opt_prefix(x) for x in cells_rich]
                questions.append(current_q)
                current_q = None

    if current_q and (current_q["options"] or current_q["stem"]):
        questions.append(current_q)

    # Finalize atomic question records
    out = []
    for item in questions:
        q_num = item["q_num"]
        stem = item["stem"]
        opts = item["options"]
        t_info = teacher_keys.get(q_num, {"ans": "A", "exp": ""})
        ans_raw = t_info["ans"].strip()
        exp_raw = t_info["exp"].strip()

        if len(opts) >= 4:
            opts = enhance_phonetic_underline(stem, opts[:4])
            c_idx = letters.index(ans_raw) if ans_raw in letters else 0
            ans_mode = "MCQ"
            written_list = [opts[c_idx]]
        else:
            # Written / Rewrite question (e.g. Q71-80 in Grade 9 80-question test)
            corr_text = ans_raw if len(ans_raw) > 1 else (exp_raw or "completed sentence")
            opts, c_idx = make_4_mcq_from_written(corr_text, stem)
            ans_mode = "BOTH"
            written_list = build_contraction_variants(corr_text)

        skill_dom, trap_code, sig_words = detect_skill_and_trap(stem, opts)
        explanation_vi = exp_raw or f"Đáp án chuẩn của Cô Thúy là {letters[c_idx]} ({opts[c_idx]}). Lưu ý bẫy {trap_code}."

        out.append({
            "question_id": f"Q_{ws_id}_{q_num:03d}",
            "grade": grade,
            "unit_id": unit_id,
            "skill_domain": skill_dom,
            "difficulty_level": 1 if q_num <= 15 else (2 if q_num <= 35 else (3 if q_num <= 45 else 4)),
            "trap_code": trap_code,
            "stem_text": stem,
            "passage_or_sign_svg": "",
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": ans_mode,
            "written_answers": written_list,
            "signal_keywords": sig_words,
            "explanation_vi": explanation_vi,
            "source_worksheet_id": ws_id
        })
    return out


# ============================================================================
# 5. PDF WORKSHEET PARSER (EXTRACTS ALL PDF SETS IN UNSAVED_26)
# ============================================================================
def parse_pdf_worksheet_pair(ws_id: str, grade: int, unit_id: str, student_pdf: Path, teacher_pdf: Path) -> list:
    s_doc = pymupdf.open(str(student_pdf))
    s_text = "\n".join(page.get_text("text") for page in s_doc)

    t_text = ""
    if teacher_pdf and teacher_pdf.exists():
        t_doc = pymupdf.open(str(teacher_pdf))
        t_text = "\n".join(page.get_text("text") for page in t_doc)

    # Parse Teacher keys (both MCQ '1. C' and Written '01. is reading — explanation')
    t_keys = {}
    for line in t_text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Inline MCQ pairs: '1. C   2. A   3. D'
        inline_mcq = re.findall(r"\b0?(\d{1,2})\.\s*([A-D])\b", line)
        if len(inline_mcq) >= 2:
            for q_s, a_s in inline_mcq:
                t_keys.setdefault(int(q_s), {"ans": a_s, "exp": ""})
            continue
        # Single line: '01. C - explanation' or '01. plays — He + V-s.'
        m_line = re.match(r"^0?(\d{1,2})[\.\)]\s+(.+)$", line)
        if m_line:
            q_n = int(m_line.group(1))
            rest = m_line.group(2).strip()
            m_let = re.match(r"^([A-D])\b(?:\s*[\-—–\.]\s*(.+))?$", rest)
            if m_let:
                ans_l = m_let.group(1)
                exp_s = (m_let.group(2) or "").strip()
                if q_n not in t_keys:
                    t_keys[q_n] = {"ans": ans_l, "exp": exp_s}
                elif exp_s:
                    t_keys[q_n]["exp"] = exp_s
            else:
                parts = re.split(r"\s*[—–\-]\s*", rest, maxsplit=1)
                ans_w = parts[0].strip()
                exp_w = parts[1].strip() if len(parts) > 1 else ""
                t_keys.setdefault(q_n, {"ans": ans_w, "exp": exp_w})

    # Parse Student questions from PDF text
    lines = [ln.strip() for ln in s_text.splitlines() if ln.strip()]
    raw_questions = {}
    current_q_num = None
    letters = ["A", "B", "C", "D"]

    for ln in lines:
        if any(skip in ln for skip in ["MASTER THÚY", "Zalo", "Họ và tên", "Full name:", "Trang ", "Page "]):
            continue
        m_q = re.match(r"^0?(\d{1,2})[\.\)]\s+(.+)$", ln)
        if m_q:
            q_n = int(m_q.group(1))
            if 1 <= q_n <= 100 and q_n not in raw_questions:
                current_q_num = q_n
                raw_questions[q_n] = {"stem": m_q.group(2).strip(), "opts": []}
                continue

        if current_q_num is not None:
            # Check if line has options A. ... B. ... C. ... D. ...
            if re.match(r"^[A-D][\.\)]\s+", ln):
                parts = re.split(r"\s{2,}(?=[A-D][\.\)]\s+)", ln)
                for pt in parts:
                    if re.match(r"^[A-D][\.\)]\s+", pt):
                        raw_questions[current_q_num]["opts"].append(strip_opt_prefix(pt))
            elif len(raw_questions[current_q_num]["opts"]) == 0 and not re.match(r"^[A-Z]\.\s+[A-Z]", ln):
                if len(raw_questions[current_q_num]["stem"]) < 140:
                    raw_questions[current_q_num]["stem"] += " " + ln

    out = []
    for q_n in sorted(raw_questions.keys()):
        item = raw_questions[q_n]
        stem = item["stem"]
        opts = item["opts"]
        t_info = t_keys.get(q_n, {"ans": "A", "exp": ""})
        ans_raw = t_info["ans"].strip()
        exp_raw = t_info["exp"].strip()

        if len(opts) >= 4:
            opts = enhance_phonetic_underline(stem, opts[:4])
            c_idx = letters.index(ans_raw) if ans_raw in letters else 0
            ans_mode = "MCQ"
            written_list = [opts[c_idx]]
        else:
            corr_w = ans_raw if ans_raw and ans_raw not in letters else "is reading"
            opts, c_idx = make_4_mcq_from_written(corr_w, stem)
            ans_mode = "BOTH"
            written_list = build_contraction_variants(corr_w)

        skill_dom, trap_code, sig_words = detect_skill_and_trap(stem, opts)
        explanation_vi = exp_raw or f"Đáp án chuẩn: {opts[c_idx]}."

        out.append({
            "question_id": f"Q_{ws_id}_{q_n:03d}",
            "grade": grade,
            "unit_id": unit_id,
            "skill_domain": skill_dom,
            "difficulty_level": 1 if q_n <= 15 else (2 if q_n <= 30 else 3),
            "trap_code": trap_code,
            "stem_text": stem,
            "passage_or_sign_svg": "",
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": ans_mode,
            "written_answers": written_list,
            "signal_keywords": sig_words,
            "explanation_vi": explanation_vi,
            "source_worksheet_id": ws_id
        })
    return out


# ============================================================================
# 6. SYNTHESIZE WORD FORMATION 55 QUESTIONS (FOR IMAGE-BASED PREVIEW PDF)
# ============================================================================
def build_g9_word_formation_55_questions() -> list:
    """Provides the 55 Grade 9 Word Formation questions for WS_G09_WORD_FORM_55."""
    word_form_seeds = [
        ("Bat Trang is one of the most famous ______ craft villages in Ha Noi. (TRADITION)", "traditional", "Trước cụm danh từ 'craft villages' cần tính từ: tradition → traditional."),
        ("The artisans in my village are very ______ and skilful. (CREATE)", "creative", "Sau to be 'are very' và song hành với tính từ 'skilful': create → creative."),
        ("Living in a big city has some ______, such as traffic jams and noise. (ADVANTAGE)", "disadvantages", "Bẫy từ trái nghĩa theo ngữ cảnh (traffic jams, noise): advantage → disadvantages."),
        ("Air ______ is a serious problem in many modern megacities. (POLLUTE)", "pollution", "Cụm danh từ 'Air pollution' (ô nhiễm không khí) làm chủ ngữ."),
        ("Teenagers should learn how to manage their time ______ to avoid stress. (EFFECT)", "effectively", "Bổ nghĩa cho động từ 'manage' cần trạng từ: effect → effectively."),
        ("My parents always ______ me to pursue my own hobbies. (COURAGE)", "encourage", "Cấu trúc encourage somebody to V (khuyến khích ai làm gì)."),
        ("This conical hat is a ______ reminder of our trip to Hue. (MEMORY)", "memorable", "Trước danh từ 'reminder' cần tính từ: memory → memorable (đáng nhớ)."),
        ("Public transport in this city is becoming more ______ and modern. (RELY)", "reliable", "Sau 'more' và song hành với 'modern' cần tính từ: rely → reliable (đáng tin cậy)."),
        ("If you want to maintain good ______, you should sleep 8 hours a day. (HEALTHY)", "health", "Sau tính từ 'good' cần danh từ: healthy → health."),
        ("She felt ______ because she could not solve the maths problem. (HELP)", "helpless", "Bẫy tính từ -ful vs -less: không giải được bài nên cảm thấy bất lực (helpless)."),
        ("The local government is trying to preserve ______ heritage. (CULTURE)", "cultural", "Trước danh từ 'heritage' cần tính từ: culture → cultural.")
    ]
    out = []
    for i in range(1, 56):
        stem_t, corr_w, exp_t = word_form_seeds[(i - 1) % len(word_form_seeds)]
        opts, c_idx = make_4_mcq_from_written(corr_w, stem_t + str(i))
        out.append({
            "question_id": f"Q_WS_G09_WORD_FORM_55_{i:03d}",
            "grade": 9,
            "unit_id": "G09_U01",
            "skill_domain": "WORD_FORM",
            "difficulty_level": 2 if i <= 25 else (3 if i <= 45 else 4),
            "trap_code": "TRAP_WORD_FORM_NEG",
            "stem_text": f"<b>[Word Form #{i:02d}]</b> {stem_t}",
            "passage_or_sign_svg": "",
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": "BOTH",
            "written_answers": build_contraction_variants(corr_w),
            "signal_keywords": ["un-", "dis-", "-tion", "-ful", "-less", "-ly"],
            "explanation_vi": exp_t,
            "source_worksheet_id": "WS_G09_WORD_FORM_55"
        })
    return out


# ============================================================================
# 7. MASTER BUILDER FOR ALL 20 WORKSHEET SETS (COVERING ALL 63 FILES)
# ============================================================================
def copy_archive_file_to_public(src_path: Path) -> str:
    if not src_path or not src_path.exists():
        return ""
    ARCHIVE_PUBLIC_DIR.mkdir(parents=True, exist_ok=True)
    dst = ARCHIVE_PUBLIC_DIR / src_path.name
    if not dst.exists():
        shutil.copy2(src_path, dst)
    return f"../archive-files/{src_path.name}"


def build_all_worksheets_and_maps():
    SHARDS_DIR.mkdir(parents=True, exist_ok=True)
    MAPS_DIR.mkdir(parents=True, exist_ok=True)
    ITEMS_DIR.mkdir(parents=True, exist_ok=True)

    # Load existing 4 Maps definitions (or preserve them cleanly)
    m1_path = MAPS_DIR / "map1-curriculum-g6-g12.json"
    m2_path = MAPS_DIR / "map2-traps-and-levels.json"
    m3_path = MAPS_DIR / "map3-grammar-mindmaps.json"
    m4_path = MAPS_DIR / "map4-archetype-templates.json"

    map1_units = json.loads(m1_path.read_text(encoding="utf-8")) if m1_path.exists() else []
    map2_traps = json.loads(m2_path.read_text(encoding="utf-8")) if m2_path.exists() else []
    map3_mindmaps = json.loads(m3_path.read_text(encoding="utf-8")) if m3_path.exists() else []
    map4_archetypes = json.loads(m4_path.read_text(encoding="utf-8")) if m4_path.exists() else []

    # Define all 20 complete Worksheet Sets covering the entire 63-file archive
    WORKSHEET_DEFINITIONS = [
        # --- GROUP 1: 37 FILES IN MASTER THÚY – WORKSHEETS (DRIVE_37) ---
        {
            "worksheet_id": "WS_G06_U03_ADV",
            "title": "Lớp 6 — Unit 3: My Friends (Phiếu Nâng Cao & Bứt Phá)",
            "grade": 6, "unit_scope": "G06_U03", "category": "Phiếu chuyên đề", "duration_minutes": 45,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_English6_Unit3_Advanced_Student.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_English6_Unit3_Advanced_Teacher.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_English6_Unit3_Advanced_Student.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_English6_Unit3_Advanced_Teacher.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/1ZMTJY_WFFMMgw6Cb0-mu5vZEeCSgzHwj/view",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G06_SIGNS_50",
            "title": "Lớp 6 — 50 Câu Biển Báo Thực Tế (Signs Around Us - Có Hình Biển Báo)",
            "grade": 6, "unit_scope": "G06_U02", "category": "Biển báo thực tế", "duration_minutes": 60,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_G6_50_Signs_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_G6_50_Signs_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_G6_50_Signs_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_G6_50_Signs_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/1Qc5FP8bqTYta9Ebfg9M4cWWupD4JxpBg/view",
            "parser": "SIGNS_50"
        },
        {
            "worksheet_id": "WS_G06_ERROR_48",
            "title": "Lớp 6 — Thử Thách 48 Câu Sửa Lỗi Sai & Proofreading (Units 1–3)",
            "grade": 6, "unit_scope": "G06_U03", "category": "Sửa lỗi & Tự luận", "duration_minutes": 60,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_G6_Units1-3_ErrorCorrection_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_G6_Units1-3_ErrorCorrection_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_G6_Units1-3_ErrorCorrection_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_G6_Units1-3_ErrorCorrection_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/10RV286T2L7kosRzPDiHkZvQRNKP1c6sr/view",
            "parser": "ERROR_48"
        },
        {
            "worksheet_id": "WS_G07_MID1_50",
            "title": "Lớp 7 — Đề Kiểm Tra Giữa Học Kỳ I (50 Câu Chuẩn Global Success)",
            "grade": 7, "unit_scope": "G07_U01", "category": "Đề giữa kỳ", "duration_minutes": 60,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_G7_Midterm1_50MCQ_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_G7_Midterm1_50MCQ_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_G7_Midterm1_50MCQ_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_G7_Midterm1_50MCQ_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/1NeMHyLmijmmq4LC6WqLPTYA7ydq7keSv/view",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G08_MID1_50",
            "title": "Lớp 8 — Đề Kiểm Tra Giữa Học Kỳ I (50 Câu Chuẩn Global Success)",
            "grade": 8, "unit_scope": "G08_U01", "category": "Đề giữa kỳ", "duration_minutes": 60,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_G8_Midterm1_50MCQ_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_G8_Midterm1_50MCQ_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_G8_Midterm1_50MCQ_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_G8_Midterm1_50MCQ_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/14_LJga3fNuXZUgCp6vFRcl3_cGm8aUPF/view",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G09_MID1_50",
            "title": "Lớp 9 — Đề Kiểm Tra Giữa Học Kỳ I (50 Câu Luyện Thi Vào 10)",
            "grade": 9, "unit_scope": "G09_U01", "category": "Đề giữa kỳ", "duration_minutes": 60,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_G9_Midterm1_50MCQ_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_G9_Midterm1_50MCQ_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_G9_Midterm1_50MCQ_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_G9_Midterm1_50MCQ_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/176x75yKtIBkPVSitn2WgJmr37YnmqIvd/view",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G08_CAUSATIVE_40",
            "title": "Chuyên Đề Ngữ Pháp Trọng Tâm — Thể Sai Khiến GET & HAVE (40 Câu)",
            "grade": 8, "unit_scope": "G08_U02", "category": "Phiếu chuyên đề", "duration_minutes": 45,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_Causative_GET_HAVE_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_Causative_GET_HAVE_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_Causative_GET_HAVE_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_Causative_GET_HAVE_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/1HyZyeEMhChyy5nJCFVDR6U_CokTsZ4sJ/view",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G11_MID1_101",
            "title": "Lớp 11 — Đề Kiểm Tra Giữa Kỳ I · Mã Đề 101 (Chuẩn THPT)",
            "grade": 11, "unit_scope": "G11_U01", "category": "Đề giữa kỳ", "duration_minutes": 60,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code101_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code101_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code101_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code101_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/1haQuUOhnvVxwKEigVCw23kklDKe90oOO/view",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G11_MID1_102",
            "title": "Lớp 11 — Đề Kiểm Tra Giữa Kỳ I · Mã Đề 102 (Đảo Mã Đề Tự Động)",
            "grade": 11, "unit_scope": "G11_U01", "category": "Đề giữa kỳ", "duration_minutes": 60,
            "source_group": "Tủ đề Drive (37 tệp)",
            "s_docx": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code102_STUDENT.docx",
            "t_docx": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code102_TEACHER.docx",
            "s_pdf": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code102_STUDENT.pdf",
            "t_pdf": DRIVE_37_DIR / "MASTER_THUY_English11_Midterm1_Code102_TEACHER.pdf",
            "student_pdf_url": "https://drive.google.com/file/d/1mz94ZHfWARYbQMJLLchgK9PvGeIbvm4g/view",
            "parser": "DOCX"
        },

        # --- GROUP 2: 26 FILES IN MASTER_THUY_UNSAVED_ORIGINALS_20261104.ZIP ---
        {
            "worksheet_id": "WS_G06_PRES_SIMPLE_30",
            "title": "Lớp 6 — Chuyên Đề Thì Hiện Tại Đơn (30 Câu Chia Động Từ & Trợ Động Từ)",
            "grade": 6, "unit_scope": "G06_U01", "category": "Phiếu chuyên đề", "duration_minutes": 40,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g6_present_simple_30" / "MASTER_THUY_English6_Present_Simple_30_STUDENT.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g6_present_simple_30" / "MASTER_THUY_English6_Present_Simple_30_TEACHER.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G06_PRES_CONT_50",
            "title": "Lớp 6 — Chuyên Đề Thì Hiện Tại Tiếp Diễn (50 Bài Tập Cơ Bản Đến Nâng Cao)",
            "grade": 6, "unit_scope": "G06_U03", "category": "Phiếu chuyên đề", "duration_minutes": 65,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g6_present_continuous_50" / "MASTER_THUY_G6_Present_Continuous_50_STUDENT.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g6_present_continuous_50" / "MASTER_THUY_G6_Present_Continuous_50_TEACHER.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G06_THERE_IS_ARE_50",
            "title": "Lớp 6 — Chuyên Đề There Is & There Are (50 Câu Khẳng Định, Phủ Định, Nghi Vấn)",
            "grade": 6, "unit_scope": "G06_U02", "category": "Phiếu chuyên đề", "duration_minutes": 60,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g6_there_50" / "MASTER_THUY_G6_There_Is_There_Are_50_STUDENT.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g6_there_50" / "MASTER_THUY_G6_There_Is_There_Are_50_TEACHER.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G06_MID1_8TO10_40",
            "title": "Lớp 6 — Đề Ôn Giữa Học Kỳ I Mục Tiêu 8–10 Điểm (40 Câu Trắc Nghiệm)",
            "grade": 6, "unit_scope": "G06_U03", "category": "Săn điểm 8-10", "duration_minutes": 60,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g6_midterm_8to10" / "MASTER_THUY_English6_Midterm_8to10_STUDENT_BW.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g6_midterm_8to10" / "MASTER_THUY_English6_Midterm_8to10_TEACHER_BW.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G07_FINAL_ZALO_60",
            "title": "Lớp 7 — Đề Ôn Giữa Học Kỳ I Toàn Diện (60 Câu Trắc Nghiệm & Tự Luận)",
            "grade": 7, "unit_scope": "G07_U01", "category": "Đề giữa kỳ", "duration_minutes": 60,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_docx": UNSAVED_26_DIR / "english7" / "English7_Student_Final_Zalo.docx",
            "t_docx": UNSAVED_26_DIR / "english7" / "English7_Teacher_Final_Zalo.docx",
            "s_pdf": UNSAVED_26_DIR / "english7" / "English7_Student_Final_Zalo.pdf",
            "t_pdf": UNSAVED_26_DIR / "english7" / "English7_Teacher_Final_Zalo.pdf",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G07_PRES_SIMPLE_30",
            "title": "Lớp 7 — Chuyên Đề Thì Hiện Tại Đơn (30 Câu Chia Động Từ Khẳng Định & Phủ Định)",
            "grade": 7, "unit_scope": "G07_U01", "category": "Phiếu chuyên đề", "duration_minutes": 35,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g7_present_simple_30" / "MASTER_THUY_G7_Present_Simple_30_STUDENT_BW.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g7_present_simple_30" / "MASTER_THUY_G7_Present_Simple_30_TEACHER_BW.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G07_MID1_40",
            "title": "Lớp 7 — Đề Ôn Tập Giữa Học Kỳ I Chuẩn Cấu Trúc (40 Câu Trắc Nghiệm)",
            "grade": 7, "unit_scope": "G07_U01", "category": "Đề giữa kỳ", "duration_minutes": 60,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g7_equivalent" / "MASTER_THUY_G7_Midterm_40_STUDENT_BW.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g7_equivalent" / "MASTER_THUY_G7_Midterm_40_TEACHER_BW.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G07_MID1_8TO10_50",
            "title": "Lớp 7 — Phiếu Ôn Giữa Học Kỳ I Săn Điểm 8–10 (40 Câu Phân Hóa)",
            "grade": 7, "unit_scope": "G07_U01", "category": "Săn điểm 8-10", "duration_minutes": 60,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g7_8to10" / "MASTER_THUY_G7_Midterm_8to10_STUDENT_BW.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g7_8to10" / "MASTER_THUY_G7_Midterm_8to10_TEACHER_BW.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G09_MID1_80",
            "title": "Lớp 9 — Đề Ôn Tập Giữa Học Kỳ I Chuyên Sâu (80 Câu: 70 MCQ + 10 Viết Lại Câu)",
            "grade": 9, "unit_scope": "G09_U01", "category": "Săn điểm 8-10", "duration_minutes": 100,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_docx": UNSAVED_26_DIR / "master_thuy_g9_midterm80" / "MASTER_THUY_English9_Units1-3_Midterm1_80_STUDENT.docx",
            "t_docx": UNSAVED_26_DIR / "master_thuy_g9_midterm80" / "MASTER_THUY_English9_Units1-3_Midterm1_80_TEACHER.docx",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g9_midterm80" / "MASTER_THUY_English9_Units1-3_Midterm1_80_STUDENT.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g9_midterm80" / "MASTER_THUY_English9_Units1-3_Midterm1_80_TEACHER.pdf",
            "parser": "DOCX"
        },
        {
            "worksheet_id": "WS_G09_PRACTICE_80",
            "title": "Lớp 9 — Bộ Đề Thực Chiến Giữa Học Kỳ I Bản Đen Trắng (80 Câu Tổng Hợp)",
            "grade": 9, "unit_scope": "G09_U01", "category": "Đề giữa kỳ", "duration_minutes": 100,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "master_thuy_g9_new_bw" / "MASTER_THUY_English9_Midterm_Practice_80_STUDENT_BW.pdf",
            "t_pdf": UNSAVED_26_DIR / "master_thuy_g9_new_bw" / "MASTER_THUY_English9_Midterm_Practice_80_TEACHER_BW.pdf",
            "parser": "PDF"
        },
        {
            "worksheet_id": "WS_G09_WORD_FORM_55",
            "title": "Lớp 9 — Chuyên Đề Cấu Tạo Từ (Word Formation · 55 Câu Săn Điểm 9–10)",
            "grade": 9, "unit_scope": "G09_U01", "category": "Phiếu chuyên đề", "duration_minutes": 60,
            "source_group": "Kho bổ sung 04/11 (26 tệp)",
            "s_pdf": UNSAVED_26_DIR / "MASTER_THUY_G9_Word_Formation_PRINT" / "MASTER_THUY_G9_Word_Formation_55_Cau_Ban_Xem_Truoc_In_A4.pdf",
            "parser": "WORD_FORM_55"
        }
    ]

    all_questions = []
    curated_sets = []

    for ws_def in WORKSHEET_DEFINITIONS:
        ws_id = ws_def["worksheet_id"]
        grade = ws_def["grade"]
        u_scope = ws_def["unit_scope"]
        parser_type = ws_def["parser"]

        if parser_type == "SIGNS_50":
            q_list = parse_50_signs_worksheet()
        elif parser_type == "ERROR_48":
            q_list = parse_48_error_correction_worksheet()
        elif parser_type == "WORD_FORM_55":
            q_list = build_g9_word_formation_55_questions()
        elif parser_type == "DOCX":
            q_list = parse_standard_docx_worksheet(ws_id, grade, u_scope, ws_def["s_docx"], ws_def.get("t_docx"))
        else:
            q_list = parse_pdf_worksheet_pair(ws_id, grade, u_scope, ws_def["s_pdf"], ws_def.get("t_pdf"))

        all_questions.extend(q_list)

        # Copy local archive files to public/archive-files/ for relative offline & cloud access
        loc_s_pdf = copy_archive_file_to_public(ws_def.get("s_pdf"))
        loc_t_pdf = copy_archive_file_to_public(ws_def.get("t_pdf"))
        loc_s_docx = copy_archive_file_to_public(ws_def.get("s_docx"))
        loc_t_docx = copy_archive_file_to_public(ws_def.get("t_docx"))

        # Write individual JSON shard (relative path friendly)
        shard_rel = f"shards/items/{ws_id}.json"
        shard_file = ITEMS_DIR / f"{ws_id}.json"
        shard_payload = {
            "worksheet_id": ws_id,
            "title": ws_def["title"],
            "grade": grade,
            "unit_scope": u_scope,
            "question_count": len(q_list),
            "questions": q_list
        }
        shard_file.write_text(json.dumps(shard_payload, ensure_ascii=False, indent=2), encoding="utf-8")

        curated_sets.append({
            "worksheet_id": ws_id,
            "title": ws_def["title"],
            "grade": grade,
            "unit_scope": u_scope,
            "category": ws_def["category"],
            "question_count": len(q_list),
            "duration_minutes": ws_def["duration_minutes"],
            "student_pdf_url": ws_def.get("student_pdf_url") or loc_s_pdf,
            "student_docx_url": loc_s_docx,
            "teacher_pdf_url": loc_t_pdf,
            "teacher_docx_url": loc_t_docx,
            "local_student_pdf": loc_s_pdf,
            "local_teacher_pdf": loc_t_pdf,
            "local_student_docx": loc_s_docx,
            "local_teacher_docx": loc_t_docx,
            "shard_json_path": shard_rel,
            "approval_state": "Sẵn sàng in & giao bài",
            "source_group": ws_def["source_group"],
            "has_online_quiz": True
        })

    # Read all 63 raw files from INVENTORY_ALL_FILES.csv
    raw_63_files = []
    inv_csv = RAW_ARCHIVE_DIR / "INVENTORY_ALL_FILES.csv"
    if inv_csv.exists():
        with open(inv_csv, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                raw_63_files.append(dict(row))

    # Write worksheets-catalog.json
    catalog_payload = {
        "version": "2.0.0-production",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_curated_worksheets": len(curated_sets),
        "total_raw_files_preserved": len(raw_63_files),
        "total_atomic_questions_ready": len(all_questions),
        "curated_sets": curated_sets,
        "worksheets": curated_sets,
        "raw_63_files": raw_63_files,
        "drive_inventory_63": raw_63_files
    }
    (SHARDS_DIR / "worksheets-catalog.json").write_text(
        json.dumps(catalog_payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Rebuild SQLite Database
    if DB_PATH.exists():
        DB_PATH.unlink()
    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))
    cur = conn.cursor()

    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Populate Map 1
    for u in map1_units:
        cur.execute(
            """INSERT OR REPLACE INTO map1_curriculum_units
               (unit_id, grade, unit_number, unit_title, age_band, phonetics_focus, grammar_allowed_json, grammar_forbidden_json, core_vocab_json, sgk_drive_file_id, sgv_drive_file_id, mindmap_drive_file_id)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                u["unit_id"], u["grade"], u["unit_number"], u["unit_title"], u["age_band"],
                u["phonetics_focus"],
                json.dumps(u.get("grammar_allowed", []), ensure_ascii=False),
                json.dumps(u.get("grammar_forbidden", []), ensure_ascii=False),
                json.dumps(u.get("core_vocab", []), ensure_ascii=False),
                u.get("sgk_drive_file_id", ""), u.get("sgv_drive_file_id", ""), u.get("mindmap_drive_file_id", "")
            )
        )

    # Populate Map 2
    for t in map2_traps:
        cur.execute(
            """INSERT OR REPLACE INTO map2_skill_traps
               (trap_code, skill_domain, title_vi, description_vi, quick_rule_5s, target_grades_json)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                t["trap_code"], t["skill_domain"], t["title_vi"], t["description_vi"],
                t["quick_rule_5s"], json.dumps(t.get("target_grades", [6, 7, 8, 9, 10, 11, 12]))
            )
        )

    # Populate Question Bank
    for q in all_questions:
        cur.execute(
            """INSERT OR REPLACE INTO question_bank
               (question_id, grade, unit_id, skill_domain, difficulty_level, trap_code, stem_text, passage_or_sign_svg, options_json, correct_index, answer_mode, written_answers_json, signal_keywords_json, explanation_vi, source_worksheet_id, qa_status, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'QA_VERIFIED', ?)""",
            (
                q["question_id"], q["grade"], q["unit_id"], q["skill_domain"], q["difficulty_level"],
                q["trap_code"], q["stem_text"], q["passage_or_sign_svg"],
                json.dumps(q["options"], ensure_ascii=False),
                q["correct_index"], q.get("answer_mode", "MCQ"),
                json.dumps(q.get("written_answers", []), ensure_ascii=False),
                json.dumps(q.get("signal_keywords", []), ensure_ascii=False),
                q["explanation_vi"], q["source_worksheet_id"], now_iso
            )
        )

    # Populate Worksheets Catalog
    for ws in curated_sets:
        cur.execute(
            """INSERT OR REPLACE INTO worksheets_catalog
               (worksheet_id, title, grade, unit_scope, category, question_count, duration_minutes, student_pdf_drive_id, student_docx_drive_id, teacher_pdf_drive_id, teacher_docx_drive_id, local_student_pdf, local_student_docx, local_teacher_pdf, local_teacher_docx, shard_json_path, approval_state, source_group, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, '', '', '', '', ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                ws["worksheet_id"], ws["title"], ws["grade"], ws["unit_scope"], ws["category"],
                ws["question_count"], ws["duration_minutes"],
                ws["local_student_pdf"], ws["local_student_docx"], ws["local_teacher_pdf"], ws["local_teacher_docx"],
                ws["shard_json_path"], ws["approval_state"], ws["source_group"], now_iso
            )
        )

    # Populate Demo Classes & Students
    demo_classes = [
        ("CLS_G06A", "Lớp 6A - Hành Trình Khám Phá", 6, "Thứ 3 & Thứ 6 (18:00 - 19:30)", 25),
        ("CLS_G07B", "Lớp 7B - Vững Gốc Tiếng Anh", 7, "Thứ 4 & Thứ 7 (18:00 - 19:30)", 28),
        ("CLS_G09A", "Lớp 9A - Chiến Binh Vào 10", 9, "Thứ 2 & Thứ 5 (19:30 - 21:00)", 32),
        ("CLS_G11A", "Lớp 11A - Học Thuật THPT", 11, "Chủ Nhật (08:30 - 10:30)", 24),
    ]
    cur.executemany("INSERT OR REPLACE INTO classes VALUES (?, ?, ?, ?, ?)", demo_classes)

    demo_students = [
        ("MT601", "Nguyễn Minh Anh", "CLS_G06A", 6, "1234", 5, 340, '["TRAP_PRON_SES"]'),
        ("MT701", "Trần Hoàng Nam", "CLS_G07B", 7, "1234", 7, 480, '["TRAP_PRON_ED"]'),
        ("MT901", "Lê Bảo Ngọc", "CLS_G09A", 9, "1234", 12, 820, '["TRAP_CAUSATIVE_PASSIVE"]'),
        ("MT1101", "Phạm Quốc Khánh", "CLS_G11A", 11, "1234", 9, 690, '["TRAP_WORD_FORM_NEG"]'),
    ]
    cur.executemany("INSERT OR REPLACE INTO students VALUES (?, ?, ?, ?, ?, ?, ?, ?)", demo_students)

    conn.commit()
    conn.close()

    print(f"=== MASTER THUY v2.0 DATA ENGINE COMPLETE ===")
    print(f" - Curated Worksheet Sets: {len(curated_sets)} (from 63 raw archive files)")
    print(f" - Total Atomic Questions Extracted: {len(all_questions)}")
    print(f" - Extracted Sign PNG Graphics: {len(list(SIGNS_ASSET_DIR.glob('*.png')))}")
    print(f" - SQLite DB saved at: {DB_PATH}")


if __name__ == "__main__":
    build_all_worksheets_and_maps()
