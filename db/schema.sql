-- ============================================================================
-- MASTER THÚY — ENGLISH STUDIO v2.0 (PRODUCTION & CLOUD-READY SCHEMA)
-- Thiết kế chịu tải: 10.000+ Phiếu bài tập | 300.000+ Câu hỏi nguyên tử | 1.000+ Học sinh
-- ============================================================================

PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

-- 1. BẢNG BẢN ĐỒ 1 (MAP 1): KHUNG CHƯƠNG TRÌNH SGK GLOBAL SUCCESS (LỚP 6 -> LỚP 12)
CREATE TABLE IF NOT EXISTS map1_curriculum_units (
    unit_id TEXT PRIMARY KEY,
    grade INTEGER NOT NULL,
    unit_number INTEGER NOT NULL,
    unit_title TEXT NOT NULL,
    age_band TEXT NOT NULL,
    phonetics_focus TEXT NOT NULL,
    grammar_allowed_json TEXT NOT NULL,
    grammar_forbidden_json TEXT NOT NULL,
    core_vocab_json TEXT NOT NULL,
    sgk_drive_file_id TEXT,
    sgv_drive_file_id TEXT,
    mindmap_drive_file_id TEXT
);
CREATE INDEX IF NOT EXISTS idx_map1_grade ON map1_curriculum_units(grade, unit_number);

-- 2. BẢNG BẢN ĐỒ 2 (MAP 2): VI KỸ NĂNG & 7 NHÓM BẪY PHÒNG THI TIẾNG ANH
CREATE TABLE IF NOT EXISTS map2_skill_traps (
    trap_code TEXT PRIMARY KEY,
    skill_domain TEXT NOT NULL,
    title_vi TEXT NOT NULL,
    description_vi TEXT NOT NULL,
    quick_rule_5s TEXT NOT NULL,
    target_grades_json TEXT NOT NULL
);

-- 3. BẢNG NGÂN HÀNG CÂU HỎI NGUYÊN TỬ (ATOMIC QUESTION BANK - BẢO TOÀN <u> & <b>)
CREATE TABLE IF NOT EXISTS question_bank (
    question_id TEXT PRIMARY KEY,
    grade INTEGER NOT NULL,
    unit_id TEXT NOT NULL,
    skill_domain TEXT NOT NULL,            -- 'PHONETICS' | 'GRAMMAR' | 'VOCABULARY' | 'SIGNS' | 'READING' | 'REWRITE' | 'WORD_FORM' | 'ERROR_FIX'
    difficulty_level INTEGER NOT NULL,     -- 1..4
    trap_code TEXT,                        -- Liên kết map2_skill_traps
    stem_text TEXT NOT NULL,               -- Nội dung câu hỏi (giữ nguyên <u>...</u> và <b>...</b>)
    passage_or_sign_svg TEXT,              -- Khung biển báo trực quan HTML/SVG hoặc đoạn văn Proofreading/Reading
    options_json TEXT NOT NULL,            -- JSON mảng 4 phương án (giữ nguyên <u>...</u> và <b>...</b>)
    correct_index INTEGER NOT NULL,        -- 0, 1, 2, 3 (tương ứng A, B, C, D)
    answer_mode TEXT DEFAULT 'MCQ',        -- 'MCQ' | 'WRITTEN' | 'BOTH'
    written_answers_json TEXT DEFAULT '[]',-- Danh sách đáp án tự luận chuẩn (hỗ trợ viết tắt tương đương don't = do not)
    signal_keywords_json TEXT,
    explanation_vi TEXT NOT NULL,
    source_worksheet_id TEXT,
    qa_status TEXT DEFAULT 'QA_VERIFIED',
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_qbank_lookup ON question_bank(grade, unit_id, skill_domain, difficulty_level, qa_status);
CREATE INDEX IF NOT EXISTS idx_qbank_trap ON question_bank(trap_code);
CREATE INDEX IF NOT EXISTS idx_qbank_ws ON question_bank(source_worksheet_id);

-- 4. BẢNG DANH MỤC TỦ ĐỀ CỦA CÔ THÚY (ĐỦ 20 BỘ ĐỀ TƯƠNG TÁC TỪ 63 TỆP GỐC)
CREATE TABLE IF NOT EXISTS worksheets_catalog (
    worksheet_id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    grade INTEGER NOT NULL,
    unit_scope TEXT NOT NULL,
    category TEXT NOT NULL,
    question_count INTEGER NOT NULL,
    duration_minutes INTEGER DEFAULT 45,
    student_pdf_drive_id TEXT,
    student_docx_drive_id TEXT,
    teacher_pdf_drive_id TEXT,
    teacher_docx_drive_id TEXT,
    local_student_pdf TEXT,
    local_student_docx TEXT,
    local_teacher_pdf TEXT,
    local_teacher_docx TEXT,
    shard_json_path TEXT,
    approval_state TEXT NOT NULL,
    source_group TEXT NOT NULL,
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_ws_filter ON worksheets_catalog(grade, category, approval_state);

-- 5. BẢNG LỚP HỌC CỦA CÔ THÚY
CREATE TABLE IF NOT EXISTS classes (
    class_id TEXT PRIMARY KEY,
    class_name TEXT NOT NULL,
    grade INTEGER NOT NULL,
    schedule_note TEXT,
    student_count INTEGER DEFAULT 0
);

-- 6. BẢNG HỌC SINH
CREATE TABLE IF NOT EXISTS students (
    student_code TEXT PRIMARY KEY,
    full_name TEXT NOT NULL,
    class_id TEXT NOT NULL,
    grade INTEGER NOT NULL,
    pin_code TEXT NOT NULL,
    streak_days INTEGER DEFAULT 1,
    xp_points INTEGER DEFAULT 100,
    weak_traps_json TEXT DEFAULT '[]',
    FOREIGN KEY (class_id) REFERENCES classes(class_id)
);
CREATE INDEX IF NOT EXISTS idx_students_class ON students(class_id, grade);

-- 7. BẢNG ĐỢT GIAO BÀI & CA THI ONLINE
CREATE TABLE IF NOT EXISTS assignments (
    assignment_id TEXT PRIMARY KEY,
    worksheet_id TEXT NOT NULL,
    class_id TEXT NOT NULL,
    title TEXT NOT NULL,
    duration_minutes INTEGER NOT NULL,
    shuffle_questions INTEGER DEFAULT 1,
    show_explanation_mode TEXT DEFAULT 'AFTER_SUBMIT',
    status TEXT DEFAULT 'OPEN',
    created_at TEXT NOT NULL
);

-- 8. BẢNG BÀI NỘP CỦA HỌC SINH & HÀNG ĐỢI IN TỪ XA (ĐỒNG BỘ 00_PRINT_QUEUE TRÊN DRIVE)
CREATE TABLE IF NOT EXISTS submissions (
    submission_id TEXT PRIMARY KEY,
    assignment_id TEXT NOT NULL,
    worksheet_id TEXT NOT NULL,
    student_code TEXT NOT NULL,
    student_name TEXT NOT NULL,
    class_id TEXT NOT NULL,
    exam_code TEXT DEFAULT '101',
    answers_json TEXT NOT NULL,
    auto_score REAL NOT NULL,
    final_score REAL,
    teacher_approved INTEGER DEFAULT 0,
    teacher_comment TEXT,
    wrong_traps_json TEXT DEFAULT '[]',
    submitted_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_sub_assign ON submissions(assignment_id, class_id);
CREATE INDEX IF NOT EXISTS idx_sub_student ON submissions(student_code);

CREATE TABLE IF NOT EXISTS print_jobs (
    job_id TEXT PRIMARY KEY,
    worksheet_id TEXT NOT NULL,
    worksheet_title TEXT NOT NULL,
    target_version TEXT NOT NULL,
    copies INTEGER DEFAULT 1,
    paper_size TEXT DEFAULT 'A4',
    duplex INTEGER DEFAULT 1,
    drive_file_id TEXT,
    status TEXT DEFAULT 'QUEUED',
    requested_from TEXT DEFAULT 'Điện thoại Cô Thúy',
    created_at TEXT NOT NULL,
    completed_at TEXT
);
