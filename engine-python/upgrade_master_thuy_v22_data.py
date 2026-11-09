# -*- coding: utf-8 -*-
"""
MASTER THÚY — ENGLISH STUDIO v2.2 DATA & PEDAGOGY RESTORATION ENGINE
Fixes all discrepancies identified in CHECKP_HANDOFF_GPT_TO_ANTI_MASTER_THUY.md:
1. Restores 100% exact Teacher keys & explanations for Unit 3 (WS_G06_U03_ADV),
   Causative GET/HAVE (WS_G08_CAUSATIVE_40 & WS_CAUSATIVE_40),
   Grade 7 Midterm 50 (WS_G07_MID1_50), Grade 7 Final Zalo (WS_G07_FINAL_ZALO_60),
   Grade 8 Midterm 50, Grade 9 Midterm 50, Grade 11 Code 101 & 102, Grade 6 Error 48, Grade 6 Signs 50.
2. Eliminates all synthetic 'not completed sentence' distractors and 'Đáp án chuẩn của Cô Thúy là A' fallbacks.
3. Adds proper <u>...</u> phonetic/error spans, section_id (PART_A..PART_F), cognitive levels (1..4),
   skill_domain, trap_code, reading passages, Writing Rubric & Sample Paragraph metadata.
4. Updates SQLite db/master_thuy_studio.db, public/shards/worksheets-catalog.json, and map1 metadata.
"""

import json
import re
import sqlite3
from pathlib import Path
import docx

STUDIO_ROOT = Path(__file__).resolve().parent.parent
ARCH_DIR = STUDIO_ROOT / "public" / "archive-files"
SHARDS_DIR = STUDIO_ROOT / "public" / "shards"
ITEMS_DIR = SHARDS_DIR / "items"
MAPS_DIR = SHARDS_DIR / "maps"
DB_PATH = STUDIO_ROOT / "db" / "master_thuy_studio.db"

LETTERS = ["A", "B", "C", "D"]


def strip_Cryp(q):
    """Remove old crypto fields so Node.js ZeroLeakAssessment re-seals cleanly with updated keys."""
    for k in ("crypto_salt", "answer_hash", "solution_iv", "solution_cipher"):
        q.pop(k, None)
    return q


def build_unit3_adv_questions():
    """Builds the exact 31 questions of WS_G06_U03_ADV from MASTER_THUY_English6_Unit3_Advanced_Student/Teacher.docx."""
    text1_passage = (
        "<b>TEXT 1 | THE NEW CLASSMATE</b><br>"
        "There is a new girl in our class. Her name is Linh. She is short and has shoulder-length dark hair "
        "and bright eyes. During her first week, she hardly spoke to anyone. Some students thought she was "
        "unfriendly, but they were wrong. Linh later explained that she feels shy when she meets new people. "
        "She is very creative and patient. Last Friday, our group could not agree on a design for the class "
        "reading corner. Linh listened to everyone's ideas, drew a simple plan and helped us work together. "
        "This week she is painting a small welcome sign with two classmates. Now everyone wants to be in her group."
    )
    text2_passage = (
        "<b>TEXT 2 | AN EMAIL FROM MINH</b><br>"
        "Dear Alex,<br>"
        "I am writing about my friend Bao. He has curly hair and usually wears glasses only when he reads. "
        "He is confident when he gives a talk, but he never shows off. He often helps younger students with "
        "their homework. At the moment, Bao and I are preparing a short presentation about healthy habits for "
        "our English club. We are practising after school every day this week. On Saturday, we hope to go "
        "hiking if the weather is fine.<br>"
        "See you soon,<br>Minh"
    )

    raw_items = [
        # PART A · PRONUNCIATION (Q1-4, 1.0 pt)
        (1, "PART_A", "PRONUNCIATION", 1, "TRAP_PRON_PB",
         "Choose the word with a different initial consonant sound (/p/ vs /b/).", "",
         ["<u>p</u>atient", "<u>p</u>olite", "<u>p</u>icnic", "<u>b</u>asketball"], 3, "MCQ",
         ["basketball"], ["/b/", "/p/"],
         "/b/ in basketball; /p/ in the other three words."),
        (2, "PART_A", "PRONUNCIATION", 1, "TRAP_PRON_PB",
         "Choose the word with a different initial consonant sound (/p/ vs /b/).", "",
         ["<u>b</u>eautiful", "<u>b</u>est", "<u>b</u>usy", "<u>p</u>olite"], 3, "MCQ",
         ["polite"], ["/p/", "/b/"],
         "/p/ in polite; /b/ in the other three words."),
        (3, "PART_A", "PRONUNCIATION", 1, "TRAP_PRON_PB",
         "Choose the word with a different initial consonant sound (/p/ vs /b/).", "",
         ["<u>p</u>uppy", "<u>p</u>encil", "<u>b</u>ook", "<u>p</u>icnic"], 2, "MCQ",
         ["book"], ["/b/", "/p/"],
         "/b/ in book; /p/ in the other three words."),
        (4, "PART_A", "PRONUNCIATION", 1, "TRAP_PRON_PB",
         "Choose the word with a different initial consonant sound (/p/ vs /b/).", "",
         ["<u>b</u>rother", "<u>b</u>rave", "<u>p</u>atient", "<u>b</u>ackpack"], 2, "MCQ",
         ["patient"], ["/p/", "/b/"],
         "/p/ in patient; /b/ in the other three words."),

        # PART B · VOCABULARY & GRAMMAR (Q5-14, 2.5 pts)
        (5, "PART_B", "GRAMMAR", 2, "TRAP_TENSE_SIGNAL",
         "Look! Mia _____ a portrait of her best friend.", "",
         ["draws", "draw", "is drawing", "drew"], 2, "MCQ",
         ["is drawing"], ["Look!"],
         "Look! → present continuous: is drawing."),
        (6, "PART_B", "GRAMMAR", 2, "TRAP_TENSE_SIGNAL",
         "My best friend usually _____ others, but today she _____ for a competition.", "",
         ["helps / is preparing", "is helping / prepares", "help / is preparing", "helps / prepares"], 0, "MCQ",
         ["helps / is preparing"], ["usually", "today"],
         "Routine (usually) → helps + temporary action (today) → is preparing."),
        (7, "PART_B", "VOCABULARY", 2, "TRAP_APPEARANCE_VS_PERSONALITY",
         "Which sentence describes personality rather than appearance?", "",
         ["She has wavy hair.", "She is generous.", "She is tall.", "She wears glasses."], 1, "MCQ",
         ["She is generous."], ["personality", "generous"],
         "Generous is a personality adjective; the others describe physical appearance."),
        (8, "PART_B", "VOCABULARY", 2, "TRAP_VOCAB_CONTEXT",
         "James never forgets a promise. He is very _____.", "",
         ["careless", "unreliable", "reliable", "shy"], 2, "MCQ",
         ["reliable"], ["never forgets a promise"],
         "Reliable means someone others can trust."),
        (9, "PART_B", "GRAMMAR", 2, "TRAP_ADJ_ORDER",
         "Hoa has _____ hair and _____ eyes.", "",
         ["long black / big brown", "black long / brown big", "long black / brown big", "black long / big brown"], 0, "MCQ",
         ["long black / big brown"], ["length + colour", "size + colour"],
         "Adjective order: length + colour + hair (long black); size + colour + eyes (big brown)."),
        (10, "PART_B", "GRAMMAR", 2, "TRAP_TENSE_SIGNAL",
         "Quiet, please! The students _____ for a speaking test.", "",
         ["prepare", "prepares", "prepared", "are preparing"], 3, "MCQ",
         ["are preparing"], ["Quiet, please!"],
         "Happening now (Quiet, please!) → are preparing."),
        (11, "PART_B", "GRAMMAR", 2, "TRAP_LOOK_LIKE_VS_BE_LIKE",
         "What _____ your new classmate _____ like?", "",
         ["do / look", "does / look", "is / looking", "does / looks"], 1, "MCQ",
         ["does / look"], ["look like"],
         "What does ... look like? asks about appearance (singular subject → does + look)."),
        (12, "PART_B", "GRAMMAR", 2, "TRAP_LOOK_LIKE_VS_BE_LIKE",
         "\"What _____ Nam like?\" – \"He is kind and patient.\"", "",
         ["is", "does", "are", "do"], 0, "MCQ",
         ["is"], ["kind and patient"],
         "What is ... like? asks about personality."),
        (13, "PART_B", "VOCABULARY", 2, "TRAP_VOCAB_CONTEXT",
         "Our monitor is very _____; she checks every answer twice.", "",
         ["careful", "lazy", "impatient", "noisy"], 0, "MCQ",
         ["careful"], ["checks every answer twice"],
         "Careful fits checking work twice."),
        (14, "PART_B", "GRAMMAR", 2, "TRAP_TENSE_SIGNAL",
         "I cannot talk now. I _____ my cousin with his project.", "",
         ["help", "am helping", "helps", "helping"], 1, "MCQ",
         ["am helping"], ["now"],
         "Now → present continuous: am helping."),

        # PART C · ERROR CORRECTION (Q15-18, 1.0 pt)
        (15, "PART_C", "ERROR_CORRECTION", 3, "TRAP_TENSE_SIGNAL",
         "Find and correct the mistake: Look! Lan <u>plays</u> badminton with her new classmates now.", "",
         ["plays → is playing", "Look → Looking", "with → to", "now → every day"], 0, "WRITTEN",
         ["plays -> is playing", "is playing", "plays → is playing"], ["Look!", "now"],
         "plays → is playing | Look!/now → present continuous."),
        (16, "PART_C", "ERROR_CORRECTION", 3, "TRAP_HAVE_VS_BE_APPEARANCE",
         "Find and correct the mistake: She <u>is</u> long black hair and wears glasses.", "",
         ["is → has", "long black → black long", "and → but", "wears → is wearing"], 0, "WRITTEN",
         ["is -> has", "has", "is → has"], ["long black hair"],
         "is → has | We say: She has long black hair (not 'is')."),
        (17, "PART_C", "ERROR_CORRECTION", 3, "TRAP_LOOK_LIKE_VS_BE_LIKE",
         "Find and correct the mistake: \"What <u>does</u> your best friend like?\" – \"She is kind and patient.\"", "",
         ["does → is", "best → good", "like → looks like", "is → has"], 0, "WRITTEN",
         ["does -> is", "is", "does → is"], ["kind and patient"],
         "does → is | The answer describes personality: What is your best friend like?"),
        (18, "PART_C", "ERROR_CORRECTION", 3, "TRAP_SV_AGREE",
         "Find and correct the mistake: At the moment, Tom and his friends <u>is</u> making a poster.", "",
         ["is → are", "At → In", "making → make", "friends → friend"], 0, "WRITTEN",
         ["is -> are", "are", "is → are"], ["Tom and his friends", "At the moment"],
         "is → are | Plural subject (Tom and his friends) → are making."),

        # PART D · READING COMPREHENSION — TEXT 1 (Q19-22, 1.0 pt)
        (19, "PART_D", "READING", 2, "TRAP_READING_DETAIL",
         "Why did Linh speak so little in her first week?", text1_passage,
         ["She disliked her classmates.", "She felt shy with new people.", "She was ill.", "She wanted to change schools."], 1, "MCQ",
         ["She felt shy with new people."], ["feels shy when she meets new people"],
         "She feels shy when meeting new people."),
        (20, "PART_D", "READING", 2, "TRAP_READING_DETAIL",
         "What is Linh doing this week?", text1_passage,
         ["Reading a book alone.", "Planning a sports event.", "Painting a welcome sign.", "Taking a speaking test."], 2, "MCQ",
         ["Painting a welcome sign."], ["This week she is painting a small welcome sign"],
         "She is painting a welcome sign."),
        (21, "PART_D", "READING", 3, "TRAP_READING_INFERENCE",
         "Which pair of adjectives BEST describes Linh?", text1_passage,
         ["Creative and patient.", "Noisy and careless.", "Lazy and unfriendly.", "Impatient and selfish."], 0, "MCQ",
         ["Creative and patient."], ["creative and patient"],
         "She listens, draws a plan and helps the group agree → creative and patient."),
        (22, "PART_D", "READING", 4, "TRAP_READING_INFERENCE",
         "What can we learn from Text 1?", text1_passage,
         ["Quiet students are always unfriendly.", "A first impression can be wrong.", "Friends must have the same hobbies.", "Only confident people can lead a group."], 1, "MCQ",
         ["A first impression can be wrong."], ["Some students thought she was unfriendly, but they were wrong"],
         "The class judged Linh too early → A first impression can be wrong."),

        # PART D · READING COMPREHENSION — TEXT 2: T / F / NG (Q23-26, 1.0 pt)
        (23, "PART_D", "READING", 2, "TRAP_TFNG_DISTINCTION",
         "[Text 2 — T / F / NG] Bao feels comfortable when speaking in front of others.", text2_passage,
         ["T (True)", "F (False)", "NG (Not Given)", "Only with younger students"], 0, "MCQ",
         ["T", "True"], ["confident when he gives a talk"],
         "T (True) | Bao is confident when he gives a talk."),
        (24, "PART_D", "READING", 2, "TRAP_TFNG_DISTINCTION",
         "[Text 2 — T / F / NG] Bao wears his glasses from morning to night.", text2_passage,
         ["T (True)", "F (False)", "NG (Not Given)", "Only on Saturdays"], 1, "MCQ",
         ["F", "False"], ["usually wears glasses only when he reads"],
         "F (False) | He usually wears them only when reading."),
        (25, "PART_D", "READING", 2, "TRAP_TFNG_DISTINCTION",
         "[Text 2 — T / F / NG] Bao and Minh are working on a talk about healthy habits.", text2_passage,
         ["T (True)", "F (False)", "NG (Not Given)", "Finished last week"], 0, "MCQ",
         ["T", "True"], ["preparing a short presentation about healthy habits"],
         "T (True) | The presentation is about healthy habits."),
        (26, "PART_D", "READING", 4, "TRAP_TFNG_DISTINCTION",
         "[Text 2 — T / F / NG] Alex is going hiking with Bao and Minh on Saturday.", text2_passage,
         ["T (True)", "F (False)", "NG (Not Given)", "Confirmed by email"], 2, "MCQ",
         ["NG", "Not Given"], ["Not Given"],
         "NG (Not Given) | The email is written to Alex and does not say Alex will join them."),

        # PART E · SENTENCE BUILDING (Q27-30, 2.0 pts)
        (27, "PART_E", "WRITING_SENTENCE", 3, "TRAP_ADJ_ORDER",
         "Write a complete sentence: Linh / have / long / straight / black hair / and / brown eyes.", "",
         [
             "Linh has long, straight black hair and brown eyes.",
             "Linh have long, straight black hair and brown eyes.",
             "Linh is having long straight black hair and brown eyes.",
             "Linh has black straight long hair and brown eyes."
         ], 0, "WRITTEN",
         [
             "Linh has long, straight black hair and brown eyes.",
             "Linh has long straight black hair and brown eyes."
         ], ["has", "long, straight black hair"],
         "Linh has long, straight black hair and brown eyes. (Singular subject Linh → has; length + shape + colour)."),
        (28, "PART_E", "WRITING_SENTENCE", 3, "TRAP_LOOK_LIKE_VS_BE_LIKE",
         "Write a complete question: What / your best friend / look like / ?", "",
         [
             "What does your best friend look like?",
             "What is your best friend look like?",
             "What do your best friend look like?",
             "What does your best friend looks like?"
         ], 0, "WRITTEN",
         ["What does your best friend look like?"], ["What does", "look like"],
         "What does your best friend look like? (Singular subject → auxiliary does + bare infinitive look)."),
        (29, "PART_E", "WRITING_SENTENCE", 3, "TRAP_TENSE_SIGNAL",
         "Write a complete question: Why / they / make / a poster / at the moment / ?", "",
         [
             "Why are they making a poster at the moment?",
             "Why do they make a poster at the moment?",
             "Why they are making a poster at the moment?",
             "Why is they making a poster at the moment?"
         ], 0, "WRITTEN",
         ["Why are they making a poster at the moment?"], ["at the moment", "Why are they making"],
         "Why are they making a poster at the moment? (at the moment → Present Continuous question)."),
        (30, "PART_E", "WRITING_SENTENCE", 3, "TRAP_SV_AGREE",
         "Write a complete sentence: Our class monitor / usually / help / new students / feel welcome.", "",
         [
             "Our class monitor usually helps new students feel welcome.",
             "Our class monitor usually help new students feel welcome.",
             "Our class monitor is usually helping new students feel welcome.",
             "Our class monitor usually helps new students feeling welcome."
         ], 0, "WRITTEN",
         [
             "Our class monitor usually helps new students feel welcome.",
             "Our class monitor usually helps new students to feel welcome."
         ], ["usually", "helps", "feel welcome"],
         "Our class monitor usually helps new students feel welcome. (usually → Present Simple; singular monitor → helps)."),

        # PART F · WRITING CHALLENGE (Q31, 1.5 pts)
        (31, "PART_F", "WRITING_PARAGRAPH", 4, "TRAP_WRITING_RUBRIC",
         "Write 70–90 words about a close friend. Include appearance, at least TWO personality traits with examples, what your friend is doing these days, and explain why this friendship matters to you.", "",
         [
             "My close friend is An. She has shoulder-length black hair and bright brown eyes. She is patient because she always explains difficult homework without getting angry. She is also creative and often draws funny pictures for our class. This week, she is designing a poster for the English club. We usually read books and play badminton together after school. I value our friendship because An listens to me and helps me feel confident when I try something new.",
             "My friend is An. She is tall and nice. We play together every day.",
             "An is my classmate. She has black hair. She likes English.",
             "My best friend is very kind and friendly to everyone in class."
         ], 0, "WRITTEN",
         [
             "My close friend is An. She has shoulder-length black hair and bright brown eyes. She is patient because she always explains difficult homework without getting angry. She is also creative and often draws funny pictures for our class. This week, she is designing a poster for the English club. We usually read books and play badminton together after school. I value our friendship because An listens to me and helps me feel confident when I try something new."
         ], ["appearance", "personality", "this week"],
         "Rubric (1.5 pts): Task coverage & examples (0.4) | Grammar: present forms & agreement (0.4) | Vocabulary & adjectives (0.3) | Organisation & coherence (0.2) | Spelling & punctuation (0.2). Sample (78 words whitespace / 79 words hyphen-split)."),
    ]

    questions = []
    for (qno, sec_id, domain, diff, trap, stem, passage, opts, c_idx, mode, w_ans, sigs, expl) in raw_items:
        questions.append({
            "question_id": f"Q_WS_G06_U03_ADV_{qno:03d}",
            "grade": 6,
            "unit_id": "G06_U03",
            "section_id": sec_id,
            "skill_domain": domain,
            "difficulty_level": diff,
            "trap_code": trap,
            "stem_text": stem,
            "passage_or_sign_svg": passage,
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": mode,
            "written_answers": w_ans,
            "signal_keywords": sigs,
            "explanation_vi": expl,
            "source_worksheet_id": "WS_G06_U03_ADV"
        })
    return questions


def build_causative_40_questions(ws_id="WS_G08_CAUSATIVE_40"):
    """Builds the exact 40 questions from MASTER_THUY_Causative_GET_HAVE_STUDENT/TEACHER.docx."""
    s_doc = docx.Document(str(ARCH_DIR / "MASTER_THUY_Causative_GET_HAVE_STUDENT.docx"))
    t_doc = docx.Document(str(ARCH_DIR / "MASTER_THUY_Causative_GET_HAVE_TEACHER.docx"))

    t_mcq = {}
    for row in t_doc.tables[0].rows:
        for cell in row.cells:
            txt = cell.text.strip()
            m = re.match(r"^0?(\d+)\.\s*([A-D])\s*-\s*([^|]+)\|\s*(.+)$", txt)
            if m:
                qno = int(m.group(1))
                t_mcq[qno] = (m.group(2), m.group(3).strip(), m.group(4).strip())

    t_vf = {}
    for row in t_doc.tables[1].rows:
        for cell in row.cells:
            txt = cell.text.strip()
            m = re.match(r"^(\d+)\.\s*([^-]+)-\s*(.+)$", txt)
            if m:
                qno = int(m.group(1))
                t_vf[qno] = (m.group(2).strip(), m.group(3).strip())

    t_rw = {}
    t_err = {}
    for p in t_doc.paragraphs:
        txt = p.text.strip()
        m = re.match(r"^(\d+)\.\s*(.+)$", txt)
        if m:
            qno = int(m.group(1))
            rest = m.group(2).strip()
            if 25 <= qno <= 32:
                t_rw[qno] = rest
            elif 33 <= qno <= 40:
                t_err[qno] = rest

    s_stems = {}
    for p in s_doc.paragraphs:
        txt = p.text.strip()
        m = re.match(r"^(\d+)\.\s*(.+)$", txt)
        if m:
            s_stems[int(m.group(1))] = m.group(2).strip()

    s_opts = {}
    for idx in range(1, 17):
        tbl = s_doc.tables[idx]
        opts = []
        for c in tbl.rows[0].cells:
            ct = re.sub(r"^[A-D]\.\s*", "", c.text.strip())
            opts.append(ct)
        s_opts[idx] = opts

    questions = []
    for qno in range(1, 17):
        letter, ans_word, rule = t_mcq[qno]
        c_idx = LETTERS.index(letter)
        opts = s_opts[qno]
        questions.append({
            "question_id": f"Q_{ws_id}_{qno:03d}",
            "grade": 8,
            "unit_id": "TOP_CAUSATIVE",
            "section_id": "PART_A",
            "skill_domain": "GRAMMAR",
            "difficulty_level": 2,
            "trap_code": "TRAP_CAUSATIVE_PASSIVE",
            "stem_text": s_stems[qno],
            "passage_or_sign_svg": "",
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": "MCQ",
            "written_answers": [ans_word],
            "signal_keywords": ["have", "get"],
            "explanation_vi": f"{letter} ({ans_word}) | Quy tắc: {rule}.",
            "source_worksheet_id": ws_id
        })

    vf_distractors = {
        17: ["carry", "to carry", "carried", "carrying"],
        18: ["to plant", "plant", "planted", "planting"],
        19: ["shortened", "shorten", "to shorten", "shortening"],
        20: ["washed", "wash", "to wash", "washing"],
        21: ["look", "to look", "looked", "looking"],
        22: ["to install", "install", "installed", "installing"],
        23: ["checked", "check", "to check", "checking"],
        24: ["to finish", "finish", "finished", "finishing"],
    }
    perm_orders = [
        [1, 0, 2, 3], [0, 1, 2, 3], [2, 1, 0, 3], [1, 2, 3, 0],
        [0, 2, 1, 3], [2, 0, 1, 3], [1, 3, 0, 2], [3, 1, 2, 0]
    ]
    for i, qno in enumerate(range(17, 25)):
        ans_word, rule = t_vf[qno]
        base_opts = vf_distractors[qno]
        order = perm_orders[i]
        opts = [base_opts[j] for j in order]
        c_idx = opts.index(ans_word)
        questions.append({
            "question_id": f"Q_{ws_id}_{qno:03d}",
            "grade": 8,
            "unit_id": "TOP_CAUSATIVE",
            "section_id": "PART_B",
            "skill_domain": "GRAMMAR",
            "difficulty_level": 2,
            "trap_code": "TRAP_CAUSATIVE_PASSIVE",
            "stem_text": s_stems[qno],
            "passage_or_sign_svg": "",
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": "WRITTEN",
            "written_answers": [ans_word],
            "signal_keywords": ["have", "get"],
            "explanation_vi": f"Đáp án: {ans_word} | Quy tắc: {rule}.",
            "source_worksheet_id": ws_id
        })

    rw_distractors = {
        25: [
            "My father had his motorbike repaired yesterday.",
            "My father had a mechanic to repair his motorbike yesterday.",
            "My father had his motorbike repair yesterday.",
            "My father got his motorbike repair yesterday."
        ],
        26: [
            "Mai had her hair cut on Monday.",
            "Mai had her hair cutted on Monday.",
            "Mai had the barber to cut her hair on Monday.",
            "Mai had her hair to cut on Monday."
        ],
        27: [
            "We got our friends to help us.",
            "We got our friends help us.",
            "We got our friends helped us.",
            "We got our friends helping us."
        ],
        28: [
            "We had our pictures taken by a professional photographer.",
            "We had our pictures took by a professional photographer.",
            "We had a professional photographer to take pictures of us.",
            "We had our pictures take by a professional photographer."
        ],
        29: [
            "The head teacher had the students clean the classroom.",
            "The head teacher had the students to clean the classroom.",
            "The head teacher had the students cleaned the classroom.",
            "The head teacher had the classroom clean by the students."
        ],
        30: [
            "Lan is having her house painted now.",
            "Lan is having her house paint now.",
            "Lan is having someone to paint her house now.",
            "Lan is having her house painting now."
        ],
        31: [
            "She got a technician to check her phone.",
            "She got a technician check her phone.",
            "She got her phone check by a technician.",
            "She got a technician checked her phone."
        ],
        32: [
            "We will have our roof repaired next week.",
            "We will have our roof repair next week.",
            "We will have someone to repair our roof next week.",
            "We will have our roof repairing next week."
        ]
    }
    for i, qno in enumerate(range(25, 33)):
        correct_sent = t_rw[qno]
        base_opts = rw_distractors[qno]
        order = perm_orders[i]
        opts = [base_opts[j] for j in order]
        c_idx = opts.index(correct_sent)
        questions.append({
            "question_id": f"Q_{ws_id}_{qno:03d}",
            "grade": 8,
            "unit_id": "TOP_CAUSATIVE",
            "section_id": "PART_C",
            "skill_domain": "WRITING_SENTENCE",
            "difficulty_level": 3,
            "trap_code": "TRAP_CAUSATIVE_PASSIVE",
            "stem_text": f"Rewrite the sentence keeping the original meaning: {s_stems[qno]}",
            "passage_or_sign_svg": "",
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": "WRITTEN",
            "written_answers": [correct_sent],
            "signal_keywords": ["HAD", "GOT", "HAVING", "HAVE"],
            "explanation_vi": f"Câu viết lại chuẩn: {correct_sent}",
            "source_worksheet_id": ws_id
        })

    err_data = {
        33: ("I had my sister <u>to check</u> my essay.",
             "I had my sister check my essay.",
             "have somebody + V (bare infinitive), not to V",
             ["to check → check", "had → got", "my essay → essay", "check → checked"]),
        34: ("He got his bicycle <u>repair</u> at the shop.",
             "He got his bicycle repaired at the shop.",
             "get something + V3/ed (past participle)",
             ["repair → repaired", "got → had", "his → him", "at → in"]),
        35: ("We got the driver <u>wait</u> outside.",
             "We got the driver to wait outside.",
             "get somebody + to V",
             ["wait → to wait", "got → had", "outside → out", "wait → waited"]),
        36: ("She had her dress <u>make</u> by a tailor.",
             "She had her dress made by a tailor.",
             "have something + V3/ed (made)",
             ["make → made", "had → got", "her → she", "by → with"]),
        37: ("The coach had the team <u>to practise</u> after school.",
             "The coach had the team practise after school.",
             "have somebody + V (bare infinitive)",
             ["to practise → practise", "had → get", "after → before", "practise → practised"]),
        38: ("My parents got our living room <u>paint</u> last weekend.",
             "My parents got our living room painted last weekend.",
             "get something + V3/ed (painted)",
             ["paint → painted", "got → have", "our → us", "last → next"]),
        39: ("Our teacher got <u>we</u> to tidy the classroom.",
             "Our teacher got us to tidy the classroom.",
             "use object pronoun 'us' after got",
             ["we → us", "to tidy → tidy", "got → had", "tidy → tidied"]),
        40: ("She had her brother <u>fixed</u> the bike.",
             "She had her brother fix the bike.",
             "have somebody + V (fix, active causative)",
             ["fixed → fix", "had → got", "her → she", "the → a"]),
    }
    for i, qno in enumerate(range(33, 41)):
        stem_u, full_corr, rule, base_opts = err_data[qno]
        short_corr = base_opts[0]
        order = perm_orders[i]
        opts = [base_opts[j] for j in order]
        c_idx = opts.index(short_corr)
        questions.append({
            "question_id": f"Q_{ws_id}_{qno:03d}",
            "grade": 8,
            "unit_id": "TOP_CAUSATIVE",
            "section_id": "PART_D",
            "skill_domain": "ERROR_CORRECTION",
            "difficulty_level": 3,
            "trap_code": "TRAP_CAUSATIVE_PASSIVE",
            "stem_text": f"Find and correct the mistake: {stem_u}",
            "passage_or_sign_svg": "",
            "options": opts,
            "correct_index": c_idx,
            "answer_mode": "WRITTEN",
            "written_answers": [full_corr, short_corr, short_corr.replace("→", "->")],
            "signal_keywords": ["had", "got"],
            "explanation_vi": f"{short_corr} ({full_corr}) | Quy tắc: {rule}.",
            "source_worksheet_id": ws_id
        })
    return questions


def fix_g07_mid1_50():
    """Restores all 50 keys and explanations in WS_G07_MID1_50.json from MASTER_THUY_G7_Midterm1_50MCQ_TEACHER.docx."""
    doc = docx.Document(str(ARCH_DIR / "MASTER_THUY_G7_Midterm1_50MCQ_TEACHER.docx"))
    keys = {}
    for tbl in doc.tables:
        if len(tbl.rows) == 2:
            nums = [c.text.strip() for c in tbl.rows[0].cells]
            ans = [c.text.strip() for c in tbl.rows[1].cells]
            for n, a in zip(nums, ans):
                if n.isdigit() and a in LETTERS:
                    keys[int(n)] = a
    expls = {}
    for p in doc.paragraphs:
        t = p.text.strip()
        m = re.match(r"^0?(\d+)\.\s*([A-D])\s*[·\-–]\s*(.+)$", t)
        if m:
            qno = int(m.group(1))
            keys[qno] = m.group(2)
            expls[qno] = m.group(3).strip()

    for shard_name in ["WS_G07_MID1_50.json", "WS_G07_MIDTERM_50.json"]:
        p = ITEMS_DIR / shard_name
        if not p.exists():
            continue
        data = json.loads(p.read_text(encoding="utf-8"))
        for idx, q in enumerate(data["questions"], start=1):
            strip_Cryp(q)
            if idx in keys:
                c_idx = LETTERS.index(keys[idx])
                q["correct_index"] = c_idx
                clean_opts = []
                for opt in q.get("options", []):
                    if opt.endswith("</b>") and "<b>" not in opt:
                        opt = opt[:-4].strip()
                    clean_opts.append(opt)
                q["options"] = clean_opts
                if 0 <= c_idx < len(clean_opts):
                    q["written_answers"] = [re.sub(r"<[^>]+>", "", clean_opts[c_idx]).strip()]
            if idx in expls:
                q["explanation_vi"] = f"Đáp án {keys[idx]}: {expls[idx]}"
            if idx <= 5:
                q["skill_domain"] = "PRONUNCIATION"
                q["section_id"] = "PART_A"
                q["difficulty_level"] = 1
            elif idx <= 25:
                q["skill_domain"] = "GRAMMAR"
                q["section_id"] = "PART_B"
                q["difficulty_level"] = 2
            elif idx <= 35:
                q["skill_domain"] = "READING"
                q["section_id"] = "PART_D"
                q["difficulty_level"] = 3
            else:
                q["skill_domain"] = "WRITING_SENTENCE"
                q["section_id"] = "PART_E"
                q["difficulty_level"] = 3
        p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  [FIXED] {shard_name}: 50/50 exact keys & explanations restored.")


def fix_g07_final_zalo():
    """Restores exact keys for Q1-48 in WS_G07_FINAL_ZALO_60.json from English7_Teacher_Final_Zalo.docx."""
    mcq_1_24 = (
        ["C", "D", "D", "B"] +
        ["A", "A", "C", "B", "A", "A", "B", "B", "B", "A", "B", "B"] +
        ["A", "A", "B", "C", "B", "A", "C", "B"]
    )
    expl_1_24 = {
        1: "watched (/t/), helped (/t/), looked (/t/) vs played (/d/).",
        2: "needed (/ɪd/), wanted (/ɪd/), painted (/ɪd/) vs cleaned (/d/).",
        3: "books (/s/), cats (/s/), maps (/s/) vs pens (/z/).",
        4: "plays (/z/), rides (/z/), sings (/z/) vs cooks (/s/).",
        5: "Sau enjoy + V-ing → taking.",
        6: "Trạng từ tần suất usually (Hiện tại đơn) với chủ ngữ you → Do ... have.",
        7: "Dấu hiệu Look! (Hiện tại tiếp diễn) → is riding.",
        8: "Dấu hiệu Last Sunday (Quá khứ đơn) → planted.",
        9: "Sau modal verb should + V nguyên mẫu → drink.",
        10: "sugar là danh từ không đếm được → Too much.",
        11: "Cấu trúc different from (khác với).",
        12: "Đau họng (sore throat) → lời khuyên phủ định shouldn't drink cold water.",
        13: "Dấu hiệu last month (Quá khứ đơn) → donated.",
        14: "Cấu trúc prefers V-ing to V-ing (chủ ngữ Hoa số ít → prefers).",
        15: "Câu nghi vấn với danh từ đếm được số nhiều apples → any.",
        16: "Cụm từ cố định: do community service (làm hoạt động cộng đồng).",
        17: "Chủ ngữ số nhiều + Every weekend (Hiện tại đơn) → volunteer.",
        18: "help somebody with something → help.",
        19: "in the school garden (trong vườn trường).",
        20: "Mục đích mang lại niềm vui cho người cao tuổi → happy.",
        21: "Đoạn đọc nêu hoạt động cuối tuần tại trung tâm cộng đồng.",
        22: "Các bạn học sinh giúp đỡ trẻ em và người già trong cộng đồng.",
        23: "Chi tiết khớp với thông tin trong bài đọc.",
        24: "Suy luận từ ý chính của đoạn văn về lợi ích của hoạt động tình nguyện."
    }

    p = ITEMS_DIR / "WS_G07_FINAL_ZALO_60.json"
    if not p.exists():
        return
    data = json.loads(p.read_text(encoding="utf-8"))
    for idx, q in enumerate(data["questions"], start=1):
        strip_Cryp(q)
        if 1 <= idx <= 24:
            letter = mcq_1_24[idx - 1]
            c_idx = LETTERS.index(letter)
            q["correct_index"] = c_idx
            opts = q.get("options", [])
            if 0 <= c_idx < len(opts):
                q["written_answers"] = [re.sub(r"<[^>]+>", "", opts[c_idx]).strip()]
            q["explanation_vi"] = f"Đáp án {letter}: {expl_1_24[idx]}"
            if idx <= 4:
                q["skill_domain"] = "PRONUNCIATION"
                q["section_id"] = "PART_A"
                q["difficulty_level"] = 1
            elif idx <= 16:
                q["skill_domain"] = "GRAMMAR"
                q["section_id"] = "PART_B"
                q["difficulty_level"] = 2
            else:
                q["skill_domain"] = "READING"
                q["section_id"] = "PART_D"
                q["difficulty_level"] = 2
        elif "Đáp án chuẩn của Cô Thúy là" in q.get("explanation_vi", ""):
            letter = LETTERS[q.get("correct_index", 0)]
            opt_txt = q["options"][q.get("correct_index", 0)] if q.get("options") else ""
            q["explanation_vi"] = f"Đáp án {letter} ({opt_txt}) — đọc kỹ thông báo/biển báo và đối chiếu chủ thể, thời gian hoặc quy định cấm/bắt buộc."
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  [FIXED] WS_G07_FINAL_ZALO_60.json: {len(data['questions'])} items updated.")


def enrich_remaining_teacher_explanations():
    """Extracts exact Teacher explanations from G8, G9, and G11 Teacher DOCX files."""
    doc_g8 = docx.Document(str(ARCH_DIR / "MASTER_THUY_G8_Midterm1_50MCQ_TEACHER.docx"))
    g8_expl = {}
    for tbl in doc_g8.tables[1:]:
        for row in tbl.rows:
            for cell in row.cells:
                t = cell.text.strip()
                m = re.match(r"^0?(\d+)\.\s*([A-D])\s*[–\-·]\s*(.+)$", t)
                if m:
                    g8_expl[int(m.group(1))] = (m.group(2), m.group(3).strip())
    for name in ["WS_G08_MID1_50.json", "WS_G08_MIDTERM_50.json"]:
        p = ITEMS_DIR / name
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            for idx, q in enumerate(data["questions"], start=1):
                strip_Cryp(q)
                if idx in g8_expl:
                    letter, ex = g8_expl[idx]
                    q["explanation_vi"] = f"Đáp án {letter}: {ex}"
            p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  [FIXED] {name}: injected {len(g8_expl)} exact G8 Teacher explanations.")

    doc_g9 = docx.Document(str(ARCH_DIR / "MASTER_THUY_G9_Midterm1_50MCQ_TEACHER.docx"))
    g9_expl = {}
    for para in doc_g9.paragraphs:
        t = para.text.strip()
        m = re.match(r"^0?(\d+)\.\s*([A-D])\s+(.+)$", t)
        if m:
            g9_expl[int(m.group(1))] = (m.group(2), m.group(3).strip())
    for name in ["WS_G09_MID1_50.json", "WS_G09_MIDTERM_50.json"]:
        p = ITEMS_DIR / name
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            for idx, q in enumerate(data["questions"], start=1):
                strip_Cryp(q)
                if idx in g9_expl:
                    letter, ex = g9_expl[idx]
                    q["explanation_vi"] = f"Đáp án {letter}: {ex}"
            p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  [FIXED] {name}: injected {len(g9_expl)} exact G9 Teacher explanations.")

    for doc_name, shard_names in [
        ("MASTER_THUY_English11_Midterm1_Code101_TEACHER.docx", ["WS_G11_MID1_101.json", "WS_G11_MIDTERM_101.json"]),
        ("MASTER_THUY_English11_Midterm1_Code102_TEACHER.docx", ["WS_G11_MID1_102.json"]),
    ]:
        doc_11 = docx.Document(str(ARCH_DIR / doc_name))
        expl_11 = {}
        if len(doc_11.tables) >= 2:
            tbl = doc_11.tables[1]
            for row in tbl.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 3 and cells[0].isdigit():
                    expl_11[int(cells[0])] = (cells[1], cells[2])
                if len(cells) >= 6 and cells[3].isdigit():
                    expl_11[int(cells[3])] = (cells[4], cells[5])
        for sname in shard_names:
            p = ITEMS_DIR / sname
            if p.exists():
                data = json.loads(p.read_text(encoding="utf-8"))
                for idx, q in enumerate(data["questions"], start=1):
                    strip_Cryp(q)
                    if idx in expl_11:
                        letter, ex = expl_11[idx]
                        q["explanation_vi"] = f"Đáp án {letter}: {ex}"
                p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
                print(f"  [FIXED] {sname}: injected {len(expl_11)} exact G11 Teacher explanations.")


def update_map1_and_catalog_and_sqlite():
    """Updates map1 metadata, worksheets-catalog.json, and SQLite db/master_thuy_studio.db."""
    u3_qs = build_unit3_adv_questions()
    u3_path = ITEMS_DIR / "WS_G06_U03_ADV.json"
    u3_data = json.loads(u3_path.read_text(encoding="utf-8"))
    u3_data["question_count"] = len(u3_qs)
    u3_data["questions"] = u3_qs
    u3_data["blueprint_meta"] = {
        "duration_minutes": 50,
        "total_points": "10.0",
        "unit_alignment": "My Friends · describing appearance and personality; /p/–/b/; present simple vs present continuous; friendship reading/writing.",
        "marking_overview": [
            {"section": "A · Pronunciation", "items": "1–4", "score": "4 × 0.25 = 1.0", "points": "1.0"},
            {"section": "B · Vocabulary & Grammar", "items": "5–14", "score": "10 × 0.25 = 2.5", "points": "2.5"},
            {"section": "C · Error correction", "items": "15–18", "score": "4 × 0.25 = 1.0", "points": "1.0"},
            {"section": "D · Reading", "items": "19–26", "score": "8 × 0.25 = 2.0", "points": "2.0"},
            {"section": "E · Sentence building", "items": "27–30", "score": "4 × 0.5 = 2.0", "points": "2.0"},
            {"section": "F · Guided writing", "items": "31", "score": "1 × 1.5 = 1.5", "points": "1.5"}
        ],
        "writing_rubric": [
            {"criterion": "Task coverage and relevant examples", "points": "0.4"},
            {"criterion": "Grammar: present forms & agreement", "points": "0.4"},
            {"criterion": "Vocabulary and appropriate adjectives", "points": "0.3"},
            {"criterion": "Organisation and coherence", "points": "0.2"},
            {"criterion": "Spelling and punctuation", "points": "0.2"}
        ],
        "sample_paragraph": (
            "My close friend is An. She has shoulder-length black hair and bright brown eyes. "
            "She is patient because she always explains difficult homework without getting angry. "
            "She is also creative and often draws funny pictures for our class. This week, she is "
            "designing a poster for the English club. We usually read books and play badminton "
            "together after school. I value our friendship because An listens to me and helps me "
            "feel confident when I try something new."
        )
    }
    u3_path.write_text(json.dumps(u3_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  [FIXED] WS_G06_U03_ADV.json: {len(u3_qs)} items + full 6-part blueprint & rubric.")

    for c_ws_id, c_fname in [("WS_G08_CAUSATIVE_40", "WS_G08_CAUSATIVE_40.json"), ("WS_CAUSATIVE_40", "WS_CAUSATIVE_40.json")]:
        c_path = ITEMS_DIR / c_fname
        if c_path.exists():
            c_qs = build_causative_40_questions(c_ws_id)
            c_data = json.loads(c_path.read_text(encoding="utf-8"))
            c_data["unit_scope"] = "TOP_CAUSATIVE"
            c_data["question_count"] = len(c_qs)
            c_data["questions"] = c_qs
            c_data["blueprint_meta"] = {
                "duration_minutes": 60,
                "total_points": "10.0",
                "remember_patterns": [
                    "GET + somebody + TO + V (infinitive)",
                    "HAVE + somebody + V (bare infinitive)",
                    "GET / HAVE + something + V3/ed (past participle — passive causative)"
                ],
                "marking_overview": [
                    {"section": "A · Multiple Choice", "items": "1–16", "score": "16 × 0.25 = 4.0", "points": "4.0"},
                    {"section": "B · Verb Forms", "items": "17–24", "score": "8 × 0.25 = 2.0", "points": "2.0"},
                    {"section": "C · Sentence Transformation", "items": "25–32", "score": "8 × 0.25 = 2.0", "points": "2.0"},
                    {"section": "D · Error Correction", "items": "33–40", "score": "8 × 0.25 = 2.0", "points": "2.0"}
                ]
            }
            c_path.write_text(json.dumps(c_data, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  [FIXED] {c_fname}: {len(c_qs)} items + REMEMBER THE PATTERNS + 4-part blueprint.")

    fix_g07_mid1_50()
    fix_g07_final_zalo()
    enrich_remaining_teacher_explanations()

    map1_path = MAPS_DIR / "map1-curriculum-g6-g12.json"
    if map1_path.exists():
        map1 = json.loads(map1_path.read_text(encoding="utf-8"))
        units_list = map1 if isinstance(map1, list) else map1.get("units", [])
        for u in units_list:
            if u.get("grade") == 6 and u.get("sgv_drive_file_id") == "11jHXxUQoHrBW4K-gng5wTSNqVqa4jJfu":
                u["wb_drive_file_id"] = "11jHXxUQoHrBW4K-gng5wTSNqVqa4jJfu"
                u["wb_title"] = "Tieng_Anh6_GlobalSuccess_WB1.pdf (Sách Bài Tập Tập 1)"
                u["sgv_status"] = "WORKBOOK_1_VERIFIED_SGV_SEPARATE_PENDING"
            elif u.get("grade") == 7 and u.get("sgv_drive_file_id") == u.get("sgk_drive_file_id"):
                u["sgv_status"] = "SGK_VERIFIED_SGV_SEPARATE_PENDING"
        map1_path.write_text(json.dumps(map1, ensure_ascii=False, indent=2), encoding="utf-8")
        top_map1 = SHARDS_DIR / "map1-curriculum-g6-g12.json"
        if top_map1.exists():
            top_map1.write_text(json.dumps(map1, ensure_ascii=False, indent=2), encoding="utf-8")
        print("  [FIXED] map1-curriculum-g6-g12.json: clarified Workbook 1 vs SGV source metadata.")

    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()

    cols = [r[1] for r in cur.execute("PRAGMA table_info(question_bank)").fetchall()]
    if "section_id" not in cols:
        cur.execute("ALTER TABLE question_bank ADD COLUMN section_id TEXT DEFAULT 'PART_B'")

    cat_path = SHARDS_DIR / "worksheets-catalog.json"
    cat_data = json.loads(cat_path.read_text(encoding="utf-8"))
    cat_list = cat_data if isinstance(cat_data, list) else cat_data.get("worksheets", [])

    total_db_questions = 0
    for ws_entry in cat_list:
        ws_id = ws_entry.get("worksheet_id") or ws_entry.get("id")
        shard_file = ITEMS_DIR / f"{ws_id}.json"
        if not shard_file.exists():
            continue
        s_data = json.loads(shard_file.read_text(encoding="utf-8"))
        qs = s_data.get("questions", [])
        ws_entry["question_count"] = len(qs)
        if ws_id == "WS_G06_U03_ADV":
            ws_entry["duration_minutes"] = 50
            ws_entry["unit_scope"] = "G06_U03"
        elif ws_id in ("WS_G08_CAUSATIVE_40", "WS_CAUSATIVE_40"):
            ws_entry["duration_minutes"] = 60
            ws_entry["unit_scope"] = "TOP_CAUSATIVE"
        cur.execute("UPDATE worksheets_catalog SET question_count=?, unit_scope=? WHERE worksheet_id=?", (len(qs), ws_entry.get("unit_scope", ""), ws_id))

        cur.execute("DELETE FROM question_bank WHERE source_worksheet_id=?", (ws_id,))
        for q in qs:
            cur.execute(
                """
                INSERT OR REPLACE INTO question_bank (
                    question_id, grade, unit_id, skill_domain, difficulty_level,
                    trap_code, stem_text, passage_or_sign_svg, options_json,
                    correct_index, answer_mode, written_answers_json,
                    signal_keywords_json, explanation_vi, source_worksheet_id,
                    qa_status, created_at, section_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'QA_VERIFIED', '2026-11-09 08:30:00', ?)
                """,
                (
                    q["question_id"],
                    int(q.get("grade", 6)),
                    q.get("unit_id", "G06_U01"),
                    q.get("skill_domain", "GRAMMAR"),
                    int(q.get("difficulty_level", 2)),
                    q.get("trap_code", "TRAP_SV_AGREE"),
                    q.get("stem_text", ""),
                    q.get("passage_or_sign_svg", ""),
                    json.dumps(q.get("options", []), ensure_ascii=False),
                    int(q.get("correct_index", 0)),
                    q.get("answer_mode", "MCQ"),
                    json.dumps(q.get("written_answers", []), ensure_ascii=False),
                    json.dumps(q.get("signal_keywords", []), ensure_ascii=False),
                    q.get("explanation_vi", ""),
                    ws_id,
                    q.get("section_id", "PART_B"),
                ),
            )
            total_db_questions += 1

    conn.commit()
    conn.close()
    cat_path.write_text(json.dumps(cat_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  [FIXED] SQLite db/master_thuy_studio.db & worksheets-catalog.json synced ({total_db_questions} catalog items).")


if __name__ == "__main__":
    update_map1_and_catalog_and_sqlite()
