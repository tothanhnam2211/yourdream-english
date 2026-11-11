/**
 * ============================================================================
 * MASTER THÚY — ENGLISH STUDIO v2.2 (24/7 CLOUD DAEMON ON GOOGLE DRIVE)
 * File: Code.gs (Chạy ngầm 24/7 trên Google Apps Script — Không cần bật máy tính)
 * ============================================================================
 * CƠ CHẾ HOẠT ĐỘNG TỰ ĐỘNG 100%:
 * 1. WebApp của Cô Thúy gửi thẳng yêu cầu + nhiều ảnh chụp (Base64) qua Webhook `doPost(e)`
 *    (hoặc cô Thúy thả ảnh/file vào thư mục 00_LENH_DIEU_KHIEN_INBOX trên Drive).
 * 2. Bộ đếm giờ (Time-driven Trigger mỗi 1-5 phút) gọi hàm `autoScanInboxEveryMinute()`
 *    để tự thức dậy, lấy yêu cầu + ảnh trong INBOX.
 * 3. Gọi Gemini API để đọc ảnh viết tay/đề mẫu, biên soạn câu hỏi mới 100% chuẩn SGK
 *    Global Success (Bản Học Sinh + Bản Giáo Viên có Zebra Key 3 cột & Writing Rubric).
 * 4. Tự động tạo Google Docs + Xuất file PDF chuẩn A4 đẩy vào `00_BAO_CAO_KET_QUA_OUTBOX`
 *    và chuyển file gốc sang `00_LENH_DIEU_KHIEN_INBOX/PROCESSED`.
 * ============================================================================
 */

const CONFIG = {
  // Đã điền sẵn ID thư mục thực tế của Cô Thúy từ hệ thống:
  INBOX_FOLDER_ID: '151fBqzjfNFzlN0YnrwvEfjyt6gxeXDAQ',  // 00_LENH_DIEU_KHIEN_INBOX
  OUTBOX_FOLDER_ID: '1bZOe2dMRkB55v3phOgKLjFxs3xlL2tsE', // 00_BAO_CAO_KET_QUA_OUTBOX
  PRINT_QUEUE_FOLDER_ID: '1YUqDPtlgHLcNgEu1cRCRXxBIN6kEmN30', // 00_PRINT_QUEUE

  // Lấy API Key miễn phí tại https://aistudio.google.com/apikey và dán vào đây
  // (Hoặc lưu trong Project Settings -> Script Properties với tên GEMINI_API_KEY)
  DEFAULT_GEMINI_API_KEY: 'PASTE_YOUR_GEMINI_API_KEY_HERE',

  // Model xử lý (Có thể đổi thành 'gemini-4-carbon' hoặc 'gemini-2.5-flash')
  GEMINI_MODEL: 'gemini-2.5-flash'
};

function getApiKey_() {
  const propKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return (propKey && propKey.trim()) ? propKey.trim() : CONFIG.DEFAULT_GEMINI_API_KEY;
}

/**
 * 1. WEBHOOK NHẬN TRỰC TIẾP TỪ WEBAPP CỦA CÔ THÚY (0 CẦN TẢI FILE THỦ CÔNG)
 * Khi Cô Thúy bấm "Đẩy vào INBOX" trên WebApp, dữ liệu + ảnh gửi thẳng vào hàm này!
 */
function doPost(e) {
  try {
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = JSON.parse(rawBody);
    const inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
    const ticketId = data.ticket_id || ('INBOX_CLOUD_' + new Date().getTime());

    // 1. Lưu các ảnh đính kèm (nếu có) vào INBOX
    const savedImageIds = [];
    const attachments = data.attachments || [];
    for (let i = 0; i < attachments.length; i++) {
      const att = attachments[i];
      if (att.data_url && att.data_url.indexOf(',') !== -1) {
        const parts = att.data_url.split(',');
        const mimeMatch = parts[0].match(/:(.*?);/);
        const mimeType = mimeMatch ? mimeMatch[1] : (att.type || 'image/jpeg');
        const bytes = Utilities.base64Decode(parts[1]);
        const blob = Utilities.newBlob(bytes, mimeType, ticketId + '_P' + (i + 1) + '_' + (att.name || 'photo.jpg'));
        const f = inboxFolder.createFile(blob);
        savedImageIds.push(f.getId());
      }
    }

    // 2. Lưu vé JSON vào INBOX
    const cleanTicket = {
      ticket_id: ticketId,
      created_at: new Date().toISOString(),
      teacher: 'MASTER THÚY · Zalo: 0836 686 874',
      grade: data.grade || 6,
      ai_target: data.ai_target || 'CLOUD_GEMINI_24_7',
      request_prompt: data.request_prompt || 'Tạo bộ phiếu bài tập chuẩn Master Thúy',
      saved_image_ids: savedImageIds,
      status: 'PENDING'
    };

    const ticketFile = inboxFolder.createFile(
      ticketId + '.json',
      JSON.stringify(cleanTicket, null, 2),
      MimeType.PLAIN_TEXT
    );

    // 3. Nếu yêu cầu xử lý ngay lập tức, kích hoạt tạo phiếu luôn!
    let immediateResult = null;
    if (data.process_immediately !== false) {
      immediateResult = processSingleInboxTicket_(ticketFile);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        ticket_id: ticketId,
        saved_images: savedImageIds.length,
        processed: !!immediateResult,
        outbox_result: immediateResult
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Cho phép WebApp kiểm tra kết nối Cloud Daemon & lấy danh sách file mới trong OUTBOX
 */
function doGet(e) {
  try {
    const outboxFolder = DriveApp.getFolderById(CONFIG.OUTBOX_FOLDER_ID);
    const files = outboxFolder.getFiles();
    const recentOutbox = [];
    while (files.hasNext() && recentOutbox.length < 20) {
      const f = files.next();
      recentOutbox.push({
        id: f.getId(),
        name: f.getName(),
        url: f.getUrl(),
        updated_at: f.getLastUpdated().toISOString()
      });
    }
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        service: 'MASTER_THUY_24_7_CLOUD_DAEMON',
        outbox_files: recentOutbox
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 2. HÀM TỰ THỨC DẬY MỖI PHÚT (CHẠY NGẦM 24/7 TRÊN ĐÁM MÂY)
 * Quét toàn bộ file .json hoặc ảnh mới cô Thúy bỏ vào 00_LENH_DIEU_KHIEN_INBOX
 */
function autoScanInboxEveryMinute() {
  const inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  const processedFolder = getOrCreateSubfolder_(inboxFolder, 'PROCESSED');

  const files = inboxFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fname = file.getName();

    // Chỉ xử lý các file vé .json chưa xử lý
    if (fname.toLowerCase().endsWith('.json') && fname.indexOf('PROCESSED_') !== 0) {
      try {
        processSingleInboxTicket_(file, processedFolder);
      } catch (err) {
        Logger.log('Error processing ticket ' + fname + ': ' + err);
      }
    }
  }
}

/**
 * 3. XỬ LÝ 1 VÉ YÊU CẦU -> GỌI GEMINI -> TẠO DOCS & PDF BẢN HỌC SINH + GIÁO VIÊN VÀO OUTBOX
 */
function processSingleInboxTicket_(ticketFile, processedFolderOpt) {
  const inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  const outboxFolder = DriveApp.getFolderById(CONFIG.OUTBOX_FOLDER_ID);
  const processedFolder = processedFolderOpt || getOrCreateSubfolder_(inboxFolder, 'PROCESSED');

  const ticketData = JSON.parse(ticketFile.getBlob().getDataAsString('UTF-8'));
  const ticketId = ticketData.ticket_id || 'MT_' + new Date().getTime();
  const promptText = ticketData.request_prompt || 'Tạo đề kiểm tra Tiếng Anh Global Success';
  const grade = ticketData.grade || 6;

  // Chuẩn bị danh sách ảnh (Multimodal) gửi cho Gemini
  const parts = [];
  const systemInstruction = buildMasterThuySystemPrompt_(grade, promptText);
  parts.push({ text: systemInstruction });

  const imageIds = ticketData.saved_image_ids || [];
  for (let i = 0; i < imageIds.length; i++) {
    try {
      const imgFile = DriveApp.getFileById(imageIds[i]);
      const blob = imgFile.getBlob();
      parts.push({
        inline_data: {
          mime_type: blob.getContentType() || 'image/jpeg',
          data: Utilities.base64Encode(blob.getBytes())
        }
      });
      imgFile.moveTo(processedFolder);
    } catch (e) {}
  }

  // Gọi Gemini API để sinh JSON cấu trúc Bản Học Sinh + Bản Giáo Viên
  const generatedJson = callGeminiStructuredWorksheet_(parts);

  // Xuất ra 2 cặp file (Google Docs + PDF A4) trong 00_BAO_CAO_KET_QUA_OUTBOX
  const exportRes = createStudentAndTeacherDocsInOutbox_(outboxFolder, ticketId, grade, generatedJson);

  // Chuyển file vé sang thư mục PROCESSED để không lặp lại
  ticketData.status = 'COMPLETED';
  ticketData.completed_at = new Date().toISOString();
  ticketData.outbox_files = exportRes;
  ticketFile.setContent(JSON.stringify(ticketData, null, 2));
  ticketFile.moveTo(processedFolder);

  return exportRes;
}

/**
 * 4. PROMPT CHUẨN SƯ PHẠM MASTER THÚY CHO GEMINI
 */
function buildMasterThuySystemPrompt_(grade, userRequest) {
  return `Bạn là Trợ lý Sư phạm Tiếng Anh cao cấp của MASTER THÚY (Thạc sĩ · Giảng viên Tiếng Anh · Zalo: 0836 686 874).
Nhiệm vụ: Dựa vào yêu cầu và hình ảnh đính kèm (nếu có), hãy biên soạn bộ phiếu bài tập Tiếng Anh Khối ${grade} theo chuẩn SGK Global Success.
YÊU CẦU CỦA CÔ THÚY: "${userRequest}"

QUY TẮC BẮT BUỘC:
1. Nếu có ảnh chụp đề mẫu: Giữ nguyên cấu trúc và độ khó tương đương nhưng PHẢI viết câu hỏi và ngữ liệu mới 100%, không chép nguyên xi câu gốc.
2. Khóa chặt ngữ pháp trong phạm vi Khối ${grade} Global Success, tuyệt đối không đưa ngữ pháp vượt cấp.
3. Trả về kết quả DUY NHẤT dưới dạng JSON hợp lệ theo schema sau (không kèm markdown):
{
  "title": "Tên phiếu bài tập (VD: UNIT 3 - MY FRIENDS - ADVANCED WORKSHEET)",
  "grade": ${grade},
  "duration_minutes": 45,
  "remember_patterns": ["Mẫu công thức 1", "Mẫu công thức 2", "Mẫu công thức 3"],
  "questions": [
    {
      "num": 1,
      "part": "PART A · PRONUNCIATION",
      "stem": "Choose the word whose underlined part is pronounced differently:",
      "options": ["A. पेंcil", "B. <u>b</u>ook", "C. <u>p</u>aper", "D. <u>p</u>en"],
      "answer": "B",
      "trap_code": "TRAP_PRON_PB",
      "explanation": "Dấu hiệu -> Quy tắc -> Đáp án B phát âm /b/, các từ còn lại phát âm /p/."
    }
  ],
  "writing_prompt": "Write a paragraph (70-90 words) about...",
  "sample_paragraph": "Đoạn văn mẫu hoàn chỉnh..."
}`;
}

/**
 * 5. GỌI GEMINI API (HỖ TRỢ ĐỌC NHIỀU ẢNH + TRẢ VỀ JSON)
 */
function callGeminiStructuredWorksheet_(partsArray) {
  const apiKey = getApiKey_();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [{ parts: partsArray }],
    generationConfig: {
      temperature: 0.3,
      responseMimeType: 'application/json'
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(resp.getContentText());
  const rawText = json.candidates[0].content.parts[0].text;
  return JSON.parse(rawText);
}

/**
 * 6. TẠO FILE GOOGLE DOCS & XUẤT PDF CHUẨN A4 VÀO 00_BAO_CAO_KET_QUA_OUTBOX
 */
function createStudentAndTeacherDocsInOutbox_(outboxFolder, ticketId, grade, wsData) {
  const title = wsData.title || ('PHIEU_BAI_TAP_KHOI_' + grade);
  const safeName = title.replace(/[^a-zA-Z0-9_\-\u00C0-\u1EF9 ]/g, '').trim();

  // --- A. TẠO BẢN HỌC SINH (STUDENT'S COPY) ---
  const stuDoc = DocumentApp.create(ticketId + '_STUDENT_' + safeName);
  const stuBody = stuDoc.getBody();
  stuBody.setMarginTop(48);
  stuBody.setMarginBottom(46);
  stuBody.setMarginLeft(60);
  stuBody.setMarginRight(58);

  stuBody.appendParagraph('MASTER THÚY · Thạc sĩ · Giảng viên Tiếng Anh · Zalo: 0836 686 874')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);
  stuBody.appendParagraph(title.toUpperCase() + ' — KHỐI ' + grade + ' (STUDENT\'S COPY | CODE: 101)')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  stuBody.appendParagraph('Full name: ....................................................   Class: ............   Date: ..../..../2026\n');

  if (wsData.remember_patterns && wsData.remember_patterns.length) {
    stuBody.appendParagraph('💡 REMEMBER THE PATTERNS: ' + wsData.remember_patterns.join('  |  ')).setBold(true);
  }

  const questions = wsData.questions || [];
  let currentPart = '';
  for (let i = 0; i < questions.length; i++) {
    const q = questions[i];
    if (q.part && q.part !== currentPart) {
      currentPart = q.part;
      stuBody.appendParagraph('\n' + currentPart).setBold(true);
    }
    stuBody.appendParagraph('Question ' + (i + 1) + ': ' + String(q.stem || '').replace(/<[^>]+>/g, ''));
    if (q.options && q.options.length) {
      stuBody.appendParagraph('   ' + q.options.map(o => String(o).replace(/<[^>]+>/g, '')).join('      '));
    }
  }

  if (wsData.writing_prompt) {
    stuBody.appendParagraph('\nPART F · WRITING CHALLENGE').setBold(true);
    stuBody.appendParagraph(wsData.writing_prompt);
    stuBody.appendParagraph('..........................................................................................................................................\n..........................................................................................................................................');
  }
  stuDoc.saveAndClose();

  // Lưu file Docs + PDF Bản Học Sinh vào OUTBOX
  const stuDocFile = DriveApp.getFileById(stuDoc.getId());
  stuDocFile.moveTo(outboxFolder);
  const stuPdfBlob = stuDocFile.getAs(MimeType.PDF).setName(ticketId + '_STUDENT_' + safeName + '.pdf');
  const stuPdfFile = outboxFolder.createFile(stuPdfBlob);

  // --- B. TẠO BẢN GIÁO VIÊN (TEACHER'S COPY + ZEBRA KEY) ---
  const teaDoc = DocumentApp.create(ticketId + '_TEACHER_KEY_' + safeName);
  const teaBody = teaDoc.getBody();
  teaBody.appendParagraph('MASTER THÚY · Thạc sĩ · Giảng viên Tiếng Anh · Zalo: 0836 686 874')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);
  teaBody.appendParagraph(title.toUpperCase() + ' — TEACHER\'S COPY (ANSWER KEY & EXPLANATIONS)')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const tableRows = [['No.', 'Key (101)', 'Explanation (Signal → Rule → Answer/Trap)']];
  for (let i = 0; i < questions.length; i++) {
    const q = questions[i];
    tableRows.push([
      String(i + 1),
      String(q.answer || 'A'),
      String((q.explanation || '') + (q.trap_code ? ' [' + q.trap_code + ']' : ''))
    ]);
  }
  teaBody.appendTable(tableRows);

  if (wsData.sample_paragraph) {
    const wc = String(wsData.sample_paragraph).trim().split(/\s+/).length;
    teaBody.appendParagraph('\nSAMPLE PARAGRAPH (' + wc + ' words):').setBold(true);
    teaBody.appendParagraph(wsData.sample_paragraph);
  }
  teaDoc.saveAndClose();

  const teaDocFile = DriveApp.getFileById(teaDoc.getId());
  teaDocFile.moveTo(outboxFolder);
  const teaPdfBlob = teaDocFile.getAs(MimeType.PDF).setName(ticketId + '_TEACHER_KEY_' + safeName + '.pdf');
  const teaPdfFile = outboxFolder.createFile(teaPdfBlob);

  return {
    student_pdf_url: stuPdfFile.getUrl(),
    student_doc_url: stuDocFile.getUrl(),
    teacher_pdf_url: teaPdfFile.getUrl(),
    teacher_doc_url: teaDocFile.getUrl()
  };
}

/**
 * 7. HÀM CÀI ĐẶT TRIGGER TỰ ĐỘNG CHẠY NGẦM 24/7 (CHỈ CẦN BẤM RUN 1 LẦN DUY NHẤT)
 */
function setup24x7AutoTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'autoScanInboxEveryMinute') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  ScriptApp.newTrigger('autoScanInboxEveryMinute')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('✅ Đã kích hoạt chế độ chạy ngầm 24/7: Quét 00_LENH_DIEU_KHIEN_INBOX mỗi 1 phút!');
}

function getOrCreateSubfolder_(parentFolder, subName) {
  const iter = parentFolder.getFoldersByName(subName);
  if (iter.hasNext()) return iter.next();
  return parentFolder.createFolder(subName);
}
