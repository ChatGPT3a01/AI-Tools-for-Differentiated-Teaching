// ============================================================
//  整合版：AI 工具協助差異化教學（3 合 1）
// ============================================================
//
//  本檔案整合三個工具的完整功能，只需貼這一份程式碼：
//
//  工具① 段考 AI 審題助手
//    - 分析段考試題 → 雙向細目表 + 難易度分析 + 檢核建議
//
//  工具② 動態評量產生器
//    - 知識點 → AI 產生三層提示題目 → Google Form → 自動計分
//
//  工具③ AI 學生回饋系統
//    - Rubric 模板 + Google Drive 作品連結 → AI 回饋 + 自動計分
//    - 全班統計報表 + 寄信給學生
//
//  【安裝步驟（5 分鐘）】
//  1. 新增 Google Sheets → 擴充功能 → Apps Script
//  2. 全選貼上此程式碼 → 儲存
//  3. 專案設定 → 指令碼屬性 → 新增 GEMINI_API_KEY
//  4. 執行 onOpen → 授權 → 重新整理
//
//  【API 金鑰】https://aistudio.google.com/apikey
//
// ============================================================

// ──────────────── 共用常數與工具函式 ────────────────

const GEMINI_MODEL = 'gemini-2.5-flash';

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🤖 AI 審題助手')
    .addItem('📊 分析段考試題', 'tool1_analyzeExam')
    .addItem('ℹ️ 使用說明', 'tool1_showHelp')
    .addToUi();

  ui.createMenu('🎯 動態評量')
    .addItem('📝 產生動態評量題目', 'tool2_generateQuestions')
    .addItem('📋 建立 Google Form', 'tool2_createForm')
    .addItem('📊 計算成績', 'tool2_calculateScores')
    .addToUi();

  ui.createMenu('💬 AI 回饋')
    .addItem('📋 初始化欄位標頭', 'tool3_setupHeaders')
    .addItem('📝 產生 Rubric 模板參考', 'tool3_createRubricTemplates')
    .addSeparator()
    .addItem('🔄 產生個人化回饋（含評分）', 'tool3_generateFeedback')
    .addItem('📧 批次產生全部回饋', 'tool3_generateAllFeedback')
    .addSeparator()
    .addItem('📊 產生全班統計報表', 'tool3_generateClassReport')
    .addItem('✉️ 寄送回饋信件給學生', 'tool3_sendFeedbackEmails')
    .addToUi();
}

/** 取得 API 金鑰，若未設定則彈窗提醒 */
function getApiKeyOrAlert_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('❌ 尚未設定 API 金鑰！\n\n請到「專案設定」→「指令碼屬性」中新增 GEMINI_API_KEY');
    return null;
  }
  return apiKey;
}

/** 共用 Gemini 純文字 API 呼叫 */
function callGeminiText_(prompt, temperature, maxOutputTokens) {
  const apiKey = getApiKeyOrAlert_();
  if (!apiKey) return null;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: temperature || 0.3,
      maxOutputTokens: maxOutputTokens || 2048
    }
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log('Gemini HTTP ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 500));
      return null;
    }

    const json = JSON.parse(res.getContentText());
    if (!json.candidates || !json.candidates.length) return null;
    return json.candidates[0].content.parts[0].text || null;
  } catch (e) {
    Logger.log('Gemini 呼叫錯誤：' + e.message);
    return null;
  }
}

/** 從 AI 回傳文字中提取 JSON 物件 */
function extractJsonObject_(text) {
  if (!text) return null;
  const cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '');
  const match = cleaned.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch (e) {
    Logger.log('JSON 解析失敗：' + e.message);
    return null;
  }
}


// ============================================================
//  工具 ① 段考 AI 審題助手
// ============================================================

function tool1_showHelp() {
  const html = HtmlService.createHtmlOutput(
    '<h3>段考 AI 審題助手 使用說明</h3>' +
    '<ol>' +
    '<li>在 Sheet1 的 A 欄貼上段考試題（每列一題）</li>' +
    '<li>點選選單「AI 審題助手」→「分析段考試題」</li>' +
    '<li>等待 AI 分析完成（約 30 秒～1 分鐘）</li>' +
    '<li>結果會自動產生在新的工作表中</li>' +
    '</ol>' +
    '<p><b>產出內容：</b></p>' +
    '<ul>' +
    '<li>雙向細目表（知識向度 × 認知層次）</li>' +
    '<li>難易度分析（簡單/中等/困難比例）</li>' +
    '<li>覆蓋率檢核與改進建議</li>' +
    '</ul>'
  ).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '使用說明');
}

function tool1_analyzeExam() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('請先在 A 欄貼上段考試題（從 A2 開始）');
    return;
  }

  const questions = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(function(r) { return r[0]; })
    .filter(function(q) { return q.toString().trim() !== ''; });

  if (questions.length === 0) {
    ui.alert('未偵測到試題，請確認 A 欄已貼上題目。');
    return;
  }

  ui.alert('已偵測到 ' + questions.length + ' 題，開始 AI 分析...\n（約需 30 秒～1 分鐘，請稍候）');

  const examText = questions.map(function(q, i) { return '第' + (i + 1) + '題：' + q; }).join('\n\n');

  const prompt = '你是一位專業的高中教育評量專家。請分析以下段考試題，並產出：\n\n'
    + '1. **雙向細目表**（JSON 格式）\n'
    + '   - 列：各知識單元/主題\n'
    + '   - 欄：認知層次（記憶、理解、應用、分析、評鑑、創造）\n'
    + '   - 值：對應的題號\n\n'
    + '2. **難易度分析**（JSON 格式）\n'
    + '   - 將每題分類為「簡單」「中等」「困難」\n'
    + '   - 統計各難度的題數與百分比\n\n'
    + '3. **覆蓋率檢核**\n'
    + '   - 哪些單元/主題出太多題？\n'
    + '   - 哪些單元/主題沒有出到？\n'
    + '   - 難度分布是否均衡？\n\n'
    + '4. **改進建議**\n'
    + '   - 具體建議如何調整出題\n\n'
    + '請用以下 JSON 格式回傳（確保是合法 JSON）：\n'
    + '{\n'
    + '  "雙向細目表": {\n'
    + '    "單元列表": ["單元A", "單元B"],\n'
    + '    "認知層次": ["記憶", "理解", "應用", "分析", "評鑑", "創造"],\n'
    + '    "對應題號": {\n'
    + '      "單元A": {"記憶": [1,2], "理解": [3]}\n'
    + '    }\n'
    + '  },\n'
    + '  "難易度分析": {\n'
    + '    "各題難度": {"1": "中等", "2": "簡單"},\n'
    + '    "統計": {"簡單": {"題數": 5, "百分比": "25%"}, "中等": {"題數": 10, "百分比": "50%"}, "困難": {"題數": 5, "百分比": "25%"}}\n'
    + '  },\n'
    + '  "覆蓋率檢核": {\n'
    + '    "出題偏多": ["..."],\n'
    + '    "未覆蓋": ["..."],\n'
    + '    "難度均衡性": "..."\n'
    + '  },\n'
    + '  "改進建議": ["建議1", "建議2"]\n'
    + '}\n\n'
    + '以下是段考試題：\n\n' + examText;

  const text = callGeminiText_(prompt, 0.2, 16384);
  const analysis = extractJsonObject_(text);
  if (!analysis) {
    ui.alert('❌ 分析失敗，請檢查 API 金鑰與輸入資料。');
    return;
  }

  tool1_outputResults_(analysis);
  ui.alert('✅ 分析完成！請查看新產生的工作表。');
}

function tool1_outputResults_(analysis) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 雙向細目表 ---
  var sheetTable = ss.getSheetByName('雙向細目表');
  if (sheetTable) sheetTable.clear();
  else sheetTable = ss.insertSheet('雙向細目表');

  const table = analysis['雙向細目表'] || {};
  const fixedLevels = ['記憶', '理解', '應用', '分析', '評鑑', '創造'];
  const headers = ['單元 \\ 認知層次'].concat(fixedLevels);
  sheetTable.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheetTable.setColumnWidth(1, 160);
  for (var c = 2; c <= headers.length; c++) {
    sheetTable.setColumnWidth(c, 80);
  }

  const units = table['單元列表'] || [];
  units.forEach(function(unit, i) {
    var row = [unit];
    fixedLevels.forEach(function(level) {
      var mapping = table['對應題號'] || {};
      var unitMap = mapping[unit] || {};
      var nums = unitMap[level] || [];
      row.push(nums.length > 0 ? nums.join(', ') : '—');
    });
    sheetTable.getRange(i + 2, 1, 1, row.length).setValues([row]);
  });

  var dataRows = units.length;
  if (dataRows > 0) {
    sheetTable.getRange(2, 1, dataRows, 1).setHorizontalAlignment('left');
    sheetTable.getRange(2, 2, dataRows, fixedLevels.length).setHorizontalAlignment('center');
  }

  // --- 難易度分析 ---
  var sheetDiff = ss.getSheetByName('難易度分析');
  if (sheetDiff) sheetDiff.clear();
  else sheetDiff = ss.insertSheet('難易度分析');

  const diff = analysis['難易度分析'] || {};
  const byQ = diff['各題難度'] || {};

  sheetDiff.getRange(1, 1).setValue('各題難度分析').setFontWeight('bold').setFontSize(12);
  sheetDiff.getRange(2, 1, 1, 2).setValues([['題號', '難度']])
    .setBackground('#34A853').setFontColor('white').setFontWeight('bold');

  var qKeys = Object.keys(byQ);
  qKeys.forEach(function(q, i) {
    sheetDiff.getRange(i + 3, 1, 1, 2).setValues([['第 ' + q + ' 題', byQ[q]]]);
  });

  var statsRow = qKeys.length + 5;
  sheetDiff.getRange(statsRow, 1).setValue('統計摘要').setFontWeight('bold').setFontSize(12);
  sheetDiff.getRange(statsRow + 1, 1, 1, 3).setValues([['難度', '題數', '百分比']])
    .setBackground('#FBBC04').setFontWeight('bold');

  var stat = diff['統計'] || {};
  ['簡單', '中等', '困難'].forEach(function(k, idx) {
    var s = stat[k] || {};
    sheetDiff.getRange(statsRow + 2 + idx, 1, 1, 3).setValues([[k, s['題數'] || 0, s['百分比'] || '0%']]);
  });
  sheetDiff.autoResizeColumns(1, 3);

  // --- 檢核與建議 ---
  var sheetAdvice = ss.getSheetByName('檢核與建議');
  if (sheetAdvice) sheetAdvice.clear();
  else sheetAdvice = ss.insertSheet('檢核與建議');

  const check = analysis['覆蓋率檢核'] || {};
  const advice = analysis['改進建議'] || [];

  sheetAdvice.getRange(1, 1).setValue('覆蓋率檢核').setFontWeight('bold').setFontSize(12);

  var r = 2;
  sheetAdvice.getRange(r, 1).setValue('出題偏多的單元：').setFontWeight('bold');
  sheetAdvice.getRange(r, 2).setValue((check['出題偏多'] || ['無']).join('、'));
  r++;
  sheetAdvice.getRange(r, 1).setValue('未覆蓋的單元：').setFontWeight('bold');
  sheetAdvice.getRange(r, 2).setValue((check['未覆蓋'] || ['無']).join('、'));
  r++;
  sheetAdvice.getRange(r, 1).setValue('難度均衡性：').setFontWeight('bold');
  sheetAdvice.getRange(r, 2).setValue(check['難度均衡性'] || '—');
  r += 2;

  sheetAdvice.getRange(r, 1).setValue('改進建議').setFontWeight('bold').setFontSize(12);
  r++;
  advice.forEach(function(a, i) {
    sheetAdvice.getRange(r + i, 1).setValue((i + 1) + '. ' + a);
  });
  sheetAdvice.autoResizeColumns(1, 2);
}


// ============================================================
//  工具 ② 動態評量產生器
// ============================================================

function tool2_generateQuestions() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('請先在 Sheet1 填入知識點（A:科目, B:單元, C:知識點）');
    return;
  }
  if (!getApiKeyOrAlert_()) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var knowledgePoints = data
    .filter(function(row) { return row[2].toString().trim() !== ''; })
    .map(function(row) { return { subject: row[0], unit: row[1], point: row[2] }; });

  if (knowledgePoints.length === 0) {
    ui.alert('C 欄沒有可用知識點，請先填入再執行。');
    return;
  }
  if (knowledgePoints.length > 30) {
    ui.alert('⚠️ 知識點數量過多（' + knowledgePoints.length + '），建議一次不超過 30 個，以免執行超時。');
    return;
  }

  ui.alert('開始為 ' + knowledgePoints.length + ' 個知識點產生動態評量題目...\n（每個知識點約需 10 秒）');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resultSheet = ss.getSheetByName('題目與提示');
  if (resultSheet) resultSheet.clear();
  else resultSheet = ss.insertSheet('題目與提示');

  var headers = ['知識點', '題目', '正確答案',
    '提示1（方向提示）', '提示1後答案',
    '提示2（給公式）', '提示2後答案',
    '提示3（帶數字）', '提示3後答案'];
  resultSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold');

  var currentRow = 2;
  knowledgePoints.forEach(function(kp) {
    var result = tool2_generateOneQuestion_(kp);
    if (result) {
      resultSheet.getRange(currentRow, 1, 1, headers.length).setValues([result]);
      currentRow++;
    }
  });

  resultSheet.autoResizeColumns(1, headers.length);
  ui.alert('✅ 已產生 ' + (currentRow - 2) + ' 道動態評量題目！\n請查看「題目與提示」工作表。');
}

function tool2_generateOneQuestion_(kp) {
  var prompt = '你是一位高中' + kp.subject + '老師，請針對「' + kp.unit + ' - ' + kp.point + '」設計一道動態評量題目。\n\n'
    + '請提供以下內容（用 JSON 格式回覆）：\n\n'
    + '{\n'
    + '  "題目": "一道計算題或應用題（含具體數字，約 50-80 字）",\n'
    + '  "正確答案": "完整的正確答案",\n'
    + '  "提示1_方向提示": "不給公式，只提示解題方向",\n'
    + '  "提示1後答案": "看了提示1之後的正確答案",\n'
    + '  "提示2_給公式": "給出需要用到的公式",\n'
    + '  "提示2後答案": "看了提示2之後的正確答案",\n'
    + '  "提示3_帶數字": "直接幫學生把數字代入公式，只差最後一步",\n'
    + '  "提示3後答案": "最後一步的答案"\n'
    + '}\n\n'
    + '要求：\n'
    + '- 題目要有具體數字，學生可以實際計算\n'
    + '- 三層提示要有明確的層次差異\n'
    + '- 提示越多，解題越容易\n'
    + '- 即使看了全部提示，學生仍需要做最後一步才能得到答案';

  var text = callGeminiText_(prompt, 0.3, 2048);
  var json = extractJsonObject_(text);
  if (!json) return null;

  return [
    kp.point,
    json['題目'] || '',
    json['正確答案'] || '',
    json['提示1_方向提示'] || '',
    json['提示1後答案'] || '',
    json['提示2_給公式'] || '',
    json['提示2後答案'] || '',
    json['提示3_帶數字'] || '',
    json['提示3後答案'] || ''
  ];
}

function tool2_createForm() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('題目與提示');

  if (!sheet) {
    ui.alert('請先產生題目（點選「產生動態評量題目」）');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('沒有可用的題目。');
    return;
  }

  var form = FormApp.create('動態評量 - ' + new Date().toLocaleDateString('zh-TW'));
  form.setDescription(
    '這是一份動態評量，答錯時會獲得提示，幫助你一步步解題。\n' +
    '每題滿分 5 分，使用提示會扣分，但你一定能得到分數！\n\n' +
    '計分規則：不用提示＝5分 ／ 1個提示＝4分 ／ 2個提示＝3分 ／ 3個提示＝1分'
  );
  form.setIsQuiz(false);

  // 學生資訊欄位
  form.addTextItem().setTitle('班級').setRequired(true)
    .setHelpText('例如：301、高一甲');
  form.addTextItem().setTitle('座號').setRequired(true)
    .setHelpText('例如：5、05');
  form.addTextItem().setTitle('姓名').setRequired(true);

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var numQuestions = data.length;

  // 第一輪：建立所有頁面和項目
  var questionPages = [];
  var navItems = [];

  data.forEach(function(row, idx) {
    var qNum = idx + 1;
    var point = row[0], question = row[1], hint1 = row[3], hint2 = row[5], hint3 = row[7];

    var questionPage = form.addPageBreakItem()
      .setTitle('第 ' + qNum + ' 題（知識點：' + point + '）');
    questionPages.push(questionPage);

    form.addParagraphTextItem()
      .setTitle(question)
      .setHelpText('請直接作答。如果不確定，可以先試試看！')
      .setRequired(true);

    var needHint = form.addMultipleChoiceItem()
      .setTitle('答完了嗎？需要提示嗎？')
      .setHelpText('不使用提示得 5 分；每使用一層提示會扣分')
      .setRequired(true);

    var hint1Page = form.addPageBreakItem()
      .setTitle('第 ' + qNum + ' 題 - 提示 1');
    form.addSectionHeaderItem().setTitle('方向提示：' + hint1);
    form.addParagraphTextItem().setTitle('看了提示後，請再試一次：').setRequired(false);
    var needHint2 = form.addMultipleChoiceItem()
      .setTitle('還需要更多提示嗎？')
      .setHelpText('再使用一層提示會再扣 1 分')
      .setRequired(true);

    var hint2Page = form.addPageBreakItem()
      .setTitle('第 ' + qNum + ' 題 - 提示 2');
    form.addSectionHeaderItem().setTitle('公式提示：' + hint2);
    form.addParagraphTextItem().setTitle('有了公式，請再算一次：').setRequired(false);
    var needHint3 = form.addMultipleChoiceItem()
      .setTitle('還需要最後一個提示嗎？')
      .setHelpText('這是最後一層提示，使用後得 1 分')
      .setRequired(true);

    var hint3Page = form.addPageBreakItem()
      .setTitle('第 ' + qNum + ' 題 - 提示 3');
    form.addSectionHeaderItem().setTitle('帶入數字：' + hint3);
    form.addParagraphTextItem().setTitle('數字都代好了，最後一步是？').setRequired(false);

    navItems.push({
      needHint: needHint,
      needHint2: needHint2,
      needHint3: needHint3,
      hint3Page: hint3Page
    });
  });

  // 第二輪：設定分支導航
  navItems.forEach(function(nav, idx) {
    var nextTarget = (idx + 1 < numQuestions)
      ? questionPages[idx + 1]
      : FormApp.PageNavigationType.SUBMIT;

    nav.needHint.setChoices([
      nav.needHint.createChoice('不需要，我已經會了', nextTarget),
      nav.needHint.createChoice('需要提示', FormApp.PageNavigationType.CONTINUE)
    ]);
    nav.needHint2.setChoices([
      nav.needHint2.createChoice('不需要了，我會了', nextTarget),
      nav.needHint2.createChoice('再給一點提示', FormApp.PageNavigationType.CONTINUE)
    ]);
    nav.needHint3.setChoices([
      nav.needHint3.createChoice('不需要了', nextTarget),
      nav.needHint3.createChoice('給我最後提示', FormApp.PageNavigationType.CONTINUE)
    ]);

    if (idx + 1 < numQuestions) {
      nav.hint3Page.setGoToPage(questionPages[idx + 1]);
    } else {
      nav.hint3Page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
    }
  });

  // 記錄 Form URL + 題目數量
  var formUrl = form.getPublishedUrl();
  var editUrl = form.getEditUrl();

  var urlSheet = ss.getSheetByName('表單連結');
  if (urlSheet) urlSheet.clear();
  else urlSheet = ss.insertSheet('表單連結');

  urlSheet.getRange(1, 1).setValue('學生作答連結：').setFontWeight('bold');
  urlSheet.getRange(1, 2).setValue(formUrl);
  urlSheet.getRange(2, 1).setValue('編輯連結：').setFontWeight('bold');
  urlSheet.getRange(2, 2).setValue(editUrl);
  urlSheet.getRange(3, 1).setValue('題目數量：').setFontWeight('bold');
  urlSheet.getRange(3, 2).setValue(numQuestions);
  urlSheet.autoResizeColumns(1, 2);

  ui.alert('✅ Google Form 已建立！\n\n學生連結：' + formUrl + '\n\n（連結也已記錄在「表單連結」工作表）');
}

function tool2_calculateScores() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 讀取表單資訊
  var urlSheet = ss.getSheetByName('表單連結');
  if (!urlSheet) {
    ui.alert('❌ 找不到「表單連結」工作表，請先建立 Google Form。');
    return;
  }

  var editUrl = urlSheet.getRange(2, 2).getValue();
  var numQuestions = urlSheet.getRange(3, 2).getValue();

  if (!editUrl || !numQuestions) {
    ui.alert('❌ 表單連結資訊不完整，請重新建立 Google Form。');
    return;
  }

  // 2. 開啟表單並取得回覆
  var form;
  try {
    form = FormApp.openByUrl(editUrl);
  } catch (e) {
    ui.alert('❌ 無法開啟表單，請確認表單連結是否正確。\n\n錯誤：' + e.message);
    return;
  }

  var responses = form.getResponses();
  if (responses.length === 0) {
    ui.alert('⚠️ 目前還沒有學生回覆，請等學生作答後再計算成績。');
    return;
  }

  // 3. 計算每位學生成績
  var studentResults = [];

  responses.forEach(function(response) {
    var itemResponses = response.getItemResponses();

    var className = '', seatNo = '', studentName = '';
    var mcAnswers = [];

    itemResponses.forEach(function(ir) {
      var itemType = ir.getItem().getType();
      var title = ir.getItem().getTitle();
      var answer = ir.getResponse();

      if (title === '班級') {
        className = answer;
      } else if (title === '座號') {
        seatNo = answer;
      } else if (title === '姓名') {
        studentName = answer;
      } else if (itemType === FormApp.ItemType.MULTIPLE_CHOICE) {
        mcAnswers.push(answer);
      }
    });

    // 每 3 個 MC 回答 = 1 道題
    var scores = [];
    var hints = [];

    for (var q = 0; q < numQuestions; q++) {
      var baseIdx = q * 3;
      var ans1 = mcAnswers[baseIdx];
      var ans2 = mcAnswers[baseIdx + 1];
      var ans3 = mcAnswers[baseIdx + 2];

      var score, hintCount;

      if (!ans1 || ans1.indexOf('不需要') >= 0) {
        score = 5; hintCount = 0;
      } else if (!ans2 || ans2.indexOf('不需要') >= 0) {
        score = 4; hintCount = 1;
      } else if (!ans3 || ans3.indexOf('不需要') >= 0) {
        score = 3; hintCount = 2;
      } else {
        score = 1; hintCount = 3;
      }

      scores.push(score);
      hints.push(hintCount);
    }

    var totalScore = scores.reduce(function(a, b) { return a + b; }, 0);
    var maxScore = numQuestions * 5;
    var percentage = Math.round(totalScore / maxScore * 100);

    var indicator;
    if (percentage >= 90) indicator = '精熟';
    else if (percentage >= 70) indicator = '接近精熟';
    else if (percentage >= 40) indicator = '發展中';
    else indicator = '需補強';

    var row = [className, seatNo, studentName];
    for (var i = 0; i < numQuestions; i++) {
      row.push(scores[i]);
      row.push(hints[i] + ' 個提示');
    }
    row.push(totalScore, maxScore, percentage + '%', indicator);

    studentResults.push({ row: row, scores: scores, percentage: percentage });
  });

  // 4. 建立成績報表
  var reportSheet = ss.getSheetByName('成績報表');
  if (reportSheet) reportSheet.clear();
  else reportSheet = ss.insertSheet('成績報表');

  var headers = ['班級', '座號', '姓名'];
  for (var q = 1; q <= numQuestions; q++) {
    headers.push('Q' + q + '得分');
    headers.push('Q' + q + '提示');
  }
  headers.push('總分', '滿分', '百分比', '能力指標');

  reportSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold');

  if (studentResults.length > 0) {
    var allRows = studentResults.map(function(s) { return s.row; });
    reportSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);

    // 得分欄位上色
    var scoreColorMap = { 5: '#34A853', 4: '#A8DAB5', 3: '#FBBC04', 1: '#F4A261' };

    for (var r = 0; r < studentResults.length; r++) {
      var rowNum = r + 2;
      for (var q = 0; q < numQuestions; q++) {
        var colNum = 4 + q * 2;
        var score = studentResults[r].scores[q];
        var color = scoreColorMap[score] || '#FFFFFF';
        reportSheet.getRange(rowNum, colNum).setBackground(color);
        if (score === 5) {
          reportSheet.getRange(rowNum, colNum).setFontColor('white');
        }
      }
    }

    // 能力指標上色
    var indicatorColors = { '精熟': '#34A853', '接近精熟': '#4285F4', '發展中': '#FBBC04', '需補強': '#EA4335' };
    for (var r2 = 0; r2 < studentResults.length; r2++) {
      var indicatorCol = headers.length;
      var indicatorVal = studentResults[r2].row[indicatorCol - 1];
      var iColor = indicatorColors[indicatorVal] || '#FFFFFF';
      reportSheet.getRange(r2 + 2, indicatorCol).setBackground(iColor)
        .setFontColor('white').setFontWeight('bold');
    }

    // 全班平均列
    var avgRow = ['', '', '全班平均'];
    var totalSumAll = 0;
    var maxScoreAll = numQuestions * 5;

    for (var q2 = 0; q2 < numQuestions; q2++) {
      var qSum = 0;
      for (var r3 = 0; r3 < studentResults.length; r3++) {
        qSum += studentResults[r3].scores[q2];
      }
      var qAvg = Math.round(qSum / studentResults.length * 10) / 10;
      totalSumAll += qSum;
      avgRow.push(qAvg);
      avgRow.push('');
    }

    var totalAvg = Math.round(totalSumAll / studentResults.length * 10) / 10;
    var avgPercentage = Math.round(totalAvg / maxScoreAll * 100);

    var avgIndicator;
    if (avgPercentage >= 90) avgIndicator = '精熟';
    else if (avgPercentage >= 70) avgIndicator = '接近精熟';
    else if (avgPercentage >= 40) avgIndicator = '發展中';
    else avgIndicator = '需補強';

    avgRow.push(totalAvg, maxScoreAll, avgPercentage + '%', avgIndicator);

    var avgRowNum = studentResults.length + 2;
    reportSheet.getRange(avgRowNum, 1, 1, headers.length).setValues([avgRow])
      .setBackground('#E8EAED').setFontWeight('bold');
  }

  reportSheet.autoResizeColumns(1, headers.length);

  ui.alert('✅ 成績報表已產生！\n\n' +
    '共 ' + studentResults.length + ' 位學生，' + numQuestions + ' 道題目\n' +
    '請查看「成績報表」工作表。');
}


// ============================================================
//  工具 ③ AI 學生回饋系統
// ============================================================

// ---------- 初始化欄位標頭 ----------
function tool3_setupHeaders() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var headers = [
    '學生姓名', 'Email', '評量形式', 'Rubric 評分規準',
    '學生作品（貼連結）', 'AI 回饋（自動）', '面向1', '面向2',
    '面向3', '面向4', '總分', '滿分', '百分比', '等第'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  ui.alert('✅ 欄位標頭已設定完成！\n\n' +
    'A~E 欄由老師填入資料，F~N 欄由程式自動產生。\n\n' +
    'E 欄支援貼入：\n' +
    '• Google Docs 連結（文字報告）\n' +
    '• Google Slides 連結（簡報）\n' +
    '• Google Drive 圖片連結（畫作、圖表）\n' +
    '• 直接貼純文字也可以');
}

// ---------- 產生 Rubric 模板參考 ----------
function tool3_createRubricTemplates() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var templateSheet = ss.getSheetByName('Rubric 模板參考');
  if (templateSheet) templateSheet.clear();
  else templateSheet = ss.insertSheet('Rubric 模板參考');

  var templates = [
    {
      title: '文字報告型 Rubric',
      color: '#4285F4',
      rows: [
        ['內容正確性', '核心概念完全正確，能清楚解釋原理並舉出恰當例子', '大部分概念正確，但有少數解釋不夠精確', '有明顯概念錯誤，或僅複製教材未加理解'],
        ['組織架構', '段落分明，有清楚的前言、本文、結論，邏輯連貫', '有基本架構但過渡生硬，部分段落邏輯不清', '缺乏架構，想到什麼寫什麼，讀者難以跟隨'],
        ['數據佐證', '引用具體數據或實驗結果支持論點，來源可信', '有提及數據但不夠具體，或來源不明', '完全沒有數據佐證，僅靠感覺論述'],
        ['語言表達', '用詞精準，句型多變，學術語言與日常說法並用', '表達尚可但用詞重複，偶有口語化問題', '錯字多、語句不通，或過度口語化']
      ]
    },
    {
      title: '視覺圖像型 Rubric',
      color: '#8B5CF6',
      rows: [
        ['概念連結', '清楚呈現概念之間的關聯性，箭頭與連結線有意義', '部分連結正確，但有遺漏或不必要的連結', '概念散亂放置，缺乏有意義的連結'],
        ['視覺呈現', '配色協調、版面整潔、圖文比例恰當，一目了然', '尚可辨識但略雜亂，部分區域過於擁擠', '難以閱讀、配色混亂或字體過小'],
        ['資訊完整性', '涵蓋所有關鍵概念，無重要遺漏', '遺漏 1-2 個重要概念或次要細節', '缺少多個主要概念，內容不完整'],
        ['創意表達', '圖像富有巧思與原創性，用獨特方式呈現知識', '有嘗試但較為制式，缺乏個人風格', '直接複製教材圖表，無任何創意加工']
      ]
    },
    {
      title: '口說 Podcast 型 Rubric',
      color: '#F59E0B',
      rows: [
        ['內容正確性', '概念正確無誤，解釋清楚且深入淺出', '大致正確但有部分解釋含糊不清', '有明顯概念錯誤或解釋嚴重偏差'],
        ['口語表達', '語速適中、咬字清晰、音量穩定、自然流暢', '偶有停頓或語速不穩，但整體尚可', '語速過快或過慢、含糊不清、大量留白'],
        ['架構邏輯', '有清楚的開場、主體、結語，過渡自然順暢', '有架構但過渡不順，部分內容跳躍', '無明顯架構，想到什麼說什麼'],
        ['聽眾互動', '善用提問、舉例、類比，引導聽眾思考', '偶有互動嘗試但不夠自然', '單向念稿式呈現，完全無互動感']
      ]
    }
  ];

  var currentRow = 1;

  // 說明區
  templateSheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('使用方式：從下方選擇適合的 Rubric 模板，複製整個表格，貼到主工作表的 D 欄（Rubric 評分規準）')
    .setBackground('#E8F0FE').setFontWeight('bold').setWrap(true);
  templateSheet.setRowHeight(currentRow, 40);
  currentRow += 2;

  templates.forEach(function(tpl) {
    templateSheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue(tpl.title)
      .setBackground(tpl.color).setFontColor('white').setFontWeight('bold')
      .setFontSize(14);
    currentRow++;

    templateSheet.getRange(currentRow, 1, 1, 4)
      .setValues([['評分面向', '優（5分）', '中（3分）', '再加強（1分）']])
      .setBackground('#E8EAED').setFontWeight('bold');
    currentRow++;

    tpl.rows.forEach(function(row) {
      templateSheet.getRange(currentRow, 1, 1, 4).setValues([row]);
      templateSheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
    });

    templateSheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue('滿分：4 個面向 × 5 分 = 20 分')
      .setFontColor('#666666').setFontStyle('italic');
    currentRow += 2;
  });

  // 可直接複製的純文字版
  currentRow++;
  templateSheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('可直接貼到 D 欄的 Rubric 純文字版')
    .setBackground('#34A853').setFontColor('white').setFontWeight('bold')
    .setFontSize(14);
  currentRow += 1;

  var rubricTexts = [
    ['文字報告型', '評分面向與標準：\n1. 內容正確性：優(5)=概念完全正確，能清楚解釋原理並舉出恰當例子｜中(3)=大部分正確，少數不精確｜再加強(1)=有明顯概念錯誤\n2. 組織架構：優(5)=段落分明，前言本文結論俱全｜中(3)=有架構但過渡生硬｜再加強(1)=缺乏架構\n3. 數據佐證：優(5)=引用具體數據支持論點｜中(3)=數據不夠具體｜再加強(1)=無數據佐證\n4. 語言表達：優(5)=用詞精準，句型多變｜中(3)=尚可但用詞重複｜再加強(1)=錯字多、語句不通'],
    ['視覺圖像型', '評分面向與標準：\n1. 概念連結：優(5)=清楚呈現概念間關聯｜中(3)=部分連結正確但有遺漏｜再加強(1)=概念散亂無連結\n2. 視覺呈現：優(5)=配色協調、版面整潔｜中(3)=尚可辨識但略雜亂｜再加強(1)=難以閱讀\n3. 資訊完整性：優(5)=涵蓋所有關鍵概念｜中(3)=遺漏1-2個重點｜再加強(1)=缺少多個主要概念\n4. 創意表達：優(5)=富有巧思與原創性｜中(3)=較制式缺乏風格｜再加強(1)=直接複製無創意'],
    ['口說 Podcast 型', '評分面向與標準：\n1. 內容正確性：優(5)=概念正確，解釋清楚深入淺出｜中(3)=大致正確但含糊｜再加強(1)=有明顯錯誤\n2. 口語表達：優(5)=語速適中、咬字清晰｜中(3)=偶有停頓但尚可｜再加強(1)=含糊不清\n3. 架構邏輯：優(5)=開場、主體、結語俱全｜中(3)=有架構但不順｜再加強(1)=無明顯架構\n4. 聽眾互動：優(5)=善用提問與舉例引導思考｜中(3)=偶有嘗試｜再加強(1)=單向念稿']
  ];

  rubricTexts.forEach(function(item) {
    templateSheet.getRange(currentRow, 1).setValue(item[0]).setFontWeight('bold');
    templateSheet.getRange(currentRow, 2, 1, 3).merge().setValue(item[1]).setWrap(true);
    templateSheet.setRowHeight(currentRow, 120);
    currentRow++;
  });

  templateSheet.autoResizeColumns(1, 4);
  templateSheet.setColumnWidth(2, 280);
  templateSheet.setColumnWidth(3, 280);
  templateSheet.setColumnWidth(4, 280);

  ui.alert('✅ Rubric 模板已建立！\n\n' +
    '請查看「Rubric 模板參考」工作表。\n' +
    '可直接複製表格內容，貼到主工作表的 D 欄。');
}

// ---------- 從 Google Drive 連結擷取內容 ----------
function tool3_extractContentFromDrive_(urlOrText) {
  if (!urlOrText) return { type: 'error', content: '欄位為空' };
  var text = urlOrText.toString().trim();
  if (text === '') return { type: 'error', content: '欄位為空' };

  // Google Docs
  var docsMatch = text.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
  if (docsMatch) {
    try {
      var doc = DocumentApp.openById(docsMatch[1]);
      return { type: 'text', content: doc.getBody().getText() };
    } catch (e) {
      return { type: 'error', content: '無法開啟 Google Doc：' + e.message };
    }
  }

  // Google Slides
  var slidesMatch = text.match(/docs\.google\.com\/presentation\/d\/([a-zA-Z0-9_-]+)/);
  if (slidesMatch) {
    try {
      var pres = SlidesApp.openById(slidesMatch[1]);
      var slides = pres.getSlides();
      var allText = '';
      slides.forEach(function(slide, idx) {
        allText += '【第 ' + (idx + 1) + ' 頁投影片】\n';
        slide.getShapes().forEach(function(shape) {
          if (shape.getText) {
            var t = shape.getText().asString().trim();
            if (t) allText += t + '\n';
          }
        });
        allText += '\n';
      });
      return { type: 'text', content: allText };
    } catch (e) {
      return { type: 'error', content: '無法開啟 Google Slides：' + e.message };
    }
  }

  // Google Drive 檔案連結
  var driveMatch = text.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/)
    || text.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (driveMatch) {
    try {
      var file = DriveApp.getFileById(driveMatch[1]);
      var mime = file.getMimeType();

      if (mime.indexOf('image/') === 0) {
        var blob = file.getBlob();
        var base64 = Utilities.base64Encode(blob.getBytes());
        return { type: 'image', content: base64, mimeType: mime };
      }

      if (mime === 'application/vnd.google-apps.document') {
        var doc2 = DocumentApp.openById(driveMatch[1]);
        return { type: 'text', content: doc2.getBody().getText() };
      }

      if (mime === 'application/vnd.google-apps.presentation') {
        var pres2 = SlidesApp.openById(driveMatch[1]);
        var slides2 = pres2.getSlides();
        var allText2 = '';
        slides2.forEach(function(slide, idx) {
          allText2 += '【第 ' + (idx + 1) + ' 頁投影片】\n';
          slide.getShapes().forEach(function(shape) {
            if (shape.getText) {
              var t = shape.getText().asString().trim();
              if (t) allText2 += t + '\n';
            }
          });
          allText2 += '\n';
        });
        return { type: 'text', content: allText2 };
      }

      if (mime.indexOf('text/') === 0) {
        return { type: 'text', content: file.getBlob().getDataAsString() };
      }

      return { type: 'error', content: '不支援的檔案類型：' + mime + '\n請改用 Google Docs 或 Slides 連結。' };
    } catch (e) {
      return { type: 'error', content: '無法開啟檔案：' + e.message + '\n請確認檔案已開啟分享權限。' };
    }
  }

  // 不是連結 → 當作純文字
  return { type: 'text', content: text };
}

// ---------- 產生單一學生的回饋（含評分）----------
function tool3_generateFeedback() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var row = SpreadsheetApp.getActiveRange().getRow();

  if (row < 2) {
    ui.alert('請選取學生資料列（第 2 列以下）再執行。');
    return;
  }

  var data = sheet.getRange(row, 1, 1, 5).getValues()[0];
  var name = data[0];
  var format = data[2];
  var rubric = data[3];
  var work = data[4];

  if (!work || work.toString().trim() === '') {
    ui.alert('該列的「學生作品」(E欄) 是空的。\n請貼上 Google Drive 連結或作品文字。');
    return;
  }
  if (!rubric || rubric.toString().trim() === '') {
    ui.alert('該列的「Rubric 評分規準」(D欄) 是空的，請先填入再執行。\n\n' +
      '提示：可點選「📝 產生 Rubric 模板參考」取得範例。');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('正在擷取 ' + name + ' 的作品內容...', 'AI 回饋', -1);
  var extracted = tool3_extractContentFromDrive_(work);

  if (extracted.type === 'error') {
    ui.alert('❌ 無法擷取作品內容：\n\n' + extracted.content);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('正在為 ' + name + ' 產生個人化回饋...', 'AI 回饋', -1);
  var feedback = tool3_callGeminiFeedback_(name, format, rubric, extracted);

  if (feedback) {
    tool3_writeFeedbackToSheet_(sheet, row, feedback);
    tool3_updateDimensionHeaders_(sheet, feedback);
    ui.alert('✅ 已為 ' + name + ' 產生回饋！\n請查看 F~N 欄。');
  } else {
    ui.alert('❌ 回饋產生失敗，請檢查 API 金鑰或稍後再試。');
  }
}

// ---------- 批次產生全部回饋 ----------
function tool3_generateAllFeedback() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('請先填入學生資料。');
    return;
  }

  var count = 0;
  var failed = [];
  var headersUpdated = false;

  for (var row = 2; row <= lastRow; row++) {
    var data = sheet.getRange(row, 1, 1, 6).getValues()[0];
    var name = data[0];
    var format = data[2];
    var rubric = data[3];
    var work = data[4];
    var existing = data[5];

    if ((existing && existing.toString().trim() !== '') || !work || work.toString().trim() === '') continue;

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '正在擷取 ' + name + ' 的作品...（第 ' + (count + 1) + ' 位）', 'AI 回饋', -1
    );

    var extracted = tool3_extractContentFromDrive_(work);
    if (extracted.type === 'error') {
      failed.push(name + '（擷取失敗）');
      continue;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '正在為 ' + name + ' 產生回饋...（第 ' + (count + 1) + ' 位）', 'AI 回饋', -1
    );

    var feedback = tool3_callGeminiFeedback_(name, format, rubric, extracted);
    if (feedback) {
      tool3_writeFeedbackToSheet_(sheet, row, feedback);
      if (!headersUpdated) {
        tool3_updateDimensionHeaders_(sheet, feedback);
        headersUpdated = true;
      }
      count++;
    } else {
      failed.push(name);
    }

    Utilities.sleep(2000);
  }

  var msg = '✅ 完成！共產生 ' + count + ' 份個人化回饋。';
  if (failed.length > 0) {
    msg += '\n\n以下學生回饋產生失敗：' + failed.join('、');
  }
  ui.alert(msg);
}

// ---------- 呼叫 Gemini API 產生回饋（支援多模態）----------
function tool3_callGeminiFeedback_(name, format, rubric, workData) {
  var apiKey = getApiKeyOrAlert_();
  if (!apiKey) return null;

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  var workSection = '';
  if (workData.type === 'image') {
    workSection = '（學生作品為圖片，請仔細觀察圖片內容後進行評分）';
  } else {
    workSection = workData.content;
  }

  var prompt = '你是一位嚴謹但溫暖的高中老師，正在批改學生的作品。\n\n'
    + '## 重要指示\n'
    + '- 你必須誠實指出作品的不足之處\n'
    + '- 不要全部都說好話\n'
    + '- 根據 Rubric 逐項評估，有缺失就要指出\n'
    + '- 回饋的目的是幫助學生進步，不是讓他開心\n\n'
    + '## 學生資訊\n'
    + '- 姓名：' + name + '\n'
    + '- 評量形式：' + format + '\n\n'
    + '## 評分規準（Rubric）\n'
    + rubric + '\n\n'
    + '## 學生作品\n'
    + workSection + '\n\n'
    + '## 請依照以下格式回饋：\n\n'
    + '### 逐項評分\n'
    + '（根據 Rubric 的每個面向，給出等級＋分數＋簡短說明）\n'
    + '格式範例：\n'
    + '- **內容正確性**：優（5/5）— 核心概念解釋清楚...\n'
    + '- **組織架構**：中（3/5）— 有基本架構但...\n\n'
    + '### 做得好的地方（2～3 點）\n'
    + '（具體指出優點，引用作品中的內容）\n\n'
    + '### 需要改進的地方（2～3 點）\n'
    + '（明確指出不足，說明為什麼這樣不好）\n\n'
    + '### 具體改進建議（2～3 點）\n'
    + '（告訴學生下一步可以怎麼做）\n\n'
    + '### 鼓勵語\n'
    + '（一句真誠的鼓勵，不要太浮誇）\n\n'
    + '## 分數摘要（必須嚴格遵守以下格式）\n\n'
    + '在上方回饋文字結束後，請另起一行，輸出以下 JSON 格式（用 <!--SCORES_JSON 和 SCORES_JSON--> 包裹）：\n\n'
    + '<!--SCORES_JSON\n'
    + '{"dimensions":[{"name":"面向名稱","score":分數,"max":5}],"total":總分,"maxTotal":滿分,"percentage":百分比整數,"grade":"等第"}\n'
    + 'SCORES_JSON-->\n\n'
    + '分數規則：\n'
    + '- 每個面向依 Rubric 等級給分：優=5、中=3、再加強=1\n'
    + '- total = 各面向分數加總\n'
    + '- maxTotal = 面向數量 × 5\n'
    + '- percentage = Math.round(total / maxTotal * 100)\n'
    + '- grade 等第：percentage >= 85 → "優"，>= 70 → "良"，>= 50 → "中"，< 50 → "再加強"\n\n'
    + '注意：請用繁體中文，語氣像老師對學生說話，親切但不失專業。';

  // 組合 API 請求（文字 or 圖片多模態）
  var parts = [{ text: prompt }];
  if (workData.type === 'image') {
    parts.push({
      inline_data: {
        mime_type: workData.mimeType,
        data: workData.content
      }
    });
  }

  var payload = {
    contents: [{ parts: parts }],
    generationConfig: { temperature: 0.4, maxOutputTokens: 4096 }
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('API HTTP ' + response.getResponseCode() + ': ' + response.getContentText().substring(0, 500));
      return null;
    }

    var json = JSON.parse(response.getContentText());
    if (!json.candidates || json.candidates.length === 0) {
      Logger.log('API 未回傳結果（可能被安全過濾器攔截）');
      return null;
    }
    return json.candidates[0].content.parts[0].text;
  } catch (e) {
    Logger.log('錯誤：' + e.message);
    return null;
  }
}

// ---------- 解析回饋中的分數 JSON（3 層備援）----------
function tool3_parseScoresFromFeedback_(feedbackText) {
  if (!feedbackText) return null;
  var text = feedbackText.toString();

  // 策略 1：找 <!--SCORES_JSON ... SCORES_JSON--> 標記
  var match = text.match(/<!--SCORES_JSON\s*([\s\S]*?)\s*SCORES_JSON-->/);
  if (match) {
    try {
      var cleanJson = match[1].replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
      var jsonMatch = cleanJson.match(/\{[\s\S]*\}/);
      if (jsonMatch) return JSON.parse(jsonMatch[0]);
    } catch (e) {
      Logger.log('SCORES_JSON 解析失敗：' + e.message);
    }
  }

  // 策略 2：找任何含 "dimensions" 的 JSON
  var jsonPatterns = text.match(/```json\s*([\s\S]*?)```/g) || [];
  for (var i = 0; i < jsonPatterns.length; i++) {
    try {
      var clean = jsonPatterns[i].replace(/```json\s*/gi, '').replace(/```\s*/g, '');
      var obj = JSON.parse(clean);
      if (obj.dimensions) return obj;
    } catch (e) { /* 繼續嘗試 */ }
  }

  // 策略 3：找回饋文字中的 X/5 分數模式
  var scorePattern = /\*\*([^*]+)\*\*[：:]\s*[^\d]*(\d)\/5/g;
  var dims = [];
  var m;
  while ((m = scorePattern.exec(text)) !== null) {
    dims.push({ name: m[1].trim(), score: parseInt(m[2]), max: 5 });
  }

  if (dims.length === 0) {
    var altPattern = /[-•]\s*\*?\*?([^：:*]+)\*?\*?[：:][^（(]*[（(](\d)\s*[/／]\s*5\s*[）)]/g;
    while ((m = altPattern.exec(text)) !== null) {
      dims.push({ name: m[1].trim(), score: parseInt(m[2]), max: 5 });
    }
  }

  if (dims.length > 0) {
    var total = dims.reduce(function(sum, d) { return sum + d.score; }, 0);
    var maxTotal = dims.length * 5;
    var percentage = Math.round(total / maxTotal * 100);
    var grade;
    if (percentage >= 85) grade = '優';
    else if (percentage >= 70) grade = '良';
    else if (percentage >= 50) grade = '中';
    else grade = '再加強';

    return {
      dimensions: dims,
      total: total,
      maxTotal: maxTotal,
      percentage: percentage,
      grade: grade
    };
  }

  Logger.log('所有分數解析策略都失敗');
  return null;
}

// ---------- 清理回饋文字（移除 JSON 標記）----------
function tool3_cleanFeedbackText_(feedbackText) {
  return feedbackText.toString()
    .replace(/<!--SCORES_JSON[\s\S]*?SCORES_JSON-->/g, '')
    .replace(/```json\s*\{[\s\S]*?"dimensions"[\s\S]*?\}\s*```/g, '')
    .trim();
}

// ---------- 將回饋和分數寫入工作表 ----------
function tool3_writeFeedbackToSheet_(sheet, row, rawFeedback) {
  var cleanedFeedback = tool3_cleanFeedbackText_(rawFeedback);
  var scores = tool3_parseScoresFromFeedback_(rawFeedback);

  // F 欄 = 回饋文字
  sheet.getRange(row, 6).setValue(cleanedFeedback);

  if (scores && scores.dimensions) {
    // G~J 欄 = 面向分數（最多 4 個）
    for (var i = 0; i < 4; i++) {
      if (i < scores.dimensions.length) {
        sheet.getRange(row, 7 + i).setValue(scores.dimensions[i].score);
        var color = tool3_getScoreColor_(scores.dimensions[i].score);
        sheet.getRange(row, 7 + i).setBackground(color);
      } else {
        sheet.getRange(row, 7 + i).setValue('');
      }
    }
    // K~N 欄
    sheet.getRange(row, 11).setValue(scores.total);
    sheet.getRange(row, 12).setValue(scores.maxTotal);
    sheet.getRange(row, 13).setValue(scores.percentage + '%');
    sheet.getRange(row, 14).setValue(scores.grade);

    var gradeColors = { '優': '#34A853', '良': '#4285F4', '中': '#FBBC04', '再加強': '#EA4335' };
    var gradeColor = gradeColors[scores.grade] || '#FFFFFF';
    sheet.getRange(row, 14).setBackground(gradeColor).setFontColor('white').setFontWeight('bold');
  }
}

// ---------- 根據分數回傳背景色 ----------
function tool3_getScoreColor_(score) {
  if (score >= 5) return '#34A853';
  if (score >= 3) return '#FBBC04';
  return '#EA4335';
}

// ---------- 用第一位學生的面向名稱更新標頭 ----------
function tool3_updateDimensionHeaders_(sheet, rawFeedback) {
  var scores = tool3_parseScoresFromFeedback_(rawFeedback);
  if (!scores || !scores.dimensions) return;

  for (var i = 0; i < Math.min(scores.dimensions.length, 4); i++) {
    sheet.getRange(1, 7 + i).setValue(scores.dimensions[i].name);
  }
}

// ---------- 產生全班統計報表 ----------
function tool3_generateClassReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var lastRow = dataSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('沒有學生資料。');
    return;
  }

  var dimNames = dataSheet.getRange(1, 7, 1, 4).getValues()[0];
  var allData = dataSheet.getRange(2, 1, lastRow - 1, 14).getValues();

  // 先嘗試從 F 欄回饋文字重新解析分數
  var reparsed = 0;
  for (var r = 0; r < allData.length; r++) {
    var feedback = allData[r][5];
    var totalScore = allData[r][10];
    if (feedback && feedback.toString().trim() !== '' &&
        (totalScore === '' || totalScore === null || totalScore === undefined)) {
      var scores = tool3_parseScoresFromFeedback_(feedback.toString());
      if (scores && scores.dimensions) {
        var sheetRow = r + 2;
        for (var d = 0; d < Math.min(scores.dimensions.length, 4); d++) {
          dataSheet.getRange(sheetRow, 7 + d).setValue(scores.dimensions[d].score);
          dataSheet.getRange(sheetRow, 7 + d).setBackground(tool3_getScoreColor_(scores.dimensions[d].score));
        }
        dataSheet.getRange(sheetRow, 11).setValue(scores.total);
        dataSheet.getRange(sheetRow, 12).setValue(scores.maxTotal);
        dataSheet.getRange(sheetRow, 13).setValue(scores.percentage + '%');
        dataSheet.getRange(sheetRow, 14).setValue(scores.grade);

        var gradeColors = { '優': '#34A853', '良': '#4285F4', '中': '#FBBC04', '再加強': '#EA4335' };
        dataSheet.getRange(sheetRow, 14).setBackground(gradeColors[scores.grade] || '#FFFFFF')
          .setFontColor('white').setFontWeight('bold');

        for (var d2 = 0; d2 < Math.min(scores.dimensions.length, 4); d2++) {
          allData[r][6 + d2] = scores.dimensions[d2].score;
        }
        allData[r][10] = scores.total;
        allData[r][11] = scores.maxTotal;
        allData[r][12] = scores.percentage + '%';
        allData[r][13] = scores.grade;

        if (reparsed === 0) {
          for (var d3 = 0; d3 < Math.min(scores.dimensions.length, 4); d3++) {
            dimNames[d3] = scores.dimensions[d3].name;
            dataSheet.getRange(1, 7 + d3).setValue(scores.dimensions[d3].name);
          }
        }
        reparsed++;
      }
    }
  }

  if (reparsed > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '已從回饋文字中重新解析 ' + reparsed + ' 位學生的分數', '分數修復', 5
    );
  }

  // 過濾有分數的學生
  var students = allData.filter(function(row) {
    return row[10] !== '' && row[10] !== null && row[10] !== undefined;
  });

  if (students.length === 0) {
    ui.alert('⚠️ 尚無已評分的學生。\n\n' +
      '可能原因：AI 回饋中未包含可解析的分數格式。\n' +
      '請確認 F 欄的回饋內容中有「逐項評分」區塊，\n' +
      '格式如：**內容正確性**：優（5/5）');
    return;
  }

  // 建立報表工作表
  var reportSheet = ss.getSheetByName('全班統計報表');
  if (reportSheet) reportSheet.clear();
  else reportSheet = ss.insertSheet('全班統計報表');

  var currentRow = 1;

  // ===== 區塊 A：全班成績總覽 =====
  var headerA = ['學生姓名', dimNames[0] || '面向1', dimNames[1] || '面向2',
                 dimNames[2] || '面向3', dimNames[3] || '面向4',
                 '總分', '滿分', '百分比', '等第'];

  reportSheet.getRange(currentRow, 1, 1, headerA.length).setValues([headerA])
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold');
  currentRow++;

  var dimSums = [0, 0, 0, 0];
  var dimCounts = [0, 0, 0, 0];
  var dimAllScores = [[], [], [], []];
  var totalSum = 0;
  var gradeDistribution = { '優': 0, '良': 0, '中': 0, '再加強': 0 };

  students.forEach(function(row) {
    var studentRow = [row[0]];

    for (var d = 0; d < 4; d++) {
      var score = Number(row[6 + d]);
      if (!isNaN(score) && row[6 + d] !== '' && row[6 + d] !== null) {
        studentRow.push(score);
        dimSums[d] += score;
        dimCounts[d]++;
        dimAllScores[d].push(score);
      } else {
        studentRow.push('');
      }
    }

    studentRow.push(row[10]);
    studentRow.push(row[11]);
    studentRow.push(row[12]);
    studentRow.push(row[13]);

    totalSum += Number(row[10]) || 0;

    var grade = row[13] ? row[13].toString() : '';
    if (gradeDistribution.hasOwnProperty(grade)) {
      gradeDistribution[grade]++;
    }

    reportSheet.getRange(currentRow, 1, 1, headerA.length).setValues([studentRow]);

    var gradeColors = { '優': '#34A853', '良': '#4285F4', '中': '#FBBC04', '再加強': '#EA4335' };
    if (gradeColors[grade]) {
      reportSheet.getRange(currentRow, 9).setBackground(gradeColors[grade])
        .setFontColor('white').setFontWeight('bold');
    }

    currentRow++;
  });

  // 全班平均列
  var avgRow = ['全班平均'];
  for (var d = 0; d < 4; d++) {
    avgRow.push(dimCounts[d] > 0 ? Math.round(dimSums[d] / dimCounts[d] * 10) / 10 : '');
  }
  var totalAvg = Math.round(totalSum / students.length * 10) / 10;
  var maxScore = students[0] ? Number(students[0][11]) || 20 : 20;
  var avgPct = Math.round(totalAvg / maxScore * 100);
  avgRow.push(totalAvg, maxScore, avgPct + '%', '');

  reportSheet.getRange(currentRow, 1, 1, headerA.length).setValues([avgRow])
    .setBackground('#E8EAED').setFontWeight('bold');
  currentRow += 3;

  // ===== 區塊 B：各面向分析 =====
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('各面向全班分析')
    .setBackground('#0D5F5C').setFontColor('white').setFontWeight('bold').setFontSize(13);
  currentRow++;

  reportSheet.getRange(currentRow, 1, 1, 6)
    .setValues([['評分面向', '全班平均', '最高分', '最低分', '滿分人數', '待加強人數']])
    .setBackground('#E8EAED').setFontWeight('bold');
  currentRow++;

  var weakestIdx = -1;
  var weakestAvg = 999;

  for (var d = 0; d < 4; d++) {
    if (dimCounts[d] === 0) continue;

    var avg = Math.round(dimSums[d] / dimCounts[d] * 10) / 10;
    var maxVal = Math.max.apply(null, dimAllScores[d]);
    var minVal = Math.min.apply(null, dimAllScores[d]);
    var fullCount = dimAllScores[d].filter(function(s) { return s >= 5; }).length;
    var lowCount = dimAllScores[d].filter(function(s) { return s <= 1; }).length;

    var dimRow = [dimNames[d] || ('面向' + (d + 1)), avg, maxVal, minVal,
                  fullCount + ' 人', lowCount + ' 人'];

    reportSheet.getRange(currentRow, 1, 1, 6).setValues([dimRow]);

    if (avg < weakestAvg) {
      weakestAvg = avg;
      weakestIdx = currentRow;
    }
    currentRow++;
  }

  if (weakestIdx > 0) {
    reportSheet.getRange(weakestIdx, 1, 1, 6)
      .setBackground('#FEE2E2').setFontColor('#DC2626');
  }

  currentRow += 2;

  // ===== 區塊 C：等第分佈 =====
  reportSheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('等第分佈統計')
    .setBackground('#0D5F5C').setFontColor('white').setFontWeight('bold').setFontSize(13);
  currentRow++;

  reportSheet.getRange(currentRow, 1, 1, 4)
    .setValues([['等第', '人數', '百分比', '分佈圖']])
    .setBackground('#E8EAED').setFontWeight('bold');
  currentRow++;

  var gradeOrder = ['優', '良', '中', '再加強'];

  gradeOrder.forEach(function(grade) {
    var count = gradeDistribution[grade] || 0;
    var pct = students.length > 0 ? Math.round(count / students.length * 100) : 0;
    var bar = '';
    for (var i = 0; i < count; i++) bar += '█';

    reportSheet.getRange(currentRow, 1, 1, 4)
      .setValues([[grade, count + ' 人', pct + '%', bar]]);
    currentRow++;
  });

  currentRow += 2;

  // ===== 區塊 D：教學建議 =====
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('教學建議')
    .setBackground('#0D5F5C').setFontColor('white').setFontWeight('bold').setFontSize(13);
  currentRow++;

  var weakDimName = '';
  var weakDimAvg = 999;
  for (var d = 0; d < 4; d++) {
    if (dimCounts[d] > 0) {
      var avg2 = dimSums[d] / dimCounts[d];
      if (avg2 < weakDimAvg) {
        weakDimAvg = avg2;
        weakDimName = dimNames[d] || ('面向' + (d + 1));
      }
    }
  }
  weakDimAvg = Math.round(weakDimAvg * 10) / 10;

  var topGrade = '良';
  var topCount = 0;
  gradeOrder.forEach(function(g) {
    if ((gradeDistribution[g] || 0) > topCount) {
      topCount = gradeDistribution[g];
      topGrade = g;
    }
  });

  var suggestion = '根據全班統計分析：\n\n'
    + '1. 全班在「' + weakDimName + '」面向表現最弱（平均 ' + weakDimAvg + ' / 5 分），'
    + '建議下次教學針對此面向加強練習或提供更多範例。\n\n'
    + '2. 等第分佈以「' + topGrade + '」佔最多（' + topCount + ' 人，'
    + Math.round(topCount / students.length * 100) + '%）';

  if (avgPct >= 70) {
    suggestion += '，全班整體表現良好。\n\n';
  } else if (avgPct >= 50) {
    suggestion += '，全班整體表現中等，仍有進步空間。\n\n';
  } else {
    suggestion += '，建議安排補強教學或調整教學策略。\n\n';
  }

  suggestion += '3. 全班平均分：' + totalAvg + ' / ' + maxScore + '（' + avgPct + '%），'
    + '共 ' + students.length + ' 位學生完成評量。';

  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue(suggestion).setWrap(true)
    .setBackground('#F0FDF4').setVerticalAlignment('top');
  reportSheet.setRowHeight(currentRow, 140);

  reportSheet.autoResizeColumns(1, 6);

  ui.alert('✅ 全班統計報表已產生！\n\n' +
    '共 ' + students.length + ' 位學生\n' +
    '全班平均：' + totalAvg + ' / ' + maxScore + '（' + avgPct + '%）\n' +
    '全班最弱面向：' + weakDimName + '\n\n' +
    '請查看「全班統計報表」工作表。');
}

// ---------- 建立 Email HTML 模板 ----------
function tool3_buildEmailHtml_(name, format, feedback, total, maxTotal, pct, grade) {
  // 移除 emoji（避免 email 亂碼 "?"）
  var cleanFeedback = feedback
    .replace(/📊/g, '[評分]').replace(/✅/g, '[優]').replace(/⚠️/g, '[注意]')
    .replace(/💡/g, '[建議]').replace(/💪/g, '[加油]').replace(/📬/g, '')
    .replace(/[\u{1F300}-\u{1F9FF}]/gu, '');

  // Markdown → HTML
  var feedbackHtml = cleanFeedback
    .replace(/### (.*)/g, '<h3 style="color:#4285F4;border-bottom:1px solid #ddd;padding-bottom:6px;margin-top:18px;">$1</h3>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/- /g, '&bull; ')
    .replace(/\n/g, '<br>');

  var gradeColors = { '優': '#34A853', '良': '#4285F4', '中': '#FBBC04', '再加強': '#EA4335' };
  var gradeColor = gradeColors[grade] || '#333';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;max-width:650px;margin:auto;padding:0;">'
    + '<div style="background:linear-gradient(135deg,#4285F4,#34A853);color:white;padding:24px 30px;border-radius:12px 12px 0 0;">'
    + '<h2 style="margin:0;font-size:22px;">學習回饋通知</h2>'
    + '<p style="margin:6px 0 0;opacity:0.9;">' + name + ' 同學 | ' + format + '</p>'
    + '</div>';

  if (total) {
    html += '<div style="background:#f8f9fa;padding:16px 30px;border-left:1px solid #ddd;border-right:1px solid #ddd;">'
      + '<table style="width:100%;border-collapse:collapse;">'
      + '<tr>'
      + '<td style="text-align:center;padding:8px;"><div style="font-size:13px;color:#666;">總分</div><div style="font-size:24px;font-weight:bold;">' + total + '<span style="font-size:14px;color:#666;">/' + maxTotal + '</span></div></td>'
      + '<td style="text-align:center;padding:8px;"><div style="font-size:13px;color:#666;">百分比</div><div style="font-size:24px;font-weight:bold;">' + pct + '</div></td>'
      + '<td style="text-align:center;padding:8px;"><div style="font-size:13px;color:#666;">等第</div><div style="font-size:24px;font-weight:bold;color:' + gradeColor + ';">' + grade + '</div></td>'
      + '</tr></table></div>';
  }

  html += '<div style="padding:20px 30px;border:1px solid #ddd;border-top:none;border-radius:0 0 12px 12px;line-height:1.8;">'
    + feedbackHtml
    + '<hr style="margin-top:24px;border:none;border-top:1px solid #eee;">'
    + '<p style="color:#999;font-size:12px;margin-top:12px;">此信件由 AI 學生回饋系統自動產生。如有疑問，請洽詢任課老師。</p>'
    + '</div></body></html>';

  return html;
}

// ---------- 寄送回饋信件給學生 ----------
function tool3_sendFeedbackEmails() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('沒有學生資料。');
    return;
  }

  var allData = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  var toSend = allData.filter(function(row) {
    return row[1] && row[1].toString().trim() !== '' && row[5] && row[5].toString().trim() !== '';
  });

  if (toSend.length === 0) {
    ui.alert('⚠️ 沒有可寄出的信件。\n\n請確認：\n' +
      '1. B 欄有填入學生 Email\n' +
      '2. F 欄已產生 AI 回饋');
    return;
  }

  var confirm = ui.alert(
    '確認寄信',
    '即將寄出 ' + toSend.length + ' 封回饋信件給學生，確定要寄出嗎？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var sent = 0;
  var failed = [];
  var skipped = 0;

  for (var row = 2; row <= lastRow; row++) {
    var data = sheet.getRange(row, 1, 1, 14).getValues()[0];
    var name = data[0];
    var email = data[1] ? data[1].toString().trim() : '';
    var format = data[2];
    var feedback = data[5];
    var total = data[10];
    var maxTotal = data[11];
    var pct = data[12];
    var grade = data[13];

    if (!email || !feedback || feedback.toString().trim() === '') {
      skipped++;
      continue;
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      failed.push(name + '（Email 格式錯誤）');
      continue;
    }

    var subject = '【學習回饋】' + name + ' 同學的' + (format || '') + '評量回饋';
    var htmlBody = tool3_buildEmailHtml_(name, format || '', feedback.toString(), total, maxTotal, pct, grade);

    try {
      GmailApp.sendEmail(email, subject, '請使用支援 HTML 的郵件程式查看此信。', {
        htmlBody: htmlBody,
        name: '學習回饋系統'
      });
      sent++;

      SpreadsheetApp.getActiveSpreadsheet().toast(
        '已寄出：' + name + '（' + sent + '/' + toSend.length + '）', '寄信中', 3
      );
    } catch (e) {
      failed.push(name + '（' + e.message + '）');
    }

    Utilities.sleep(500);
  }

  var msg = '✅ 寄信完成！\n\n'
    + '成功寄出：' + sent + ' 封\n'
    + '跳過（無 Email 或無回饋）：' + skipped + ' 人';
  if (failed.length > 0) {
    msg += '\n\n❌ 寄送失敗：\n' + failed.join('\n');
  }
  ui.alert(msg);
}
