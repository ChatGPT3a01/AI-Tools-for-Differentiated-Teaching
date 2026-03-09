// ============================================================
// 整合版：AI 工具協助差異化教學（3 合 1）
// 目的：避免多檔案 onOpen() 衝突，提供單一可共用版本
// ============================================================

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
    .addItem('🔄 產生個人化回饋', 'tool3_generateFeedback')
    .addItem('📧 批次產生全部回饋', 'tool3_generateAllFeedback')
    .addToUi();
}

function getApiKeyOrAlert_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('請先設定 GEMINI_API_KEY（專案設定 → 指令碼屬性）');
    return null;
  }
  return apiKey;
}

function callGeminiText_(prompt, temperature, maxOutputTokens) {
  const apiKey = getApiKeyOrAlert_();
  if (!apiKey) return null;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
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
// 工具 1：段考 AI 審題助手
// ============================================================

function tool1_showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <h3>段考 AI 審題助手 使用說明</h3>
    <ol>
      <li>在 Sheet1 的 A 欄貼上段考試題（每列一題）</li>
      <li>點選「AI 審題助手」→「分析段考試題」</li>
      <li>等待 30 秒～1 分鐘</li>
      <li>查看新產生的：雙向細目表、難易度分析、檢核與建議</li>
    </ol>
  `).setWidth(420).setHeight(300);
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
    .map(r => r[0])
    .filter(v => String(v).trim() !== '');

  if (!questions.length) {
    ui.alert('未偵測到試題，請確認 A 欄已貼上題目。');
    return;
  }

  const examText = questions.map((q, i) => `第${i + 1}題：${q}`).join('\n\n');
  ui.alert(`已偵測到 ${questions.length} 題，開始 AI 分析...`);

  const prompt = `你是一位專業的高中教育評量專家。請分析以下段考試題，並產出：

1. 雙向細目表（JSON）
2. 難易度分析（JSON）
3. 覆蓋率檢核
4. 改進建議

請用以下 JSON 格式回傳（必須合法 JSON）：
{
  "雙向細目表": {
    "單元列表": ["單元A"],
    "認知層次": ["記憶", "理解", "應用", "分析", "評鑑", "創造"],
    "對應題號": {"單元A": {"記憶": [1]}}
  },
  "難易度分析": {
    "各題難度": {"1": "中等"},
    "統計": {"簡單": {"題數": 0, "百分比": "0%"}, "中等": {"題數": 1, "百分比": "100%"}, "困難": {"題數": 0, "百分比": "0%"}}
  },
  "覆蓋率檢核": {
    "出題偏多": [],
    "未覆蓋": [],
    "難度均衡性": ""
  },
  "改進建議": []
}

以下是段考試題：
${examText}`;

  const text = callGeminiText_(prompt, 0.2, 8192);
  const analysis = extractJsonObject_(text);
  if (!analysis) {
    ui.alert('❌ 分析失敗，請檢查 API 金鑰與輸入資料。');
    return;
  }

  tool1_outputResults_(analysis);
  ui.alert('✅ 分析完成！請查看新工作表。');
}

function tool1_outputResults_(analysis) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetTable = ss.getSheetByName('雙向細目表');
  if (sheetTable) sheetTable.clear();
  else sheetTable = ss.insertSheet('雙向細目表');

  const table = analysis['雙向細目表'] || {};
  const levels = table['認知層次'] || ['記憶', '理解', '應用', '分析', '評鑑', '創造'];
  const headers = ['單元 \\ 認知層次'].concat(levels);
  sheetTable.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold');

  const units = table['單元列表'] || [];
  units.forEach((unit, i) => {
    const row = [unit];
    levels.forEach(level => {
      const nums = (table['對應題號'] && table['對應題號'][unit] && table['對應題號'][unit][level]) || [];
      row.push(nums.length ? nums.join(', ') : '—');
    });
    sheetTable.getRange(i + 2, 1, 1, row.length).setValues([row]);
  });
  sheetTable.autoResizeColumns(1, headers.length);

  let sheetDiff = ss.getSheetByName('難易度分析');
  if (sheetDiff) sheetDiff.clear();
  else sheetDiff = ss.insertSheet('難易度分析');

  const diff = analysis['難易度分析'] || {};
  const byQ = diff['各題難度'] || {};
  sheetDiff.getRange(1, 1).setValue('各題難度分析').setFontWeight('bold').setFontSize(12);
  sheetDiff.getRange(2, 1, 1, 2).setValues([['題號', '難度']])
    .setBackground('#34A853').setFontColor('white').setFontWeight('bold');
  Object.keys(byQ).forEach((q, i) => {
    sheetDiff.getRange(i + 3, 1, 1, 2).setValues([[`第 ${q} 題`, byQ[q]]]);
  });

  const statStart = Object.keys(byQ).length + 5;
  sheetDiff.getRange(statStart, 1).setValue('統計摘要').setFontWeight('bold').setFontSize(12);
  sheetDiff.getRange(statStart + 1, 1, 1, 3).setValues([['難度', '題數', '百分比']])
    .setBackground('#FBBC04').setFontWeight('bold');
  const stat = diff['統計'] || {};
  ['簡單', '中等', '困難'].forEach((k, idx) => {
    const s = stat[k] || {};
    sheetDiff.getRange(statStart + 2 + idx, 1, 1, 3).setValues([[k, s['題數'] || 0, s['百分比'] || '0%']]);
  });
  sheetDiff.autoResizeColumns(1, 3);

  let sheetAdvice = ss.getSheetByName('檢核與建議');
  if (sheetAdvice) sheetAdvice.clear();
  else sheetAdvice = ss.insertSheet('檢核與建議');
  const check = analysis['覆蓋率檢核'] || {};
  const advice = analysis['改進建議'] || [];
  let r = 1;
  sheetAdvice.getRange(r++, 1).setValue('覆蓋率檢核').setFontWeight('bold').setFontSize(12);
  sheetAdvice.getRange(r, 1).setValue('出題偏多的單元：').setFontWeight('bold');
  sheetAdvice.getRange(r++, 2).setValue((check['出題偏多'] || ['無']).join('、'));
  sheetAdvice.getRange(r, 1).setValue('未覆蓋的單元：').setFontWeight('bold');
  sheetAdvice.getRange(r++, 2).setValue((check['未覆蓋'] || ['無']).join('、'));
  sheetAdvice.getRange(r, 1).setValue('難度均衡性：').setFontWeight('bold');
  sheetAdvice.getRange(r + 1, 2).setValue(check['難度均衡性'] || '—');
  r += 3;
  sheetAdvice.getRange(r++, 1).setValue('改進建議').setFontWeight('bold').setFontSize(12);
  advice.forEach((v, i) => sheetAdvice.getRange(r + i, 1).setValue(`${i + 1}. ${v}`));
  sheetAdvice.autoResizeColumns(1, 2);
}

// ============================================================
// 工具 2：動態評量產生器
// ============================================================

function tool2_generateQuestions() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('請先在 Sheet1 填入知識點（A:科目, B:單元, C:知識點）');
    return;
  }
  if (!getApiKeyOrAlert_()) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const points = data
    .filter(r => String(r[2]).trim() !== '')
    .map(r => ({ subject: r[0], unit: r[1], point: r[2] }));

  if (!points.length) {
    ui.alert('C 欄沒有可用知識點，請先填入再執行。');
    return;
  }
  if (points.length > 30) {
    ui.alert(`⚠️ 知識點數量過多（${points.length}），建議一次不超過 30 個。`);
    return;
  }

  let rs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('題目與提示');
  if (rs) rs.clear();
  else rs = SpreadsheetApp.getActiveSpreadsheet().insertSheet('題目與提示');

  const headers = ['知識點', '題目', '正確答案', '提示1（方向提示）', '提示1後答案', '提示2（給公式）', '提示2後答案', '提示3（帶數字）', '提示3後答案'];
  rs.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#4285F4').setFontColor('white').setFontWeight('bold');

  let row = 2;
  points.forEach(kp => {
    const item = tool2_generateOneQuestion_(kp);
    if (item) rs.getRange(row++, 1, 1, headers.length).setValues([item]);
  });

  rs.autoResizeColumns(1, headers.length);
  ui.alert(`✅ 已產生 ${row - 2} 道動態評量題目。`);
}

function tool2_generateOneQuestion_(kp) {
  const prompt = `你是一位高中${kp.subject}老師，請針對「${kp.unit} - ${kp.point}」設計一道動態評量題目。
請用 JSON 回覆：
{
  "題目": "",
  "正確答案": "",
  "提示1_方向提示": "",
  "提示1後答案": "",
  "提示2_給公式": "",
  "提示2後答案": "",
  "提示3_帶數字": "",
  "提示3後答案": ""
}
要求：題目有具體數字，三層提示有層次。`;

  const text = callGeminiText_(prompt, 0.3, 2048);
  const json = extractJsonObject_(text);
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
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('題目與提示');
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('請先產生題目（題目與提示工作表）。');
    return;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const form = FormApp.create('動態評量 - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'));
  form.setDescription('這是一份動態評量，答錯時可逐層獲得提示。\n不用提示 5 分；提示越多分數越低。');
  form.setIsQuiz(false);

  const qPages = [];
  const navs = [];

  data.forEach((row, idx) => {
    const n = idx + 1;
    const point = row[0], question = row[1], hint1 = row[3], hint2 = row[5], hint3 = row[7];

    const qPage = form.addPageBreakItem().setTitle(`第 ${n} 題（${point}）`);
    qPages.push(qPage);

    form.addParagraphTextItem().setTitle(question).setRequired(true);
    const need1 = form.addMultipleChoiceItem().setTitle('需要提示嗎？').setRequired(true);

    form.addPageBreakItem().setTitle(`第 ${n} 題 - 提示 1`);
    form.addSectionHeaderItem().setTitle(`提示 1：${hint1}`);
    form.addParagraphTextItem().setTitle('再試一次').setRequired(false);
    const need2 = form.addMultipleChoiceItem().setTitle('還需要更多提示嗎？').setRequired(true);

    form.addPageBreakItem().setTitle(`第 ${n} 題 - 提示 2`);
    form.addSectionHeaderItem().setTitle(`提示 2：${hint2}`);
    form.addParagraphTextItem().setTitle('再算一次').setRequired(false);
    const need3 = form.addMultipleChoiceItem().setTitle('還要最後提示嗎？').setRequired(true);

    const p3 = form.addPageBreakItem().setTitle(`第 ${n} 題 - 提示 3`);
    form.addSectionHeaderItem().setTitle(`提示 3：${hint3}`);
    form.addParagraphTextItem().setTitle('完成最後一步').setRequired(false);

    navs.push({ need1, need2, need3, p3 });
  });

  navs.forEach((nav, idx) => {
    const next = idx + 1 < qPages.length ? qPages[idx + 1] : FormApp.PageNavigationType.SUBMIT;
    nav.need1.setChoices([
      nav.need1.createChoice('不需要，我會了', next),
      nav.need1.createChoice('需要提示', FormApp.PageNavigationType.CONTINUE)
    ]);
    nav.need2.setChoices([
      nav.need2.createChoice('不需要了', next),
      nav.need2.createChoice('再給一點提示', FormApp.PageNavigationType.CONTINUE)
    ]);
    nav.need3.setChoices([
      nav.need3.createChoice('不需要了', next),
      nav.need3.createChoice('給我最後提示', FormApp.PageNavigationType.CONTINUE)
    ]);
    if (idx + 1 < qPages.length) nav.p3.setGoToPage(qPages[idx + 1]);
    else nav.p3.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  });

  let linkSheet = ss.getSheetByName('表單連結');
  if (linkSheet) linkSheet.clear();
  else linkSheet = ss.insertSheet('表單連結');
  linkSheet.getRange(1, 1, 2, 2).setValues([
    ['學生作答連結', form.getPublishedUrl()],
    ['編輯連結', form.getEditUrl()]
  ]);
  linkSheet.autoResizeColumns(1, 2);
  ui.alert('✅ Google Form 已建立，連結已寫入「表單連結」工作表。');
}

function tool2_calculateScores() {
  SpreadsheetApp.getUi().alert(
    '計分規則：\n' +
    '不用提示=5 分\n1 層提示=4 分\n2 層提示=3 分\n3 層提示=1 分\n全部提示後仍錯=0 分'
  );
}

// ============================================================
// 工具 3：AI 學生回饋系統
// ============================================================

function tool3_generateFeedback() {
  const ui = SpreadsheetApp.getUi();
  if (!getApiKeyOrAlert_()) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const row = SpreadsheetApp.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('請選取學生資料列（第 2 列以下）再執行。');
    return;
  }

  const [name, format, rubric, work] = sheet.getRange(row, 1, 1, 4).getValues()[0];
  if (String(work).trim() === '') {
    ui.alert('D 欄（學生作品）是空的。');
    return;
  }
  if (String(rubric).trim() === '') {
    ui.alert('C 欄（Rubric）是空的。');
    return;
  }

  const feedback = tool3_callGeminiFeedback_(name, format, rubric, work);
  if (!feedback) {
    ui.alert('❌ 回饋產生失敗。');
    return;
  }
  sheet.getRange(row, 5).setValue(feedback);
  ui.alert(`✅ 已為 ${name} 產生回饋。`);
}

function tool3_generateAllFeedback() {
  const ui = SpreadsheetApp.getUi();
  if (!getApiKeyOrAlert_()) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('請先填入學生資料。');
    return;
  }

  let ok = 0;
  const failed = [];
  for (let row = 2; row <= lastRow; row++) {
    const [name, format, rubric, work, existing] = sheet.getRange(row, 1, 1, 5).getValues()[0];
    if (String(existing).trim() !== '') continue;
    if (String(work).trim() === '' || String(rubric).trim() === '') continue;

    const fb = tool3_callGeminiFeedback_(name, format, rubric, work);
    if (fb) {
      sheet.getRange(row, 5).setValue(fb);
      ok++;
    } else {
      failed.push(name || `第${row}列`);
    }
    Utilities.sleep(1500);
  }

  let msg = `✅ 完成！共產生 ${ok} 份回饋。`;
  if (failed.length) msg += `\n失敗：${failed.join('、')}`;
  ui.alert(msg);
}

function tool3_callGeminiFeedback_(name, format, rubric, work) {
  const prompt = `你是一位嚴謹但溫暖的高中老師，請根據 Rubric 回饋學生作品。

學生姓名：${name}
評量形式：${format}

Rubric：
${rubric}

學生作品：
${work}

請用繁體中文，依下列格式輸出：
1) 逐項評分
2) 做得好的地方（2-3點）
3) 需要改進的地方（2-3點）
4) 具體改進建議（2-3點）
5) 鼓勵語`;

  return callGeminiText_(prompt, 0.4, 2048);
}

