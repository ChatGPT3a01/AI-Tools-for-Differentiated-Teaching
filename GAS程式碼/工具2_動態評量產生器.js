// ============================================================
//  工具 ② 動態評量產生器
// ============================================================
//
//  【這個工具做什麼？】
//  你輸入幾個「知識點」→ AI 自動產生：
//    1. 每個知識點一道題目
//    2. 每道題目附三層提示（方向→公式→帶數字）
//    3. 自動建立 Google Form 讓學生作答
//
//  學生答題時：答對得 5 分，用 1 個提示得 4 分，
//  用 2 個提示得 3 分，用 3 個提示得 1 分。
//  → 所有學生都能得分、都能邊做邊學！
//
//  【安裝步驟（5 分鐘）】
//  1. 新增一份 Google Sheets，命名為「動態評量產生器」
//  2. 點選上方選單「擴充功能」→「Apps Script」
//  3. 把這整份程式碼「全選 → 複製 → 貼上」取代編輯器裡的內容
//  4. 左邊齒輪「專案設定」→ 拉到最下面「指令碼屬性」
//     → 新增一筆：屬性 = GEMINI_API_KEY，值 = 你的 API 金鑰
//  5. 按 Ctrl+S 儲存 → 點上方「執行 ▶」→ 函式選 onOpen → 執行
//  6. 會跳出授權視窗 → 照著點「允許」就好
//  7. 回到 Sheets 按 F5 重新整理 → 上方選單出現「動態評量」
//
//  【使用方式】
//  1. 在 Sheet1 填入知識點：
//     A 欄 = 科目，B 欄 = 單元，C 欄 = 知識點
//     例如：物理 | 牛頓運動定律 | F=ma 的應用
//  2. 點選「🎯 動態評量」→「📝 產生動態評量題目」
//  3. 等 AI 產生完（每個知識點約 10 秒）
//  4. 檢查「題目與提示」工作表的內容
//  5. 點選「🎯 動態評量」→「📋 建立 Google Form」
//  6. 把產生的表單連結發給學生！
//
//  【API 金鑰哪裡拿？】
//  到 https://aistudio.google.com/apikey → 建立 API 金鑰 → 複製
//
// ============================================================

// ---------- 建立自訂選單 ----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎯 動態評量')
    .addItem('📝 產生動態評量題目', 'generateQuestions')
    .addItem('📋 建立 Google Form', 'createForm')
    .addItem('📊 計算成績', 'calculateScores')
    .addToUi();
}

// ---------- 主功能：產生動態評量題目 + 三層提示 ----------
function generateQuestions() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('請先在 Sheet1 填入知識點（A:科目, B:單元, C:知識點）');
    return;
  }

  // 檢查 API Key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    ui.alert('❌ 尚未設定 API 金鑰！\n\n請到「專案設定」→「指令碼屬性」中新增 GEMINI_API_KEY');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const knowledgePoints = data
    .filter(row => row[2].toString().trim() !== '')
    .map(row => ({ subject: row[0], unit: row[1], point: row[2] }));

  if (knowledgePoints.length > 30) {
    ui.alert('⚠️ 知識點數量過多（' + knowledgePoints.length + '），建議一次不超過 30 個，以免執行超時。');
    return;
  }

  ui.alert(`開始為 ${knowledgePoints.length} 個知識點產生動態評量題目...\n（每個知識點約需 10 秒）`);

  // 建立或清除結果工作表
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let resultSheet = ss.getSheetByName('題目與提示');
  if (resultSheet) resultSheet.clear();
  else resultSheet = ss.insertSheet('題目與提示');

  // 設定表頭
  const headers = ['知識點', '題目', '正確答案',
                   '提示1（方向提示）', '提示1後答案',
                   '提示2（給公式）', '提示2後答案',
                   '提示3（帶數字）', '提示3後答案'];
  resultSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold');

  // 逐一產生題目
  let currentRow = 2;
  knowledgePoints.forEach(kp => {
    const result = generateOneQuestion(kp);
    if (result) {
      resultSheet.getRange(currentRow, 1, 1, headers.length).setValues([result]);
      currentRow++;
    }
  });

  resultSheet.autoResizeColumns(1, headers.length);
  ui.alert(`✅ 已產生 ${currentRow - 2} 道動態評量題目！\n請查看「題目與提示」工作表。`);
}

// ---------- 為單一知識點產生題目（呼叫 Gemini）----------
function generateOneQuestion(kp) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const prompt = `你是一位高中${kp.subject}老師，請針對「${kp.unit} - ${kp.point}」設計一道動態評量題目。

請提供以下內容（用 JSON 格式回覆）：

{
  "題目": "一道計算題或應用題（含具體數字，約 50-80 字）",
  "正確答案": "完整的正確答案",
  "提示1_方向提示": "不給公式，只提示解題方向（如：這題要用到牛頓第二定律喔）",
  "提示1後答案": "看了提示1之後的正確答案（與原答案相同）",
  "提示2_給公式": "給出需要用到的公式（如：F = ma，其中 F 是力，m 是質量，a 是加速度）",
  "提示2後答案": "看了提示2之後的正確答案",
  "提示3_帶數字": "直接幫學生把數字代入公式，只差最後一步計算（如：F = 5 × 2 = ?）",
  "提示3後答案": "最後一步的答案"
}

要求：
- 題目要有具體數字，學生可以實際計算
- 三層提示要有明確的層次差異
- 提示越多，解題越容易
- 即使看了全部提示，學生仍需要做最後一步才能得到答案`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.3, maxOutputTokens: 2048 }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('API HTTP ' + responseCode + ': ' + response.getContentText().substring(0, 500));
      return null;
    }

    const json = JSON.parse(response.getContentText());
    if (!json.candidates || json.candidates.length === 0) {
      Logger.log('API 未回傳結果');
      return null;
    }
    const text = json.candidates[0].content.parts[0].text;
    const cleanText = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '');
    const match = cleanText.match(/\{[\s\S]*\}/);

    if (match) {
      const data = JSON.parse(match[0]);
      return [
        kp.point,
        data['題目'],
        data['正確答案'],
        data['提示1_方向提示'],
        data['提示1後答案'],
        data['提示2_給公式'],
        data['提示2後答案'],
        data['提示3_帶數字'],
        data['提示3後答案']
      ];
    }
  } catch (e) {
    Logger.log('錯誤：' + e.message);
  }
  return null;
}

// ---------- 建立 Google Form（含分支邏輯）----------
//
// 每道題的表單結構：
//   題目頁 → 「需要提示嗎？」
//     ├─ 不需要 → 跳到下一題（得 5 分）
//     └─ 需要   → 提示 1 頁 → 「還需要嗎？」
//                    ├─ 不需要 → 跳到下一題（得 4 分）
//                    └─ 需要   → 提示 2 頁 → 「還需要嗎？」
//                                   ├─ 不需要 → 跳到下一題（得 3 分）
//                                   └─ 需要   → 提示 3 頁（得 1 分）→ 下一題
//
function createForm() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('題目與提示');

  if (!sheet) {
    ui.alert('請先產生題目（點選「產生動態評量題目」）');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('沒有可用的題目。');
    return;
  }

  // 建立 Google Form
  const form = FormApp.create('動態評量 - ' + new Date().toLocaleDateString('zh-TW'));
  form.setDescription(
    '這是一份動態評量，答錯時會獲得提示，幫助你一步步解題。\n' +
    '每題滿分 5 分，使用提示會扣分，但你一定能得到分數！\n\n' +
    '計分規則：不用提示＝5分 ／ 1個提示＝4分 ／ 2個提示＝3分 ／ 3個提示＝1分'
  );
  form.setIsQuiz(false); // 用自訂計分邏輯

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const numQuestions = data.length;

  // ===== 第一輪：建立所有頁面和項目，儲存參照 =====
  const questionPages = []; // 每題的起始 PageBreakItem
  const navItems = [];      // 需要設定導航的 MultipleChoiceItem

  data.forEach((row, idx) => {
    const qNum = idx + 1;
    const [point, question, answer, hint1, ans1, hint2, ans2, hint3, ans3] = row;

    // ── 題目頁 ──
    const questionPage = form.addPageBreakItem()
      .setTitle(`第 ${qNum} 題（知識點：${point}）`);
    questionPages.push(questionPage);

    form.addParagraphTextItem()
      .setTitle(question)
      .setHelpText('請直接作答。如果不確定，可以先試試看！')
      .setRequired(true);

    const needHint = form.addMultipleChoiceItem()
      .setTitle(`答完了嗎？需要提示嗎？`)
      .setHelpText('不使用提示得 5 分；每使用一層提示會扣分')
      .setRequired(true);

    // ── 提示 1 頁 ──
    const hint1Page = form.addPageBreakItem()
      .setTitle(`第 ${qNum} 題 - 💡 提示 1`);

    form.addSectionHeaderItem()
      .setTitle(`💡 方向提示：${hint1}`);

    form.addParagraphTextItem()
      .setTitle('看了提示後，請再試一次：')
      .setRequired(false);

    const needHint2 = form.addMultipleChoiceItem()
      .setTitle('還需要更多提示嗎？')
      .setHelpText('再使用一層提示會再扣 1 分')
      .setRequired(true);

    // ── 提示 2 頁 ──
    const hint2Page = form.addPageBreakItem()
      .setTitle(`第 ${qNum} 題 - 📐 提示 2`);

    form.addSectionHeaderItem()
      .setTitle(`📐 公式提示：${hint2}`);

    form.addParagraphTextItem()
      .setTitle('有了公式，請再算一次：')
      .setRequired(false);

    const needHint3 = form.addMultipleChoiceItem()
      .setTitle('還需要最後一個提示嗎？')
      .setHelpText('這是最後一層提示，使用後得 1 分')
      .setRequired(true);

    // ── 提示 3 頁 ──
    const hint3Page = form.addPageBreakItem()
      .setTitle(`第 ${qNum} 題 - 🔢 提示 3`);

    form.addSectionHeaderItem()
      .setTitle(`🔢 帶入數字：${hint3}`);

    form.addParagraphTextItem()
      .setTitle('數字都代好了，最後一步是？')
      .setRequired(false);

    // 儲存導航項目參照
    navItems.push({
      needHint, hint1Page,
      needHint2, hint2Page,
      needHint3, hint3Page
    });
  });

  // ===== 第二輪：設定分支導航 =====
  navItems.forEach((nav, idx) => {
    // 「跳過」的目標：下一題的起始頁 or 送出表單
    const nextTarget = (idx + 1 < numQuestions)
      ? questionPages[idx + 1]
      : FormApp.PageNavigationType.SUBMIT;

    // 「需要提示嗎？」
    nav.needHint.setChoices([
      nav.needHint.createChoice('不需要，我已經會了', nextTarget),
      nav.needHint.createChoice('需要提示', FormApp.PageNavigationType.CONTINUE)
    ]);

    // 「還需要更多提示嗎？」
    nav.needHint2.setChoices([
      nav.needHint2.createChoice('不需要了，我會了', nextTarget),
      nav.needHint2.createChoice('再給一點提示', FormApp.PageNavigationType.CONTINUE)
    ]);

    // 「還需要最後一個提示嗎？」
    nav.needHint3.setChoices([
      nav.needHint3.createChoice('不需要了', nextTarget),
      nav.needHint3.createChoice('給我最後提示', FormApp.PageNavigationType.CONTINUE)
    ]);

    // 提示 3 頁結束後，自動跳到下一題或送出
    if (idx + 1 < numQuestions) {
      nav.hint3Page.setGoToPage(questionPages[idx + 1]);
    } else {
      nav.hint3Page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
    }
  });

  // ===== 記錄 Form URL =====
  const formUrl = form.getPublishedUrl();
  const editUrl = form.getEditUrl();

  let urlSheet = ss.getSheetByName('表單連結');
  if (urlSheet) urlSheet.clear();
  else urlSheet = ss.insertSheet('表單連結');

  urlSheet.getRange(1, 1).setValue('學生作答連結：').setFontWeight('bold');
  urlSheet.getRange(1, 2).setValue(formUrl);
  urlSheet.getRange(2, 1).setValue('編輯連結：').setFontWeight('bold');
  urlSheet.getRange(2, 2).setValue(editUrl);
  urlSheet.autoResizeColumns(1, 2);

  ui.alert(`✅ Google Form 已建立！\n\n學生連結：${formUrl}\n\n（連結也已記錄在「表單連結」工作表）`);
}

// ---------- 計算成績（說明版）----------
function calculateScores() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('📊 計分功能說明：\n\n' +
    '請在 Google Form 收到回覆後：\n' +
    '1. 開啟表單回覆的 Google Sheets\n' +
    '2. 根據學生使用提示的情況計分：\n' +
    '   • 不用提示答對 = 5 分\n' +
    '   • 用 1 個提示答對 = 4 分\n' +
    '   • 用 2 個提示答對 = 3 分\n' +
    '   • 用 3 個提示答對 = 1 分\n' +
    '   • 全部提示後仍答錯 = 0 分\n\n' +
    '（未來版本將支援全自動計分）');
}
