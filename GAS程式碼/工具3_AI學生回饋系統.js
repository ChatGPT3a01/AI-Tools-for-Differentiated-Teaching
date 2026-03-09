// ============================================================
//  工具 ③ AI 學生回饋系統
// ============================================================
//
//  【這個工具做什麼？】
//  貼上「學生作品」+「Rubric 評分規準」→ AI 自動產生：
//    - 逐項評分（根據 Rubric 每個面向給等級）
//    - 做得好的地方（引用作品內容，不是空泛稱讚）
//    - 需要改進的地方（誠實指出不足）
//    - 具體改進建議（告訴學生下一步怎麼做）
//    - 鼓勵語（真誠但不浮誇）
//
//  【跟直接丟給 AI 改有什麼不同？】
//  直接問 AI「幫我評分」→ AI 會討好你，全部說好話。
//  這個工具用特殊 Prompt 設計，要求 AI 必須指出缺點，
//  搭配 Rubric 逐項評分，回饋才有品質。
//
//  【安裝步驟（5 分鐘）】
//  1. 新增一份 Google Sheets，命名為「AI 學生回饋系統」
//  2. 點選上方選單「擴充功能」→「Apps Script」
//  3. 把這整份程式碼「全選 → 複製 → 貼上」取代編輯器裡的內容
//  4. 左邊齒輪「專案設定」→ 拉到最下面「指令碼屬性」
//     → 新增一筆：屬性 = GEMINI_API_KEY，值 = 你的 API 金鑰
//  5. 按 Ctrl+S 儲存 → 點上方「執行 ▶」→ 函式選 onOpen → 執行
//  6. 會跳出授權視窗 → 照著點「允許」就好
//  7. 回到 Sheets 按 F5 重新整理 → 上方選單出現「AI 回饋」
//
//  【使用方式】
//  1. Google Sheets 格式如下：
//     A 欄 = 學生姓名（或座號）
//     B 欄 = 評量形式（文字報告 / 概念圖 / Podcast）
//     C 欄 = Rubric 評分規準（從 Gemini 產出的貼過來）
//     D 欄 = 學生作品內容（把學生交的東西貼過來）
//     E 欄 = AI 回饋（這欄留空，程式會自動填入）
//
//  2. 改一個學生：選取那一列 → 點「💬 AI 回饋」→「🔄 產生個人化回饋」
//  3. 改全班：點「💬 AI 回饋」→「📧 批次產生全部回饋」
//
//  【搭配 NotebookLM 使用效果更好】
//  先用 NotebookLM 上傳課程資料（課本、講義、標準答案），
//  確認該單元的核心概念，再把 Rubric 和作品丟進這個工具。
//  → 詳見「NotebookLM設定指引.md」
//
//  【API 金鑰哪裡拿？】
//  到 https://aistudio.google.com/apikey → 建立 API 金鑰 → 複製
//
// ============================================================

// ---------- 建立自訂選單 ----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💬 AI 回饋')
    .addItem('🔄 產生個人化回饋', 'generateFeedback')
    .addItem('📧 批次產生全部回饋', 'generateAllFeedback')
    .addToUi();
}

// ---------- 產生單一學生的回饋 ----------
function generateFeedback() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const row = SpreadsheetApp.getActiveRange().getRow();

  if (row < 2) {
    ui.alert('請選取學生資料列（第 2 列以下）再執行。');
    return;
  }

  const data = sheet.getRange(row, 1, 1, 4).getValues()[0];
  const [name, format, rubric, work] = data;

  if (!work || work.toString().trim() === '') {
    ui.alert('該列的「學生作品內容」(D欄) 是空的。');
    return;
  }

  if (!rubric || rubric.toString().trim() === '') {
    ui.alert('該列的「Rubric 評分規準」(C欄) 是空的，請先填入再執行。');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`正在為 ${name} 產生個人化回饋...`, 'AI 回饋', -1);

  const feedback = callGeminiFeedback(name, format, rubric, work);

  if (feedback) {
    sheet.getRange(row, 5).setValue(feedback);
    ui.alert(`✅ 已為 ${name} 產生回饋！請查看 E 欄。`);
  } else {
    ui.alert('❌ 回饋產生失敗。');
  }
}

// ---------- 批次產生全部回饋 ----------
function generateAllFeedback() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('請先填入學生資料。');
    return;
  }

  let count = 0;
  let failed = [];
  for (let row = 2; row <= lastRow; row++) {
    const data = sheet.getRange(row, 1, 1, 5).getValues()[0];
    const [name, format, rubric, work, existing] = data;

    // 跳過已有回饋的或空白的
    if (existing.toString().trim() !== '' || work.toString().trim() === '') continue;

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `正在處理 ${name}（${count + 1} / ${lastRow - 1}）...`, 'AI 回饋', -1
    );

    const feedback = callGeminiFeedback(name, format, rubric, work);
    if (feedback) {
      sheet.getRange(row, 5).setValue(feedback);
      count++;
    } else {
      failed.push(name);
    }

    // 避免 API 速率限制（每個學生間隔 2 秒）
    Utilities.sleep(2000);
  }

  let msg = `✅ 完成！共產生 ${count} 份個人化回饋。`;
  if (failed.length > 0) {
    msg += `\n\n以下學生回饋產生失敗：${failed.join('、')}`;
  }
  ui.alert(msg);
}

// ---------- 呼叫 Gemini API 產生回饋 ----------
function callGeminiFeedback(name, format, rubric, work) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const prompt = `你是一位嚴謹但溫暖的高中老師，正在批改學生的作品。

## 重要指示
- 你必須誠實指出作品的不足之處
- 不要全部都說好話
- 根據 Rubric 逐項評估，有缺失就要指出
- 回饋的目的是幫助學生進步，不是讓他開心

## 學生資訊
- 姓名：${name}
- 評量形式：${format}

## 評分規準（Rubric）
${rubric}

## 學生作品
${work}

## 請依照以下格式回饋：

### 逐項評分
（根據 Rubric 的每個面向，給出等級和簡短說明）

### 做得好的地方（2～3 點）
（具體指出優點，引用作品中的內容）

### 需要改進的地方（2～3 點）
（明確指出不足，說明為什麼這樣不好）

### 具體改進建議（2～3 點）
（告訴學生下一步可以怎麼做）

### 鼓勵語
（一句真誠的鼓勵，不要太浮誇）

注意：請用繁體中文，語氣像老師對學生說話，親切但不失專業。`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.4, maxOutputTokens: 2048 }
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
      Logger.log('API 未回傳結果（可能被安全過濾器攔截）');
      return null;
    }
    return json.candidates[0].content.parts[0].text;
  } catch (e) {
    Logger.log('錯誤：' + e.message);
    return null;
  }
}
