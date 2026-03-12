// ============================================================
//  工具 ① 段考 AI 審題助手
// ============================================================
//
//  【這個工具做什麼？】
//  把你的段考試題貼進 Google Sheets → AI 自動產出：
//    1. 雙向細目表（哪個單元考了哪些認知層次）
//    2. 難易度分析（簡單 / 中等 / 困難 各幾題）
//    3. 檢核與建議（哪些單元沒考到、怎麼改）
//
//  【安裝步驟（5 分鐘）】
//  1. 新增一份 Google Sheets，命名為「段考AI審題助手」
//  2. 點選上方選單「擴充功能」→「Apps Script」
//  3. 把這整份程式碼「全選 → 複製 → 貼上」取代編輯器裡的內容
//  4. 左邊齒輪「專案設定」→ 拉到最下面「指令碼屬性」
//     → 新增一筆：屬性 = GEMINI_API_KEY，值 = 你的 API 金鑰
//  5. 按 Ctrl+S 儲存 → 點上方「執行 ▶」→ 函式選 onOpen → 執行
//  6. 會跳出授權視窗 → 照著點「允許」就好
//  7. 回到 Sheets 按 F5 重新整理 → 上方選單就會出現「AI 審題助手」
//
//  【使用方式】
//  1. 在 Sheet1 的 A 欄貼上段考題目（A1 寫標題，A2 開始每列一題）
//  2. 點選「🤖 AI 審題助手」→「📊 分析段考試題」
//  3. 等 30 秒～1 分鐘 → 自動產生三個新工作表
//
//  【API 金鑰哪裡拿？】
//  到 https://aistudio.google.com/apikey → 建立 API 金鑰 → 複製
//
// ============================================================

// ---------- 取得 API 金鑰 ----------
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

// ---------- 建立自訂選單（Sheets 開啟時自動執行）----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 AI 審題助手')
    .addItem('📊 分析段考試題', 'analyzeExam')
    .addItem('ℹ️ 使用說明', 'showHelp')
    .addToUi();
}

// ---------- 使用說明彈窗 ----------
function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <h3>段考 AI 審題助手 使用說明</h3>
    <ol>
      <li>在 Sheet1 的 A 欄貼上段考試題（每列一題）</li>
      <li>點選選單「AI 審題助手」→「分析段考試題」</li>
      <li>等待 AI 分析完成（約 30 秒～1 分鐘）</li>
      <li>結果會自動產生在新的工作表中</li>
    </ol>
    <p><b>產出內容：</b></p>
    <ul>
      <li>雙向細目表（知識向度 × 認知層次）</li>
      <li>難易度分析（簡單/中等/困難比例）</li>
      <li>覆蓋率檢核與改進建議</li>
    </ul>
  `).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '使用說明');
}

// ---------- 主功能：分析段考試題 ----------
function analyzeExam() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1')
                || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // 讀取 A 欄試題
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('請先在 A 欄貼上段考試題（從 A2 開始）');
    return;
  }

  const questions = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(q => q.toString().trim() !== '');

  if (questions.length === 0) {
    ui.alert('未偵測到試題，請確認 A 欄已貼上題目。');
    return;
  }

  ui.alert(`已偵測到 ${questions.length} 題，開始 AI 分析...\n（約需 30 秒～1 分鐘，請稍候）`);

  // 組合試題文字
  const examText = questions.map((q, i) => `第${i+1}題：${q}`).join('\n\n');

  // 呼叫 Gemini API
  const analysis = callGemini(examText);

  if (analysis) {
    outputResults(analysis);
    ui.alert('✅ 分析完成！請查看新產生的工作表。');
  } else {
    ui.alert('❌ 分析失敗，請檢查 API 金鑰設定。');
  }
}

// ---------- 呼叫 Gemini API ----------
function callGemini(examText) {
  const apiKey = getApiKey();
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('請先設定 GEMINI_API_KEY（專案設定 → 指令碼屬性）');
    return null;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;

  const prompt = `你是一位專業的高中教育評量專家。請分析以下段考試題，並產出：

1. **雙向細目表**（JSON 格式）
   - 列：各知識單元/主題
   - 欄：認知層次（記憶、理解、應用、分析、評鑑、創造）
   - 值：對應的題號

2. **難易度分析**（JSON 格式）
   - 將每題分類為「簡單」「中等」「困難」
   - 統計各難度的題數與百分比

3. **覆蓋率檢核**
   - 哪些單元/主題出太多題？
   - 哪些單元/主題沒有出到？
   - 難度分布是否均衡？

4. **改進建議**
   - 具體建議如何調整出題

請用以下 JSON 格式回傳（確保是合法 JSON）：
{
  "雙向細目表": {
    "單元列表": ["單元A", "單元B", ...],
    "認知層次": ["記憶", "理解", "應用", "分析", "評鑑", "創造"],
    "對應題號": {
      "單元A": {"記憶": [1,2], "理解": [3], ...},
      ...
    }
  },
  "難易度分析": {
    "各題難度": {"1": "中等", "2": "簡單", ...},
    "統計": {"簡單": {"題數": 5, "百分比": "25%"}, "中等": {...}, "困難": {...}}
  },
  "覆蓋率檢核": {
    "出題偏多": ["..."],
    "未覆蓋": ["..."],
    "難度均衡性": "..."
  },
  "改進建議": ["建議1", "建議2", ...]
}

以下是段考試題：

${examText}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 16384
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
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
    const text = json.candidates[0].content.parts[0].text;

    // 從回應中提取 JSON（移除 Markdown 程式碼區塊標記）
    const cleanText = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '');
    const jsonMatch = cleanText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
    return null;
  } catch (e) {
    Logger.log('Gemini API 錯誤：' + e.message);
    return null;
  }
}

// ---------- 將分析結果輸出到新工作表 ----------
function outputResults(analysis) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 工作表：雙向細目表 ---
  let sheetTable = ss.getSheetByName('雙向細目表');
  if (sheetTable) sheetTable.clear();
  else sheetTable = ss.insertSheet('雙向細目表');

  const table = analysis['雙向細目表'];
  if (table) {
    // 固定 6 個認知層次，避免 AI 漏掉某些層次導致欄位不一致
    const fixedLevels = ['記憶', '理解', '應用', '分析', '評鑑', '創造'];
    const headers = ['單元 \\ 認知層次', ...fixedLevels];
    sheetTable.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#4285F4').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center');

    // A 欄（單元名稱）設定足夠寬度，避免遮住 B 欄標題
    sheetTable.setColumnWidth(1, 160);
    // B~G 欄（認知層次）設定統一寬度
    for (let c = 2; c <= headers.length; c++) {
      sheetTable.setColumnWidth(c, 80);
    }

    const units = table['單元列表'] || [];
    units.forEach((unit, i) => {
      const row = [unit];
      fixedLevels.forEach(level => {
        const nums = table['對應題號']?.[unit]?.[level] || [];
        row.push(nums.length > 0 ? nums.join(', ') : '—');
      });
      sheetTable.getRange(i + 2, 1, 1, row.length).setValues([row]);
    });

    // 單元名稱欄靠左，認知層次欄置中
    const dataRows = units.length;
    if (dataRows > 0) {
      sheetTable.getRange(2, 1, dataRows, 1).setHorizontalAlignment('left');
      sheetTable.getRange(2, 2, dataRows, fixedLevels.length).setHorizontalAlignment('center');
    }
  }

  // --- 工作表：難易度分析 ---
  let sheetDiff = ss.getSheetByName('難易度分析');
  if (sheetDiff) sheetDiff.clear();
  else sheetDiff = ss.insertSheet('難易度分析');

  const diff = analysis['難易度分析'];
  if (diff) {
    // 各題難度
    sheetDiff.getRange(1, 1).setValue('各題難度分析')
      .setFontWeight('bold').setFontSize(12);
    sheetDiff.getRange(2, 1, 1, 2).setValues([['題號', '難度']])
      .setBackground('#34A853').setFontColor('white').setFontWeight('bold');

    const items = diff['各題難度'] || {};
    Object.entries(items).forEach(([num, level], i) => {
      sheetDiff.getRange(i + 3, 1, 1, 2).setValues([[`第 ${num} 題`, level]]);
    });

    // 統計
    const statsRow = Object.keys(items).length + 5;
    sheetDiff.getRange(statsRow, 1).setValue('統計摘要')
      .setFontWeight('bold').setFontSize(12);
    sheetDiff.getRange(statsRow + 1, 1, 1, 3).setValues([['難度', '題數', '百分比']])
      .setBackground('#FBBC04').setFontWeight('bold');

    const stats = diff['統計'] || {};
    let row = statsRow + 2;
    ['簡單', '中等', '困難'].forEach(level => {
      const s = stats[level] || {};
      sheetDiff.getRange(row, 1, 1, 3).setValues([[level, s['題數'] || 0, s['百分比'] || '0%']]);
      row++;
    });

    sheetDiff.autoResizeColumns(1, 3);
  }

  // --- 工作表：檢核與建議 ---
  let sheetAdvice = ss.getSheetByName('檢核與建議');
  if (sheetAdvice) sheetAdvice.clear();
  else sheetAdvice = ss.insertSheet('檢核與建議');

  const check = analysis['覆蓋率檢核'] || {};
  const advice = analysis['改進建議'] || [];

  sheetAdvice.getRange(1, 1).setValue('覆蓋率檢核')
    .setFontWeight('bold').setFontSize(12);

  let r = 2;
  sheetAdvice.getRange(r, 1).setValue('出題偏多的單元：').setFontWeight('bold');
  sheetAdvice.getRange(r, 2).setValue((check['出題偏多'] || ['無']).join('、'));
  r++;
  sheetAdvice.getRange(r, 1).setValue('未覆蓋的單元：').setFontWeight('bold');
  sheetAdvice.getRange(r, 2).setValue((check['未覆蓋'] || ['無']).join('、'));
  r++;
  sheetAdvice.getRange(r, 1).setValue('難度均衡性：').setFontWeight('bold');
  sheetAdvice.getRange(r, 2).setValue(check['難度均衡性'] || '—');
  r += 2;

  sheetAdvice.getRange(r, 1).setValue('改進建議')
    .setFontWeight('bold').setFontSize(12);
  r++;
  advice.forEach((a, i) => {
    sheetAdvice.getRange(r + i, 1).setValue(`${i + 1}. ${a}`);
  });

  sheetAdvice.autoResizeColumns(1, 2);
}
