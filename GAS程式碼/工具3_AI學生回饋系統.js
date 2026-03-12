// ============================================================
//  工具 ③ AI 學生回饋系統
// ============================================================
//
//  【這個工具做什麼？】
//  貼上「學生作品」+「Rubric 評分規準」→ AI 自動產生：
//    - 逐項評分（根據 Rubric 每個面向給等級＋分數）
//    - 做得好的地方（引用作品內容，不是空泛稱讚）
//    - 需要改進的地方（誠實指出不足）
//    - 具體改進建議（告訴學生下一步怎麼做）
//    - 鼓勵語（真誠但不浮誇）
//    - 自動計分 + 等第判定
//    - 全班統計報表 + 弱點分析
//    - 一鍵寄信給學生
//
//  【安裝步驟（5 分鐘）】
//  1. 新增一份 Google Sheets，命名為「AI 學生回饋系統」
//  2. 點選上方選單「擴充功能」→「Apps Script」
//  3. 把這整份程式碼「全選 → 複製 → 貼上」取代編輯器裡的內容
//  4. 左邊齒輪「專案設定」→ 拉到最下面「指令碼屬性」
//     → 新增一筆：屬性 = GEMINI_API_KEY，值 = 你的 API 金鑰
//  5. 按 Ctrl+S 儲存 → 點上方「執行 ▶」→ 函式選 onOpen → 執行
//  6. 會跳出授權視窗 → 照著點「允許」就好
//  7. 回到 Sheets 按 F5 重新整理 → 上方選單出現「💬 AI 回饋」
//
//  【使用方式】
//  1. 點「📋 初始化欄位標頭」設定表頭
//  2. 點「📝 產生 Rubric 模板參考」取得 Rubric 範例
//  3. 在主工作表填入：
//     A 欄 = 學生姓名   B 欄 = Email（選填）
//     C 欄 = 評量形式   D 欄 = Rubric
//     E 欄 = 學生作品（貼 Google Drive 連結，自動擷取內容）
//           支援：Google Docs / Slides / 圖片 / 純文字
//  4. 點「🔄 產生個人化回饋」或「📧 批次產生全部回饋」
//  5. 點「📊 產生全班統計報表」查看全班弱點分析
//  6. 點「✉️ 寄送回饋信件給學生」將回饋寄出
//
//  【API 金鑰哪裡拿？】
//  到 https://aistudio.google.com/apikey → 建立 API 金鑰 → 複製
//
// ============================================================

// ---------- 建立自訂選單 ----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💬 AI 回饋')
    .addItem('📋 初始化欄位標頭', 'setupHeaders')
    .addItem('📝 產生 Rubric 模板參考', 'createRubricTemplates')
    .addSeparator()
    .addItem('🔄 產生個人化回饋（含評分）', 'generateFeedback')
    .addItem('📧 批次產生全部回饋', 'generateAllFeedback')
    .addSeparator()
    .addItem('📊 產生全班統計報表', 'generateClassReport')
    .addItem('✉️ 寄送回饋信件給學生', 'sendFeedbackEmails')
    .addToUi();
}

// ---------- 初始化欄位標頭 ----------
function setupHeaders() {
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
function createRubricTemplates() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var templateSheet = ss.getSheetByName('Rubric 模板參考');
  if (templateSheet) templateSheet.clear();
  else templateSheet = ss.insertSheet('Rubric 模板參考');

  var templates = [
    {
      title: '📝 文字報告型 Rubric',
      color: '#4285F4',
      rows: [
        ['內容正確性', '核心概念完全正確，能清楚解釋原理並舉出恰當例子', '大部分概念正確，但有少數解釋不夠精確', '有明顯概念錯誤，或僅複製教材未加理解'],
        ['組織架構', '段落分明，有清楚的前言、本文、結論，邏輯連貫', '有基本架構但過渡生硬，部分段落邏輯不清', '缺乏架構，想到什麼寫什麼，讀者難以跟隨'],
        ['數據佐證', '引用具體數據或實驗結果支持論點，來源可信', '有提及數據但不夠具體，或來源不明', '完全沒有數據佐證，僅靠感覺論述'],
        ['語言表達', '用詞精準，句型多變，學術語言與日常說法並用', '表達尚可但用詞重複，偶有口語化問題', '錯字多、語句不通，或過度口語化']
      ]
    },
    {
      title: '🎨 視覺圖像型 Rubric',
      color: '#8B5CF6',
      rows: [
        ['概念連結', '清楚呈現概念之間的關聯性，箭頭與連結線有意義', '部分連結正確，但有遺漏或不必要的連結', '概念散亂放置，缺乏有意義的連結'],
        ['視覺呈現', '配色協調、版面整潔、圖文比例恰當，一目了然', '尚可辨識但略雜亂，部分區域過於擁擠', '難以閱讀、配色混亂或字體過小'],
        ['資訊完整性', '涵蓋所有關鍵概念，無重要遺漏', '遺漏 1-2 個重要概念或次要細節', '缺少多個主要概念，內容不完整'],
        ['創意表達', '圖像富有巧思與原創性，用獨特方式呈現知識', '有嘗試但較為制式，缺乏個人風格', '直接複製教材圖表，無任何創意加工']
      ]
    },
    {
      title: '🎙️ 口說 Podcast 型 Rubric',
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
    .setValue('💡 使用方式：從下方選擇適合的 Rubric 模板，複製整個表格，貼到主工作表的 D 欄（Rubric 評分規準）')
    .setBackground('#E8F0FE').setFontWeight('bold').setWrap(true);
  templateSheet.setRowHeight(currentRow, 40);
  currentRow += 2;

  templates.forEach(function(tpl) {
    // 類型標題
    templateSheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue(tpl.title)
      .setBackground(tpl.color).setFontColor('white').setFontWeight('bold')
      .setFontSize(14);
    currentRow++;

    // 表頭
    templateSheet.getRange(currentRow, 1, 1, 4)
      .setValues([['評分面向', '優（5分）', '中（3分）', '再加強（1分）']])
      .setBackground('#E8EAED').setFontWeight('bold');
    currentRow++;

    // 內容
    tpl.rows.forEach(function(row) {
      templateSheet.getRange(currentRow, 1, 1, 4).setValues([row]);
      templateSheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
    });

    // 滿分說明
    templateSheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue('滿分：4 個面向 × 5 分 = 20 分')
      .setFontColor('#666666').setFontStyle('italic');
    currentRow += 2;
  });

  // 可直接複製的 Rubric 文字版
  currentRow++;
  templateSheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('📋 可直接貼到 D 欄的 Rubric 純文字版')
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
//
// 支援的連結格式：
//   - Google Docs:   https://docs.google.com/document/d/FILE_ID/...
//   - Google Slides: https://docs.google.com/presentation/d/FILE_ID/...
//   - Google Drive:  https://drive.google.com/file/d/FILE_ID/...
//   - Google Drive:  https://drive.google.com/open?id=FILE_ID
//
// 回傳格式：{ type: 'text'|'image'|'error', content: string, mimeType?: string }
//
function extractContentFromDrive(urlOrText) {
  if (!urlOrText) return { type: 'error', content: '欄位為空' };
  var text = urlOrText.toString().trim();
  if (text === '') return { type: 'error', content: '欄位為空' };

  // --- Google Docs ---
  var docsMatch = text.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
  if (docsMatch) {
    try {
      var doc = DocumentApp.openById(docsMatch[1]);
      return { type: 'text', content: doc.getBody().getText() };
    } catch (e) {
      return { type: 'error', content: '無法開啟 Google Doc：' + e.message };
    }
  }

  // --- Google Slides ---
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

  // --- Google Drive 檔案連結 ---
  var driveMatch = text.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/)
    || text.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (driveMatch) {
    try {
      var file = DriveApp.getFileById(driveMatch[1]);
      var mime = file.getMimeType();

      // 圖片 → base64（給 Gemini 多模態分析）
      if (mime.indexOf('image/') === 0) {
        var blob = file.getBlob();
        var base64 = Utilities.base64Encode(blob.getBytes());
        return { type: 'image', content: base64, mimeType: mime };
      }

      // Google Docs（從 Drive 開啟）
      if (mime === 'application/vnd.google-apps.document') {
        var doc2 = DocumentApp.openById(driveMatch[1]);
        return { type: 'text', content: doc2.getBody().getText() };
      }

      // Google Slides（從 Drive 開啟）
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

      // 純文字檔
      if (mime.indexOf('text/') === 0) {
        return { type: 'text', content: file.getBlob().getDataAsString() };
      }

      return { type: 'error', content: '不支援的檔案類型：' + mime + '\n請改用 Google Docs 或 Slides 連結。' };
    } catch (e) {
      return { type: 'error', content: '無法開啟檔案：' + e.message + '\n請確認檔案已開啟分享權限。' };
    }
  }

  // --- 不是連結 → 當作純文字 ---
  return { type: 'text', content: text };
}

// ---------- 產生單一學生的回饋（含評分）----------
function generateFeedback() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var row = SpreadsheetApp.getActiveRange().getRow();

  if (row < 2) {
    ui.alert('請選取學生資料列（第 2 列以下）再執行。');
    return;
  }

  var data = sheet.getRange(row, 1, 1, 5).getValues()[0];
  var name = data[0];    // A 欄
  var format = data[2];  // C 欄
  var rubric = data[3];  // D 欄
  var work = data[4];    // E 欄

  if (!work || work.toString().trim() === '') {
    ui.alert('該列的「學生作品」(E欄) 是空的。\n請貼上 Google Drive 連結或作品文字。');
    return;
  }

  if (!rubric || rubric.toString().trim() === '') {
    ui.alert('該列的「Rubric 評分規準」(D欄) 是空的，請先填入再執行。\n\n' +
      '提示：可點選「📝 產生 Rubric 模板參考」取得範例。');
    return;
  }

  // 擷取作品內容
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '正在擷取 ' + name + ' 的作品內容...', 'AI 回饋', -1
  );
  var extracted = extractContentFromDrive(work);

  if (extracted.type === 'error') {
    ui.alert('❌ 無法擷取作品內容：\n\n' + extracted.content);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '正在為 ' + name + ' 產生個人化回饋...', 'AI 回饋', -1
  );

  var feedback = callGeminiFeedback(name, format, rubric, extracted);

  if (feedback) {
    writeFeedbackToSheet(sheet, row, feedback);
    updateDimensionHeaders(sheet, feedback);
    ui.alert('✅ 已為 ' + name + ' 產生回饋！\n請查看 F~N 欄。');
  } else {
    ui.alert('❌ 回饋產生失敗，請檢查 API 金鑰或稍後再試。');
  }
}

// ---------- 批次產生全部回饋 ----------
function generateAllFeedback() {
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
    var name = data[0];       // A 欄
    var format = data[2];     // C 欄
    var rubric = data[3];     // D 欄
    var work = data[4];       // E 欄
    var existing = data[5];   // F 欄（已有回饋）

    // 跳過已有回饋的或空白的
    if ((existing && existing.toString().trim() !== '') || !work || work.toString().trim() === '') continue;

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '正在擷取 ' + name + ' 的作品...（第 ' + (count + 1) + ' 位）', 'AI 回饋', -1
    );

    var extracted = extractContentFromDrive(work);
    if (extracted.type === 'error') {
      failed.push(name + '（擷取失敗）');
      continue;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '正在為 ' + name + ' 產生回饋...（第 ' + (count + 1) + ' 位）', 'AI 回饋', -1
    );

    var feedback = callGeminiFeedback(name, format, rubric, extracted);
    if (feedback) {
      writeFeedbackToSheet(sheet, row, feedback);

      // 只更新一次面向標頭
      if (!headersUpdated) {
        updateDimensionHeaders(sheet, feedback);
        headersUpdated = true;
      }
      count++;
    } else {
      failed.push(name);
    }

    // 避免 API 速率限制
    Utilities.sleep(2000);
  }

  var msg = '✅ 完成！共產生 ' + count + ' 份個人化回饋。';
  if (failed.length > 0) {
    msg += '\n\n以下學生回饋產生失敗：' + failed.join('、');
  }
  ui.alert(msg);
}

// ---------- 呼叫 Gemini API 產生回饋 ----------
// workData 格式：{ type: 'text'|'image', content: string, mimeType?: string }
function callGeminiFeedback(name, format, rubric, workData) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('❌ 尚未設定 API 金鑰！\n\n請到「專案設定」→「指令碼屬性」中新增 GEMINI_API_KEY');
    return null;
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=' + apiKey;

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
    + '### 📊 逐項評分\n'
    + '（根據 Rubric 的每個面向，給出等級＋分數＋簡短說明）\n'
    + '格式範例：\n'
    + '- **內容正確性**：優（5/5）— 核心概念解釋清楚...\n'
    + '- **組織架構**：中（3/5）— 有基本架構但...\n\n'
    + '### ✅ 做得好的地方（2～3 點）\n'
    + '（具體指出優點，引用作品中的內容）\n\n'
    + '### ⚠️ 需要改進的地方（2～3 點）\n'
    + '（明確指出不足，說明為什麼這樣不好）\n\n'
    + '### 💡 具體改進建議（2～3 點）\n'
    + '（告訴學生下一步可以怎麼做）\n\n'
    + '### 💪 鼓勵語\n'
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

    var responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('API HTTP ' + responseCode + ': ' + response.getContentText().substring(0, 500));
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

// ---------- 解析回饋中的分數 JSON ----------
function parseScoresFromFeedback(feedbackText) {
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

  // 策略 2：找任何含 "dimensions" 的 JSON（AI 可能沒用標記）
  var jsonPatterns = text.match(/```json\s*([\s\S]*?)```/g) || [];
  for (var i = 0; i < jsonPatterns.length; i++) {
    try {
      var clean = jsonPatterns[i].replace(/```json\s*/gi, '').replace(/```\s*/g, '');
      var obj = JSON.parse(clean);
      if (obj.dimensions) return obj;
    } catch (e) { /* 繼續嘗試 */ }
  }

  // 策略 3：找回饋文字中的 X/5 分數模式（最可靠的 fallback）
  var scorePattern = /\*\*([^*]+)\*\*[：:]\s*[^\d]*(\d)\/5/g;
  var dims = [];
  var m;
  while ((m = scorePattern.exec(text)) !== null) {
    dims.push({ name: m[1].trim(), score: parseInt(m[2]), max: 5 });
  }

  // 如果 策略3 也找不到，試另一種格式：面向名（X/5）
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
function cleanFeedbackText(feedbackText) {
  return feedbackText.toString()
    .replace(/<!--SCORES_JSON[\s\S]*?SCORES_JSON-->/g, '')
    .replace(/```json\s*\{[\s\S]*?"dimensions"[\s\S]*?\}\s*```/g, '')
    .trim();
}

// ---------- 將回饋和分數寫入工作表 ----------
function writeFeedbackToSheet(sheet, row, rawFeedback) {
  var cleanedFeedback = cleanFeedbackText(rawFeedback);
  var scores = parseScoresFromFeedback(rawFeedback);

  // F 欄 = 回饋文字
  sheet.getRange(row, 6).setValue(cleanedFeedback);

  if (scores && scores.dimensions) {
    // G~J 欄 = 面向分數（最多 4 個）
    for (var i = 0; i < 4; i++) {
      if (i < scores.dimensions.length) {
        sheet.getRange(row, 7 + i).setValue(scores.dimensions[i].score);
        // 上色
        var color = getScoreColor(scores.dimensions[i].score);
        sheet.getRange(row, 7 + i).setBackground(color);
      } else {
        sheet.getRange(row, 7 + i).setValue('');
      }
    }
    // K 欄 = 總分, L 欄 = 滿分, M 欄 = 百分比, N 欄 = 等第
    sheet.getRange(row, 11).setValue(scores.total);
    sheet.getRange(row, 12).setValue(scores.maxTotal);
    sheet.getRange(row, 13).setValue(scores.percentage + '%');
    sheet.getRange(row, 14).setValue(scores.grade);

    // 等第上色
    var gradeColors = { '優': '#34A853', '良': '#4285F4', '中': '#FBBC04', '再加強': '#EA4335' };
    var gradeColor = gradeColors[scores.grade] || '#FFFFFF';
    sheet.getRange(row, 14).setBackground(gradeColor).setFontColor('white').setFontWeight('bold');
  }
}

// ---------- 根據分數回傳背景色 ----------
function getScoreColor(score) {
  if (score >= 5) return '#34A853';  // 綠
  if (score >= 3) return '#FBBC04';  // 黃
  return '#EA4335';                   // 紅
}

// ---------- 用第一位學生的面向名稱更新標頭 ----------
function updateDimensionHeaders(sheet, rawFeedback) {
  var scores = parseScoresFromFeedback(rawFeedback);
  if (!scores || !scores.dimensions) return;

  for (var i = 0; i < Math.min(scores.dimensions.length, 4); i++) {
    sheet.getRange(1, 7 + i).setValue(scores.dimensions[i].name);
  }
}

// ---------- 產生全班統計報表 ----------
function generateClassReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var lastRow = dataSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('沒有學生資料。');
    return;
  }

  // 讀取面向名稱
  var dimNames = dataSheet.getRange(1, 7, 1, 4).getValues()[0];
  // 讀取所有資料
  var allData = dataSheet.getRange(2, 1, lastRow - 1, 14).getValues();

  // 先嘗試從 F 欄回饋文字重新解析分數（修復先前解析失敗的情況）
  var reparsed = 0;
  for (var r = 0; r < allData.length; r++) {
    var feedback = allData[r][5]; // F 欄
    var totalScore = allData[r][10]; // K 欄
    // 如果有回饋但沒有分數 → 重新解析
    if (feedback && feedback.toString().trim() !== '' &&
        (totalScore === '' || totalScore === null || totalScore === undefined)) {
      var scores = parseScoresFromFeedback(feedback.toString());
      if (scores && scores.dimensions) {
        // 寫回工作表
        var sheetRow = r + 2;
        for (var d = 0; d < Math.min(scores.dimensions.length, 4); d++) {
          dataSheet.getRange(sheetRow, 7 + d).setValue(scores.dimensions[d].score);
          dataSheet.getRange(sheetRow, 7 + d).setBackground(getScoreColor(scores.dimensions[d].score));
        }
        dataSheet.getRange(sheetRow, 11).setValue(scores.total);
        dataSheet.getRange(sheetRow, 12).setValue(scores.maxTotal);
        dataSheet.getRange(sheetRow, 13).setValue(scores.percentage + '%');
        dataSheet.getRange(sheetRow, 14).setValue(scores.grade);

        var gradeColors = { '優': '#34A853', '良': '#4285F4', '中': '#FBBC04', '再加強': '#EA4335' };
        dataSheet.getRange(sheetRow, 14).setBackground(gradeColors[scores.grade] || '#FFFFFF')
          .setFontColor('white').setFontWeight('bold');

        // 更新記憶體中的資料
        for (var d2 = 0; d2 < Math.min(scores.dimensions.length, 4); d2++) {
          allData[r][6 + d2] = scores.dimensions[d2].score;
        }
        allData[r][10] = scores.total;
        allData[r][11] = scores.maxTotal;
        allData[r][12] = scores.percentage + '%';
        allData[r][13] = scores.grade;

        // 更新面向標頭（用第一筆成功的）
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

  // 寫入學生資料
  var dimSums = [0, 0, 0, 0];
  var dimCounts = [0, 0, 0, 0];
  var dimAllScores = [[], [], [], []];
  var totalSum = 0;
  var gradeDistribution = { '優': 0, '良': 0, '中': 0, '再加強': 0 };

  students.forEach(function(row) {
    var studentRow = [row[0]]; // 姓名

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

    studentRow.push(row[10]); // 總分
    studentRow.push(row[11]); // 滿分
    studentRow.push(row[12]); // 百分比
    studentRow.push(row[13]); // 等第

    totalSum += Number(row[10]) || 0;

    var grade = row[13] ? row[13].toString() : '';
    if (gradeDistribution.hasOwnProperty(grade)) {
      gradeDistribution[grade]++;
    }

    reportSheet.getRange(currentRow, 1, 1, headerA.length).setValues([studentRow]);

    // 等第上色
    var gradeColors = { '優': '#34A853', '良': '#4285F4', '中': '#FBBC04', '再加強': '#EA4335' };
    if (gradeColors[grade]) {
      reportSheet.getRange(currentRow, 9).setBackground(gradeColors[grade])
        .setFontColor('white').setFontWeight('bold');
    }

    currentRow++;
  });

  // 全班平均列
  var avgRow = ['📊 全班平均'];
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
    .setValue('📈 各面向全班分析')
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

  // 標出最弱面向
  if (weakestIdx > 0) {
    reportSheet.getRange(weakestIdx, 1, 1, 6)
      .setBackground('#FEE2E2').setFontColor('#DC2626');
  }

  currentRow += 2;

  // ===== 區塊 C：等第分佈 =====
  reportSheet.getRange(currentRow, 1, 1, 4).merge()
    .setValue('📊 等第分佈統計')
    .setBackground('#0D5F5C').setFontColor('white').setFontWeight('bold').setFontSize(13);
  currentRow++;

  reportSheet.getRange(currentRow, 1, 1, 4)
    .setValues([['等第', '人數', '百分比', '分佈圖']])
    .setBackground('#E8EAED').setFontWeight('bold');
  currentRow++;

  var gradeOrder = ['優', '良', '中', '再加強'];
  var gradeEmojis = { '優': '🟢', '良': '🔵', '中': '🟡', '再加強': '🔴' };

  gradeOrder.forEach(function(grade) {
    var count = gradeDistribution[grade] || 0;
    var pct = students.length > 0 ? Math.round(count / students.length * 100) : 0;
    var bar = '';
    for (var i = 0; i < count; i++) bar += '█';

    reportSheet.getRange(currentRow, 1, 1, 4)
      .setValues([[gradeEmojis[grade] + ' ' + grade, count + ' 人', pct + '%', bar]]);
    currentRow++;
  });

  currentRow += 2;

  // ===== 區塊 D：教學建議 =====
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('💡 教學建議')
    .setBackground('#0D5F5C').setFontColor('white').setFontWeight('bold').setFontSize(13);
  currentRow++;

  // 找出最弱面向名稱
  var weakDimName = '';
  var weakDimAvg = 999;
  for (var d = 0; d < 4; d++) {
    if (dimCounts[d] > 0) {
      var avg = dimSums[d] / dimCounts[d];
      if (avg < weakDimAvg) {
        weakDimAvg = avg;
        weakDimName = dimNames[d] || ('面向' + (d + 1));
      }
    }
  }
  weakDimAvg = Math.round(weakDimAvg * 10) / 10;

  // 找最多人的等第
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
function buildEmailHtml(name, format, feedback, total, maxTotal, pct, grade) {
  // 移除回饋文字中的 emoji（避免 email 亂碼 "?"）
  var cleanFeedback = feedback
    .replace(/📊/g, '[評分]').replace(/✅/g, '[優]').replace(/⚠️/g, '[注意]')
    .replace(/💡/g, '[建議]').replace(/💪/g, '[加油]').replace(/📬/g, '')
    .replace(/[\u{1F300}-\u{1F9FF}]/gu, '');

  // 將 markdown 標題轉為 HTML
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

  // 成績摘要
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
function sendFeedbackEmails() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('沒有學生資料。');
    return;
  }

  // 計算有多少封要寄
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
    var name = data[0];          // A 欄
    var email = data[1] ? data[1].toString().trim() : '';  // B 欄
    var format = data[2];        // C 欄
    var feedback = data[5];      // F 欄
    var total = data[10];        // K 欄
    var maxTotal = data[11];     // L 欄
    var pct = data[12];          // M 欄
    var grade = data[13];        // N 欄

    // 跳過沒有 Email 或沒有回饋的
    if (!email || !feedback || feedback.toString().trim() === '') {
      skipped++;
      continue;
    }

    // 簡易 Email 格式驗證
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      failed.push(name + '（Email 格式錯誤）');
      continue;
    }

    var subject = '【學習回饋】' + name + ' 同學的' + (format || '') + '評量回饋';
    var htmlBody = buildEmailHtml(name, format || '', feedback.toString(), total, maxTotal, pct, grade);

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

    // 避免 Gmail 速率限制
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
