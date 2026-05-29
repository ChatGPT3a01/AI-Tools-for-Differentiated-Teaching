// AI 差異化教學工具 · 全域設定（學員自帶金鑰版）
// =============================================================
// 安全說明：
//   本版本「不再內建任何 API Key」，也不再經過 Cloudflare Worker。
//   每位學員在右上角「🔑 設定金鑰」輸入自己的 OpenAI 或 Google Gemini Key，
//   金鑰只會存在「你自己這台電腦的瀏覽器 localStorage」，
//   並由瀏覽器「直接」傳給 OpenAI / Google，不會經過任何第三方伺服器。
//   要清除金鑰：右上角 →「清除金鑰」，或清除瀏覽器資料即可。
// =============================================================

(function () {
  "use strict";

  // ---------- localStorage 金鑰管理 ----------
  const LS = {
    provider: "ai_provider",          // "openai" | "gemini"
    keyOpenai: "ai_key_openai",
    keyGemini: "ai_key_gemini",
    modelOpenai: "ai_model_openai",
    modelGemini: "ai_model_gemini",
  };

  // 模型選項（依使用者全域規則，2026 最新；學員可在設定面板自選）
  const MODELS = {
    openai: [
      { id: "gpt-4o-mini",        label: "gpt-4o-mini（推薦：便宜快速，適合研習）" },
      { id: "gpt-5.4",            label: "gpt-5.4（2026/3 最新，較貴）" },
      { id: "gpt-5.4-pro",        label: "gpt-5.4-pro（最強，最貴）" },
      { id: "gpt-5.3-chat-latest",label: "gpt-5.3-chat-latest（快速）" },
    ],
    gemini: [
      { id: "gemini-3-flash-preview", label: "gemini-3-flash-preview（推薦：速度快、免費額度高）" },
      { id: "gemini-3.1-pro-preview", label: "gemini-3.1-pro-preview（最強，較慢）" },
    ],
  };

  const cfg = {
    getProvider() { return localStorage.getItem(LS.provider) || "openai"; },
    setProvider(p) { localStorage.setItem(LS.provider, p); },
    getKey(p) { return localStorage.getItem(p === "gemini" ? LS.keyGemini : LS.keyOpenai) || ""; },
    setKey(p, v) { localStorage.setItem(p === "gemini" ? LS.keyGemini : LS.keyOpenai, v); },
    getModel(p) {
      const saved = localStorage.getItem(p === "gemini" ? LS.modelGemini : LS.modelOpenai);
      return saved || MODELS[p][0].id;
    },
    setModel(p, v) { localStorage.setItem(p === "gemini" ? LS.modelGemini : LS.modelOpenai, v); },
    clearAll() {
      Object.values(LS).forEach((k) => localStorage.removeItem(k));
    },
    isReady() {
      const p = this.getProvider();
      return !!this.getKey(p).trim();
    },
  };
  window.aiConfig = cfg;
  window.AI_MODELS = MODELS;

  // ---------- 系統 Prompt（原本放在 Worker，現在搬到前端） ----------
  const SHEN_TI_SYSTEM = `你是台灣高中自然科段考命題顧問，熟悉「台北市段考雙向細目表檢核」標準。
任務：對使用者提供的試題（單題或多題），產出符合教師備課需求的結構化分析。

對每一題，請判斷：
1. 內容向度（題目對應到哪個物理/化學/生物/地科概念與單元，請具體寫到「課綱主題–子主題」層級）
2. 認知向度（依 Anderson 修訂版 Bloom：記憶/理解/應用/分析/評鑑/創造，擇一）
3. 難易度（易/中/難，並用一句話說明判斷理由——是計算量？概念深度？跨單元？情境陌生？）
4. 命題建議（如果是好題目，指出好在哪；如果有缺陷，指出可優化處）

輸出格式：嚴格 JSON，schema：
{
  "items": [
    {
      "number": 題號,
      "stem": "題幹摘要（30 字內）",
      "content_dimension": "課綱主題–子主題",
      "cognitive_dimension": "記憶|理解|應用|分析|評鑑|創造",
      "difficulty": "易|中|難",
      "difficulty_reason": "判斷理由（一句話）",
      "suggestion": "命題建議"
    }
  ],
  "summary": {
    "total": 題數,
    "by_difficulty": {"易": n, "中": n, "難": n},
    "by_cognitive": {"記憶": n, "理解": n, ...},
    "overall_comment": "整份卷的整體評語（兩三句）"
  }
}
只輸出 JSON，不要前後加任何說明文字。`;

  const DONG_TAI_SYSTEM = `你是台灣高中自然科差異化教學設計顧問，專長動態評量（Dynamic Assessment）。
任務：對使用者提供的試題或概念，產出三個程度層次的「鷹架提示」，協助不同學習狀態的學生靠自己想出來。

對每一題或每一個概念，請產出：
- 高層次提示（給高程度學生）：只給最小提示，引導他「想清楚自己為什麼會錯」
- 中層次提示（給中程度學生）：點出關鍵公式或概念連結，但不直接給答案
- 低層次提示（給低程度學生）：把問題拆成更小的子問題、給範例對照

設計原則：
- 提示是「動詞」而非「答案」——引導學生動腦，不是替學生答
- 每層提示之間有清楚的鷹架程度差異
- 用學生看得懂的口語，不要太正式

輸出格式：嚴格 JSON，schema：
{
  "items": [
    {
      "number": 題號或編號,
      "topic": "概念主題（10 字內）",
      "hints": {
        "high": "高程度學生的提示（≤80 字）",
        "mid": "中程度學生的提示（≤80 字）",
        "low": "低程度學生的提示（≤120 字，可拆步驟）"
      },
      "key_concept": "這題核心概念（一句話）"
    }
  ]
}
只輸出 JSON，不要前後加任何說明文字。`;

  const GRADE_SYSTEM = `你是台灣高中自然科老師，正在和學生做動態評量互動。
學生會給你一道題目以及他的作答，你要：
1. 判斷他是否答對（如果題目有明確答案）
2. 給予溫暖、鼓勵口吻的回饋（不要太嚴肅）
3. 簡短解釋為什麼（讓他理解，不只記答案）
4. 如果答錯，給一個「下一步可以怎麼想」的引導（不直接給答案）

輸出嚴格 JSON：
{
  "correct": true|false|null,
  "correct_answer": "這題的正確答案（一句話，例如 '(B) 15 m/s' 或 '光合作用 + 呼吸作用同時進行'）",
  "feedback": "給學生的回饋（溫暖鼓勵口吻，50 字內，正向開頭如「不錯喔！」「再想想看～」）",
  "explanation": "為什麼這樣答（簡短解釋核心概念，60 字內）",
  "hint_for_retry": "如果答錯，給下一步引導；如果答對，可放空字串。30 字內"
}

注意：
- correct 為 null 表示題目本身是開放問題、無單一正確答案，請給予 constructive feedback
- 不要直接抄學生的答案到 explanation
- feedback 用對學生說話的口氣，不要用「該學生」這種第三人稱
只輸出 JSON，不要前後加任何說明文字。`;

  // 給「Prompt 對比實驗」用：模擬使用者沒給好 prompt 的情境
  const DEMO_BAD_SYSTEM = "你是一個 AI 助理，請回答使用者的問題。";

  // endpoint → system prompt 對照（沿用原本 Worker 的路徑命名，工具頁不必改）
  const ENDPOINT_PROMPT = {
    "/api/shen-ti": SHEN_TI_SYSTEM,
    "/api/dong-tai": DONG_TAI_SYSTEM,
    "/api/grade": GRADE_SYSTEM,
  };

  // ---------- 直接呼叫 AI 供應商 ----------
  async function callOpenAI(systemPrompt, userContent, jsonMode) {
    const key = cfg.getKey("openai").trim();
    const model = cfg.getModel("openai");
    const body = {
      model,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userContent },
      ],
      temperature: 0.3,
    };
    if (jsonMode) body.response_format = { type: "json_object" };

    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${key}`,
      },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(friendlyError("openai", resp.status, errText));
    }
    const data = await resp.json();
    return data.choices?.[0]?.message?.content || (jsonMode ? "{}" : "");
  }

  async function callGemini(systemPrompt, userContent, jsonMode) {
    const key = cfg.getKey("gemini").trim();
    const model = cfg.getModel("gemini");
    const generationConfig = { temperature: 0.3 };
    if (jsonMode) generationConfig.responseMimeType = "application/json";

    const body = {
      systemInstruction: { parts: [{ text: systemPrompt }] },
      contents: [{ role: "user", parts: [{ text: userContent }] }],
      generationConfig,
    };

    const url =
      `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=` +
      encodeURIComponent(key);
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(friendlyError("gemini", resp.status, errText));
    }
    const data = await resp.json();
    const parts = data.candidates?.[0]?.content?.parts || [];
    return parts.map((p) => p.text || "").join("") || (jsonMode ? "{}" : "");
  }

  function friendlyError(provider, status, errText) {
    const head = provider === "gemini" ? "Google Gemini" : "OpenAI";
    if (status === 401 || status === 403) {
      return `${head} 金鑰無效或沒有權限（HTTP ${status}）。請按右上角「🔑 設定金鑰」確認 Key 是否正確、是否已開通。`;
    }
    if (status === 429) {
      return `${head} 額度用盡或請求太頻繁（HTTP 429）。請稍等一下再試，或確認你的帳戶額度。`;
    }
    if (status === 400) {
      return `${head} 請求被拒（HTTP 400）。可能是模型名稱不適用，請到設定面板換一個模型試試。\n細節：${String(errText).slice(0, 200)}`;
    }
    return `${head} 回應錯誤 HTTP ${status}：${String(errText).slice(0, 200)}`;
  }

  // ---------- 對外：與舊版 callAPI 介面相容 ----------
  // 回傳「已解析的 JSON 物件」；解析失敗時回 { _parse_error:true, _raw:"..." }
  window.callAPI = async function (endpoint, input) {
    if (!cfg.isReady()) {
      openSettings();
      throw new Error("尚未設定 API 金鑰。請按右上角「🔑 設定金鑰」，輸入你自己的 OpenAI 或 Google Gemini Key。");
    }
    const systemPrompt = ENDPOINT_PROMPT[endpoint];
    if (!systemPrompt) throw new Error("未知的功能：" + endpoint);

    const userContent = String(input || "").slice(0, 8000);
    if (!userContent.trim()) throw new Error("輸入內容是空的");

    const provider = cfg.getProvider();
    const raw = provider === "gemini"
      ? await callGemini(systemPrompt, userContent, true)
      : await callOpenAI(systemPrompt, userContent, true);

    try {
      return JSON.parse(raw);
    } catch {
      // 有些模型會在 JSON 外包 ```json ... ```，嘗試擷取
      const m = raw.match(/\{[\s\S]*\}/);
      if (m) {
        try { return JSON.parse(m[0]); } catch (_) { /* fallthrough */ }
      }
      return { _parse_error: true, _raw: raw };
    }
  };

  // 「Prompt 對比實驗 — 爛 prompt」用：回傳純文字字串
  window.callDemoBad = async function (input) {
    if (!cfg.isReady()) {
      openSettings();
      throw new Error("尚未設定 API 金鑰。請按右上角「🔑 設定金鑰」設定。");
    }
    const userContent = "幫我分析這題：\n" + String(input || "").slice(0, 8000);
    const provider = cfg.getProvider();
    return provider === "gemini"
      ? await callGemini(DEMO_BAD_SYSTEM, userContent, false)
      : await callOpenAI(DEMO_BAD_SYSTEM, userContent, false);
  };

  // =============================================================
  // 金鑰設定 UI（自動注入到每個有引用 config.js 的頁面）
  // =============================================================
  function injectUI() {
    if (document.getElementById("ai-key-fab")) return;

    const style = document.createElement("style");
    style.textContent = `
      #ai-key-fab{position:fixed;top:12px;right:12px;z-index:9998;
        background:#fff;color:#c2410c;border:2px solid #c2410c;border-radius:999px;
        padding:7px 14px;font-size:13px;font-weight:700;cursor:pointer;
        box-shadow:0 2px 10px rgba(0,0,0,.15);font-family:inherit;line-height:1;}
      #ai-key-fab.ready{color:#15803d;border-color:#15803d;}
      #ai-key-fab:hover{filter:brightness(.97);}
      #ai-key-overlay{position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.5);
        display:none;align-items:flex-start;justify-content:center;padding:24px 12px;overflow-y:auto;}
      #ai-key-overlay.show{display:flex;}
      .ai-key-box{background:#fff;border-radius:16px;max-width:480px;width:100%;
        padding:24px;box-shadow:0 12px 40px rgba(0,0,0,.3);font-family:inherit;color:#1f2937;margin-top:24px;}
      .ai-key-box h3{margin:0 0 4px;font-size:20px;color:#9a3412;}
      .ai-key-box p.sub{margin:0 0 16px;font-size:13px;color:#6b7280;line-height:1.6;}
      .ai-key-box label{display:block;font-size:13px;font-weight:700;margin:14px 0 6px;color:#374151;}
      .ai-key-box input[type=password],.ai-key-box input[type=text],.ai-key-box select{
        width:100%;box-sizing:border-box;padding:10px 12px;border:2px solid #e5e7eb;
        border-radius:10px;font-size:14px;font-family:inherit;}
      .ai-prov-row{display:flex;gap:8px;margin-bottom:4px;}
      .ai-prov-btn{flex:1;padding:10px;border:2px solid #e5e7eb;border-radius:10px;background:#fff;
        cursor:pointer;font-size:14px;font-weight:700;color:#6b7280;}
      .ai-prov-btn.active{border-color:#c2410c;background:#fff7ed;color:#9a3412;}
      .ai-key-note{background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;
        padding:10px 12px;font-size:12px;color:#075985;line-height:1.7;margin-top:14px;}
      .ai-key-note a{color:#0369a1;font-weight:700;}
      .ai-key-actions{display:flex;gap:8px;margin-top:18px;}
      .ai-key-actions button{flex:1;padding:11px;border-radius:10px;font-size:14px;font-weight:700;
        cursor:pointer;border:none;font-family:inherit;}
      .ai-btn-save{background:#c2410c;color:#fff;}
      .ai-btn-clear{background:#fff;color:#b91c1c;border:2px solid #fecaca;}
      .ai-btn-close{background:#f3f4f6;color:#374151;}
      .ai-key-status{font-size:12px;margin-top:10px;font-weight:700;}
      .ai-key-status.ok{color:#15803d;} .ai-key-status.no{color:#b45309;}
    `;
    document.head.appendChild(style);

    const fab = document.createElement("button");
    fab.id = "ai-key-fab";
    fab.type = "button";
    fab.onclick = openSettings;
    document.body.appendChild(fab);

    const overlay = document.createElement("div");
    overlay.id = "ai-key-overlay";
    overlay.innerHTML = `
      <div class="ai-key-box" role="dialog" aria-modal="true">
        <h3>🔑 設定你的 AI 金鑰</h3>
        <p class="sub">本工具不提供共用金鑰。請選一個供應商、貼上<strong>你自己的</strong> API Key。
        金鑰只會存在<strong>你這台電腦的瀏覽器</strong>，直接傳給 OpenAI／Google，<strong>不會經過任何中介伺服器</strong>。</p>

        <label>選擇供應商</label>
        <div class="ai-prov-row">
          <button type="button" class="ai-prov-btn" data-prov="openai">OpenAI</button>
          <button type="button" class="ai-prov-btn" data-prov="gemini">Google Gemini</button>
        </div>

        <label id="ai-key-label">API Key</label>
        <input type="password" id="ai-key-input" placeholder="貼上你的 API Key" autocomplete="off" spellcheck="false">

        <label>使用模型</label>
        <select id="ai-model-select"></select>

        <div class="ai-key-note" id="ai-key-help"></div>
        <div class="ai-key-status" id="ai-key-status"></div>

        <div class="ai-key-actions">
          <button type="button" class="ai-btn-save"  id="ai-key-save">💾 儲存</button>
          <button type="button" class="ai-btn-clear" id="ai-key-clear">🗑️ 清除金鑰</button>
          <button type="button" class="ai-btn-close" id="ai-key-close-btn">關閉</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);

    overlay.addEventListener("click", (e) => { if (e.target === overlay) closeSettings(); });
    overlay.querySelectorAll(".ai-prov-btn").forEach((b) => {
      b.addEventListener("click", () => selectProvider(b.dataset.prov));
    });
    document.getElementById("ai-key-save").onclick = saveSettings;
    document.getElementById("ai-key-clear").onclick = clearSettings;
    document.getElementById("ai-key-close-btn").onclick = closeSettings;

    refreshFab();
  }

  let uiProvider = "openai"; // 設定面板當前選的供應商（尚未儲存前的暫存）

  const HELP = {
    openai: '沒有 Key？到 <a href="https://platform.openai.com/api-keys" target="_blank" rel="noopener">platform.openai.com/api-keys</a> 申請（需綁定付款，研習用量極少、花費通常不到幾塊台幣）。Key 開頭為 <code>sk-</code>。',
    gemini: '沒有 Key？到 <a href="https://aistudio.google.com/apikey" target="_blank" rel="noopener">aistudio.google.com/apikey</a> 免費申請（Google 帳號即可，有免費額度）。Key 開頭為 <code>AIza</code>。',
  };

  function renderModelOptions(p) {
    const sel = document.getElementById("ai-model-select");
    const current = cfg.getModel(p);
    sel.innerHTML = MODELS[p].map((m) =>
      `<option value="${m.id}" ${m.id === current ? "selected" : ""}>${m.label}</option>`
    ).join("");
  }

  function selectProvider(p) {
    uiProvider = p;
    document.querySelectorAll("#ai-key-overlay .ai-prov-btn").forEach((b) => {
      b.classList.toggle("active", b.dataset.prov === p);
    });
    document.getElementById("ai-key-input").value = cfg.getKey(p);
    document.getElementById("ai-key-label").textContent =
      p === "gemini" ? "Google Gemini API Key" : "OpenAI API Key";
    document.getElementById("ai-key-help").innerHTML = HELP[p];
    renderModelOptions(p);
    updateStatus();
  }

  function updateStatus() {
    const el = document.getElementById("ai-key-status");
    const has = !!cfg.getKey(uiProvider).trim();
    el.className = "ai-key-status " + (has ? "ok" : "no");
    el.textContent = has
      ? `✅ 已儲存 ${uiProvider === "gemini" ? "Gemini" : "OpenAI"} 金鑰`
      : `⚠️ 尚未儲存 ${uiProvider === "gemini" ? "Gemini" : "OpenAI"} 金鑰`;
  }

  function refreshFab() {
    const fab = document.getElementById("ai-key-fab");
    if (!fab) return;
    if (cfg.isReady()) {
      const p = cfg.getProvider();
      fab.classList.add("ready");
      fab.textContent = `🔑 ${p === "gemini" ? "Gemini" : "OpenAI"} 已就緒`;
    } else {
      fab.classList.remove("ready");
      fab.textContent = "🔑 設定金鑰";
    }
  }

  function openSettings() {
    if (!document.getElementById("ai-key-overlay")) injectUI();
    selectProvider(cfg.getProvider());
    document.getElementById("ai-key-overlay").classList.add("show");
  }
  function closeSettings() {
    const o = document.getElementById("ai-key-overlay");
    if (o) o.classList.remove("show");
  }
  window.openAIKeySettings = openSettings;

  function saveSettings() {
    const key = document.getElementById("ai-key-input").value.trim();
    const model = document.getElementById("ai-model-select").value;
    cfg.setKey(uiProvider, key);
    cfg.setModel(uiProvider, model);
    cfg.setProvider(uiProvider);
    updateStatus();
    refreshFab();
    if (key) {
      closeSettings();
    }
  }

  function clearSettings() {
    if (!confirm("確定要清除已儲存的金鑰嗎？")) return;
    cfg.clearAll();
    document.getElementById("ai-key-input").value = "";
    selectProvider("openai");
    refreshFab();
  }

  // 初始化
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", injectUI);
  } else {
    injectUI();
  }
})();
