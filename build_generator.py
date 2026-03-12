#!/usr/bin/env python3
"""Build generator.html by embedding GAS code templates."""

import os

# Read the GAS source code
gas_path = os.path.join('GAS程式碼', '整合版_三工具_GAS.js')
with open(gas_path, 'r', encoding='utf-8') as f:
    gas_lines = f.readlines()

gas_code = ''.join(gas_lines)

# Split into sections
# Shared: lines 1-118 (index 0-117)
# Tool1: lines 120-308 (index 119-307)
# Tool2: lines 311-722 (index 310-721)
# Tool3: lines 725-1676 (index 724-end)
shared_code = ''.join(gas_lines[0:118])
tool1_code = ''.join(gas_lines[119:308])
tool2_code = ''.join(gas_lines[310:722])
tool3_code = ''.join(gas_lines[724:])

# The onOpen menu sections (for conditional inclusion)
# Tool1 menu: lines 34-37
# Tool2 menu: lines 39-43
# Tool3 menu: lines 45-54
tool1_menu = ''.join(gas_lines[33:37])
tool2_menu = ''.join(gas_lines[38:43])
tool3_menu = ''.join(gas_lines[44:54])

html = r'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>GAS 程式碼生成器 — AI 工具協助差異化教學</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html { scroll-behavior: smooth; }
        body {
            font-family: 'Microsoft JhengHei', 'Segoe UI', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            color: #333;
        }

        /* ===== Nav ===== */
        .top-nav {
            position: sticky; top: 0; z-index: 1000;
            height: 52px;
            background: rgba(10, 22, 40, 0.95);
            backdrop-filter: blur(12px);
            display: flex; align-items: center; justify-content: space-between;
            padding: 0 24px;
            border-bottom: 1px solid rgba(79,172,254,0.2);
        }
        .top-nav .nav-title { color: #4facfe; font-size: 16px; font-weight: 700; }
        .top-nav .nav-links { display: flex; gap: 6px; }
        .top-nav .nav-links a {
            color: white; text-decoration: none; padding: 6px 14px;
            border-radius: 8px; font-size: 13px;
            border: 1px solid rgba(255,255,255,0.15);
            background: rgba(255,255,255,0.06);
            transition: all 0.25s;
        }
        .top-nav .nav-links a:hover {
            background: rgba(79,172,254,0.2); border-color: #4facfe;
        }

        /* ===== Container ===== */
        .container {
            max-width: 1000px; margin: 0 auto; padding: 30px 20px 60px;
        }

        /* ===== Hero ===== */
        .hero {
            text-align: center; padding: 40px 20px 30px; color: white;
        }
        .hero h1 { font-size: 42px; font-weight: 800; margin-bottom: 12px; text-shadow: 0 4px 12px rgba(0,0,0,0.3); }
        .hero p { font-size: 20px; opacity: 0.9; }

        /* ===== Card ===== */
        .card {
            background: rgba(255,255,255,0.98); border-radius: 24px;
            padding: 40px 44px; margin-bottom: 30px;
            box-shadow: 0 25px 70px rgba(0,0,0,0.25);
        }
        .card h2 {
            font-size: 28px; color: #1e3c72; margin-bottom: 8px; font-weight: 700;
        }
        .card .section-desc { font-size: 16px; color: #888; margin-bottom: 24px; }

        /* ===== Tool Selection ===== */
        .tool-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin: 20px 0; }
        .tool-card {
            border: 2px solid #e2e8f0; border-radius: 16px; padding: 20px; text-align: center;
            cursor: pointer; transition: all 0.3s; position: relative;
        }
        .tool-card:hover { border-color: #667eea; transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.08); }
        .tool-card.active { border-color: #667eea; background: linear-gradient(135deg, rgba(102,126,234,0.06), rgba(118,75,162,0.04)); }
        .tool-card .tool-icon { font-size: 36px; margin-bottom: 8px; }
        .tool-card .tool-name { font-size: 15px; font-weight: 700; color: #1e3c72; margin-bottom: 4px; }
        .tool-card .tool-desc { font-size: 13px; color: #888; }
        .tool-card .tool-check {
            position: absolute; top: 12px; right: 12px; width: 24px; height: 24px;
            border: 2px solid #ccc; border-radius: 6px; display: flex; align-items: center; justify-content: center;
            font-size: 14px; color: white; transition: all 0.3s;
        }
        .tool-card.active .tool-check { background: #667eea; border-color: #667eea; }

        /* ===== Form Controls ===== */
        .form-section { margin-bottom: 28px; }
        .form-section h3 { font-size: 20px; color: #2a5298; margin-bottom: 16px; font-weight: 700; }
        .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
        .form-group { margin-bottom: 16px; }
        .form-label { display: block; font-size: 14px; font-weight: 600; color: #555; margin-bottom: 6px; }
        .form-hint { font-size: 12px; color: #999; margin-top: 4px; }
        .form-input, .form-select {
            width: 100%; padding: 10px 14px; border: 2px solid #e2e8f0; border-radius: 10px;
            font-size: 15px; font-family: inherit; transition: border-color 0.3s; background: white;
        }
        .form-input:focus, .form-select:focus {
            outline: none; border-color: #667eea; box-shadow: 0 0 0 3px rgba(102,126,234,0.12);
        }
        .range-group { display: flex; align-items: center; gap: 12px; }
        .range-group input[type="range"] {
            flex: 1; -webkit-appearance: none; height: 6px; border-radius: 3px; background: #e2e8f0;
        }
        .range-group input[type="range"]::-webkit-slider-thumb {
            -webkit-appearance: none; width: 20px; height: 20px; border-radius: 50%;
            background: #667eea; cursor: pointer; box-shadow: 0 2px 6px rgba(102,126,234,0.4);
        }
        .range-value {
            min-width: 36px; text-align: center; font-weight: 700; color: #667eea; font-size: 15px;
            background: rgba(102,126,234,0.08); padding: 4px 8px; border-radius: 6px;
        }

        /* ===== Accordion ===== */
        .accordion { margin: 20px 0; }
        .acc-item { border: 2px solid #e2e8f0; border-radius: 14px; margin-bottom: 10px; overflow: hidden; transition: border-color 0.3s; }
        .acc-item.open { border-color: #667eea; }
        .acc-header {
            padding: 16px 20px; display: flex; align-items: center; justify-content: space-between;
            cursor: pointer; user-select: none; background: #fafbfc; transition: background 0.3s;
            font-weight: 600; font-size: 16px; color: #333;
        }
        .acc-header:hover { background: #f1f5f9; }
        .acc-header .acc-arrow { transition: transform 0.3s; font-size: 14px; color: #999; }
        .acc-item.open .acc-arrow { transform: rotate(90deg); }
        .acc-body { max-height: 0; overflow: hidden; transition: max-height 0.4s ease; }
        .acc-item.open .acc-body { max-height: 2000px; }
        .acc-body-inner { padding: 0 20px 20px; }
        .acc-item.disabled { opacity: 0.4; pointer-events: none; }

        /* ===== Buttons ===== */
        .btn-primary {
            display: inline-flex; align-items: center; gap: 8px;
            background: linear-gradient(135deg, #667eea, #764ba2); color: white;
            padding: 14px 36px; border: none; border-radius: 12px;
            font-size: 18px; font-weight: 700; font-family: inherit;
            cursor: pointer; transition: all 0.3s;
            box-shadow: 0 4px 15px rgba(102,126,234,0.4);
        }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 8px 25px rgba(102,126,234,0.5); }
        .btn-group { display: flex; gap: 12px; flex-wrap: wrap; margin-top: 16px; }
        .btn-action {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 10px 22px; border: 2px solid #e2e8f0; border-radius: 10px;
            background: white; font-size: 14px; font-weight: 600; font-family: inherit;
            cursor: pointer; transition: all 0.3s; color: #555;
        }
        .btn-action:hover { border-color: #667eea; color: #667eea; background: rgba(102,126,234,0.04); }
        .btn-action.success { border-color: #34A853; color: #34A853; }

        /* ===== Code Preview ===== */
        .code-section { display: none; }
        .code-section.visible { display: block; }
        .code-toolbar {
            background: #2d2d44; padding: 12px 20px; border-radius: 16px 16px 0 0;
            display: flex; justify-content: space-between; align-items: center; color: #cdd6f4;
            font-size: 13px;
        }
        .code-toolbar .code-info { display: flex; gap: 16px; }
        .code-toolbar .code-badge {
            background: rgba(255,255,255,0.1); padding: 3px 10px; border-radius: 6px; font-size: 12px;
        }
        .code-body {
            background: #1e1e2e; padding: 20px; border-radius: 0 0 16px 16px;
            max-height: 500px; overflow: auto;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 13px; line-height: 1.7; color: #cdd6f4;
            white-space: pre; tab-size: 2;
        }
        .code-body .line-num { display: inline-block; width: 45px; color: #585b70; text-align: right; margin-right: 16px; user-select: none; }
        .hl-kw { color: #cba6f7; } /* keyword */
        .hl-str { color: #a6e3a1; } /* string */
        .hl-cmt { color: #6c7086; font-style: italic; } /* comment */
        .hl-num { color: #fab387; } /* number */
        .hl-fn { color: #89b4fa; } /* function name */

        /* ===== Install Tutorial ===== */
        .steps-container { margin: 20px 0; }
        .step-item {
            display: flex; gap: 16px; margin-bottom: 20px; align-items: flex-start;
        }
        .step-num {
            width: 40px; height: 40px; min-width: 40px;
            background: linear-gradient(135deg, #667eea, #764ba2); color: white;
            border-radius: 50%; display: flex; align-items: center; justify-content: center;
            font-size: 18px; font-weight: 700;
        }
        .step-content { flex: 1; }
        .step-content h4 { font-size: 18px; color: #1e3c72; margin-bottom: 6px; }
        .step-content p { font-size: 15px; color: #555; line-height: 1.7; }
        .step-content code {
            background: #f1f5f9; padding: 2px 8px; border-radius: 4px;
            font-family: 'Consolas', monospace; font-size: 14px; color: #667eea;
        }
        .step-content .info-box {
            background: rgba(102,126,234,0.06); border-left: 4px solid #667eea;
            padding: 12px 16px; border-radius: 0 8px 8px 0; margin-top: 10px;
            font-size: 14px; color: #444;
        }
        .step-content .warning-box {
            background: rgba(243,156,18,0.08); border-left: 4px solid #f59e0b;
            padding: 12px 16px; border-radius: 0 8px 8px 0; margin-top: 10px;
            font-size: 14px; color: #7d5a00;
        }

        /* ===== Toast ===== */
        .toast {
            position: fixed; bottom: 30px; left: 50%; transform: translateX(-50%) translateY(100px);
            background: #1e1e2e; color: white; padding: 14px 28px; border-radius: 12px;
            font-size: 15px; font-weight: 600; z-index: 9999;
            box-shadow: 0 8px 30px rgba(0,0,0,0.3); transition: transform 0.4s cubic-bezier(0.175,0.885,0.32,1.275);
            pointer-events: none;
        }
        .toast.show { transform: translateX(-50%) translateY(0); }

        /* ===== Footer ===== */
        .footer {
            text-align: center; padding: 30px; color: rgba(255,255,255,0.7); font-size: 14px;
        }
        .footer a { color: rgba(255,255,255,0.9); text-decoration: none; }

        /* ===== Responsive ===== */
        @media (max-width: 768px) {
            .container { padding: 20px 12px; }
            .card { padding: 28px 20px; border-radius: 18px; }
            .hero h1 { font-size: 28px; }
            .tool-grid { grid-template-columns: 1fr; }
            .form-grid { grid-template-columns: 1fr; }
            .btn-group { flex-direction: column; }
            .code-body { font-size: 11px; max-height: 350px; }
            .top-nav .nav-links a { padding: 4px 8px; font-size: 11px; }
        }
    </style>
</head>
<body>

<!-- Nav -->
<nav class="top-nav">
    <span class="nav-title">GAS 程式碼生成器</span>
    <div class="nav-links">
        <a href="index.html">&#127968; 回首頁</a>
        <a href="Part1_動機理論與段考AI審題助手.html">Part1</a>
        <a href="Part3_動態評量產生器.html">Part3</a>
        <a href="Part4_多元評量與AI學生回饋系統.html">Part4</a>
        <a href="certification.html">&#128218; 認證</a>
    </div>
</nav>

<!-- Hero -->
<div class="hero">
    <h1>&#9881;&#65039; GAS 程式碼生成器</h1>
    <p>自訂參數，一鍵生成你的 AI 教學工具程式碼</p>
</div>

<div class="container">

<!-- ==================== 設定區 ==================== -->
<div class="card">
    <h2>Step 1 &#8212; 選擇工具</h2>
    <p class="section-desc">勾選你需要的工具，取消勾選的工具不會包含在生成的程式碼中</p>

    <div class="tool-grid">
        <div class="tool-card active" data-tool="1" onclick="toggleTool(1)">
            <div class="tool-check">&#10003;</div>
            <div class="tool-icon">&#128202;</div>
            <div class="tool-name">工具 &#9312; 段考 AI 審題助手</div>
            <div class="tool-desc">雙向細目表 + 難易度分析</div>
        </div>
        <div class="tool-card active" data-tool="2" onclick="toggleTool(2)">
            <div class="tool-check">&#10003;</div>
            <div class="tool-icon">&#127919;</div>
            <div class="tool-name">工具 &#9313; 動態評量產生器</div>
            <div class="tool-desc">三層提示 + Google Form + 計分</div>
        </div>
        <div class="tool-card active" data-tool="3" onclick="toggleTool(3)">
            <div class="tool-check">&#10003;</div>
            <div class="tool-icon">&#128172;</div>
            <div class="tool-name">工具 &#9314; AI 學生回饋系統</div>
            <div class="tool-desc">Rubric + AI 回饋 + 寄信</div>
        </div>
    </div>
</div>

<div class="card">
    <h2>Step 2 &#8212; 自訂設定</h2>
    <p class="section-desc">不修改任何設定也可以直接生成，預設值即為整合版的原始設定</p>

    <!-- 全域設定 -->
    <div class="form-section">
        <h3>&#127760; 全域設定</h3>
        <div class="form-grid">
            <div class="form-group">
                <label class="form-label">AI 模型</label>
                <select class="form-select" id="cfg-model">
                    <option value="gemini-3-flash-preview" selected>gemini-3-flash-preview (推薦)</option>
                    <option value="gemini-3.1-pro-preview">gemini-3.1-pro-preview (最新)</option>
                    <option value="gemini-3.1-flash-lite-preview">gemini-3.1-flash-lite-preview (輕量)</option>
                    <option value="gemini-2.5-pro">gemini-2.5-pro (穩定)</option>
                </select>
                <div class="form-hint">選擇 Gemini API 使用的模型</div>
            </div>
        </div>
    </div>

    <!-- 各工具設定（摺疊面板）-->
    <div class="accordion">
        <!-- 工具① -->
        <div class="acc-item" id="acc-tool1">
            <div class="acc-header" onclick="toggleAccordion('acc-tool1')">
                <span>&#128202; 工具 &#9312; 段考 AI 審題助手</span>
                <span class="acc-arrow">&#9654;</span>
            </div>
            <div class="acc-body"><div class="acc-body-inner">
                <div class="form-group">
                    <label class="form-label">AI 溫度（越低越穩定）</label>
                    <div class="range-group">
                        <input type="range" id="cfg-t1-temp" min="0" max="1" step="0.1" value="0.2" oninput="updateRange(this)">
                        <span class="range-value">0.2</span>
                    </div>
                </div>
            </div></div>
        </div>

        <!-- 工具② -->
        <div class="acc-item" id="acc-tool2">
            <div class="acc-header" onclick="toggleAccordion('acc-tool2')">
                <span>&#127919; 工具 &#9313; 動態評量產生器</span>
                <span class="acc-arrow">&#9654;</span>
            </div>
            <div class="acc-body"><div class="acc-body-inner">
                <div class="form-group">
                    <label class="form-label">AI 溫度</label>
                    <div class="range-group">
                        <input type="range" id="cfg-t2-temp" min="0" max="1" step="0.1" value="0.3" oninput="updateRange(this)">
                        <span class="range-value">0.3</span>
                    </div>
                </div>
                <div class="form-group">
                    <label class="form-label">計分規則</label>
                    <div class="form-grid">
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">不用提示得分</label>
                            <input class="form-input" type="number" id="cfg-t2-s0" value="5" min="0" max="10">
                        </div>
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">1 個提示得分</label>
                            <input class="form-input" type="number" id="cfg-t2-s1" value="4" min="0" max="10">
                        </div>
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">2 個提示得分</label>
                            <input class="form-input" type="number" id="cfg-t2-s2" value="3" min="0" max="10">
                        </div>
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">3 個提示得分</label>
                            <input class="form-input" type="number" id="cfg-t2-s3" value="1" min="0" max="10">
                        </div>
                    </div>
                </div>
                <div class="form-group">
                    <label class="form-label">能力指標門檻（百分比）</label>
                    <div class="form-grid">
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">精熟 &ge;</label>
                            <input class="form-input" type="number" id="cfg-t2-m1" value="90" min="0" max="100">
                        </div>
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">接近精熟 &ge;</label>
                            <input class="form-input" type="number" id="cfg-t2-m2" value="70" min="0" max="100">
                        </div>
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">發展中 &ge;</label>
                            <input class="form-input" type="number" id="cfg-t2-m3" value="40" min="0" max="100">
                        </div>
                    </div>
                </div>
            </div></div>
        </div>

        <!-- 工具③ -->
        <div class="acc-item" id="acc-tool3">
            <div class="acc-header" onclick="toggleAccordion('acc-tool3')">
                <span>&#128172; 工具 &#9314; AI 學生回饋系統</span>
                <span class="acc-arrow">&#9654;</span>
            </div>
            <div class="acc-body"><div class="acc-body-inner">
                <div class="form-group">
                    <label class="form-label">AI 溫度</label>
                    <div class="range-group">
                        <input type="range" id="cfg-t3-temp" min="0" max="1" step="0.1" value="0.4" oninput="updateRange(this)">
                        <span class="range-value">0.4</span>
                    </div>
                </div>
                <div class="form-group">
                    <label class="form-label">等第門檻（百分比）</label>
                    <div class="form-grid">
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">優 &ge;</label>
                            <input class="form-input" type="number" id="cfg-t3-g1" value="85" min="0" max="100">
                        </div>
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">良 &ge;</label>
                            <input class="form-input" type="number" id="cfg-t3-g2" value="70" min="0" max="100">
                        </div>
                        <div class="form-group">
                            <label class="form-label" style="font-size:12px;">中 &ge;</label>
                            <input class="form-input" type="number" id="cfg-t3-g3" value="50" min="0" max="100">
                        </div>
                    </div>
                </div>
            </div></div>
        </div>
    </div>

    <div style="text-align: center; margin-top: 28px;">
        <button class="btn-primary" onclick="generateCode()">&#9881;&#65039; 生成程式碼</button>
    </div>
</div>

<!-- ==================== 程式碼預覽區 ==================== -->
<div class="card code-section" id="code-section">
    <h2>&#128196; 生成結果</h2>
    <div class="code-toolbar">
        <div class="code-info">
            <span class="code-badge" id="code-tools">3 個工具</span>
            <span class="code-badge" id="code-lines">0 行</span>
            <span class="code-badge" id="code-model">gemini-3-flash-preview</span>
        </div>
    </div>
    <div class="code-body" id="code-output"></div>
    <div class="btn-group">
        <button class="btn-action" onclick="copyCode()">&#128203; 複製到剪貼簿</button>
        <button class="btn-action" onclick="downloadCode()">&#128229; 下載 .js 檔</button>
        <button class="btn-action" onclick="resetAll()">&#128260; 重新設定</button>
    </div>
</div>

<!-- ==================== 安裝教學 ==================== -->
<div class="card">
    <h2>&#128218; 安裝教學（5 分鐘搞定）</h2>
    <p class="section-desc">跟著以下步驟，把程式碼安裝到你的 Google Sheets</p>

    <div class="steps-container">
        <div class="step-item">
            <div class="step-num">1</div>
            <div class="step-content">
                <h4>建立 Google Sheets</h4>
                <p>前往 <a href="https://sheets.google.com" target="_blank" style="color:#667eea;font-weight:600;">sheets.google.com</a>，建立一個新的空白試算表。</p>
                <div class="info-box">建議命名為「AI 教學工具」或你喜歡的名稱。</div>
            </div>
        </div>

        <div class="step-item">
            <div class="step-num">2</div>
            <div class="step-content">
                <h4>開啟 Apps Script 編輯器</h4>
                <p>在 Google Sheets 中，點選上方選單：<br>
                <code>擴充功能</code> &#8594; <code>Apps Script</code><br>
                系統會在新分頁開啟 GAS 編輯器。</p>
            </div>
        </div>

        <div class="step-item">
            <div class="step-num">3</div>
            <div class="step-content">
                <h4>貼上程式碼</h4>
                <p>在上方「程式碼生成器」產生程式碼後：</p>
                <p>1. 點「&#128203; 複製到剪貼簿」按鈕<br>
                2. 回到 GAS 編輯器，按 <code>Ctrl+A</code> 全選預設程式碼<br>
                3. 按 <code>Ctrl+V</code> 貼上<br>
                4. 按 <code>Ctrl+S</code> 儲存</p>
            </div>
        </div>

        <div class="step-item">
            <div class="step-num">4</div>
            <div class="step-content">
                <h4>取得 Gemini API 金鑰</h4>
                <p>前往 <a href="https://aistudio.google.com/apikey" target="_blank" style="color:#667eea;font-weight:600;">aistudio.google.com/apikey</a>，點「Create API Key」建立一組金鑰，複製下來。</p>
                <div class="info-box">API 金鑰免費，每分鐘有免費額度，足夠教學使用。</div>
            </div>
        </div>

        <div class="step-item">
            <div class="step-num">5</div>
            <div class="step-content">
                <h4>設定 API 金鑰</h4>
                <p>在 GAS 編輯器左側，點齒輪圖示 &#9881;&#65039;「專案設定」：</p>
                <p>1. 滾到最下方「指令碼屬性」<br>
                2. 點「新增指令碼屬性」<br>
                3. 屬性：<code>GEMINI_API_KEY</code><br>
                4. 值：貼上剛才複製的 API 金鑰<br>
                5. 點「儲存指令碼屬性」</p>
                <div class="warning-box"><strong>注意：</strong>屬性名稱必須是 <code>GEMINI_API_KEY</code>（全大寫），不可打錯。</div>
            </div>
        </div>

        <div class="step-item">
            <div class="step-num">6</div>
            <div class="step-content">
                <h4>首次授權</h4>
                <p>回到 GAS 編輯器的程式碼頁面：</p>
                <p>1. 上方函式下拉選 <code>onOpen</code> &#8594; 點「&#9654; 執行」<br>
                2. 出現授權視窗 &#8594; 點「審查權限」<br>
                3. 選擇你的 Google 帳號<br>
                4. 點「進階」&#8594;「前往 XXX（不安全）」&#8594;「允許」</p>
                <div class="warning-box"><strong>為什麼顯示「不安全」？</strong>因為這是你自己寫的腳本，Google 尚未審核，但程式碼完全透明安全。</div>
            </div>
        </div>

        <div class="step-item">
            <div class="step-num">7</div>
            <div class="step-content">
                <h4>開始使用！</h4>
                <p>回到 Google Sheets，<strong>重新整理頁面</strong>（F5 或 Ctrl+R），選單列就會出現 AI 工具的選單：</p>
                <p>&#128202; <strong>AI 審題助手</strong> — 貼上段考題 &#8594; 分析<br>
                &#127919; <strong>動態評量</strong> — 填入知識點 &#8594; 產生題目 &#8594; 建 Form &#8594; 計分<br>
                &#128172; <strong>AI 回饋</strong> — 填入 Rubric + 作品連結 &#8594; 產生回饋 &#8594; 寄信</p>
            </div>
        </div>
    </div>
</div>

</div><!-- end container -->

<!-- Footer -->
<div class="footer">
    <p>AI 工具協助差異化教學 &#8212; <a href="https://github.com/ChatGPT3a01/AI-Tools-for-Differentiated-Teaching" target="_blank">GitHub</a></p>
    <p style="margin-top:6px;">&#169; 2026 曾慶良（阿亮老師） &#8212; 高中優質化數位前導計畫</p>
</div>

<!-- Toast -->
<div class="toast" id="toast"></div>

<!-- ==================== GAS 程式碼模板 ==================== -->
''' + f'''<script type="text/template" id="tpl-shared">
{shared_code}</script>

<script type="text/template" id="tpl-menu1">
{tool1_menu}</script>

<script type="text/template" id="tpl-menu2">
{tool2_menu}</script>

<script type="text/template" id="tpl-menu3">
{tool3_menu}</script>

<script type="text/template" id="tpl-tool1">
{tool1_code}</script>

<script type="text/template" id="tpl-tool2">
{tool2_code}</script>

<script type="text/template" id="tpl-tool3">
{tool3_code}</script>
''' + r'''
<!-- ==================== Main JS ==================== -->
<script>
    // ===== State =====
    const state = { tool1: true, tool2: true, tool3: true, generatedCode: '' };

    // ===== Tool Toggle =====
    function toggleTool(n) {
        state['tool' + n] = !state['tool' + n];
        const card = document.querySelector('.tool-card[data-tool="' + n + '"]');
        card.classList.toggle('active', state['tool' + n]);
        const acc = document.getElementById('acc-tool' + n);
        acc.classList.toggle('disabled', !state['tool' + n]);
        if (!state['tool' + n]) acc.classList.remove('open');
    }

    // ===== Accordion =====
    function toggleAccordion(id) {
        const el = document.getElementById(id);
        if (el.classList.contains('disabled')) return;
        el.classList.toggle('open');
    }

    // ===== Range Slider =====
    function updateRange(el) {
        el.parentElement.querySelector('.range-value').textContent = el.value;
    }

    // ===== Code Generation =====
    function generateCode() {
        const model = document.getElementById('cfg-model').value;
        const t1Temp = document.getElementById('cfg-t1-temp').value;
        const t2Temp = document.getElementById('cfg-t2-temp').value;
        const t3Temp = document.getElementById('cfg-t3-temp').value;
        const s0 = document.getElementById('cfg-t2-s0').value;
        const s1 = document.getElementById('cfg-t2-s1').value;
        const s2 = document.getElementById('cfg-t2-s2').value;
        const s3 = document.getElementById('cfg-t2-s3').value;
        const m1 = document.getElementById('cfg-t2-m1').value;
        const m2 = document.getElementById('cfg-t2-m2').value;
        const m3 = document.getElementById('cfg-t2-m3').value;
        const g1 = document.getElementById('cfg-t3-g1').value;
        const g2 = document.getElementById('cfg-t3-g2').value;
        const g3 = document.getElementById('cfg-t3-g3').value;

        // Build shared section with conditional menus
        let shared = document.getElementById('tpl-shared').textContent;

        // Replace model
        shared = shared.replace(
            "const GEMINI_MODEL = 'gemini-3-flash-preview';",
            "const GEMINI_MODEL = '" + model + "';"
        );

        // Build onOpen with selected menus only
        let menuCode = '';
        if (state.tool1) menuCode += document.getElementById('tpl-menu1').textContent;
        if (state.tool2) menuCode += '\n' + document.getElementById('tpl-menu2').textContent;
        if (state.tool3) menuCode += '\n' + document.getElementById('tpl-menu3').textContent;

        // Replace the full menu block in shared code
        const menuStart = "  const ui = SpreadsheetApp.getUi();";
        const menuEnd = "}";
        // Rebuild onOpen function
        const onOpenReplacement = "function onOpen() {\n  const ui = SpreadsheetApp.getUi();\n\n"
            + menuCode + "\n}";
        shared = shared.replace(/function onOpen\(\) \{[\s\S]*?\n\}/,  onOpenReplacement);

        let code = shared;

        // Tool 1
        if (state.tool1) {
            let t1 = document.getElementById('tpl-tool1').textContent;
            t1 = t1.replace(
                'callGeminiText_(prompt, 0.2, 16384)',
                'callGeminiText_(prompt, ' + t1Temp + ', 16384)'
            );
            code += '\n\n' + t1;
        }

        // Tool 2
        if (state.tool2) {
            let t2 = document.getElementById('tpl-tool2').textContent;
            t2 = t2.replace(
                'callGeminiText_(prompt, 0.3, 2048)',
                'callGeminiText_(prompt, ' + t2Temp + ', 2048)'
            );
            // Score rules
            t2 = t2.replace('score = 5; hintCount = 0', 'score = ' + s0 + '; hintCount = 0');
            t2 = t2.replace('score = 4; hintCount = 1', 'score = ' + s1 + '; hintCount = 1');
            t2 = t2.replace('score = 3; hintCount = 2', 'score = ' + s2 + '; hintCount = 2');
            t2 = t2.replace('score = 1; hintCount = 3', 'score = ' + s3 + '; hintCount = 3');
            // Form description scoring text
            t2 = t2.replace(
                '計分規則：不用提示＝5分 ／ 1個提示＝4分 ／ 2個提示＝3分 ／ 3個提示＝1分',
                '計分規則：不用提示＝' + s0 + '分 ／ 1個提示＝' + s1 + '分 ／ 2個提示＝' + s2 + '分 ／ 3個提示＝' + s3 + '分'
            );
            // Mastery thresholds (appears twice in tool2)
            t2 = replaceAll(t2, 'percentage >= 90) indicator = ', 'percentage >= ' + m1 + ') indicator = ');
            t2 = replaceAll(t2, 'percentage >= 70) indicator = ', 'percentage >= ' + m2 + ') indicator = ');
            t2 = replaceAll(t2, 'percentage >= 40) indicator = ', 'percentage >= ' + m3 + ') indicator = ');
            t2 = replaceAll(t2, 'avgPercentage >= 90) avgIndicator = ', 'avgPercentage >= ' + m1 + ') avgIndicator = ');
            t2 = replaceAll(t2, 'avgPercentage >= 70) avgIndicator = ', 'avgPercentage >= ' + m2 + ') avgIndicator = ');
            t2 = replaceAll(t2, 'avgPercentage >= 40) avgIndicator = ', 'avgPercentage >= ' + m3 + ') avgIndicator = ');
            // Max score per question
            t2 = replaceAll(t2, 'numQuestions * 5', 'numQuestions * ' + s0);
            code += '\n\n' + t2;
        }

        // Tool 3
        if (state.tool3) {
            let t3 = document.getElementById('tpl-tool3').textContent;
            t3 = t3.replace(
                'temperature: 0.4, maxOutputTokens: 4096',
                'temperature: ' + t3Temp + ', maxOutputTokens: 4096'
            );
            // Grade thresholds (in prompt text)
            t3 = t3.replace(
                "percentage >= 85 → \"優\"，>= 70 → \"良\"，>= 50 → \"中\"，< 50 → \"再加強\"",
                "percentage >= " + g1 + " → \"優\"，>= " + g2 + " → \"良\"，>= " + g3 + " → \"中\"，< " + g3 + " → \"再加強\""
            );
            // Grade thresholds in parseScores (and report)
            t3 = replaceAll(t3, 'percentage >= 85) grade = ', 'percentage >= ' + g1 + ') grade = ');
            t3 = replaceAll(t3, 'percentage >= 70) grade = ', 'percentage >= ' + g2 + ') grade = ');
            t3 = replaceAll(t3, 'percentage >= 50) grade = ', 'percentage >= ' + g3 + ') grade = ');
            code += '\n\n' + t3;
        }

        state.generatedCode = code.trim();
        displayCode(state.generatedCode);
    }

    function replaceAll(str, search, replacement) {
        return str.split(search).join(replacement);
    }

    // ===== Display Code =====
    function displayCode(code) {
        const lines = code.split('\n');
        let html = '';
        for (let i = 0; i < lines.length; i++) {
            const highlighted = highlightLine(escapeHtml(lines[i]));
            html += '<div><span class="line-num">' + (i + 1) + '</span>' + highlighted + '</div>';
        }
        document.getElementById('code-output').innerHTML = html;

        // Update badges
        const toolCount = (state.tool1 ? 1 : 0) + (state.tool2 ? 1 : 0) + (state.tool3 ? 1 : 0);
        document.getElementById('code-tools').textContent = toolCount + ' 個工具';
        document.getElementById('code-lines').textContent = lines.length + ' 行';
        document.getElementById('code-model').textContent = document.getElementById('cfg-model').value;

        // Show section
        document.getElementById('code-section').classList.add('visible');
        document.getElementById('code-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function escapeHtml(str) {
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function highlightLine(line) {
        // Comments
        if (line.trimStart().startsWith('//')) {
            return '<span class="hl-cmt">' + line + '</span>';
        }
        // Apply highlighting
        return line
            .replace(/\b(const|var|let|function|return|if|else|for|try|catch|new|this|null|true|false|typeof|instanceof|throw|break|continue)\b/g, '<span class="hl-kw">$1</span>')
            .replace(/('(?:[^'\\]|\\.)*')/g, '<span class="hl-str">$1</span>')
            .replace(/("(?:[^"\\]|\\.)*")/g, '<span class="hl-str">$1</span>')
            .replace(/\b(\d+\.?\d*)\b/g, '<span class="hl-num">$1</span>');
    }

    // ===== Copy & Download =====
    function copyCode() {
        if (!state.generatedCode) return;
        navigator.clipboard.writeText(state.generatedCode).then(function() {
            showToast('&#10003; 已複製到剪貼簿！');
        }).catch(function() {
            // Fallback
            const ta = document.createElement('textarea');
            ta.value = state.generatedCode;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showToast('&#10003; 已複製到剪貼簿！');
        });
    }

    function downloadCode() {
        if (!state.generatedCode) return;
        const names = [];
        if (state.tool1) names.push('審題');
        if (state.tool2) names.push('評量');
        if (state.tool3) names.push('回饋');
        const filename = names.length === 3
            ? '整合版_三工具_GAS.js'
            : '自訂版_' + names.join('+') + '_GAS.js';

        const blob = new Blob([state.generatedCode], { type: 'text/javascript;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
        showToast('&#128229; 已下載 ' + filename);
    }

    function resetAll() {
        document.getElementById('code-section').classList.remove('visible');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // ===== Toast =====
    function showToast(msg) {
        const t = document.getElementById('toast');
        t.innerHTML = msg;
        t.classList.add('show');
        setTimeout(function() { t.classList.remove('show'); }, 2500);
    }
</script>
</body>
</html>'''

# Write generator.html
out_path = os.path.join('簡報', 'generator.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Done! generator.html created ({len(html)} bytes)")
print(f"GAS sections: shared={len(shared_code)}, tool1={len(tool1_code)}, tool2={len(tool2_code)}, tool3={len(tool3_code)}")
