<div align="center">

# 🤖 AI 工具協助差異化教學

### 以 SDT 動機理論為架構，搭配 3 個 AI 教學工具

[![適用對象](https://img.shields.io/badge/適用對象-高中教師-blue?style=for-the-badge)](.)
[![核心理論](https://img.shields.io/badge/核心理論-SDT_自我決定理論-purple?style=for-the-badge)](.)
[![工具數量](https://img.shields.io/badge/工具-3_個_GAS_工具-green?style=for-the-badge)](.)

以**自我決定理論（SDT）**的三大心理需求——**勝任感、自主性、關係**——為架構，<br>
搭配 Google Apps Script + Gemini API，實現可落地的差異化教學。

**所有工具都是「複製貼上就能用」，不需要寫程式基礎。**

<br>

### 👉 [📺 點我開啟研習簡報（線上版）](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/index.html) 👈

<br>

[![開啟簡報](https://img.shields.io/badge/🎬_開啟研習簡報-點我開始-ff6b35?style=for-the-badge&labelColor=1e3c72)](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/index.html)
[![下載教材](https://img.shields.io/badge/📄_教材_PDF-下載-27ae60?style=for-the-badge&labelColor=0d6938)](教材/研習教材_AI工具協助差異化教學.pdf)

</div>

---

## 📺 研習簡報（線上瀏覽）

<div align="center">

> 💡 點擊下方連結，直接在瀏覽器中開啟簡報，不需要下載！

| Part | 主題 | 線上開啟 |
|:---:|------|:---:|
| 🏠 | **統整頁（從這裡開始）** | [**▶ 開啟簡報首頁**](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/index.html) |
| 1 | 動機理論基礎 + 段考 AI 審題助手 | [開啟 Part 1](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/Part1_動機理論與段考AI審題助手.html) |
| 2 | 教材差異化：AI 改廠商簡報 | [開啟 Part 2](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/Part2_教材差異化AI改簡報.html) |
| 3 | 動態評量產生器 | [開啟 Part 3](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/Part3_動態評量產生器.html) |
| 4 | 多元評量 Rubric + AI 學生回饋系統 | [開啟 Part 4](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/Part4_多元評量與AI學生回饋系統.html) |
| 5 | 分享與總結 | [開啟 Part 5](https://chatgpt3a01.github.io/AI-Tools-for-Differentiated-Teaching/簡報/Part5_分享與總結.html) |

</div>

---

## 💻 程式碼下載

<div align="center">

| 工具 | 下載 |
|------|:---:|
| ① 段考 AI 審題助手 | [📥 下載 .js](GAS程式碼/工具1_段考AI審題助手.js) |
| ② 動態評量產生器 | [📥 下載 .js](GAS程式碼/工具2_動態評量產生器.js) |
| ③ AI 學生回饋系統 | [📥 下載 .js](GAS程式碼/工具3_AI學生回饋系統.js) |

</div>

---

## 📋 三個工具一覽

<div align="center">

| | 工具名稱 | 對應需求 | 功能說明 |
|:---:|------|:---:|------|
| ① | **段考 AI 審題助手** | 💪 勝任感 | 貼上段考題 → AI 自動產出雙向細目表、難易度分析、改進建議 |
| ② | **動態評量產生器** | 💪 勝任感 | 輸入知識點 → AI 產出題目 + 三層提示 → 自動建立 Google Form |
| ③ | **AI 學生回饋系統** | 💬 關係 | 貼上學生作品 + Rubric → AI 產出逐項評分 + 個人化回饋 |

</div>

---

## 🚀 快速開始

### 前置準備（一次就好）

1. **取得 Gemini API Key**
   - 前往 [Google AI Studio](https://aistudio.google.com/)
   - 點「Get API Key」→ 建立金鑰 → 複製備用

### 安裝任一工具（5 分鐘）

1. 新增 Google Sheets，命名為工具名稱
2. 點「擴充功能」→「Apps Script」
3. 左邊齒輪「專案設定」→「指令碼屬性」→ 新增 `GEMINI_API_KEY`
4. 回到編輯器 → 全選刪除 → 貼上對應的 `.js` 程式碼 → Ctrl+S
5. 上方「執行 ▶」→ 函式選 `onOpen` → 點「執行」→ 完成授權
6. 回到 Sheets → 按 F5 重新整理 → 上方出現工具選單

> 每個 `.js` 檔案最上面都有完整的安裝步驟說明，可以直接照著做。

---

## 📁 檔案結構

```
output/
├── 📁 簡報/                              ← HTML 研習簡報
│   ├── index.html                        （統整頁，點這裡開始）
│   ├── Part1_動機理論與段考AI審題助手.html
│   ├── Part2_教材差異化AI改簡報.html
│   ├── Part3_動態評量產生器.html
│   ├── Part4_多元評量與AI學生回饋系統.html
│   ├── Part5_分享與總結.html
│   └── 📁 img/                           （簡報用截圖）
│
├── 📁 教材/
│   ├── 研習教材_AI工具協助差異化教學.md    ← 完整教材 Markdown
│   └── 研習教材_AI工具協助差異化教學.pdf   ← 完整教材 PDF
│
├── 📁 GAS程式碼/                          ← 三個工具的程式碼
│   ├── README.md                          （安裝說明總整理）
│   ├── 工具1_段考AI審題助手.js
│   ├── 工具2_動態評量產生器.js
│   └── 工具3_AI學生回饋系統.js
│
├── 📁 指引/
│   └── NotebookLM設定指引.md              （工具③ 搭配使用）
│
└── 📁 截圖/                               ← 操作步驟截圖
    ├── 1-1_新建Google_Sheets.png ~ 1-12
    ├── 2-1_Gemini改寫對比.png ~ 2-4
    ├── 3-1_知識點輸入.png ~ 3-5
    └── 4-1_三種Rubric.png ~ 4-6
```

---

## 🔧 技術棧

| 技術 | 用途 |
|------|------|
| Google Apps Script (GAS) | 三個工具的後端邏輯 |
| Gemini API (gemini-2.5-flash) | AI 分析、產題、回饋 |
| Google Sheets | 工具的操作介面 |
| Google Forms | 動態評量的學生作答介面 |
| NotebookLM | 課程知識庫（選配） |

---

## 👨‍🏫 關於作者

<div align="center">

### 曾慶良 主任（阿亮老師）

<table>
<tr>
<td width="50%">

**📌 現任職務**

🎓 新興科技推廣中心主任<br>
🎓 教育部學科中心研究教師

</td>
<td width="50%">

**🏆 獲獎紀錄**

🥇 2025年 SETEAM教學專業講師認證<br>
🥇 2024年 教育部人工智慧講師認證<br>
🥇 2022、2023年 指導學生XR專題競賽特優<br>
🥇 2022年 VR教材開發教師組特優<br>
🥇 2019年 百大資訊人才獎<br>
🥇 2018、2019年 親子天下創新100教師<br>
🥇 2018年 臺北市特殊優良教師<br>
🥇 2017年 教育部行動學習優等

</td>
</tr>
</table>

<br>

### 📞 聯絡方式

[![YouTube](https://img.shields.io/badge/YouTube-@Liang--yt02-red?style=for-the-badge&logo=youtube)](https://www.youtube.com/@Liang-yt02)
[![Facebook](https://img.shields.io/badge/Facebook-3A科技研究社-blue?style=for-the-badge&logo=facebook)](https://www.facebook.com/groups/2754139931432955)
[![Email](https://img.shields.io/badge/Email-3a01chatgpt@gmail.com-green?style=for-the-badge&logo=gmail)](mailto:3a01chatgpt@gmail.com)

</div>

---

## 📜 授權聲明

**© 2026 阿亮老師 版權所有**

本專案僅供「阿亮老師課程學員」學習使用。

### ⚠️ 禁止事項

- ❌ 禁止修改本專案內容
- ❌ 禁止轉傳或散布
- ❌ 禁止商業使用
- ❌ 禁止未經授權之任何形式使用

如有任何授權需求，請聯繫作者。

---

<div align="center">

**Made with ❤️ by 阿亮老師**

[⬆️ 回到頂部](#-ai-工具協助差異化教學)

</div>
