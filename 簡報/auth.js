/**
 * 簡報密碼保護系統
 * - 預設密碼：aitool2026
 * - 無敵密碼：可進入管理面板更改密碼
 * - 同一 session 只需輸入一次
 */
(function () {
  const MASTER_KEY = 'Aa@0981737608';
  const DEFAULT_PWD = 'aitool2026';
  const STORAGE_KEY = 'slide_pwd';
  const SESSION_KEY = 'slide_auth_ok';

  // 已驗證過就跳過
  if (sessionStorage.getItem(SESSION_KEY) === '1') return;

  // 取得目前密碼（可能被管理員改過）
  function getCurrentPwd() {
    return localStorage.getItem(STORAGE_KEY) || DEFAULT_PWD;
  }

  // 隱藏頁面內容
  document.documentElement.style.visibility = 'hidden';
  document.documentElement.style.overflow = 'hidden';

  window.addEventListener('DOMContentLoaded', function () {
    document.body.style.visibility = 'hidden';

    // 建立遮罩
    var overlay = document.createElement('div');
    overlay.id = 'auth-overlay';
    overlay.innerHTML = `
      <style>
        #auth-overlay {
          position: fixed; inset: 0; z-index: 999999;
          background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
          display: flex; align-items: center; justify-content: center;
          font-family: 'Microsoft JhengHei', 'Segoe UI', Arial, sans-serif;
          visibility: visible !important;
        }
        #auth-overlay * { visibility: visible !important; }
        .auth-box {
          background: rgba(255,255,255,0.07);
          backdrop-filter: blur(20px);
          border: 1px solid rgba(255,255,255,0.15);
          border-radius: 24px;
          padding: 48px 44px;
          width: 420px; max-width: 90vw;
          text-align: center;
          box-shadow: 0 25px 60px rgba(0,0,0,0.5);
          animation: authFadeIn 0.5s ease;
        }
        @keyframes authFadeIn {
          from { opacity: 0; transform: translateY(30px) scale(0.95); }
          to   { opacity: 1; transform: translateY(0) scale(1); }
        }
        .auth-icon { font-size: 56px; margin-bottom: 16px; }
        .auth-title {
          color: #fff; font-size: 26px; font-weight: 700;
          margin-bottom: 8px;
        }
        .auth-subtitle {
          color: rgba(255,255,255,0.6); font-size: 14px;
          margin-bottom: 32px;
        }
        .auth-input-wrap {
          position: relative; margin-bottom: 20px;
        }
        .auth-input {
          width: 100%; padding: 14px 48px 14px 20px;
          border: 2px solid rgba(255,255,255,0.2);
          border-radius: 14px; background: rgba(255,255,255,0.08);
          color: #fff; font-size: 16px; outline: none;
          transition: border-color 0.3s;
        }
        .auth-input::placeholder { color: rgba(255,255,255,0.35); }
        .auth-input:focus { border-color: #667eea; }
        .auth-input.error {
          border-color: #e74c3c;
          animation: authShake 0.4s ease;
        }
        @keyframes authShake {
          0%,100% { transform: translateX(0); }
          20%     { transform: translateX(-8px); }
          40%     { transform: translateX(8px); }
          60%     { transform: translateX(-6px); }
          80%     { transform: translateX(6px); }
        }
        .auth-toggle-pwd {
          position: absolute; right: 14px; top: 50%; transform: translateY(-50%);
          background: none; border: none; color: rgba(255,255,255,0.5);
          cursor: pointer; font-size: 18px; padding: 4px;
        }
        .auth-toggle-pwd:hover { color: rgba(255,255,255,0.8); }
        .auth-btn {
          width: 100%; padding: 14px;
          background: linear-gradient(135deg, #667eea, #764ba2);
          color: #fff; border: none; border-radius: 14px;
          font-size: 17px; font-weight: 600; cursor: pointer;
          transition: transform 0.2s, box-shadow 0.2s;
        }
        .auth-btn:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 25px rgba(102,126,234,0.4);
        }
        .auth-btn:active { transform: translateY(0); }
        .auth-error-msg {
          color: #e74c3c; font-size: 14px; margin-top: 14px;
          min-height: 20px;
        }
        /* 管理面板 */
        .admin-panel { display: none; }
        .admin-panel.show { display: block; }
        .admin-panel .auth-title { font-size: 22px; color: #f59e0b; }
        .admin-label {
          color: rgba(255,255,255,0.7); font-size: 13px;
          text-align: left; display: block; margin-bottom: 6px; margin-top: 14px;
        }
        .admin-btn-row {
          display: flex; gap: 12px; margin-top: 20px;
        }
        .admin-btn-row .auth-btn { flex: 1; }
        .admin-btn-secondary {
          background: rgba(255,255,255,0.1) !important;
          border: 1px solid rgba(255,255,255,0.2) !important;
        }
        .admin-btn-secondary:hover {
          background: rgba(255,255,255,0.18) !important;
          box-shadow: none !important;
        }
        .admin-success {
          color: #2ecc71; font-size: 14px; margin-top: 14px;
          min-height: 20px;
        }
      </style>

      <!-- 登入面板 -->
      <div class="auth-box" id="login-panel">
        <div class="auth-icon">🔒</div>
        <div class="auth-title">請輸入研習密碼</div>
        <div class="auth-subtitle">AI 工具協助差異化教學｜研習簡報</div>
        <div class="auth-input-wrap">
          <input type="password" class="auth-input" id="auth-pwd"
                 placeholder="請輸入密碼" autocomplete="off">
          <button class="auth-toggle-pwd" id="toggle-pwd" title="顯示/隱藏密碼">👁</button>
        </div>
        <button class="auth-btn" id="auth-submit">進入簡報</button>
        <div class="auth-error-msg" id="auth-error"></div>
      </div>

      <!-- 管理面板 -->
      <div class="auth-box admin-panel" id="admin-panel">
        <div class="auth-icon">⚙️</div>
        <div class="auth-title">管理員面板</div>
        <div class="auth-subtitle">使用無敵密碼登入｜可更改研習密碼</div>
        <label class="admin-label">目前密碼</label>
        <div class="auth-input-wrap">
          <input type="text" class="auth-input" id="admin-current" readonly
                 style="opacity:0.6; cursor:default;">
        </div>
        <label class="admin-label">設定新密碼</label>
        <div class="auth-input-wrap">
          <input type="text" class="auth-input" id="admin-new"
                 placeholder="輸入新密碼（至少 4 個字元）" autocomplete="off">
        </div>
        <div class="admin-btn-row">
          <button class="auth-btn admin-btn-secondary" id="admin-enter">直接進入簡報</button>
          <button class="auth-btn" id="admin-save">儲存新密碼</button>
        </div>
        <div class="admin-success" id="admin-msg"></div>
      </div>
    `;
    document.body.appendChild(overlay);

    // DOM refs
    var pwdInput   = document.getElementById('auth-pwd');
    var toggleBtn  = document.getElementById('toggle-pwd');
    var submitBtn  = document.getElementById('auth-submit');
    var errorEl    = document.getElementById('auth-error');
    var loginPanel = document.getElementById('login-panel');
    var adminPanel = document.getElementById('admin-panel');
    var adminCur   = document.getElementById('admin-current');
    var adminNew   = document.getElementById('admin-new');
    var adminEnter = document.getElementById('admin-enter');
    var adminSave  = document.getElementById('admin-save');
    var adminMsg   = document.getElementById('admin-msg');

    pwdInput.focus();

    // 顯示/隱藏密碼
    toggleBtn.addEventListener('click', function () {
      var isPassword = pwdInput.type === 'password';
      pwdInput.type = isPassword ? 'text' : 'password';
      toggleBtn.textContent = isPassword ? '🙈' : '👁';
    });

    // 驗證
    function tryAuth() {
      var val = pwdInput.value.trim();
      if (!val) {
        errorEl.textContent = '請輸入密碼';
        pwdInput.classList.add('error');
        setTimeout(function () { pwdInput.classList.remove('error'); }, 400);
        return;
      }
      // 無敵密碼 → 管理面板
      if (val === MASTER_KEY) {
        loginPanel.style.display = 'none';
        adminPanel.classList.add('show');
        adminCur.value = getCurrentPwd();
        adminNew.focus();
        return;
      }
      // 一般密碼
      if (val === getCurrentPwd()) {
        unlock();
      } else {
        errorEl.textContent = '密碼錯誤，請重新輸入';
        pwdInput.classList.add('error');
        pwdInput.value = '';
        setTimeout(function () { pwdInput.classList.remove('error'); }, 400);
        pwdInput.focus();
      }
    }

    submitBtn.addEventListener('click', tryAuth);
    pwdInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') tryAuth();
    });

    // 管理面板 — 直接進入
    adminEnter.addEventListener('click', unlock);

    // 管理面板 — 儲存新密碼
    adminSave.addEventListener('click', function () {
      var newPwd = adminNew.value.trim();
      if (newPwd.length < 4) {
        adminMsg.style.color = '#e74c3c';
        adminMsg.textContent = '密碼至少需要 4 個字元';
        return;
      }
      localStorage.setItem(STORAGE_KEY, newPwd);
      adminCur.value = newPwd;
      adminNew.value = '';
      adminMsg.style.color = '#2ecc71';
      adminMsg.textContent = '✅ 密碼已更新為：' + newPwd;
    });

    // 解鎖
    function unlock() {
      sessionStorage.setItem(SESSION_KEY, '1');
      overlay.style.transition = 'opacity 0.4s ease';
      overlay.style.opacity = '0';
      setTimeout(function () {
        overlay.remove();
        document.documentElement.style.visibility = '';
        document.documentElement.style.overflow = '';
        document.body.style.visibility = '';
      }, 400);
    }
  });
})();
