// AI 差異化教學工具 · 全域設定
// -----------------------------------------------------------
// 部署 Cloudflare Worker 後，請把下面這行的 URL 改成你的 Worker 網址
// 例如：https://ai-diff-teaching-proxy.your-subdomain.workers.dev
//
// 取得 URL 的方法：
//   wrangler deploy 完成後，終端機會顯示「Published ... at <URL>」
// -----------------------------------------------------------

window.WORKER_BASE = "https://ai-diff-teaching-proxy.aliangschool2026.workers.dev";

// 共用：呼叫 Worker 的 helper
window.callAPI = async function(endpoint, input) {
  const url = window.WORKER_BASE.replace(/\/$/, "") + endpoint;
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ input })
  });
  const data = await resp.json();
  if (!resp.ok) {
    throw new Error(data.error || `HTTP ${resp.status}`);
  }
  return data.data;
};
