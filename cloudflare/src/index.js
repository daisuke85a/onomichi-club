/**
 * おのみち部活さがし — Cloudflare Worker
 *
 *   GET /            … public/index.html（静的アセットとして自動配信）
 *   GET /api/clubs   … スプレッドシートの一覧を JSON で返す（CACHE_SECONDS 秒キャッシュ）
 *   GET /api/clubs?refresh=1 … キャッシュを無視して取り直す（更新直後の確認用）
 */

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (url.pathname === "/api/clubs" && request.method === "GET") {
      return handleClubs(url, env, ctx);
    }
    return Response.json({ error: "not found" }, { status: 404 });
  },
};

async function handleClubs(url, env, ctx) {
  const cacheSeconds = Number(env.CACHE_SECONDS) || 300;
  const cache = caches.default;
  const cacheKey = new Request(`${url.origin}/api/clubs`);
  const refresh = url.searchParams.get("refresh") === "1";

  if (!refresh) {
    const hit = await cache.match(cacheKey);
    if (hit) return hit;
  }

  let rows;
  try {
    rows = await fetchRows(env.DATA_URL);
  } catch (err) {
    console.error(JSON.stringify({ event: "fetch_rows_failed", message: String(err && err.message || err) }));
    // 取得に失敗したら、古いキャッシュがあればそれを返す
    const stale = await cache.match(cacheKey);
    if (stale) return stale;
    return Response.json({ error: "データを取得できませんでした" }, { status: 502 });
  }

  const body = { updatedAt: new Date().toISOString(), count: rows.length, rows };
  const response = Response.json(body, {
    headers: {
      "cache-control": `public, max-age=60, s-maxage=${cacheSeconds}`,
      "access-control-allow-origin": "*",
    },
  });
  ctx.waitUntil(cache.put(cacheKey, response.clone()));
  return response;
}

/** DATA_URL から一覧を取得。JSON でも CSV でも受け付け、「公開」が FALSE の行は除く */
async function fetchRows(dataUrl) {
  if (!dataUrl) throw new Error("DATA_URL が設定されていません");
  const res = await fetch(dataUrl, { redirect: "follow", headers: { accept: "application/json, text/csv;q=0.9, */*;q=0.1" } });
  if (!res.ok) throw new Error(`upstream ${res.status}`);
  const text = await res.text(); // 一覧は数十〜数百行なので全文読み込みで問題ない

  let rows;
  const head = text.trimStart();
  if (head.startsWith("[") || head.startsWith("{")) {
    const json = JSON.parse(head);
    rows = Array.isArray(json) ? json : json.rows;
  } else if (head.startsWith("<")) {
    throw new Error("upstream returned HTML (ログインが必要か、URL が違う可能性があります)");
  } else {
    rows = parseCSV(text);
  }
  return rows.filter((r) => String(r["公開"] ?? "").toUpperCase() !== "FALSE");
}

function parseCSV(text) {
  const out = [];
  let row = [], cell = "", inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else inQ = false; }
      else cell += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === ",") { row.push(cell); cell = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cell); out.push(row); row = []; cell = "";
    } else cell += ch;
  }
  if (cell !== "" || row.length) { row.push(cell); out.push(row); }
  const headers = (out.shift() || []).map((h) => h.trim());
  return out
    .filter((r) => r.some((c) => c.trim() !== ""))
    .map((r) => Object.fromEntries(headers.map((h, i) => [h, (r[i] || "").trim()])));
}
