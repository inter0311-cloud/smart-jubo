/* Dasibon bulletin worker, 2026-11-03 (fast-open, Safari 12 compatible). Keep this URL stable. */
const ROOT = new URL("./", self.registration.scope);
const HOME = new URL("index.html", ROOT).href;
const CACHE = "dasibon-bulletin-recovery-v1-" + encodeURIComponent(ROOT.pathname);
const OWNED = /^(dasibon-bulletin-shell-|dasibon-pwa-engine-)/;
const ASSETS = ["manifest.json", "logo.jpg", "icon-192.png", "icon-512.png"].map((p) => new URL(p, ROOT).href);
const EXTERNAL = /* @__PURE__ */ new Set(["fonts.googleapis.com", "fonts.gstatic.com", "cdnjs.cloudflare.com"]);
const BUILD = "20261103-fastopen-5";
const TIMEOUT = 5e3;
const FETCH_LIMIT = 9e4;
const RECENT = 6e4;
let lastFresh = 0, lastVersion = "", cachedVersion = null;
function isHome(url) {
  return url.origin === ROOT.origin && [ROOT.pathname, new URL(HOME).pathname].includes(url.pathname);
}
function version(text) {
  const release = /data-bulletin-release=["']([^"']+)/.exec(text);
  const build = /data-pwa-build=["']([^"']+)/.exec(text);
  return (release ? release[1] : "") + ":" + (build ? build[1] : "legacy");
}
function navigationResponse(response) {
  const headers = new Headers(response.headers);
  headers.delete("content-encoding");
  headers.delete("content-length");
  headers.set("Cache-Control", "no-store");
  return new Response(response.body, { status: response.status, statusText: response.statusText, headers });
}
function validText(text) {
  return /id=["']dasibonApp["']/.test(text) && /id=["']panel-home["']/.test(text) && text.includes("/*STORE_END*/") && /<\/html>/i.test(text);
}
async function valid(response) {
  if (!response || !response.ok || !/text\/html/i.test(response.headers.get("content-type") || "")) return false;
  return validText(await response.clone().text());
}
async function read(cacheName, key) {
  try {
    return await (await caches.open(cacheName)).match(key);
  } catch (_) {
    return void 0;
  }
}
async function save(key, response) {
  try {
    await (await caches.open(CACHE)).put(key, response.clone());
    return true;
  } catch (_) {
    return false;
  }
}
async function savedCopy() {
  const current = await read(CACHE, HOME);
  if (!current) return null;
  const text = await current.clone().text();
  if (!validText(text)) return null;
  cachedVersion = version(text.slice(0, 4096));
  return current;
}
async function fallback() {
  const current = await savedCopy();
  if (current) return navigationResponse(current);
  let names = [];
  try {
    names = await caches.keys();
  } catch (_) {
  }
  for (const name of names.filter((n) => OWNED.test(n)).reverse()) {
    for (const key of [HOME, ROOT.href]) {
      const old = await read(name, key);
      if (await valid(old)) {
        const safe = navigationResponse(old);
        await save(HOME, safe);
        cachedVersion = null;
        return safe;
      }
    }
  }
  return null;
}
async function cleanup() {
  if (!await savedCopy()) return;
  try {
    for (const name of (await caches.keys()).filter((n) => OWNED.test(n))) {
      const cache = await caches.open(name), keys = await cache.keys();
      if (keys.every((k) => {
        const u = new URL(k.url);
        return EXTERNAL.has(u.hostname) || u.origin === ROOT.origin && u.pathname.startsWith(ROOT.pathname);
      })) await caches.delete(name);
    }
  } catch (_) {
  }
}
async function inform(message, clientId) {
  try {
    const clients = clientId ? [await self.clients.get(clientId)] : await self.clients.matchAll({ type: "window" });
    clients.filter(Boolean).filter((c) => isHome(new URL(c.url))).forEach((c) => c.postMessage(message));
  } catch (_) {
  }
}
async function fresh() {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), FETCH_LIMIT);
  try {
    if (cachedVersion === null) await savedCopy();
    const url = new URL(HOME);
    url.searchParams.set("_bulletin_fetch", Date.now().toString());
    const response = await fetch(url.href, { cache: "no-store", signal: controller.signal, credentials: "same-origin" });
    if (response.redirected && !isHome(new URL(response.url))) throw new Error("Unexpected redirect");
    if (!response.ok || !/text\/html/i.test(response.headers.get("content-type") || "") || !response.body) throw new Error("Invalid bulletin response");
    const reader = response.body.getReader(), decoder = new TextDecoder(), chunks = [];
    let head = "", size = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      size += value.length;
      if (head.length < 4096) head += decoder.decode(value, { stream: true });
      if (head.length >= 4096 || /<head[\s>]/i.test(head)) break;
    }
    const latest = version(head);
    if (cachedVersion && latest === cachedVersion) {
      reader.cancel().catch(() => {
      });
      controller.abort();
      lastFresh = Date.now();
      lastVersion = latest;
      await inform({ type: "BULLETIN_FRESH", version: latest });
      return null;
    }
    await inform({ type: "BULLETIN_DOWNLOADING" });
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const full = new Response(new Blob(chunks), { status: response.status, statusText: response.statusText, headers: response.headers });
    if (!await valid(full)) throw new Error("Invalid or incomplete bulletin");
    const safe = navigationResponse(full);
    await save(HOME, safe);
    cachedVersion = latest;
    lastFresh = Date.now();
    lastVersion = latest;
    await cleanup();
    await inform({ type: "BULLETIN_FRESH", version: latest });
    return safe;
  } finally {
    clearTimeout(timer);
  }
}
let pending;
function refresh() {
  if (!pending) pending = fresh().finally(() => {
    pending = null;
  });
  return pending;
}
function bounded(promise, ms = TIMEOUT) {
  let timer;
  return Promise.race([promise, new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error("Timed out")), ms);
  })]).finally(() => clearTimeout(timer));
}
function guidance() {
  return new Response("<!doctype html><html lang=\"ko\"><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>\uB2E4\uC2DC\uBCF8\uAD50\uD68C \uC8FC\uBCF4 \uC5F0\uACB0 \uC548\uB0B4</title><style>body{font:18px/1.7 sans-serif;background:#FAF7F2;color:#1E1F4B;margin:10vh auto;padding:24px;max-width:520px}a{display:inline-block;padding:12px 20px;background:#1E1F4B;color:white;border-radius:10px}</style><h1>\uC8FC\uBCF4\uB97C \uBD88\uB7EC\uC624\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4</h1><p>\uC778\uD130\uB137 \uC5F0\uACB0\uC744 \uD655\uC778\uD574 \uC8FC\uC138\uC694. \uC544\uC9C1 \uC774 \uAE30\uAE30\uC5D0 \uC800\uC7A5\uB41C \uC8FC\uBCF4\uAC00 \uC5C6\uAC70\uB098 \uC800\uC7A5 \uACF5\uAC04\uC744 \uC0AC\uC6A9\uD560 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.</p><p>\uC5F0\uACB0\uC774 \uB3CC\uC544\uC624\uBA74 \uB2E4\uC2DC \uBD88\uB7EC\uC635\uB2C8\uB2E4.</p><a href=\"./index.html\">\uB2E4\uC2DC \uC2DC\uB3C4</a><script>addEventListener('online',()=>location.reload());if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js',{scope:'./',updateViaCache:'none'}).then(r=>r.update()).catch(()=>{});<\/script></html>", { status: 503, headers: { "Content-Type": "text/html;charset=utf-8", "Cache-Control": "no-store" } });
}
self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    try {
      await bounded(refresh(), 15e3);
    } catch (_) {
      await fallback();
    }
    await self.skipWaiting();
  })());
});
self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    await fallback();
    await cleanup();
    if (self.registration.navigationPreload) await self.registration.navigationPreload.disable().catch(() => {
    });
    await self.clients.claim();
  })());
});
self.addEventListener("message", (event) => {
  var _a, _b, _c, _d, _e;
  if (((_a = event.data) == null ? void 0 : _a.type) === "GET_BULLETIN_BUILD") (_c = (_b = event.ports) == null ? void 0 : _b[0]) == null ? void 0 : _c.postMessage({ build: BUILD });
  if (((_d = event.data) == null ? void 0 : _d.type) === "SKIP_WAITING") event.waitUntil(self.skipWaiting());
  if (((_e = event.data) == null ? void 0 : _e.type) === "CHECK_BULLETIN") event.waitUntil((async () => {
    var _a2, _b2;
    try {
      if (pending) await pending;
      else if (Date.now() - lastFresh < RECENT) await inform({ type: "BULLETIN_FRESH", version: lastVersion }, (_a2 = event.source) == null ? void 0 : _a2.id);
      else await refresh();
    } catch (_) {
      await inform({ type: "BULLETIN_OFFLINE" }, (_b2 = event.source) == null ? void 0 : _b2.id);
    }
  })());
});
self.addEventListener("fetch", (event) => {
  const request = event.request, url = new URL(request.url);
  if (request.method !== "GET") return;
  if (request.mode === "navigate" && isHome(url)) {
    const savedP = fallback().catch(() => null);
    event.waitUntil(savedP.then((s) => s && Date.now() - lastFresh >= RECENT ? refresh().catch(() => {
    }) : null));
    event.respondWith((async () => {
      const saved = await savedP;
      if (saved) return saved;
      try {
        const r = await refresh();
        if (r) return r.clone();
        return await fallback() || guidance();
      } catch (_) {
        return await fallback() || guidance();
      }
    })());
    return;
  }
  if (!ASSETS.includes(url.href) && !EXTERNAL.has(url.hostname)) return;
  const network = fetch(request, { cache: "no-cache" }).then(async (r) => {
    if (r.ok || r.type === "opaque") await save(request, r);
    return r;
  });
  event.waitUntil(network.catch(() => {
  }));
  event.respondWith((async () => {
    try {
      const r = await bounded(network);
      if (r.ok || r.type === "opaque") return r;
    } catch (_) {
    }
    return await read(CACHE, request) || new Response("", { status: 503 });
  })());
});
