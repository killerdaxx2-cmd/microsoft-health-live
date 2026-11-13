import http from 'node:http';
import {readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';

const SOURCES = [
  {key: 'mac', name: 'Microsoft 365 corporativo / Admin Center', url: 'https://status.cloud.microsoft/api/posts/mac'},
  {key: 'azure', name: 'Azure', url: 'https://status.cloud.microsoft/api/posts/azure'},
  {key: 'ppac', name: 'Power Platform', url: 'https://status.cloud.microsoft/api/posts/ppac'},
  {key: 'm365Consumer', name: 'Microsoft 365 (consumidor)', url: 'https://status.cloud.microsoft/api/posts/m365Consumer'}
];
const HEALTHY = new Set(['Available', 'Operational', 'Normal service']);

const clean = v => String(v ?? '')
  .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&#0?39;/g, "'")
  .replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();

// ---- Tradução pt-BR (MyMemory, público e sem chave; em falha mantém o original em inglês) ----
const MM_URL = 'https://api.mymemory.translated.net/get?langpair=en|pt&q=';
const trCache = new Map(); // hash do trecho -> tradução (só retraduz quando o texto muda)

const unescape = s => String(s ?? '')
  .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&#0?39;/g, "'").replace(/&#x27;/g, "'");

function chunksOf(text, size = 480) {
  const out = [];
  let rest = text;
  while (rest.length > size) {
    let cut = rest.lastIndexOf(' ', size);
    if (cut < size * 0.5) cut = size;
    out.push(rest.slice(0, cut));
    rest = rest.slice(cut).trimStart();
  }
  if (rest) out.push(rest);
  return out;
}

async function mmChunk(chunk) {
  const r = await fetch(MM_URL + encodeURIComponent(chunk), {
    signal: AbortSignal.timeout(12000),
    headers: {Accept: 'application/json'}
  });
  if (!r.ok) throw Error('mymemory HTTP ' + r.status);
  const j = await r.json();
  const t = unescape(j?.responseData?.translatedText || '').trim();
  if (Number(j?.responseStatus) !== 200 || !t || /MYMEMORY WARNING/i.test(t)) {
    throw Error('mymemory ' + String(j?.responseDetails || 'sem tradução'));
  }
  return t;
}

// Google Translate sem chave (endpoint público usado por extensões; pode ser bloqueado por IP)
async function gtxChunk(chunk) {
  const url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=pt&dt=t&q=' + encodeURIComponent(chunk);
  const r = await fetch(url, {
    signal: AbortSignal.timeout(8000),
    headers: {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)', Accept: '*/*'}
  });
  if (!r.ok) throw Error('gtx HTTP ' + r.status);
  const j = await r.json();
  const t = Array.isArray(j?.[0]) ? j[0].map(s => (Array.isArray(s) && s[0]) || '').join('') : '';
  if (!t.trim()) throw Error('gtx resposta vazia');
  return unescape(t);
}

const PROVIDERS = [gtxChunk, mmChunk];

async function translateChunk(chunk) {
  const errs = [];
  for (const p of PROVIDERS) {
    try { return await p(chunk); } catch (e) { errs.push(String(e?.message || e)); }
  }
  throw Error(errs.join(' | '));
}

async function translate(text, budgetMs = 12000) {
  if (!text) return text;
  const parts = chunksOf(text);
  const out = [];
  const start = Date.now();
  for (const c of parts) {
    const h = createHash('md5').update(c).digest('hex');
    if (trCache.has(h)) { out.push(trCache.get(h)); continue; }
    if (Date.now() - start > budgetMs) throw Error('orçamento de tempo esgotado');
    const t = await translateChunk(c);
    trCache.set(h, t);
    if (trCache.size > 500) trCache.delete(trCache.keys().next().value); // remove o mais antigo
    out.push(t);
    await new Promise(res => setTimeout(res, 120));
  }
  return out.join(' ');
}

const trErrors = new Map(); // msg -> última vez logada (evita repetir o mesmo erro no log)
const safeTranslate = async text => {
  try { return await translate(text); } catch (e) {
    const msg = String(e?.message || e);
    const last = trErrors.get(msg) || 0;
    if (Date.now() - last > 10 * 60 * 1000) {
      trErrors.set(msg, Date.now());
      console.warn('tradução indisponível, exibindo original:', msg);
    }
    return text;
  }
};

async function fetchWithRetry(url) {
  let lastErr;
  for (let i = 0; i < 2; i++) {
    try {
      const r = await fetch(url, {signal: AbortSignal.timeout(15000), headers: {Accept: 'application/json'}});
      if (!r.ok) throw Error('HTTP ' + r.status);
      return await r.json();
    } catch (e) { lastErr = e; if (i === 0) await new Promise(res => setTimeout(res, 1500)); }
  }
  throw lastErr;
}

async function collect() {
  const checkedAt = new Date().toISOString();
  const parts = await Promise.all(SOURCES.map(async s => {
    try {
      return {s, data: await fetchWithRetry(s.url)};
    } catch (e) { return {s, err: e.message}; }
  }));

  const entries = [], services = {}, sources = {};
  let anyOk = false;

  for (const {s, data, err} of parts) {
    if (err) { sources[s.key] = {ok: false, name: s.name, error: err}; continue; }
    anyOk = true;
    sources[s.key] = {
      ok: true,
      name: s.name,
      status: typeof data.Status === 'string' ? data.Status : null,
      lastUpdatedTime: typeof data.LastUpdatedTime === 'string' ? data.LastUpdatedTime : null
    };
    const items = Array.isArray(data) ? data : [data];
    for (const it of items) {
      const st = typeof it.Status === 'string' ? it.Status.trim() : '';
      const title = String(it.Title || '').trim();
      const date = it.LastUpdatedTime || it.AuthoredDateTime || null;

      if (Array.isArray(data)) { // feed do consumidor: uma entrada por serviço
        services[it.ServiceWorkloadName || it.ServiceDisplayName || 'outro'] = {
          name: it.ServiceDisplayName || it.ServiceWorkloadName,
          status: st || null,
          lastUpdatedTime: date,
          incidentTitle: title || null
        };
      }

      // saudável sem título (ex.: "Available" do Azure/Power Platform) não gera aviso
      if (HEALTHY.has(st) && !title) continue;
      if (!title && !st) continue;

      entries.push({
        source: s.key,
        sourceName: s.name,
        title: title || s.name,
        status: st || null,
        date,
        description: clean(it.Message).slice(0, 4000)
      });
    }
  }

  entries.sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  // traduz avisos em paralelo (cache por trecho: só retraduz o que mudou desde a última atualização)
  await Promise.all(entries.map(async e => {
    e.title = await safeTranslate(e.title);
    e.description = await safeTranslate(e.description);
  }));
  await Promise.all(Object.values(services).map(async s => {
    if (s.incidentTitle) s.incidentTitle = await safeTranslate(s.incidentTitle);
  }));

  // feed do consumidor é uma lista por serviço: deriva um status agregado
  if (sources.m365Consumer && sources.m365Consumer.ok) {
    const svc = Object.values(services);
    const bad = svc.filter(x => x.status && !HEALTHY.has(x.status));
    const last = svc.map(x => x.lastUpdatedTime || '').sort().pop() || null;
    sources.m365Consumer.status = bad.length ? (bad[0].status) : (svc.length ? 'Operational' : null);
    sources.m365Consumer.lastUpdatedTime = last;
  }
  return {ok: anyOk, checkedAt, sources, services, entries};
}

let cache, pending;
async function snapshot() {
  if (cache && Date.now() - cache.time < 30000) return cache.data;
  if (!pending) pending = collect().then(data => {cache = {time: Date.now(), data}; return data;}).finally(() => pending = null);
  return pending;
}

http.createServer(async (req, res) => {
  try {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    if (req.url === '/api/status') {
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.setHeader('Cache-Control', 'no-store');
      return res.end(JSON.stringify(await snapshot()));
    }
    if (req.url === '/healthz') return res.end('ok');
    if (req.url === '/' || req.url === '/index.html') {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      return res.end(await readFile(new URL('./index.html', import.meta.url)));
    }
    res.writeHead(404); res.end('Not found');
  } catch {
    res.writeHead(500); res.end('Erro interno');
  }
}).listen(Number(process.env.PORT || 3000), '0.0.0.0');
