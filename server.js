import http from 'node:http';
import {readFile} from 'node:fs/promises';

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

async function collect() {
  const checkedAt = new Date().toISOString();
  const parts = await Promise.all(SOURCES.map(async s => {
    try {
      const r = await fetch(s.url, {signal: AbortSignal.timeout(15000), headers: {Accept: 'application/json'}});
      if (!r.ok) throw Error('HTTP ' + r.status);
      return {s, data: await r.json()};
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
