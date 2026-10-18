// Local test harness: serves ../client-sign statically and mocks the Edge Function at /fn
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

const ROOT = process.argv[2];
const PORT = Number(process.argv[3] || 8765);
const OUT = process.argv[4] || '/tmp/mock-out';
fs.mkdirSync(OUT, { recursive: true });
const sessions = new Map();
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.json': 'application/json' };

function readBody(req) { return new Promise(r => { const c = []; req.on('data', d => c.push(d)); req.on('end', () => r(Buffer.concat(c))); }); }
function send(res, status, obj) { res.writeHead(status, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }); res.end(JSON.stringify(obj)); }

http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://x');
  if (url.pathname === '/fn') {
    if (req.method === 'OPTIONS') { res.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': '*' }); return res.end(); }
    const body = await readBody(req);
    const ct = req.headers['content-type'] || '';
    try {
      if (ct.includes('multipart/form-data')) {
        const form = await new Response(body, { headers: { 'content-type': ct } }).formData();
        const id = String(form.get('id'));
        const s = sessions.get(id);
        if (!s) return send(res, 404, { ok: false, error: 'This link is not valid.' });
        if (s.status === 'signed') return send(res, 409, { ok: false, error: 'already submitted' });
        const files = [];
        for (const [k, v] of form.entries()) {
          if (!/^file\d+$/.test(k)) continue;
          const buf = Buffer.from(await v.arrayBuffer());
          const p = path.join(OUT, v.name);
          fs.writeFileSync(p, buf);
          files.push({ name: v.name, size: buf.length, pdf: buf.slice(0, 4).toString() === '%PDF', path: p });
        }
        s.status = 'signed'; s.signedAt = new Date().toISOString(); s.files = files;
        s.snapshot = JSON.parse(String(form.get('snapshot') || 'null'));
        console.log('[mock] submit', id, files.map(f => `${f.name} ${(f.size / 1024).toFixed(0)}KB pdf=${f.pdf}`).join(', '));
        return send(res, 200, { ok: true, signedAt: s.signedAt, emailedTo: [s.participantEmail || '', 'agedcare@onedreamcommunity.com.au'].filter(Boolean), files: files.map(f => ({ name: f.name, webUrl: null })) });
      }
      const b = JSON.parse(body.toString() || '{}');
      if (b.action === 'create') {
        if (b.pin !== '1234') return send(res, 403, { ok: false, error: 'Incorrect coordinator PIN.' });
        const id = crypto.randomUUID();
        sessions.set(id, { ...b, status: 'sent' });
        console.log('[mock] create', id, b.page, b.participantName, 'payload', JSON.stringify(b.payload).length, 'bytes');
        return send(res, 200, { ok: true, id, expiresAt: new Date(Date.now() + 30 * 864e5).toISOString() });
      }
      if (b.action === 'get') {
        const s = sessions.get(String(b.id));
        if (!s) return send(res, 404, { ok: false, error: 'This link is not valid.' });
        if (s.status === 'signed') return send(res, 200, { ok: true, page: s.page, status: 'signed', participant: { name: s.participantName, email: s.participantEmail }, signedAt: s.signedAt, emailedTo: [] });
        s.status = 'opened';
        return send(res, 200, { ok: true, page: s.page, status: 'opened', participant: { name: s.participantName, email: s.participantEmail }, payload: s.payload });
      }
      if (b.action === '_dump') return send(res, 200, Object.fromEntries(sessions));
      return send(res, 400, { ok: false, error: 'Unknown action.' });
    } catch (e) { console.error(e); return send(res, 500, { ok: false, error: String(e.message || e) }); }
  }
  // static
  let p = decodeURIComponent(url.pathname);
  if (p.endsWith('/')) p += 'index.html';
  if (p === '/sign/config.js') {
    res.writeHead(200, { 'Content-Type': 'text/javascript' });
    return res.end(`window.ODC_SIGN_CONFIG={functionUrl:'http://localhost:${PORT}/fn',anonKey:'',officePhone:'02 9161 7567',officeEmail:'agedcare@onedreamcommunity.com.au',linkValidDays:30};`);
  }
  const f = path.join(ROOT, p);
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); return res.end('404 ' + p); }
  res.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
}).listen(PORT, () => console.log('[mock] listening on', PORT, 'root', ROOT));
