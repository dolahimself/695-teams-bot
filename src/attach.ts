import { createHmac, randomBytes, randomUUID, timingSafeEqual } from 'crypto';
import { createWriteStream, promises as fs } from 'fs';
import type { Express, Request } from 'express';
import * as path from 'path';

/**
 * Photo upload for the story card.
 *
 * An Adaptive Card cannot hold a file picker, so the card has an "Add Photos"
 * button. In Teams the button opens a dialog that shows /attach (a small web page with a
 * real file picker). The page uploads each file here, then hands the list back to the bot,
 * which redraws the card with the files in its collection box.
 * The browser preview opens the same page in a pop-up.
 */

export const MAX_UPLOAD_BYTES = 150 * 1024 * 1024;
const TOKEN_LIFETIME_MS = 2 * 60 * 60 * 1000;

export interface AttachOptions {
  /** Public https address of this server (the same tunnel Teams uses). */
  baseUrl: string;
  /** Folder where uploads are kept until the story is saved. */
  dir: string;
  /** Signs the short-lived links so only people who opened the dialog can upload. */
  secret?: string;
}

const clean = (n: string) => n.replace(/[^\w.\- ()]+/g, '_').slice(0, 120) || 'file';

export function createAttach(app: Express, opts: AttachOptions) {
  const secret = opts.secret || randomBytes(32).toString('hex');
  const sign = (exp: string) => createHmac('sha256', secret).update(exp).digest('hex').slice(0, 32);

  const token = (now = Date.now()) => {
    const exp = String(now + TOKEN_LIFETIME_MS);
    return `${exp}.${sign(exp)}`;
  };
  const valid = (t: unknown, now = Date.now()) => {
    if (typeof t !== 'string') return false;
    const [exp, mac] = t.split('.');
    if (!exp || !mac || Number(exp) < now) return false;
    const want = Buffer.from(sign(exp));
    const got = Buffer.from(mac);
    return want.length === got.length && timingSafeEqual(want, got);
  };

  /** The address the dialog (or preview pop-up) opens. */
  const attachUrl = (mode: 'teams' | 'preview' = 'teams') =>
    `${opts.baseUrl.replace(/\/$/, '')}/attach?t=${encodeURIComponent(token())}${mode === 'preview' ? '&mode=preview' : ''}`;

  app.get('/attach', (req, res) => {
    if (!valid(req.query.t)) return res.status(403).send('This link has expired. Close this window and tap Add Photos again.');
    res.type('html').send(ATTACH_PAGE);
  });

  app.put('/api/upload', async (req: Request, res) => {
    if (!valid(req.query.t)) return res.status(403).json({ error: 'expired' });
    const name = clean(String(req.query.name ?? 'file'));
    const type = String(req.query.type ?? '').slice(0, 100) || 'application/octet-stream';
    if (!/^image\//.test(type)) return res.status(415).json({ error: 'Only photos can be added.' });
    const id = randomUUID();
    await fs.mkdir(opts.dir, { recursive: true });
    const file = path.join(opts.dir, id);
    let bytes = 0;
    let tooBig = false;
    const out = createWriteStream(file);
    req.on('data', (c: Buffer) => {
      bytes += c.length;
      if (bytes > MAX_UPLOAD_BYTES && !tooBig) {
        tooBig = true;
        req.unpipe(out);
        out.destroy();
        res.status(413).json({ error: 'That file is over 150 MB.' });
        void fs.rm(file, { force: true });
      }
    });
    req.pipe(out);
    out.on('finish', async () => {
      if (tooBig) return;
      await fs.writeFile(file + '.json', JSON.stringify({ name, type }));
      res.json({ name, contentType: type, url: `${opts.baseUrl.replace(/\/$/, '')}/files/${id}/${encodeURIComponent(name)}` });
    });
    out.on('error', () => !res.headersSent && res.status(500).json({ error: 'Upload failed.' }));
  });

  app.get('/files/:id/:name', async (req, res) => {
    const id = req.params.id;
    if (!/^[0-9a-f-]{36}$/.test(id)) return res.sendStatus(404);
    try {
      const meta = JSON.parse(await fs.readFile(path.join(opts.dir, id + '.json'), 'utf8')) as { type: string };
      res.type(meta.type).sendFile(path.resolve(opts.dir, id));
    } catch {
      res.sendStatus(404);
    }
  });

  return { attachUrl, valid, token };
}

const ATTACH_PAGE = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Add Photos</title>
<script src="https://res.cdn.office.net/teams-js/2.31.0/js/MicrosoftTeams.min.js"></script>
<style>
  :root { --brand:#5B5FC7; --ink:#242424; --mute:#616161; --font:-apple-system,"SF Pro Text","Segoe UI",system-ui,sans-serif; }
  * { box-sizing:border-box; } html,body { margin:0; min-height:100%; }
  body { font-family:var(--font); color:var(--ink); background:#fff; padding:16px; display:flex; flex-direction:column; gap:14px; min-height:100vh; }
  h1 { font-size:18px; margin:0; font-weight:600; } p { margin:0; color:var(--mute); font-size:14px; line-height:1.4; }
  .pick { border:2px dashed #B8BAE6; border-radius:12px; background:#F5F5FC; color:var(--brand); min-height:96px; font:600 16px var(--font); display:flex; align-items:center; justify-content:center; text-align:center; padding:12px; cursor:pointer; }
  .pick.over { background:#E8EBFA; }
  ul { list-style:none; margin:0; padding:0; display:flex; flex-direction:column; gap:8px; flex:1; }
  li { display:flex; gap:10px; align-items:center; border:1px solid #E0E0E0; border-radius:10px; padding:8px; }
  li .th { width:52px; height:52px; border-radius:8px; object-fit:cover; background:#EEE; flex:none; display:flex; align-items:center; justify-content:center; font-size:11px; color:var(--mute); }
  li .m { flex:1; min-width:0; } li .n { font-size:14px; font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; } li .s { font-size:12px; color:var(--mute); }
  li .bar { height:4px; background:#E8E8E8; border-radius:2px; margin-top:6px; overflow:hidden; } li .bar i { display:block; height:100%; width:0; background:var(--brand); }
  li.err .s { color:#C4314B; }
  li button { border:1px solid var(--brand); background:#fff; color:var(--brand); border-radius:8px; min-height:34px; padding:0 10px; font:600 13px var(--font); cursor:pointer; }
  .row { display:flex; gap:8px; } .row button { flex:1; min-height:46px; border-radius:8px; font:600 16px var(--font); cursor:pointer; border:1.5px solid var(--brand); background:#fff; color:var(--brand); }
  .row button.go { background:var(--brand); color:#fff; } .row button:disabled { opacity:.45; cursor:default; }
</style></head><body>
<h1>Add Photos</h1>
<p>Up to 10 per story. Please avoid faces or names where you can.</p>
<label class="pick" id="drop">Choose photos<input id="f" type="file" accept="image/*" multiple hidden></label>
<ul id="list"></ul>
<div class="row"><button type="button" id="cancel">Cancel</button><button type="button" class="go" id="done" disabled>Add to story</button></div>
<script>
const q = new URLSearchParams(location.search), tok = q.get('t'), preview = q.get('mode') === 'preview';
const MAX = 10, list = document.getElementById('list'), done = document.getElementById('done');
const items = []; // {name,url,contentType} once uploaded
let busy = 0;
const size = (n) => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
function refresh() { done.disabled = busy > 0 || items.length === 0; done.textContent = items.length ? 'Add ' + items.length + ' to story' : 'Add to story'; }
function addFile(file) {
  if (list.children.length >= MAX) return;
  const li = document.createElement('li');
  const kind = file.type.startsWith('image/') ? 'img' : file.type.startsWith('video/') ? 'video' : 'div';
  const th = document.createElement(kind); th.className = 'th';
  if (kind === 'div') th.textContent = 'audio'; else { th.src = URL.createObjectURL(file); if (kind === 'video') { th.muted = true; th.preload = 'metadata'; } }
  const m = document.createElement('div'); m.className = 'm';
  m.innerHTML = '<div class="n"></div><div class="s">Uploading...</div><div class="bar"><i></i></div>';
  m.querySelector('.n').textContent = file.name;
  const rm = document.createElement('button'); rm.type = 'button'; rm.textContent = 'Remove';
  li.append(th, m, rm); list.appendChild(li);
  let rec = null, xhr = new XMLHttpRequest(); busy++; refresh();
  rm.onclick = () => { xhr.abort(); if (rec) items.splice(items.indexOf(rec), 1); else busy--; li.remove(); refresh(); };
  xhr.open('PUT', '/api/upload?t=' + encodeURIComponent(tok) + '&name=' + encodeURIComponent(file.name) + '&type=' + encodeURIComponent(file.type));
  xhr.upload.onprogress = (e) => { if (e.lengthComputable) m.querySelector('i').style.width = (100 * e.loaded / e.total) + '%'; };
  xhr.onload = () => {
    busy--;
    if (xhr.status === 200) { rec = JSON.parse(xhr.responseText); items.push(rec); m.querySelector('.s').textContent = size(file.size); m.querySelector('.bar').remove(); }
    else { li.classList.add('err'); let msg = 'Could not upload'; try { msg = JSON.parse(xhr.responseText).error || msg; } catch (_) {} m.querySelector('.s').textContent = msg; m.querySelector('.bar').remove(); }
    refresh();
  };
  xhr.onerror = () => { busy--; li.classList.add('err'); m.querySelector('.s').textContent = 'Could not upload'; refresh(); };
  xhr.send(file);
}
const input = document.getElementById('f');
input.onchange = () => { for (const f of input.files) if (f.type.startsWith('image/')) addFile(f); input.value = ''; };
const drop = document.getElementById('drop');
drop.ondragover = (e) => { e.preventDefault(); drop.classList.add('over'); }; drop.ondragleave = () => drop.classList.remove('over');
drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove('over'); for (const f of e.dataTransfer.files) if (f.type.startsWith('image/')) addFile(f); };

let teams = false;
if (!preview && window.microsoftTeams) { microsoftTeams.app.initialize().then(() => { teams = true; }).catch(() => {}); }
function finish(result) {
  if (preview) parent.postMessage({ bearAttach: result }, '*');
  else if (teams) microsoftTeams.dialog.url.submit(result);
}
done.onclick = () => finish({ files: items });
document.getElementById('cancel').onclick = () => finish({ files: [] });
</script></body></html>`;
