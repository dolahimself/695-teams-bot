import { randomUUID } from 'crypto';
import { promises as fs } from 'fs';
import * as path from 'path';
import express from 'express';
import type { Express, NextFunction, Request, Response } from 'express';
import { TEXT_FIELDS, TextField, isTextField } from './flow';
import type { Transcriber } from './speech';
import type { LocalFile } from './store';

/**
 * Voice answers for the story boxes.
 *
 * An Adaptive Card cannot use the microphone, so each story box has a microphone button beside
 * its question. In Teams it opens a dialog showing /record (a small page that records). The page
 * sends the recording to /api/transcribe, shows the words so the worker can fix them, then hands
 * the text and the recording back to the bot, which puts the words in that box and shows the
 * recording on the card as a play button and its length.
 * The browser preview opens the same page in a pop-up.
 *
 * Recordings are kept in `dir` only so the worker can play them back. They are never sent with the story.
 */

export const MAX_RECORDING_SECONDS = 180;
const MAX_AUDIO_BYTES = 20 * 1024 * 1024;
const AUDIO_TYPE = /^(audio|video)\/(webm|ogg|mp4|mpeg|wav|x-wav|wave|aac|x-m4a|m4a)\b/i;
const ID = /^[0-9a-f-]{36}$/;

export interface VoiceOptions {
  /** Public address of this server. Empty gives links relative to this server (the preview). */
  baseUrl: string;
  /** The same short-lived link signing the photo dialog uses. */
  signer: { token: () => string; valid: (t: unknown) => boolean };
  /** The speech service. Without it, only the preview can record (with the browser's own recogniser). */
  transcriber?: Transcriber;
  /** Folder where recordings are kept for playing back. */
  dir: string;
  /**
   * Let the preview fall back to the browser's own speech recognition when there is no speech service
   * (default true). Turned off when collecting real stories: it sends the audio to Google or Apple.
   */
  fallback?: boolean;
}

const EXT: Record<string, string> = { 'audio/webm': 'webm', 'video/webm': 'webm', 'audio/mp4': 'm4a', 'audio/x-m4a': 'm4a', 'audio/ogg': 'ogg', 'audio/mpeg': 'mp3', 'audio/wav': 'wav', 'audio/x-wav': 'wav' };

export function createVoice(app: Express, opts: VoiceOptions) {
  const base = opts.baseUrl.replace(/\/$/, '');

  /** The address the dialog (or preview pop-up) opens for one story box. `again` replaces the last recording. */
  const recordUrl = (mode: 'teams' | 'preview', field: TextField, again = false) =>
    `${base}/record?t=${encodeURIComponent(opts.signer.token())}&field=${field}${again ? '&again=1' : ''}${mode === 'preview' ? '&mode=preview' : ''}`;

  /** A small player page for a recording made here; undefined for any other address. */
  const listenUrl = (recordingUrl: string) => {
    const m = /\/recordings\/([0-9a-f-]{36})$/.exec(recordingUrl);
    return m ? `${base}/listen?t=${encodeURIComponent(opts.signer.token())}&id=${m[1]}` : undefined;
  };

  app.get('/record', (req, res) => {
    if (!opts.signer.valid(req.query.t)) return res.status(403).send('This link has expired. Close this window and tap the microphone again.');
    const field = req.query.field;
    if (!isTextField(field)) return res.status(400).send('Unknown question.');
    const cfg = {
      field, question: TEXT_FIELDS[field].label, server: !!opts.transcriber, maxSeconds: MAX_RECORDING_SECONDS,
      again: req.query.again === '1', fallback: opts.fallback !== false,
    };
    res.type('html').send(RECORD_PAGE.replace('__CFG__', JSON.stringify(cfg).replace(/</g, '\\u003c')));
  });

  // Checked before the recording is read, so a bad request costs nothing.
  const guard = (req: Request, res: Response, next: NextFunction) => {
    if (!opts.signer.valid(req.query.t)) return res.status(403).json({ error: 'This link has expired. Close this window and tap the microphone again.' });
    if (!opts.transcriber) return res.status(501).json({ error: 'Voice typing is not set up on this server.' });
    if (!AUDIO_TYPE.test(String(req.headers['content-type'] ?? ''))) return res.status(415).json({ error: 'That is not a recording.' });
    next();
  };

  app.put('/api/transcribe', guard, express.raw({ type: () => true, limit: MAX_AUDIO_BYTES }), async (req, res) => {
    const audio = req.body as Buffer;
    if (!Buffer.isBuffer(audio) || !audio.length) return res.status(400).json({ error: 'The recording was empty. Please try again.' });
    const type = String(req.headers['content-type']);
    let text: string;
    try {
      text = await opts.transcriber!.transcribe(audio, type);
    } catch (err) {
      console.error('transcription failed', err);
      return res.status(502).json({ error: 'I could not turn that into text. Please try again, or type your answer instead.' });
    }
    if (!text) return res.json({ text });
    // Keep the recording so the worker can play it back from the card.
    try {
      const id = randomUUID();
      await fs.mkdir(opts.dir, { recursive: true });
      await fs.writeFile(path.join(opts.dir, id), audio);
      await fs.writeFile(path.join(opts.dir, id + '.json'), JSON.stringify({ type: type.split(';')[0].trim() }));
      res.json({ text, audio: { url: `${base}/recordings/${id}` } });
    } catch (err) {
      console.error('could not keep the recording', err);
      res.json({ text }); // the words still go on the card, just without playback
    }
  });

  app.get('/recordings/:id', async (req, res) => {
    const id = req.params.id;
    if (!ID.test(id)) return res.sendStatus(404);
    try {
      const meta = JSON.parse(await fs.readFile(path.join(opts.dir, id + '.json'), 'utf8')) as { type: string };
      res.type(meta.type).sendFile(path.resolve(opts.dir, id), (err) => { if (err && !res.headersSent) res.sendStatus(404); });
    } catch {
      res.sendStatus(404);
    }
  });

  app.get('/listen', (req, res) => {
    if (!opts.signer.valid(req.query.t)) return res.status(403).send('This link has expired. Close this window and tap play again.');
    const id = String(req.query.id ?? '');
    if (!ID.test(id)) return res.sendStatus(404);
    res.type('html').send(LISTEN_PAGE.replace('__SRC__', `${base}/recordings/${id}`));
  });

  /** The recording behind one of this server's /recordings links, or undefined for any other address. */
  const ours = (url: string) => {
    const m = /^(.*)\/recordings\/([0-9a-f-]{36})$/.exec(url);
    return m && (m[1] === '' || m[1] === base) ? m[2] : undefined;
  };

  /** A kept recording, so a store can send it on when the story is submitted. */
  const read = async (url: string): Promise<LocalFile | undefined> => {
    const id = ours(url);
    if (!id) return undefined;
    try {
      const meta = JSON.parse(await fs.readFile(path.join(opts.dir, id + '.json'), 'utf8')) as { type: string };
      return { bytes: await fs.readFile(path.join(opts.dir, id)), type: meta.type, name: `recording.${EXT[meta.type] ?? 'webm'}` };
    } catch {
      return undefined;
    }
  };

  /** Deletes a kept recording once its story is saved elsewhere. */
  const remove = async (url: string) => {
    const id = ours(url);
    if (!id) return;
    await fs.rm(path.join(opts.dir, id + '.json'), { force: true }); // first, so the link stops working before the file goes
    await fs.rm(path.join(opts.dir, id), { force: true });
  };

  return { recordUrl, listenUrl, available: !!opts.transcriber, read, remove };
}

/** Teams cannot play sound inside a card, so the play button opens this. */
const LISTEN_PAGE = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Your recording</title>
<style>
  body { margin:0; padding:16px; font-family:-apple-system,"SF Pro Text","Segoe UI",system-ui,sans-serif; color:#242424; display:flex; flex-direction:column; gap:12px; }
  h1 { font-size:18px; margin:0; font-weight:600; } audio { width:100%; }
</style></head><body>
<h1>Your recording</h1>
<audio controls autoplay preload="auto" src="__SRC__"></audio>
</body></html>`;

const RECORD_PAGE = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Record your answer</title>
<script src="https://res.cdn.office.net/teams-js/2.31.0/js/MicrosoftTeams.min.js"></script>
<style>
  :root { --brand:#5B5FC7; --ink:#242424; --mute:#616161; --red:#C4314B; --font:-apple-system,"SF Pro Text","Segoe UI",system-ui,sans-serif; }
  * { box-sizing:border-box; } html,body { margin:0; min-height:100%; }
  body { font-family:var(--font); color:var(--ink); background:#fff; padding:16px; display:flex; flex-direction:column; gap:12px; min-height:100vh; }
  h1 { font-size:18px; margin:0; font-weight:600; } .q { font-size:16px; font-weight:600; color:var(--ink); }
  p { margin:0; color:var(--mute); font-size:14px; line-height:1.4; }
  [hidden] { display:none !important; }
  .stage { flex:1; min-height:200px; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:12px; text-align:center; }
  .mic { width:88px; height:88px; border-radius:50%; border:0; background:var(--brand); color:#fff; cursor:pointer; display:flex; align-items:center; justify-content:center; }
  .mic:focus-visible { outline:3px solid var(--ink); outline-offset:3px; } .mic:disabled { opacity:.4; cursor:default; }
  .mic.on { background:var(--red); animation:pulse 1.4s ease-in-out infinite; }
  @keyframes pulse { 0%,100% { box-shadow:0 0 0 0 rgba(196,49,75,.45); } 50% { box-shadow:0 0 0 14px rgba(196,49,75,0); } }
  .spin { width:32px; height:32px; border:3px solid #E0E0E0; border-top-color:var(--brand); border-radius:50%; animation:spin .9s linear infinite; }
  @keyframes spin { to { transform:rotate(360deg); } }
  @media (prefers-reduced-motion: reduce) { .mic.on, .spin { animation:none; } }
  .time { font:600 20px var(--font); font-variant-numeric:tabular-nums; }
  .status { font-size:14px; color:var(--mute); max-width:420px; line-height:1.4; } .status.err { color:var(--red); }
  .note { font-size:12.5px; }
  .words { flex:1; display:flex; flex-direction:column; gap:6px; }
  .words label { font-size:14px; font-weight:600; }
  textarea { flex:1; width:100%; min-height:160px; border:1px solid #C8C8C8; border-radius:6px; padding:10px 12px; font:15px/1.45 var(--font); color:var(--ink); resize:vertical; }
  textarea:focus { outline:2px solid var(--brand); outline-offset:-1px; }
  .row { display:flex; gap:8px; } .row button { flex:1; min-height:46px; border-radius:8px; font:600 15px var(--font); white-space:nowrap; cursor:pointer; border:1.5px solid var(--brand); background:#fff; color:var(--brand); padding:0 6px; }
  .row button.go { background:var(--brand); color:#fff; } .row button:disabled { opacity:.45; cursor:default; }
</style></head><body>
<h1>Record your answer</h1>
<p class="q" id="q"></p>
<p>Speak naturally. You can read and fix the words before they go on your card. Please leave out names and identifying details.</p>
<p class="note" id="againNote" hidden>Recording again replaces your last recording, and the words it gave if you have not changed them.</p>
<div class="stage" id="stage">
  <button type="button" class="mic" id="mic" aria-label="Start recording"></button>
  <div class="time" id="time" hidden>0:00</div>
  <div class="spin" id="spin" hidden></div>
  <div class="status" id="status" role="status" aria-live="polite">Tap the microphone to start.</div>
  <p class="note" id="note" hidden>Preview only: your browser's own speech recognition turns your words into text.</p>
</div>
<div class="words" id="words" hidden>
  <label for="text">Your words (fix anything that is wrong)</label>
  <textarea id="text"></textarea>
</div>
<div class="row"><button type="button" id="cancel">Cancel</button><button type="button" id="again" hidden>Record again</button><button type="button" class="go" id="done" disabled>Add to story</button></div>
<script>
const CFG = __CFG__;
const q = new URLSearchParams(location.search), tok = q.get('t'), preview = q.get('mode') === 'preview';
const $ = (id) => document.getElementById(id);
const MIC = '<svg viewBox="0 0 24 24" width="36" height="36" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0014 0M12 18v3"/></svg>';
const STOP = '<svg viewBox="0 0 24 24" width="30" height="30" aria-hidden="true"><rect x="6" y="6" width="12" height="12" rx="2" fill="currentColor"/></svg>';
const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
const canRecord = !!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia && window.MediaRecorder);
// Real use sends the recording to the agency's speech service. Only the preview may fall back to the
// browser's own recogniser, so a demo works without one.
const how = CFG.server && canRecord ? 'server' : preview && SR && CFG.fallback ? 'browser' : null;
const BLOCKED = 'Microphone access is blocked. Allow it in your browser or Teams settings, or type your answer instead.';
let rec = null, stream = null, chunks = [], timer = null, started = 0, seconds = 0, recognizer = null, closing = false;
let clip = null; // { blob, url, seconds }: the recording, for playing back from the card
let clipReady = Promise.resolve();

$('q').textContent = CFG.question;
$('note').hidden = how !== 'browser';
$('againNote').hidden = !CFG.again;
const fmt = (s) => Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0');
function status(t, err) { $('status').textContent = t; $('status').classList.toggle('err', !!err); }
function view(state) { // ready | recording | working | words
  $('stage').hidden = state === 'words';
  $('words').hidden = state !== 'words';
  $('mic').hidden = state === 'working';
  $('mic').classList.toggle('on', state === 'recording');
  $('mic').innerHTML = state === 'recording' ? STOP : MIC;
  $('mic').setAttribute('aria-label', state === 'recording' ? 'Stop recording' : 'Start recording');
  $('spin').hidden = state !== 'working';
  $('time').hidden = state !== 'recording';
  $('again').hidden = state !== 'words';
  $('done').disabled = state !== 'words' || !$('text').value.trim();
}
function tick() {
  const s = Math.floor((Date.now() - started) / 1000);
  $('time').textContent = fmt(s) + ' / ' + fmt(CFG.maxSeconds);
  if (s >= CFG.maxSeconds) stop();
}
function startClock() { started = Date.now(); tick(); timer = setInterval(tick, 250); }
function showWords(t) { $('text').value = t; view('words'); $('text').focus(); }

// Records the microphone; onDone gets the finished recording.
async function record(onDone) {
  stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  const type = ['audio/webm;codecs=opus', 'audio/mp4', 'audio/ogg;codecs=opus', 'audio/webm'].find((t) => MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(t));
  rec = new MediaRecorder(stream, type ? { mimeType: type } : undefined);
  chunks = [];
  rec.ondataavailable = (e) => { if (e.data && e.data.size) chunks.push(e.data); };
  clipReady = new Promise((resolve) => {
    rec.onstop = () => { const blob = new Blob(chunks, { type: rec.mimeType || 'audio/webm' }); onDone(blob); resolve(); };
  });
  rec.start(1000);
}

async function start() {
  status('');
  clip = null;
  if (how === 'browser') return listen();
  try { await record(upload); }
  catch (e) { return status(e && e.name === 'NotAllowedError' ? BLOCKED : 'No microphone was found. You can type your answer instead.', true); }
  startClock(); view('recording'); status('Recording. Tap the button again when you are done.');
}

function stop() {
  if (timer) { clearInterval(timer); timer = null; seconds = Math.max(1, Math.round((Date.now() - started) / 1000)); }
  if (recognizer) recognizer.stop();
  if (rec && rec.state !== 'inactive') rec.stop();
  if (stream) { stream.getTracks().forEach((t) => t.stop()); stream = null; }
}

async function upload(blob) {
  if (closing) return;
  if (!blob.size) { view('ready'); return status('Nothing was recorded. Please try again.', true); }
  view('working'); status('Turning your recording into text...');
  try {
    const res = await fetch('/api/transcribe?t=' + encodeURIComponent(tok), { method: 'PUT', headers: { 'Content-Type': blob.type }, body: blob });
    const out = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(out.error || (res.status === 413 ? 'That recording is too long. Please keep it under ' + Math.round(CFG.maxSeconds / 60) + ' minutes.' : 'I could not turn that into text. Please try again, or type your answer instead.'));
    if (!out.text) { view('ready'); return status('I could not hear any words. Try again a little closer to the microphone.', true); }
    if (out.audio && out.audio.url) clip = { blob, url: out.audio.url, seconds };
    showWords(out.text);
  } catch (e) { view('ready'); status(e.message, true); }
}

// Preview fallback: live dictation with the browser's own speech recognition.
async function listen() {
  let done = '', heard = '', problem = '';
  // Keep the sound too, so it can be played back from the card. Dictation still works without it.
  if (canRecord) {
    try { await record((blob) => { if (blob.size) clip = { blob, url: 'local:' + Date.now().toString(36) + Math.random().toString(36).slice(2), seconds }; }); }
    catch (e) { /* no playback, words only */ }
  }
  recognizer = new SR();
  recognizer.lang = 'en-CA'; recognizer.continuous = true; recognizer.interimResults = true;
  recognizer.onresult = (e) => {
    let interim = '';
    for (let i = e.resultIndex; i < e.results.length; i++) {
      if (e.results[i].isFinal) done += e.results[i][0].transcript + ' '; else interim += e.results[i][0].transcript;
    }
    heard = (done + interim).replace(/\\s+/g, ' ').trim();
    status(heard ? '"' + (heard.length > 140 ? '...' + heard.slice(-140) : heard) + '"' : 'Listening...');
  };
  recognizer.onerror = (e) => { problem = e.error; };
  recognizer.onend = () => {
    recognizer = null;
    stop(); // stops the recording and the clock as well
    if (closing) return;
    if (heard) return showWords(heard);
    view('ready');
    status(problem === 'not-allowed' || problem === 'service-not-allowed' ? BLOCKED : 'I could not hear any words. Please try again.', true);
  };
  recognizer.start();
  startClock(); view('recording'); status('Listening...');
}

let teams = false;
if (!preview && window.microsoftTeams) { microsoftTeams.app.initialize().then(() => { teams = true; }).catch(() => {}); }
async function finish(voice) {
  closing = true; stop();
  if (voice) { await clipReady; if (clip) voice.audio = { url: clip.url, seconds: clip.seconds }; }
  // The preview gets the recording itself too, so it plays from the browser's own copy.
  if (preview) parent.postMessage({ bearVoice: voice, audio: voice && clip ? clip.blob : null }, location.origin);
  else if (teams) microsoftTeams.dialog.url.submit(voice ? { voice } : {});
}

$('mic').onclick = () => ((rec && rec.state === 'recording') || recognizer ? stop() : start());
$('again').onclick = () => { view('ready'); status('Tap the microphone to start.'); };
$('text').oninput = () => { $('done').disabled = !$('text').value.trim(); };
$('done').onclick = () => finish({ field: CFG.field, text: $('text').value.trim() });
$('cancel').onclick = () => finish(null);
view('ready');
if (!how) {
  $('mic').disabled = true;
  status(canRecord || SR
    ? 'Voice typing is not set up yet. Please type your answer, or use the microphone key on your phone keyboard.'
    : 'This device cannot record here. Please type your answer, or use the microphone key on your phone keyboard.', true);
}
</script></body></html>`;
