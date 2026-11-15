import test from 'node:test';
import assert from 'node:assert/strict';
import express from 'express';
import { promises as fs } from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { AddressInfo } from 'net';
import { createAttach } from '../src/attach';
import type { Transcriber } from '../src/speech';
import { createVoice } from '../src/voice';

async function serve(transcriber?: Transcriber) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'bfv-'));
  const app = express();
  const attach = createAttach(app, { baseUrl: '', dir, secret: 'test-secret' });
  const voice = createVoice(app, { baseUrl: '', signer: attach, transcriber, dir: path.join(dir, 'recordings') });
  const server = app.listen(0);
  const base = `http://localhost:${(server.address() as AddressInfo).port}`;
  const t = encodeURIComponent(attach.token());
  const put = (type: string, bytes: number, tok = t) =>
    fetch(`${base}/api/transcribe?t=${tok}`, { method: 'PUT', headers: { 'Content-Type': type }, body: new Uint8Array(bytes).fill(3) });
  return { base, voice, server, put, attach };
}

test('record page: needs a valid link and a known question; the recording comes back as text', async () => {
  const got: [number, string][] = [];
  const { base, voice, server, put } = await serve({ transcribe: async (a, type) => { got.push([a.length, type]); return 'They read aloud now.'; } });
  try {
    assert.equal(voice.available, true);
    const url = voice.recordUrl('preview', 'whatChanged');
    assert.match(url, /^\/record\?t=.+&field=whatChanged&mode=preview$/);
    const page = await fetch(base + url);
    assert.equal(page.status, 200);
    const html = await page.text();
    assert.match(html, /What changed for the child or family\?/);
    assert.match(html, /"server":true/);
    assert.equal((await fetch(`${base}/record?t=bad&field=whatChanged`)).status, 403);
    assert.equal((await fetch(base + url.replace('field=whatChanged', 'field=program'))).status, 400);

    const ok = await put('audio/webm;codecs=opus', 500);
    assert.equal(ok.status, 200);
    const body = (await ok.json()) as { text: string; audio: { url: string } };
    assert.equal(body.text, 'They read aloud now.');
    assert.match(body.audio.url, /^\/recordings\/[0-9a-f-]{36}$/); // kept, so the worker can play it back
    const back = await fetch(base + body.audio.url);
    assert.equal(back.status, 200);
    assert.equal(back.headers.get('content-type'), 'audio/webm');
    assert.equal((await back.arrayBuffer()).byteLength, 500);

    const listen = voice.listenUrl(body.audio.url)!;
    assert.match(listen, /^\/listen\?t=.+&id=[0-9a-f-]{36}$/);
    const player = await (await fetch(base + listen)).text();
    assert.ok(player.includes(`src="${body.audio.url}"`));
    assert.equal(voice.listenUrl('https://evil.example/recordings/x.mp3'), undefined);
    assert.equal((await fetch(`${base}/listen?t=bad&id=${body.audio.url.split('/').pop()}`)).status, 403);
    assert.equal((await fetch(`${base}/recordings/..%2F..%2Fetc`)).status, 404);
    assert.match(voice.recordUrl('teams', 'whatHappened', true), /&again=1$/);
    assert.match(await (await fetch(base + voice.recordUrl('preview', 'whatHappened', true))).text(), /"again":true/);
    assert.deepEqual(got, [[500, 'audio/webm;codecs=opus']]);
    assert.equal((await put('audio/mp4', 10)).status, 200); // Safari

    assert.equal((await put('audio/webm', 5, 'bad')).status, 403);
    assert.equal((await put('image/png', 5)).status, 415);
    assert.equal((await put('audio/webm', 0)).status, 400);
    assert.equal(got.length, 2); // refused requests never reach the speech service
  } finally {
    server.close();
  }
});

test('record page without a speech service: the page knows, and transcribe says it is not set up', async () => {
  const { base, voice, server, put } = await serve();
  try {
    assert.equal(voice.available, false);
    assert.match(await (await fetch(base + voice.recordUrl('teams', 'whatHappened'))).text(), /"server":false/);
    assert.equal((await put('audio/webm', 5)).status, 501);
  } finally {
    server.close();
  }
});

test('a failing speech service gives the worker a plain message', async () => {
  const { server, put } = await serve({ transcribe: async () => { throw new Error('down'); } });
  const quiet = console.error; console.error = () => {};
  try {
    const res = await put('audio/webm', 5);
    assert.equal(res.status, 502);
    assert.match(((await res.json()) as { error: string }).error, /type your answer instead/);
  } finally {
    console.error = quiet;
    server.close();
  }
});

test('no words, no kept recording', async () => {
  const { server, put } = await serve({ transcribe: async () => '' });
  try {
    assert.deepEqual(await (await put('audio/webm', 50)).json(), { text: '' });
  } finally {
    server.close();
  }
});
