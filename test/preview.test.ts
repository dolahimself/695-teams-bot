import test from 'node:test';
import assert from 'node:assert/strict';
import express from 'express';
import { promises as fs } from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { AddressInfo } from 'net';
import type { Draft } from '../src/flow';
import { createPreviewApp } from '../src/preview';
import type { StoryStore } from '../src/store';

const STORY = { program: 'Camp', whatHappened: 'A child read their notes to the group.', whatChanged: 'They ask to read aloud every week.' };
const listen = (app: express.Express) => {
  const server = app.listen(0);
  return { server, base: `http://localhost:${(server.address() as AddressInfo).port}` };
};

async function serve(store?: StoryStore) {
  delete process.env.STORE;
  delete process.env.SPEECH_KEY;
  const dataDir = await fs.mkdtemp(path.join(os.tmpdir(), 'bfp-'));
  const { server, base } = listen(createPreviewApp(store ? { store, dataDir } : { dataDir }));
  const sid = 'test-' + Math.random();
  const event = async (e: object) =>
    (await (await fetch(`${base}/event`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ sid, event: e }) })).json()) as any;
  return { server, base, event };
}

test('pilot: a submitted story is saved once, with its recordings; the page is told it is a pilot', async () => {
  const saved: Draft[] = [];
  const { server, base, event } = await serve({ save: async (d) => { saved.push(d); return 'ok'; } });
  try {
    const config = (await (await fetch(`${base}/config`)).json()) as any;
    assert.equal(config.collecting, true);
    assert.match(config.notice, /stories you submit here are real/);
    assert.equal(config.reviewTeam, 'The project team');

    const start = await event({ type: 'start' });
    assert.doesNotMatch(JSON.stringify(start.card), /"action":"record"/); // no speech service: no microphones in a pilot
    await event({ type: 'voice', field: 'whatHappened', text: STORY.whatHappened, audio: { url: '/recordings/abc', seconds: 12 }, values: STORY });
    await event({ type: 'review', values: STORY });
    const unticked = await event({ type: 'send' });
    assert.equal(unticked.sent, false);
    assert.equal(saved.length, 0); // the permission box still comes first

    const sent = await event({ type: 'send', consent: true });
    assert.equal(sent.sent, true);
    assert.match(JSON.stringify(sent.card), /Sent\. The project team will review it/);
    assert.equal(saved.length, 1);
    assert.equal(saved[0].program, 'Camp');
    assert.deepEqual(saved[0].recordings.whatHappened, { url: '/recordings/abc', seconds: 12, text: STORY.whatHappened });
    await event({ type: 'send', consent: true });
    assert.equal(saved.length, 1); // a second tap does not save it twice
  } finally {
    server.close();
  }
});

test('pilot: if saving fails, the story comes back to review with a message, and Submit works on a retry', async () => {
  let attempts = 0;
  const saved: Draft[] = [];
  const quiet = console.error;
  console.error = () => {};
  const { server, event } = await serve({ save: async (d) => { if (++attempts === 1) throw new Error('sheet down'); saved.push(d); return 'ok'; } });
  try {
    await event({ type: 'start' });
    await event({ type: 'review', values: STORY });
    const failed = await event({ type: 'send', consent: true });
    assert.equal(failed.sent, false);
    assert.match(failed.say, /could not save that just now/);
    assert.match(JSON.stringify(failed.card), /Check your story/);
    assert.equal(saved.length, 0);
    const retry = await event({ type: 'send', consent: true });
    assert.equal(retry.sent, true);
    assert.equal(saved.length, 1);
  } finally {
    console.error = quiet;
    server.close();
  }
});

test('demo: nothing is saved, the page says so, and the browser fallback for voice stays on', async () => {
  const { server, base, event } = await serve();
  try {
    const config = (await (await fetch(`${base}/config`)).json()) as any;
    assert.equal(config.collecting, false);
    assert.equal(config.notice, null);
    const start = await event({ type: 'start' });
    assert.match(JSON.stringify(start.card), /"action":"record"/);
    await event({ type: 'review', values: STORY });
    assert.match(JSON.stringify((await event({ type: 'send', consent: true })).card), /Sent\. The Foundation team will review it/);
    const { url } = (await (await fetch(`${base}/record-url?field=whatHappened`)).json()) as any;
    assert.match(await (await fetch(base + url)).text(), /"fallback":true/);
  } finally {
    server.close();
  }
});

test('pilot: the record page never falls back to the browser recogniser', async () => {
  const { server, base } = await serve({ save: async () => 'ok' });
  try {
    const { url } = (await (await fetch(`${base}/record-url?field=whatChanged`)).json()) as any;
    assert.match(await (await fetch(base + url)).text(), /"fallback":false/);
  } finally {
    server.close();
  }
});

test('pilot with STORE=sheet: the photo reaches the sheet, then this server deletes its copy', async () => {
  // A stand-in for the Apps Script web app, including its redirect to the reply.
  const posts: any[] = [];
  const collector = express();
  collector.use(express.json({ limit: '10mb' }));
  collector.post('/exec', (req, res) => {
    posts.push(req.body);
    res.redirect(302, `/echo/${posts.length}`);
  });
  collector.get('/echo/:n', (req, res) => {
    const body = posts[Number(req.params.n) - 1];
    res.json(body.secret !== 'test-secret' ? { ok: false, error: 'wrong secret' } : body.action === 'story' ? { ok: true, id: 'S-9' } : { ok: true });
  });
  const sheet = listen(collector);

  delete process.env.SPEECH_KEY;
  const dataDir = await fs.mkdtemp(path.join(os.tmpdir(), 'bfp-'));
  Object.assign(process.env, { STORE: 'sheet', SHEET_URL: `${sheet.base}/exec`, SHEET_SECRET: 'test-secret' });
  const { server, base } = listen(createPreviewApp({ dataDir }));
  delete process.env.STORE;
  delete process.env.SHEET_URL;
  delete process.env.SHEET_SECRET;
  const sid = 'sheet-' + Math.random();
  const event = async (e: object) =>
    (await (await fetch(`${base}/event`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ sid, event: e }) })).json()) as any;
  try {
    const { url: dialog } = (await (await fetch(`${base}/attach-url`)).json()) as any;
    const t = new URL(dialog, base).searchParams.get('t')!;
    const up = await fetch(`${base}/api/upload?t=${encodeURIComponent(t)}&name=drawing.png&type=image/png`, { method: 'PUT', body: new Uint8Array([1, 2, 3, 4]) });
    const photo = (await up.json()) as { name: string; url: string; contentType: string };

    await event({ type: 'start' });
    await event({ type: 'file', files: [photo], values: STORY });
    await event({ type: 'review', values: STORY });
    const sent = await event({ type: 'send', consent: true });
    assert.equal(sent.sent, true);

    assert.deepEqual(posts.map((p) => p.action), ['story', 'file']);
    assert.equal(posts[0].story.program, 'Camp');
    assert.equal(posts[1].storyId, 'S-9');
    assert.equal(posts[1].name, 'drawing.png');
    assert.deepEqual([...Buffer.from(posts[1].base64, 'base64')], [1, 2, 3, 4]);

    let gone = false;
    for (let i = 0; i < 20 && !gone; i++) {
      gone = (await fetch(base + photo.url)).status === 404;
      if (!gone) await new Promise((r) => setTimeout(r, 25));
    }
    assert.ok(gone, 'the server deletes its copy once the story is saved');
  } finally {
    server.close();
    sheet.server.close();
  }
});
