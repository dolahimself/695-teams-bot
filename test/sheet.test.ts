import test from 'node:test';
import assert from 'node:assert/strict';
import type { Draft } from '../src/flow';
import { MAX_FILE_BYTES, SheetStore } from '../src/sheet';
import { createStore } from '../src/store';

const draft = (over: Partial<Draft> = {}): Draft => ({
  program: 'Camp', whatHappened: 'A child read aloud.', whatChanged: 'More confident.', files: [], startedAt: 1000, recordings: {}, ...over,
});

/** A stand-in for the Apps Script web app. `reply` decides the answer to each POST. */
function fakeSheet(reply = (b: any): object => (b.action === 'story' ? { ok: true, id: 'S-1' } : { ok: true, url: 'https://drive.example/f' })) {
  const calls: { url: string; body?: any }[] = [];
  const fetchFn = async (url: unknown, init?: RequestInit) => {
    if (init?.method === 'POST') {
      const body = JSON.parse(String(init.body));
      calls.push({ url: String(url), body });
      return new Response(JSON.stringify(reply(body)), { status: 200 });
    }
    calls.push({ url: String(url) }); // fetching a file by its address
    return new Response(new Uint8Array([9, 9]), { status: 200, headers: { 'content-type': 'image/jpeg' } });
  };
  return { calls, fetchFn: fetchFn as typeof fetch };
}

test('SheetStore: sends the story with the secret, then each photo and recording to that story', async () => {
  const { calls, fetchFn } = fakeSheet();
  const kept = new Map([
    ['/files/a/one.jpg', { bytes: Buffer.from('photo-1'), type: 'image/jpeg', name: 'one.jpg' }],
    ['/recordings/r1', { bytes: Buffer.from('audio'), type: 'audio/webm', name: 'recording.webm' }],
  ]);
  const store = new SheetStore({ url: 'https://script.google.com/macros/s/X/exec', secret: 'shh', readLocal: async (u) => kept.get(u) }, fetchFn);
  const id = await store.save(draft({
    files: [{ name: 'one.jpg', url: '/files/a/one.jpg', contentType: 'image/jpeg' }],
    recordings: { whatChanged: { url: '/recordings/r1', seconds: 71, text: 'More confident.' } },
  }), 61_000);

  assert.equal(id, 'sheet:S-1');
  assert.equal(calls.length, 3);
  assert.equal(calls[0].url, 'https://script.google.com/macros/s/X/exec');
  const [story, photo, rec] = calls.map((c) => c.body);
  assert.equal(story.secret, 'shh');
  assert.equal(story.action, 'story');
  assert.deepEqual(story.story, {
    submittedAt: new Date(61_000).toISOString(), program: 'Camp', whatHappened: 'A child read aloud.', whatChanged: 'More confident.',
    durationSeconds: 60, permissionConfirmed: true,
  });
  assert.deepEqual({ ...photo, base64: Buffer.from(photo.base64, 'base64').toString() },
    { secret: 'shh', action: 'file', storyId: 'S-1', name: 'one.jpg', type: 'image/jpeg', base64: 'photo-1', kind: 'photo' });
  assert.equal(rec.kind, 'recording');
  assert.equal(rec.storyId, 'S-1');
  assert.equal(rec.field, 'What changed for the child or family?');
  assert.equal(rec.seconds, 71);
  assert.equal(rec.type, 'audio/webm');
  assert.equal(Buffer.from(rec.base64, 'base64').toString(), 'audio');
});

test('SheetStore: a missing, oversized or refused file is skipped and the story still saves', async () => {
  const { calls, fetchFn } = fakeSheet((b) => (b.action === 'story' ? { ok: true, id: 'S-2' } : b.name === 'bad.jpg' ? { ok: false, error: 'drive full' } : { ok: true }));
  const kept = new Map([
    ['/files/b/big.jpg', { bytes: Buffer.alloc(MAX_FILE_BYTES + 1), type: 'image/jpeg', name: 'big.jpg' }],
    ['/files/c/bad.jpg', { bytes: Buffer.from('x'), type: 'image/jpeg', name: 'bad.jpg' }],
    ['/files/d/good.jpg', { bytes: Buffer.from('y'), type: 'image/jpeg', name: 'good.jpg' }],
  ]);
  const quiet = console.warn;
  console.warn = () => {};
  try {
    const store = new SheetStore({ url: 'u', secret: 's', readLocal: async (u) => kept.get(u) }, fetchFn);
    const files = ['/files/gone/x.jpg', '/files/b/big.jpg', '/files/c/bad.jpg', '/files/d/good.jpg'].map((url) => ({ name: 'f', url }));
    assert.equal(await store.save(draft({ files }), 2000), 'sheet:S-2');
  } finally {
    console.warn = quiet;
  }
  assert.deepEqual(calls.map((c) => c.body.name ?? c.body.action), ['story', 'bad.jpg', 'good.jpg']); // missing and oversized never sent
});

test('SheetStore: the story itself failing throws, so the worker can tap Submit again', async () => {
  const refused = fakeSheet(() => ({ ok: false, error: 'wrong secret' }));
  await assert.rejects(new SheetStore({ url: 'u', secret: 'bad' }, refused.fetchFn).save(draft(), 0), /wrong secret/);
  const down = async () => new Response('oops', { status: 500 });
  await assert.rejects(new SheetStore({ url: 'u', secret: 's' }, down as unknown as typeof fetch).save(draft(), 0), /500/);
});

test('SheetStore: a photo from elsewhere (a Teams attachment) is fetched by its address', async () => {
  const { calls, fetchFn } = fakeSheet();
  await new SheetStore({ url: 'u', secret: 's' }, fetchFn).save(draft({ files: [{ name: 'p.jpg', url: 'https://teams.example/files/p.jpg' }] }), 0);
  assert.equal(calls[1].url, 'https://teams.example/files/p.jpg');
  assert.equal(calls[2].body.name, 'p.jpg');
  assert.equal(calls[2].body.type, 'image/jpeg');
});

test('createStore: STORE=sheet needs the address and the secret', () => {
  assert.throws(() => createStore({ STORE: 'sheet' }), /SHEET_URL/);
  assert.throws(() => createStore({ STORE: 'sheet', SHEET_URL: 'u' }), /SHEET_SECRET/);
  assert.ok(createStore({ STORE: 'sheet', SHEET_URL: 'u', SHEET_SECRET: 's' }) instanceof SheetStore);
});
