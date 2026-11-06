import test from 'node:test';
import assert from 'node:assert/strict';
import express from 'express';
import { promises as fs } from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { AddressInfo } from 'net';
import { ActivityTypes, TestAdapter, TurnContext } from 'botbuilder';
import { createAttach } from '../src/attach';
import { BearFriendsBot } from '../src/bot';
import type { Draft } from '../src/flow';

async function serve() {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'bfd-'));
  const app = express();
  const holder: { url?: string } = {};
  const attach = createAttach(app, { baseUrl: 'http://placeholder', dir, secret: 'test-secret' });
  const server = app.listen(0);
  const base = `http://localhost:${(server.address() as AddressInfo).port}`;
  holder.url = base;
  return { base, attach, server, dir };
}

test('upload page: needs a valid link, accepts photos only, serves them back', async () => {
  const { base, attach, server } = await serve();
  try {
    const t = attach.token();
    assert.equal((await fetch(`${base}/attach?t=bad`)).status, 403);
    assert.equal((await fetch(`${base}/attach?t=${encodeURIComponent(t)}`)).status, 200);
    assert.equal(attach.valid(attach.token(Date.now() - 3 * 60 * 60 * 1000)), false); // expired

    const bytes = new Uint8Array(50_000).fill(7);
    const up = await fetch(`${base}/api/upload?t=${encodeURIComponent(t)}&name=camp%20day.jpg&type=image/jpeg`, { method: 'PUT', body: bytes });
    assert.equal(up.status, 200);
    const rec = (await up.json()) as { name: string; url: string; contentType: string };
    assert.equal(rec.name, 'camp day.jpg');
    assert.equal(rec.contentType, 'image/jpeg');

    const back = await fetch(rec.url.replace('http://placeholder', base));
    assert.equal(back.status, 200);
    assert.equal(back.headers.get('content-type')?.startsWith('image/jpeg'), true);
    assert.equal((await back.arrayBuffer()).byteLength, 50_000);

    assert.equal((await fetch(`${base}/api/upload?t=nope&name=a.jpg&type=image/jpeg`, { method: 'PUT', body: bytes })).status, 403);
    assert.equal((await fetch(`${base}/api/upload?t=${encodeURIComponent(t)}&name=a.exe&type=application/x-msdownload`, { method: 'PUT', body: bytes })).status, 415);
    assert.equal((await fetch(`${base}/api/upload?t=${encodeURIComponent(t)}&name=clip.mp4&type=video/mp4`, { method: 'PUT', body: bytes })).status, 415); // no videos
    assert.equal((await fetch(`${base}/api/upload?t=${encodeURIComponent(t)}&name=a.m4a&type=audio/mp4`, { method: 'PUT', body: bytes })).status, 415);
  } finally {
    server.close();
  }
});

test('Teams: the button opens the upload dialog; files from the dialog land in the card without losing typed text', async () => {
  const saved: Draft[] = [];
  const bot = new BearFriendsBot({ save: async (d) => { saved.push(d); return 'ok'; } }, ['Camp', 'Other'], () => 'https://tunnel.example/attach?t=abc');
  const sent: any[] = [];
  const updated: any[] = [];
  const typed = { action: 'addMedia', program: 'Camp', whatHappened: 'Half typed', whatChanged: '' };
  const files = [{ name: 'a.jpg', url: 'https://x/a', contentType: 'image/jpeg' }, { name: 'b.jpg', url: 'https://x/b', contentType: 'image/jpeg' }];
  let fetchRes: any;
  const adapter = new TestAdapter(async (ctx: TurnContext) => {
    if (ctx.activity.name === 'task/fetch') fetchRes = await (bot as any).handleTeamsTaskModuleFetch(ctx, { data: typed });
    else if (ctx.activity.name === 'task/submit') await (bot as any).handleTeamsTaskModuleSubmit(ctx, { data: { files } });
    else await bot.run(ctx);
  });
  const origSend = adapter.sendActivities.bind(adapter);
  adapter.sendActivities = async (ctx, acts) => { sent.push(...acts); return origSend(ctx, acts); };
  adapter.updateActivity = async (_ctx, act) => { updated.push(act); };

  await adapter.send('hi').startTest();
  await adapter.send({ type: ActivityTypes.Invoke, name: 'task/fetch', value: {} } as any).startTest();
  assert.equal(fetchRes.task.type, 'continue');
  assert.equal(fetchRes.task.value.url, 'https://tunnel.example/attach?t=abc');

  await adapter.send({ type: ActivityTypes.Invoke, name: 'task/submit', value: {} } as any).startTest();
  const cards = [...sent, ...updated].filter((a) => a.attachments?.length).map((a) => JSON.stringify(a.attachments[0].content));
  const card = cards.find((c) => c.includes('Photos (Reading notes, drawings - anonymized) (2)')) ?? cards[cards.length - 1];
  assert.match(card, /Photos \(Reading notes, drawings - anonymized\) \(2\)/);
  assert.match(card, /a\.jpg/);
  assert.match(card, /Half typed/);
  assert.match(card, /Add more photos/);
  assert.equal(saved.length, 0);
});
