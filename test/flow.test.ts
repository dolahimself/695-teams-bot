import test from 'node:test';
import assert from 'node:assert/strict';
import { findIdentifyingDetail, newSession, parseText, step } from '../src/flow';
import { buildCardJson } from '../src/card';
import { GraphStore } from '../src/store';

const PROGRAMS = ['Camp', 'Family support', 'Other'];
const HAPPENED = 'A child finished a week at camp and spoke to the group for the first time.';
const CHANGED = 'Their family says they are more confident at school.';
const go = (s: ReturnType<typeof newSession>, e: Parameters<typeof step>[1], now = 0) => step(s, e, PROGRAMS, now);

test('first message is the card, with the program choice already on it', () => {
  const s = newSession();
  const out = go(s, parseText('hi'), 1000);
  assert.equal(out.newMessage, true);
  assert.equal(out.view!.kind, 'form');
  assert.equal(out.view!.greeting, true);
  const card = buildCardJson(out.view!) as any;
  const choice = card.body.find((b: any) => b.type === 'Input.ChoiceSet');
  assert.deepEqual(choice.choices.map((c: any) => c.value), PROGRAMS);
  assert.equal(choice.style, 'expanded');
  assert.equal(card.msteams.width, 'Full');
});

test('choosing a program and writing the story keeps everything visible', () => {
  const s = newSession();
  go(s, { type: 'start' }, 1000);
  const out = go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  assert.equal(out.view!.kind, 'review');
  assert.equal(out.newMessage, false); // the same card is updated in place
  const text = JSON.stringify(buildCardJson(out.view!));
  for (const must of ['Camp', HAPPENED, CHANGED]) assert.ok(text.includes(must), `card should still show: ${must}`);
});

test('a privacy problem keeps every value and shows a message', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const out = go(s, { type: 'review', values: { program: 'Camp', whatHappened: 'Her name is Amara, call 519-555-0100', whatChanged: CHANGED } });
  assert.equal(out.view!.kind, 'form');
  assert.match(out.view!.error!, /may include/);
  assert.equal(out.view!.errorField, 'whatHappened');
  const card = buildCardJson(out.view!) as any;
  const happened = card.body.find((b: any) => b.id === 'whatHappened');
  assert.equal(happened.value, 'Her name is Amara, call 519-555-0100'); // not wiped
  assert.equal(card.body.find((b: any) => b.id === 'program').value, 'Camp');
  assert.equal(card.body.find((b: any) => b.id === 'whatChanged').value, CHANGED);
});

test('missing fields are asked for, values kept', () => {
  const s = newSession();
  go(s, { type: 'start' });
  assert.match(go(s, { type: 'review', values: { whatHappened: HAPPENED, whatChanged: CHANGED } }).view!.error!, /choose a program/);
  assert.match(go(s, { type: 'review', values: { program: 'Camp', whatChanged: CHANGED } }).view!.error!, /what happened/);
  assert.match(go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED } }).view!.error!, /what changed/);
  assert.match(go(s, { type: 'review', values: { program: 'Nope', whatHappened: HAPPENED, whatChanged: CHANGED } }).view!.error!, /choose a program/);
});

test('edit goes back to the form with values kept; send submits once', () => {
  const s = newSession();
  go(s, { type: 'start' }, 1000);
  go(s, { type: 'review', values: { program: 'Family support', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const back = go(s, { type: 'edit' });
  assert.equal(back.view!.kind, 'form');
  assert.equal(back.view!.values.program, 'Family support');
  go(s, { type: 'review', values: back.view!.values });
  const sent = go(s, { type: 'send' }, 91_000);
  assert.ok(sent.submit);
  assert.equal(sent.submit!.program, 'Family support');
  assert.equal(sent.submit!.startedAt, 1000);
  assert.equal(sent.view!.kind, 'sent');
  assert.match(JSON.stringify(buildCardJson(sent.view!)), /Sent\. The Foundation team will review/);
  assert.match(JSON.stringify(buildCardJson(sent.view!)), new RegExp(HAPPENED)); // story stays on screen
  assert.equal(go(s, { type: 'send' }).submit, undefined); // double tap does nothing
  assert.equal(go(s, { type: 'again' }).newMessage, true);
});

const f = (name: string) => ({ name, url: `https://example.invalid/${name}` });

test('several photos can be attached and show in the collection box', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const typed = { program: 'Camp', whatHappened: 'Half typed', whatChanged: '' };
  const out = go(s, { type: 'file', files: [f('a.jpg'), f('b.mp4'), f('c.heic')], values: typed });
  assert.equal(out.view!.kind, 'form');
  assert.equal(out.view!.files.length, 2); // the video is left out
  assert.match(out.say!, /Only photos can be added/);
  assert.equal(out.view!.values.whatHappened, 'Half typed'); // typed text survives the redraw
  assert.equal(out.view!.values.program, 'Camp');
  const json = JSON.stringify(buildCardJson(out.view!));
  assert.match(json, /Photos \(Reading notes, drawings - anonymized\) \(2\)/);
  assert.match(json, /a\.jpg/);
  assert.doesNotMatch(json, /Video|VIDEOS/);
  assert.match(json, /Add more photos/);
  // a second batch adds to the first and ignores repeats
  const more = go(s, { type: 'file', files: [f('a.jpg'), f('d.png')], values: typed });
  assert.deepEqual(more.view!.files.map((x) => x.name), ['a.jpg', 'c.heic', 'd.png']);
});

test('attaching in the chat while the form is open does not redraw it (typed text would be lost)', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const out = go(s, { type: 'file', files: [f('a.jpg'), f('b.jpg')] });
  assert.equal(out.view, undefined);
  assert.match(out.say!, /2 photos: 2 so far/);
});

test('files can be removed, attached during review, and are saved with the story', () => {
  const s = newSession();
  go(s, { type: 'start' });
  go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const added = go(s, { type: 'file', files: [f('camp-photo.jpg'), f('drawing.png')] });
  assert.equal(added.view!.kind, 'review');
  assert.match(JSON.stringify(buildCardJson(added.view!)), /camp-photo\.jpg/);
  const removed = go(s, { type: 'removeFile', index: 0 });
  assert.deepEqual(removed.view!.files.map((x) => x.name), ['drawing.png']);
  const sent = go(s, { type: 'send' });
  assert.deepEqual(sent.submit!.files.map((x) => x.name), ['drawing.png']);
  assert.match(JSON.stringify(buildCardJson(sent.view!)), /drawing\.png/);
  assert.doesNotMatch(JSON.stringify(buildCardJson(sent.view!)), /Remove/); // no removing after sending
});

test('no more than ten files per story', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const many = Array.from({ length: 12 }, (_, i) => f(`p${i}.jpg`));
  const out = go(s, { type: 'file', files: many, values: {} });
  assert.equal(out.view!.files.length, 10);
  assert.match(out.say!, /2 were left out/);
});

test('typing while a card is open does not create a second card; start over resets in place', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const typed = go(s, parseText('hello?'));
  assert.equal(typed.view, undefined);
  assert.match(typed.say!, /card is above/);
  const reset = go(s, parseText('start over'));
  assert.equal(reset.view!.kind, 'form');
  assert.equal(reset.newMessage, false);
});

test('detector examples', () => {
  assert.equal(findIdentifyingDetail('write me at a.b@mail.com'), 'an email address');
  assert.equal(findIdentifyingDetail('call 519 555 0100'), 'a phone number');
  assert.equal(findIdentifyingDetail('see https://x.org'), 'a link');
  assert.equal(findIdentifyingDetail('N3T 1A1 is nearby'), 'a postal code');
  assert.equal(findIdentifyingDetail('my name is Dee'), 'a name');
  assert.equal(findIdentifyingDetail('Her name is Amara.'), 'a name');
  assert.equal(findIdentifyingDetail('a boy named Sam joined'), 'a name');
  assert.equal(findIdentifyingDetail('her name is not shared here'), null);
  assert.equal(findIdentifyingDetail('A child made a friend at camp.'), null);
});

test('GraphStore: gets a token, then creates the list item with Unreviewed status', async () => {
  const calls: { url: string; init?: RequestInit }[] = [];
  const fake = (async (url: string, init?: RequestInit) => {
    calls.push({ url, init });
    if (url.includes('oauth2')) return new Response(JSON.stringify({ access_token: 't', expires_in: 3600 }));
    return new Response(JSON.stringify({ id: '42' }), { status: 201 });
  }) as unknown as typeof fetch;
  const store = new GraphStore({ tenantId: 'T', clientId: 'C', clientSecret: 'S', siteId: 'SITE', listId: 'LIST' }, fake);
  const id = await store.save({ program: 'Camp', whatHappened: 'a', whatChanged: 'b', files: [], startedAt: 0 }, 60_000);
  assert.equal(id, '42');
  const item = calls.find((c) => c.url.includes('/lists/LIST/items'))!;
  const body = JSON.parse(item.init!.body as string);
  assert.equal(body.fields.Status, 'Unreviewed');
  assert.equal(body.fields.DurationSeconds, 60);
});

test('GraphStore: uploads photos simply and big videos in chunks, one link per file', async () => {
  const calls: { url: string; method?: string; range?: string; body?: unknown }[] = [];
  const big = new Uint8Array(7 * 1024 * 1024);
  const fake = (async (url: string, init?: RequestInit) => {
    const method = init?.method ?? 'GET';
    calls.push({ url, method, body: init?.body, range: (init?.headers as Record<string, string> | undefined)?.['Content-Range'] });
    if (url.includes('oauth2')) return new Response(JSON.stringify({ access_token: 't', expires_in: 3600 }));
    if (url.startsWith('https://src/photo')) return new Response(new Uint8Array(1000));
    if (url.startsWith('https://src/video')) return new Response(big);
    if (url.includes('createUploadSession')) return new Response(JSON.stringify({ uploadUrl: 'https://upload/session' }));
    if (url === 'https://upload/session') return new Response(JSON.stringify({ webUrl: 'https://sp/video' }));
    if (url.includes(':/content')) return new Response(JSON.stringify({ webUrl: 'https://sp/photo' }));
    return new Response(JSON.stringify({ id: '7' }), { status: 201 });
  }) as unknown as typeof fetch;
  const store = new GraphStore({ tenantId: 'T', clientId: 'C', clientSecret: 'S', siteId: 'S', listId: 'L', driveId: 'D' }, fake);
  await store.save(
    { program: 'Camp', whatHappened: 'a', whatChanged: 'b', startedAt: 0,
      files: [{ name: 'p.jpg', url: 'https://src/photo' }, { name: 'v.mp4', url: 'https://src/video' }] },
    1000,
  );
  assert.equal(calls.filter((c) => c.url === 'https://upload/session').length, 3); // 7 MB in 3.2 MB chunks
  const item = calls.find((c) => c.url.includes('/lists/L/items'))!;
  const links = JSON.parse(item.body as string).fields.FileLink as string;
  assert.equal(links, 'https://sp/photo\nhttps://sp/video');
});

test('card wording: placeholders, Submit button, Other label, no videos', () => {
  const s = newSession();
  const form = go(s, { type: 'start' });
  const j: any = buildCardJson(form.view!);
  const inputs = j.body.filter((b: any) => b.type === 'Input.Text');
  assert.equal(inputs[0].placeholder, 'Tell us about the impact moment');
  assert.equal(inputs[1].placeholder, 'Tell us a story (4-5 sentences work great).');
  const choice = j.body.find((b: any) => b.type === 'Input.ChoiceSet');
  const other = choice.choices.find((c: any) => c.value === 'Other');
  assert.equal(other.title, 'Other (specify what program and what happened)');
  const text = JSON.stringify(j);
  assert.match(text, /Add Photos/);
  assert.doesNotMatch(text, /video/i);
  go(s, { type: 'review', values: { program: 'Other', whatHappened: 'Something happened here today', whatChanged: 'Things are better now for them' } } as any);
  const rev = go(s, { type: 'edit' });
  assert.ok(rev.view);
});

test('review card says Submit, and submitting still saves "Other" as the program', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const r = go(s, { type: 'review', values: { program: 'Other', whatHappened: 'Something happened here today', whatChanged: 'Things are better now for them' } } as any);
  const j = JSON.stringify(buildCardJson(r.view!));
  assert.match(j, /"title":"Submit"/);
  assert.match(j, /tap Submit\./);
  assert.doesNotMatch(j, /Send for review/);
  const done = go(s, { type: 'send' });
  assert.equal(done.submit!.program, 'Other');
});
