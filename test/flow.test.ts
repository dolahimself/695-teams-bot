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
  assert.equal(choice.style, 'compact'); // a dropdown
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
  const sent = go(s, { type: 'send', consent: true }, 91_000);
  assert.ok(sent.submit);
  assert.equal(sent.submit!.program, 'Family support');
  assert.equal(sent.submit!.startedAt, 1000);
  assert.equal(sent.view!.kind, 'sent');
  assert.match(JSON.stringify(buildCardJson(sent.view!)), /Sent\. The Foundation team will review/);
  assert.match(JSON.stringify(buildCardJson(sent.view!)), new RegExp(HAPPENED)); // story stays on screen
  assert.equal(go(s, { type: 'send', consent: true }).submit, undefined); // double tap does nothing
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
  const sent = go(s, { type: 'send', consent: true });
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
  const id = await store.save({ program: 'Camp', whatHappened: 'a', whatChanged: 'b', files: [], startedAt: 0, recordings: {} }, 60_000);
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
    { program: 'Camp', whatHappened: 'a', whatChanged: 'b', startedAt: 0, recordings: {},
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
  const done = go(s, { type: 'send', consent: true });
  assert.equal(done.submit!.program, 'Other');
});

test('default program list has the programs from research, both "own story" choices, and Other last', () => {
  process.env.PROGRAMS = '';
  delete process.env.PROGRAMS;
  const { PROGRAMS: defaults } = require('../src/config');
  for (const p of ['Season of Sharing', 'Grocery vouchers']) assert.ok(defaults.includes(p), p);
  assert.deepEqual(defaults.slice(-3), ['My own story as a Social Worker', 'My own story working with the agency', 'Other']);
});

test('Edit and Submit still work when the server forgot the session: the review card carries the story', () => {
  const s = newSession();
  go(s, { type: 'start' }, 1000);
  go(s, { type: 'file', files: [f('drawing.png')] });
  const review = go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const card = buildCardJson(review.view!) as any;
  const tap = (title: string) => card.actions.find((a: any) => a.title === title).data;

  // A fresh session stands in for a restart or a new serverless instance.
  const edit = tap('Edit');
  const back = go(newSession(), { type: edit.action, story: edit.story }, 5000);
  assert.equal(back.view!.kind, 'form');
  assert.deepEqual(back.view!.values, { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED });
  assert.deepEqual(back.view!.files.map((x) => x.name), ['drawing.png']);
  assert.match(JSON.stringify(buildCardJson(back.view!)), new RegExp(HAPPENED));

  const send = tap('Submit');
  const sent = go(newSession(), { type: send.action, story: send.story, consent: true }, 9000);
  assert.equal(sent.view!.kind, 'sent');
  assert.equal(sent.submit!.whatChanged, CHANGED);
  assert.equal(sent.submit!.startedAt, 1000);
  assert.deepEqual(sent.submit!.files.map((x) => x.name), ['drawing.png']);
});

test('a carried story is checked again before it is saved, and junk is ignored', () => {
  const story = { values: { program: 'Camp', whatHappened: 'Call me at 780-555-1234 any time', whatChanged: CHANGED }, files: [], startedAt: 0 };
  const out = go(newSession(), { type: 'send', story, consent: true });
  assert.equal(out.submit, undefined);
  assert.equal(out.view!.kind, 'form');
  assert.match(out.view!.error!, /phone number/);

  const junk = go(newSession(), { type: 'edit', story: { values: null, files: 'nope', startedAt: 'x' } } as any);
  assert.equal(junk.view!.kind, 'form');
  assert.deepEqual(junk.view!.files, []);
});

test('a voice recording fills the story box it was made for, keeping everything else typed', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const out = go(s, { type: 'voice', field: 'whatChanged', text: '  They read aloud now.  ', values: { program: 'Camp', whatHappened: 'Half typed', whatChanged: '' } });
  assert.equal(out.view!.kind, 'form');
  assert.deepEqual(out.view!.values, { program: 'Camp', whatHappened: 'Half typed', whatChanged: 'They read aloud now.' });

  const more = go(s, { type: 'voice', field: 'whatChanged', text: 'Every week.', values: out.view!.values });
  assert.equal(more.view!.values.whatChanged, 'They read aloud now.\nEvery week.'); // added after what was there
  assert.equal(more.view!.values.whatHappened, 'Half typed');

  const long = go(s, { type: 'voice', field: 'whatChanged', text: 'x'.repeat(2000) });
  assert.equal(long.view!.values.whatChanged!.length, 1000); // never more than the box holds
  assert.match(long.say!, /cut off/);

  assert.match(go(s, { type: 'voice', field: 'whatHappened', text: '   ' }).say!, /could not hear/);
  assert.equal(go(s, { type: 'voice', field: 'program' as any, text: 'Camp' }).view, undefined); // only the story boxes

  go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  assert.match(go(s, { type: 'voice', field: 'whatHappened', text: 'more' }).say!, /Tap Edit/);
  assert.match(go(newSession(), { type: 'voice', field: 'whatHappened', text: 'hello' }).say!, /Open a story card/);
});

test('Submit needs the permission box: without it nothing is saved and the review card asks for it', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const review = go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const box = (buildCardJson(review.view!) as any).body.find((b: any) => b.type === 'Input.Toggle');
  assert.equal(box.id, 'consent');
  assert.equal(box.value, 'false'); // never ticked for them
  assert.match(box.title, /permission to share this story/);
  assert.match(box.title, /parent or guardian/);
  const form = buildCardJson(go(newSession(), { type: 'start' }).view!, { voice: true }) as any;
  assert.equal(form.body.some((b: any) => b.type === 'Input.Toggle'), false); // only on the review card

  const no = go(s, { type: 'send' });
  assert.equal(no.submit, undefined);
  assert.equal(no.view!.kind, 'review');
  const asked = JSON.stringify(buildCardJson(no.view!));
  assert.match(asked, /tick the permission box/);
  assert.match(asked, new RegExp(HAPPENED)); // the story stays on the card

  assert.equal(go(s, { type: 'send', consent: 'true' } as any).submit, undefined); // only a real tick counts

  const yes = go(s, { type: 'send', consent: true });
  assert.ok(yes.submit);
  const sent = JSON.stringify(buildCardJson(yes.view!));
  assert.match(sent, /PERMISSION TO SHARE/);
  assert.doesNotMatch(sent, /tick the permission box/);
});

test('Edit clears the permission box, so a changed story is confirmed again', () => {
  const s = newSession();
  go(s, { type: 'start' });
  go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  s.consent = true; // ticked, then the save failed and the card came back
  assert.equal((buildCardJson(go(s, { type: 'removeFile', index: 0 }).view!) as any).body.find((b: any) => b.type === 'Input.Toggle').value, 'true');
  go(s, { type: 'edit' });
  const back = go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED + ' More.', whatChanged: CHANGED } });
  assert.equal(back.view!.consent, false);
  assert.equal(back.view!.error, undefined);
});

test('a lost session still needs the permission box before a carried story is saved', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const review = go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const send = (buildCardJson(review.view!) as any).actions.find((a: any) => a.title === 'Submit').data;
  const no = go(newSession(), { type: 'send', story: send.story });
  assert.equal(no.submit, undefined);
  assert.match(no.view!.error!, /tick the permission box/);
  assert.ok(go(newSession(), { type: 'send', story: send.story, consent: true }).submit);
});

test('form card: with voice on, each question shares a row with a microphone; never on the review card', () => {
  const s = newSession();
  const view = go(s, { type: 'start' }).view!;
  const plain = buildCardJson(view) as any;
  assert.equal(plain.body.find((b: any) => b.id === 'whatHappened').label, 'What happened?'); // voice off: as before
  assert.doesNotMatch(JSON.stringify(plain), /"action":"record"/);

  const card = buildCardJson(view, { voice: true, assetBase: 'https://tunnel.example' }) as any;
  for (const field of ['whatHappened', 'whatChanged']) {
    const at = card.body.findIndex((b: any) => b.id === field);
    const row = card.body[at - 1];
    assert.equal(row.type, 'ColumnSet'); // the question row sits right above its box
    assert.equal(row.columns[0].items[0].text, field === 'whatHappened' ? 'What happened?' : 'What changed for the child or family?');
    const mic = row.columns[1].items[0];
    assert.equal(mic.id, `mic-${field}`);
    assert.equal(mic.url, 'https://tunnel.example/assets/mic.png');
    assert.deepEqual(mic.selectAction.data, { action: 'record', field, msteams: { type: 'task/fetch' } });
    assert.ok(mic.altText);
    assert.equal(card.body[at].label, undefined); // the question is in the row instead
    assert.ok(card.body[at].placeholder);
  }
  assert.equal(card.body.find((b: any) => b.id === 'whatChanged').maxLength, 1000);

  const review = go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } }).view!;
  assert.doesNotMatch(JSON.stringify(buildCardJson(review, { voice: true })), /"action":"record"/);
});

test('after recording, the microphone becomes a pill: play plays it, the time records again', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const out = go(s, { type: 'voice', field: 'whatHappened', text: 'A child read aloud.', audio: { url: 'https://tunnel.example/recordings/abc', seconds: 71.4 } });
  assert.deepEqual(out.view!.recordings.whatHappened, { url: 'https://tunnel.example/recordings/abc', seconds: 71, text: 'A child read aloud.' });
  const card = buildCardJson(out.view!, { voice: true }) as any;
  const row = card.body[card.body.findIndex((b: any) => b.id === 'whatHappened') - 1];
  const pill = row.columns[1].items[0];
  assert.equal(pill.id, 'rec-whatHappened');
  const [playCol, timeCol] = pill.items[0].columns;
  assert.equal(playCol.items[0].url, '/assets/play.png');
  assert.deepEqual(playCol.items[0].selectAction.data, { action: 'play', field: 'whatHappened', msteams: { type: 'task/fetch' } });
  assert.match(playCol.items[0].selectAction.title, /Play your recording \(1:11\)/);
  assert.equal(timeCol.items[0].text, '1:11');
  assert.deepEqual(timeCol.selectAction.data, { action: 'record', field: 'whatHappened', again: true, msteams: { type: 'task/fetch' } });
  // the other box still has its microphone
  assert.equal(card.body[card.body.findIndex((b: any) => b.id === 'whatChanged') - 1].columns[1].items[0].id, 'mic-whatChanged');
});

test('recording again replaces the words the last recording gave, unless the worker changed them', () => {
  const s = newSession();
  go(s, { type: 'start' });
  const audio = (n: number) => ({ url: `/recordings/${n}`, seconds: 10 });
  go(s, { type: 'voice', field: 'whatChanged', text: 'They read aloud nw.', audio: audio(1), values: { program: 'Camp', whatHappened: 'Typed', whatChanged: 'Intro.' } });
  assert.equal(s.values.whatChanged, 'Intro.\nThey read aloud nw.');
  const again = go(s, { type: 'voice', field: 'whatChanged', text: 'They read aloud now.', audio: audio(2), values: { ...s.values } });
  assert.equal(again.view!.values.whatChanged, 'Intro.\nThey read aloud now.'); // replaced, not added twice
  assert.equal(again.view!.recordings.whatChanged!.url, '/recordings/2');

  const edited = go(s, { type: 'voice', field: 'whatChanged', text: 'Every week.', audio: audio(3), values: { ...s.values, whatChanged: 'Intro. They read aloud now, always.' } });
  assert.equal(edited.view!.values.whatChanged, 'Intro. They read aloud now, always.\nEvery week.'); // edited words are kept
  assert.equal(edited.view!.values.whatHappened, 'Typed');

  const noAudio = go(s, { type: 'voice', field: 'whatChanged', text: 'More.' });
  assert.equal(noAudio.view!.recordings.whatChanged, undefined); // words without a recording: no pill
  assert.equal(go(s, { type: 'voice', field: 'whatHappened', text: 'x', audio: { url: 'javascript:alert(1)'.repeat(40), seconds: 3 } }).view!.recordings.whatHappened, undefined);

  go(s, { type: 'startOver' });
  assert.deepEqual(s.recordings, {});
});

test('a lost session keeps the recordings the review card carried', () => {
  const s = newSession();
  go(s, { type: 'start' });
  go(s, { type: 'voice', field: 'whatHappened', text: HAPPENED, audio: { url: '/recordings/1', seconds: 9 }, values: { program: 'Camp', whatChanged: CHANGED } });
  const review = go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const edit = (buildCardJson(review.view!) as any).actions.find((a: any) => a.title === 'Edit').data;
  const back = go(newSession(), { type: 'edit', story: edit.story });
  assert.deepEqual(back.view!.recordings.whatHappened, { url: '/recordings/1', seconds: 9, text: HAPPENED });
});

test('the submitted story carries its voice recordings, for stores that keep audio', () => {
  const s = newSession();
  go(s, { type: 'start' });
  go(s, { type: 'voice', field: 'whatChanged', text: CHANGED, audio: { url: '/recordings/7', seconds: 30 }, values: { program: 'Camp', whatHappened: HAPPENED } });
  go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const done = go(s, { type: 'send', consent: true });
  assert.deepEqual(done.submit!.recordings, { whatChanged: { url: '/recordings/7', seconds: 30, text: CHANGED } });
});
