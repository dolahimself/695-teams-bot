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
  assert.match(JSON.stringify(buildCardJson(sent.view!)), /Sent\. Laura will review/);
  assert.match(JSON.stringify(buildCardJson(sent.view!)), new RegExp(HAPPENED)); // story stays on screen
  assert.equal(go(s, { type: 'send' }).submit, undefined); // double tap does nothing
  assert.equal(go(s, { type: 'again' }).newMessage, true);
});

test('a file attached during review shows on the card', () => {
  const s = newSession();
  go(s, { type: 'start' });
  go(s, { type: 'review', values: { program: 'Camp', whatHappened: HAPPENED, whatChanged: CHANGED } });
  const out = go(s, { type: 'file', file: { name: 'camp-photo.jpg', url: 'https://example.invalid/p' } });
  assert.equal(out.view!.kind, 'review');
  assert.match(JSON.stringify(buildCardJson(out.view!)), /camp-photo\.jpg/);
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
  const id = await store.save({ program: 'Camp', whatHappened: 'a', whatChanged: 'b', startedAt: 0 }, 60_000);
  assert.equal(id, '42');
  const item = calls.find((c) => c.url.includes('/lists/LIST/items'))!;
  const body = JSON.parse(item.init!.body as string);
  assert.equal(body.fields.Status, 'Unreviewed');
  assert.equal(body.fields.DurationSeconds, 60);
});
