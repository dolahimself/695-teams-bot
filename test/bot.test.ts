import test from 'node:test';
import assert from 'node:assert/strict';
import { ActivityTypes, TestAdapter, TurnContext } from 'botbuilder';
import { BearFriendsBot } from '../src/bot';
import type { Draft } from '../src/flow';
import type { StoryStore } from '../src/store';

const PROGRAMS = ['Camp', 'Family support', 'Other'];

function setup() {
  const saved: Draft[] = [];
  const store: StoryStore = { save: async (d) => { saved.push(d); return 'ok'; } };
  const bot = new BearFriendsBot(store, PROGRAMS);
  const adapter = new TestAdapter(async (ctx: TurnContext) => bot.run(ctx));
  const sent: any[] = [];
  const updated: any[] = [];
  const origSend = adapter.sendActivities.bind(adapter);
  adapter.sendActivities = async (ctx, acts) => { sent.push(...acts); return origSend(ctx, acts); };
  adapter.updateActivity = async (_ctx, act) => { updated.push(act); };
  return { adapter, saved, sent, updated };
}

test('Teams: hello opens ONE message with the greeting and the card; card buttons update it in place', async () => {
  const { adapter, saved, sent, updated } = setup();
  await adapter.send('hi').startTest();
  const first = sent.find((a) => a.attachments?.length);
  assert.ok(first, 'a card message is sent');
  assert.match(first.text, /Bear Friends Stories/);
  assert.equal(first.attachments[0].contentType, 'application/vnd.microsoft.card.adaptive');

  const id = 'card-1';
  const tap = (value: object) => adapter.send({ type: ActivityTypes.Message, text: '', value, replyToId: id } as any).startTest();
  await tap({ action: 'review', program: 'Camp', whatHappened: 'A child spoke up at camp.', whatChanged: 'More confident.' });
  assert.equal(updated.length, 1);
  assert.equal(updated[0].id, id);
  assert.match(JSON.stringify(updated[0].attachments[0].content), /Check your story/);
  assert.match(JSON.stringify(updated[0].attachments[0].content), /A child spoke up at camp\./);

  await tap({ action: 'send' });
  assert.equal(saved.length, 1);
  assert.equal(saved[0].program, 'Camp');
  assert.match(JSON.stringify(updated[1].attachments[0].content), /Sent\. The Foundation team will review/);
  assert.match(JSON.stringify(updated[1].attachments[0].content), /A child spoke up at camp\./);
});

test('Teams: a privacy problem updates the same card and keeps what was typed', async () => {
  const { adapter, saved, updated } = setup();
  await adapter.send('hi').startTest();
  await adapter.send({ type: ActivityTypes.Message, text: '', replyToId: 'c', value: { action: 'review', program: 'Camp', whatHappened: 'Her name is Amara', whatChanged: 'ok' } } as any).startTest();
  const body = JSON.stringify(updated[0].attachments[0].content);
  assert.match(body, /may include a name/);
  assert.match(body, /Her name is Amara/);
  assert.equal(saved.length, 0);
});
