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

  await tap({ action: 'send', consent: 'true' });
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

test('Teams: Edit on a review card still opens the filled-in form after the bot restarted', async () => {
  const first = setup();
  await first.adapter.send('hi').startTest();
  await first.adapter.send({ type: ActivityTypes.Message, text: '', replyToId: 'c', value: { action: 'review', program: 'Camp', whatHappened: 'A child spoke up at camp.', whatChanged: 'More confident.' } } as any).startTest();
  const edit = first.updated[0].attachments[0].content.actions.find((a: any) => a.title === 'Edit').data;

  const after = setup(); // a new bot: nothing in memory
  await after.adapter.send({ type: ActivityTypes.Message, text: '', replyToId: 'c', value: edit } as any).startTest();
  assert.equal(after.updated.length, 1);
  const body = JSON.stringify(after.updated[0].attachments[0].content);
  assert.match(body, /Share a story/);
  assert.match(body, /A child spoke up at camp\./);
});

test('Teams: Submit without the permission box saves nothing and asks on the same card', async () => {
  const { adapter, saved, updated } = setup();
  await adapter.send('hi').startTest();
  const tap = (value: object) => adapter.send({ type: ActivityTypes.Message, text: '', value, replyToId: 'c' } as any).startTest();
  await tap({ action: 'review', program: 'Camp', whatHappened: 'A child spoke up at camp.', whatChanged: 'More confident.' });
  await tap({ action: 'send', consent: 'false' });
  assert.equal(saved.length, 0);
  assert.equal(updated[1].id, 'c');
  assert.match(JSON.stringify(updated[1].attachments[0].content), /tick the permission box/);
  await tap({ action: 'send', consent: 'true' });
  assert.equal(saved.length, 1);
});

test('Teams: the microphone opens the recorder; the words and the recording land on the card; play opens a player', async () => {
  const saved: Draft[] = [];
  const voice = {
    recordUrl: (field: string, again?: boolean) => `https://tunnel.example/record?field=${field}${again ? '&again=1' : ''}`,
    listenUrl: (url: string) => (url.startsWith('https://tunnel.example/recordings/') ? `https://tunnel.example/listen?id=${url.split('/').pop()}` : undefined),
    assetBase: 'https://tunnel.example',
  };
  const bot = new BearFriendsBot({ save: async (d) => { saved.push(d); return 'ok'; } }, PROGRAMS, undefined, voice);
  const sent: any[] = [];
  const updated: any[] = [];
  let fetchData: any = {};
  let fetchRes: any;
  let submitData: any = {};
  const adapter = new TestAdapter(async (ctx: TurnContext) => {
    if (ctx.activity.name === 'task/fetch') fetchRes = await (bot as any).handleTeamsTaskModuleFetch(ctx, { data: fetchData });
    else if (ctx.activity.name === 'task/submit') await (bot as any).handleTeamsTaskModuleSubmit(ctx, { data: submitData });
    else await bot.run(ctx);
  });
  const origSend = adapter.sendActivities.bind(adapter);
  adapter.sendActivities = async (ctx, acts) => { sent.push(...acts); return origSend(ctx, acts); };
  adapter.updateActivity = async (_ctx, act) => { updated.push(act); };
  const fetchTask = async (data: object) => { fetchData = data; await adapter.send({ type: ActivityTypes.Invoke, name: 'task/fetch', value: {} } as any).startTest(); return fetchRes; };
  const lastCard = () => [...sent, ...updated].filter((a) => a.attachments?.length).map((a) => a.attachments[0].content).pop();

  await adapter.send('hi').startTest();
  assert.match(JSON.stringify(lastCard()), /mic-whatHappened/);
  assert.match(JSON.stringify(lastCard()), /https:\/\/tunnel\.example\/assets\/mic\.png/);

  const rec = await fetchTask({ action: 'record', field: 'whatHappened', program: 'Camp', whatHappened: '', whatChanged: 'Typed already' });
  assert.equal(rec.task.value.title, 'Record your answer');
  assert.equal(rec.task.value.url, 'https://tunnel.example/record?field=whatHappened');

  submitData = { voice: { field: 'whatHappened', text: 'A child read to the group.', audio: { url: 'https://tunnel.example/recordings/1111', seconds: 71 } } };
  await adapter.send({ type: ActivityTypes.Invoke, name: 'task/submit', value: {} } as any).startTest();
  const card = lastCard();
  assert.equal(card.body.find((b: any) => b.id === 'whatHappened').value, 'A child read to the group.');
  assert.equal(card.body.find((b: any) => b.id === 'whatChanged').value, 'Typed already');
  assert.equal(card.body.find((b: any) => b.id === 'program').value, 'Camp');
  assert.match(JSON.stringify(card), /"id":"rec-whatHappened"/);
  assert.match(JSON.stringify(card), /"text":"1:11"/);

  const play = await fetchTask({ action: 'play', field: 'whatHappened' });
  assert.equal(play.task.value.title, 'Your recording');
  assert.equal(play.task.value.url, 'https://tunnel.example/listen?id=1111');
  assert.equal((await fetchTask({ action: 'play', field: 'whatChanged' })).task.type, 'message'); // nothing recorded there

  const again = await fetchTask({ action: 'record', field: 'whatHappened', again: true });
  assert.equal(again.task.value.url, 'https://tunnel.example/record?field=whatHappened&again=1');
  assert.equal(saved.length, 0);
});

test('Teams: a recording from somewhere else is never played', async () => {
  const voice = { recordUrl: () => 'x', listenUrl: (_url: string) => undefined, assetBase: '' };
  const bot = new BearFriendsBot({ save: async () => 'ok' }, PROGRAMS, undefined, voice);
  let res: any;
  const adapter = new TestAdapter(async (ctx: TurnContext) => {
    if (ctx.activity.name === 'task/fetch') res = await (bot as any).handleTeamsTaskModuleFetch(ctx, { data: { action: 'play', field: 'whatHappened' } });
    else if (ctx.activity.name === 'task/submit') await (bot as any).handleTeamsTaskModuleSubmit(ctx, { data: { voice: { field: 'whatHappened', text: 'Hi there.', audio: { url: 'https://evil.example/x.mp3', seconds: 3 } } } });
    else await bot.run(ctx);
  });
  adapter.updateActivity = async () => {};
  await adapter.send('hi').startTest();
  await adapter.send({ type: ActivityTypes.Invoke, name: 'task/submit', value: {} } as any).startTest();
  await adapter.send({ type: ActivityTypes.Invoke, name: 'task/fetch', value: {} } as any).startTest();
  assert.equal(res.task.type, 'message');
});

test('Teams: without a speech service the card has no microphones', async () => {
  const { adapter, sent } = setup();
  await adapter.send('hi').startTest();
  const card = JSON.stringify(sent.find((a) => a.attachments?.length).attachments[0].content);
  assert.doesNotMatch(card, /"action":"record"/);
  assert.match(card, /"label":"What happened\?"/);
});
