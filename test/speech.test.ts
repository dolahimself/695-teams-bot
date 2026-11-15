import test from 'node:test';
import assert from 'node:assert/strict';
import { AzureSpeech, createTranscriber } from '../src/speech';

test('AzureSpeech: sends the recording to fast transcription in the set region and joins the words', async () => {
  const calls: { url: string; init: any }[] = [];
  const fake = async (url: unknown, init: any) => {
    calls.push({ url: String(url), init });
    return new Response(JSON.stringify({ combinedPhrases: [{ text: 'They read aloud now.' }, { text: '  Every week. ' }] }), { status: 200 });
  };
  const speech = new AzureSpeech({ key: 'k', region: 'canadacentral', locales: ['en-CA'] }, fake as typeof fetch);
  assert.equal(await speech.transcribe(Buffer.from([1, 2, 3]), 'audio/webm;codecs=opus'), 'They read aloud now. Every week.');
  assert.equal(calls[0].url, 'https://canadacentral.api.cognitive.microsoft.com/speechtotext/transcriptions:transcribe?api-version=2024-11-15');
  assert.equal(calls[0].init.method, 'POST');
  assert.equal(calls[0].init.headers['Ocp-Apim-Subscription-Key'], 'k');
  const form = calls[0].init.body as FormData;
  assert.deepEqual(JSON.parse(String(form.get('definition'))), { locales: ['en-CA'] });
  const audio = form.get('audio') as File;
  assert.equal(audio.name, 'recording.webm');
  assert.equal(audio.size, 3);

  await speech.transcribe(Buffer.from([1]), 'audio/mp4');
  assert.equal((calls[1].init.body.get('audio') as File).name, 'recording.m4a'); // Safari records mp4
});

test('AzureSpeech: a service error is raised, so the page can tell the worker', async () => {
  const fake = async () => new Response('bad key', { status: 401 });
  const speech = new AzureSpeech({ key: 'k', region: 'canadacentral', locales: ['en-CA'] }, fake as unknown as typeof fetch);
  await assert.rejects(speech.transcribe(Buffer.from([1]), 'audio/webm'), /401/);
});

test('createTranscriber: off without a key and region; languages come from SPEECH_LOCALES', () => {
  assert.equal(createTranscriber({}), undefined);
  assert.equal(createTranscriber({ SPEECH_KEY: 'k' }), undefined);
  const t = createTranscriber({ SPEECH_KEY: 'k', SPEECH_REGION: 'canadacentral', SPEECH_LOCALES: 'en-CA, fr-CA' });
  assert.ok(t instanceof AzureSpeech);
  assert.deepEqual((t as any).cfg.locales, ['en-CA', 'fr-CA']);
  assert.deepEqual((createTranscriber({ SPEECH_KEY: 'k', SPEECH_REGION: 'canadacentral' }) as any).cfg.locales, ['en-CA']);
});
