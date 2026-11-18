import express from 'express';
import * as path from 'path';
import { buildCardJson } from './card';
import { createAttach } from './attach';
import { PROGRAMS } from './config';
import { Event, GREETING, Out, Session, isTextField, newSession, step, viewOf } from './flow';
import { createTranscriber } from './speech';
import { StoryStore, createStore } from './store';
import { createVoice } from './voice';

/** Stores that keep stories somewhere real. STORE=file (or none) leaves the preview a demo that saves nothing. */
const COLLECTING_STORES = ['sheet', 'graph', 'flow'];

export const DEFAULT_PILOT_NOTICE =
  'Pilot: the stories you submit here are real and go to the project team for review. Please leave out names and identifying details.';

export interface PreviewOptions {
  /** Where submitted stories go. Defaults to the store named by STORE (sheet, graph or flow); without one nothing is saved. */
  store?: StoryStore;
  /** Folder for uploads and recordings until a story is submitted (default data/, or /tmp on Vercel). */
  dataDir?: string;
}

/**
 * Browser preview of the exact same flow and the exact same Adaptive Card the bot sends to Teams.
 * It does not need Teams, Azure or SharePoint.
 *
 * As a demo it saves nothing. With a store it is a pilot: it says so on screen, and saves each submitted
 * story with its photos and recordings, then deletes its own copies of them.
 */
export function createPreviewApp(opts: PreviewOptions = {}) {
  const env = process.env;
  const onVercel = !!env.VERCEL;
  const root = onVercel ? process.cwd() : path.join(__dirname, '..', '..');
  const app = express();
  const sessions = new Map<string, Session>();
  const data = opts.dataDir ?? (onVercel ? '/tmp' : path.join(root, 'data'));

  // Relative links, so they point at whichever address the preview is served from.
  const attach = createAttach(app, { baseUrl: '', dir: path.join(data, 'preview-uploads') });
  const readLocal = async (url: string) => (await attach.read(url)) ?? (await voice.read(url));
  const store = opts.store ?? (COLLECTING_STORES.includes((env.STORE ?? '').toLowerCase()) ? createStore(env, readLocal) : undefined);
  const collecting = !!store;
  if (collecting && onVercel) console.warn('Collecting stories on Vercel: photos and recordings can be lost between requests. Host the pilot on Render (see README).');
  const reviewTeam = collecting ? env.REVIEW_TEAM || 'The project team' : 'The Foundation team';

  app.get('/attach-url', (_req, res) => res.json({ url: attach.attachUrl('preview') }));
  // With SPEECH_KEY and SPEECH_REGION set, the preview records through the speech service like Teams.
  // Without them a demo falls back to the browser's own speech recognition; a pilot does not, since that
  // sends real people's audio to Google or Apple, so it hides the microphones instead.
  const voice = createVoice(app, {
    baseUrl: '',
    signer: attach,
    transcriber: createTranscriber(),
    dir: path.join(data, 'preview-recordings'),
    fallback: !collecting,
  });
  app.get('/record-url', (req, res) => {
    const field = req.query.field;
    if (!isTextField(field)) return res.sendStatus(400);
    res.json({ url: voice.recordUrl('preview', field, req.query.again === '1') });
  });
  const cardOpts = { voice: collecting ? voice.available : true, assetBase: '', reviewTeam };

  app.use(express.json());
  app.use('/assets', express.static(path.join(root, 'assets')));
  app.get('/ac.js', (_req, res) =>
    res.sendFile(path.join(root, 'node_modules', 'adaptivecards', 'dist', 'adaptivecards.min.js')),
  );
  app.get('/', (_req, res) => res.sendFile(path.join(root, 'public', 'preview.html')));
  // Tells the page whether this is a demo or a pilot collecting real stories, so it can say so.
  app.get('/config', (_req, res) => res.json({ collecting, notice: collecting ? env.PILOT_NOTICE || DEFAULT_PILOT_NOTICE : null, reviewTeam }));

  app.post('/event', async (req, res) => {
    try {
      const { sid, event } = req.body as { sid: string; event: Event };
      let session = sessions.get(sid);
      if (!session || (event as Event).type === 'start') {
        session = newSession();
        sessions.set(sid, session);
      }
      let out: Out = step(session, event, PROGRAMS);
      if (out.submit && store) {
        const draft = out.submit;
        try {
          await store.save(draft, Date.now());
          // Saved elsewhere now, so this server's copies of the photos and recordings can go.
          const recs = Object.values(draft.recordings).flatMap((r) => (r ? [r.url] : []));
          void Promise.all([...draft.files.map((f) => attach.remove(f.url)), ...recs.map((u) => voice.remove(u))]).catch(() => {});
        } catch (err) {
          console.error('save failed', err);
          // Put the story back in review so the worker can tap Submit again. Nothing is lost.
          session.phase = 'review';
          out = { view: viewOf(session, PROGRAMS), say: 'I could not save that just now. Nothing was lost. Please tap Submit again in a minute.' };
        }
      }
      res.json({
        newMessage: !!out.newMessage,
        say: out.say ?? null,
        sent: !!out.submit,
        greeting: out.view?.greeting ? GREETING : null,
        card: out.view ? buildCardJson(out.view, cardOpts) : null,
        files: out.view ? out.view.files.map((f) => f.url) : null,
        // So the page can play a recording from the pill (from its own copy when it has one).
        recordings: out.view ? out.view.recordings : null,
      });
    } catch (err) {
      console.error('preview event failed', err);
      res.status(500).json({ say: 'Something went wrong on my side. Please try again.' });
    }
  });

  app.locals.collecting = collecting;
  return app;
}
