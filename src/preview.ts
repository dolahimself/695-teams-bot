import express from 'express';
import * as path from 'path';
import { buildCardJson } from './card';
import { createAttach } from './attach';
import { PROGRAMS } from './config';
import { Event, GREETING, Session, isTextField, newSession, step } from './flow';
import { createTranscriber } from './speech';
import { createVoice } from './voice';

/**
 * Browser preview of the exact same flow and the exact same Adaptive Card the bot sends to Teams.
 * It does not need Teams, Azure or SharePoint, and it saves nothing.
 */
export function createPreviewApp() {
  const onVercel = !!process.env.VERCEL;
  const root = onVercel ? process.cwd() : path.join(__dirname, '..', '..');
  const app = express();
  const sessions = new Map<string, Session>();
  // Relative links, so they point at whichever address the preview is served from (localhost or Vercel).
  const attach = createAttach(app, { baseUrl: '', dir: onVercel ? '/tmp/preview-uploads' : path.join(root, 'data', 'preview-uploads') });
  app.get('/attach-url', (_req, res) => res.json({ url: attach.attachUrl('preview') }));
  // With SPEECH_KEY and SPEECH_REGION set, the preview records through the speech service like Teams;
  // without them the recording page falls back to the browser's own speech recognition.
  const voice = createVoice(app, {
    baseUrl: '',
    signer: attach,
    transcriber: createTranscriber(),
    dir: onVercel ? '/tmp/preview-recordings' : path.join(root, 'data', 'preview-recordings'),
  });
  app.get('/record-url', (req, res) => {
    const field = req.query.field;
    if (!isTextField(field)) return res.sendStatus(400);
    res.json({ url: voice.recordUrl('preview', field, req.query.again === '1') });
  });

  app.use(express.json());
  app.use('/assets', express.static(path.join(root, 'assets')));
  app.get('/ac.js', (_req, res) =>
    res.sendFile(path.join(root, 'node_modules', 'adaptivecards', 'dist', 'adaptivecards.min.js')),
  );
  app.get('/', (_req, res) => res.sendFile(path.join(root, 'public', 'preview.html')));

  app.post('/event', (req, res) => {
    const { sid, event } = req.body as { sid: string; event: Event };
    let session = sessions.get(sid);
    if (!session || (event as Event).type === 'start') {
      session = newSession();
      sessions.set(sid, session);
    }
    const out = step(session, event, PROGRAMS);
    res.json({
      newMessage: !!out.newMessage,
      say: out.say ?? null,
      sent: !!out.submit,
      greeting: out.view?.greeting ? GREETING : null,
      card: out.view ? buildCardJson(out.view, { voice: true, assetBase: '' }) : null,
      files: out.view ? out.view.files.map((f) => f.url) : null,
      // So the page can play a recording from the pill (from its own copy when it has one).
      recordings: out.view ? out.view.recordings : null,
    });
  });

  return app;
}
