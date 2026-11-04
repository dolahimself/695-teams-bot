import express from 'express';
import * as path from 'path';
import { buildCardJson } from './card';
import { createAttach } from './attach';
import { PROGRAMS } from './config';
import { Event, GREETING, Session, newSession, step } from './flow';

/**
 * Browser preview of the exact same flow and the exact same Adaptive Card the bot sends to Teams.
 * It does not need Teams, Azure or SharePoint, and it saves nothing.
 */
export function createPreviewApp() {
  const root = path.join(__dirname, '..', '..');
  const app = express();
  const sessions = new Map<string, Session>();
  const port = Number(process.env.PREVIEW_PORT ?? 3979);
  const attach = createAttach(app, { baseUrl: `http://localhost:${port}`, dir: path.join(root, 'data', 'preview-uploads') });
  app.get('/attach-url', (_req, res) => res.json({ url: attach.attachUrl('preview').replace(/^https?:\/\/[^/]+/, '') }));

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
      card: out.view ? buildCardJson(out.view) : null,
      files: out.view ? out.view.files.map((f) => f.url) : null,
    });
  });

  return app;
}
