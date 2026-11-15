import express from 'express';
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
} from 'botbuilder';
import { BearFriendsBot } from './bot';
import * as path from 'path';
import { createAttach } from './attach';
import { PORT, PROGRAMS, PUBLIC_BASE_URL } from './config';
import { createStore } from './store';
import { createTranscriber } from './speech';
import { createVoice } from './voice';

const auth = new ConfigurationBotFrameworkAuthentication(process.env as Record<string, string>);
const adapter = new CloudAdapter(auth);

adapter.onTurnError = async (context: TurnContext, error: Error) => {
  console.error('turn error', error);
  await context.sendActivity('Something went wrong on my side. Nothing was saved. Please try again.');
};

const app = express();
const attach = createAttach(app, {
  baseUrl: PUBLIC_BASE_URL,
  dir: path.join(process.cwd(), 'data', 'uploads'),
  secret: process.env.MicrosoftAppPassword,
});
const voice = createVoice(app, {
  baseUrl: PUBLIC_BASE_URL,
  signer: attach,
  transcriber: createTranscriber(),
  dir: path.join(process.cwd(), 'data', 'recordings'),
});
const bot = new BearFriendsBot(
  createStore(),
  PROGRAMS,
  () => attach.attachUrl('teams'),
  voice.available
    ? { recordUrl: (field, again) => voice.recordUrl('teams', field, again), listenUrl: voice.listenUrl, assetBase: PUBLIC_BASE_URL }
    : undefined,
);

// The card's microphone and play icons load from here.
app.use('/assets', express.static(path.join(__dirname, '..', '..', 'assets')));
app.use(express.json());
app.get('/health', (_req, res) => res.send('ok'));
app.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

app.listen(PORT, () => {
  console.log(`Bear Friends Stories listening on :${PORT}  (store: ${process.env.STORE ?? 'file'}, voice: ${voice.available ? 'on' : 'off, set SPEECH_KEY and SPEECH_REGION'})`);
});
