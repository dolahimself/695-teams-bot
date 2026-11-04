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
const bot = new BearFriendsBot(createStore(), PROGRAMS, () => attach.attachUrl('teams'));

app.use(express.json());
app.get('/health', (_req, res) => res.send('ok'));
app.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

app.listen(PORT, () => {
  console.log(`Bear Friends Desk listening on :${PORT}  (store: ${process.env.STORE ?? 'file'})`);
});
