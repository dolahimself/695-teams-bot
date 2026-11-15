# Bear Friends Stories bot (Microsoft Teams)

A Teams bot that lets a worker share a short, anonymous story in about two minutes, mostly from a phone. Stories go to a private list for the Foundation team to review before anything is used.

**Design:** Figma file "Bear Friends Desk: Teams chat design" (https://www.figma.com/design/Fea8JyDrKHCHEyuivm50aA). Six phone screens and two desktop screens in the Teams look.

## How it works for the worker

1. The first message is the **story card**. The program choice is already on it, so there is no separate step for choosing a program.
2. The worker picks a program from a dropdown and types what happened and what changed, or taps the **microphone** beside either question to speak the answer (see Voice below). **Everything stays on the card.** Nothing disappears as they go.
3. **Review story** turns the same card into a read-only summary. They can attach photos in the chat, and it shows on the card. **Edit** goes back with everything kept.
4. Before **Submit**, the worker ticks the **Permission to share** box: the child, family or person the story is about (or their parent or guardian) agreed that the Foundation can use it. Submit does nothing without it; the card asks them to tick the box. Going back with **Edit** clears it, so a changed story is confirmed again.
5. **Submit** saves the story. The card stays on screen and says the Foundation team will review it.
6. If the privacy check finds a name, phone number, email, link or postal code, the same card shows a message and keeps everything typed. Nothing is saved.

The card is one Adaptive Card that the bot updates in place. Teams uses the full width on phone and desktop.

## What is in here

| Path | What it does |
|---|---|
| `src/flow.ts` | The conversation as a state machine. No Teams or SharePoint code. Edit the wording here. |
| `src/card.ts` | Builds the Adaptive Card for each state (form, review, sent). |
| `src/bot.ts` | Connects it to Teams: card taps, typed messages, file attachments, in-place updates. |
| `src/attach.ts` | The Add Photos dialog page and uploads. |
| `src/voice.ts`, `src/speech.ts` | The recording dialog and player pages, and turning a recording into text with Azure AI Speech. |
| `src/store.ts` | Where stories go: `file` (local test), `graph` (SharePoint list), `flow` (Power Automate). |
| `src/local.ts` | Run the same flow in your terminal. |
| `src/preview.ts`, `src/preview-server.ts`, `public/preview.html` | The browser preview (`npm run preview`). |
| `src/index.ts` | The web server Teams talks to (`/api/messages`). |
| `appPackage/` | Teams app manifest and the bear icon. `npm run package` zips them. |
| `assets/bear.png` | The bear image. |
| `test/` | Tests for the flow, the card, the privacy check, the Teams layer, photo upload, voice, and the SharePoint save. |

## 0. See the UI in your browser (no accounts needed)

```bash
cd bear-friends-desk-bot
npm install
npm run preview
```

Then open **http://localhost:3979**. This plays the same flow and the same card the Teams bot uses, in a phone-sized chat (switch to **Desktop** at the top). Nothing is saved. Press Ctrl+C in the terminal to stop it.

## 1. Try it in the terminal (2 minutes)

```bash
npm install
npm test          # builds and runs the tests
npm run local     # plays the card in the terminal; stories go to data/stories.json
```

In `npm run local`, type: `hi`, then `1` (program), `happened ...`, `changed ...`, `review`, `attach photo.jpg`, `send`.

## 2. Create the SharePoint list (for STORE=graph)

Create a list called **Unreviewed stories** with these columns. Create each one **without spaces first**, then rename the display name if you like (the internal name stays the same):

`Program` (single line), `WhatHappened` (multiple lines), `WhatChanged` (multiple lines), `Status` (choice: Unreviewed, Approved, Not usable), `FileLink` (multiple lines of text, one link per attachment), `DurationSeconds` (number).

`DurationSeconds` is how long the worker took, so you can test the "about two minutes" target.

For photos, add a document library and put its drive id in `SHAREPOINT_DRIVE_ID`. Photos up to 150 MB each are uploaded; anything larger, or any upload that fails, is skipped and the story still saves.

## 3. Register the bot and connect it to Teams

1. In Teams, open **Developer Portal for Teams** (dev.teams.microsoft.com), then **Tools, Bot management, New bot**. Copy the bot id and create a client secret. (Azure Bot also works.)
2. Copy `.env.example` to `.env` and fill in `MicrosoftAppId`, `MicrosoftAppTenantId`, `MicrosoftAppPassword`.
3. Run the bot: `npm start` (port 3978).
4. Expose it with a tunnel, for example Visual Studio dev tunnels: `devtunnel host -p 3978 --allow-anonymous`. Set the bot's messaging endpoint to `https://<your-tunnel>/api/messages`.
5. Build the app package and upload it: `BOT_ID=<your bot id> npm run package`, then in Teams: Apps, Manage your apps, Upload an app. If upload is blocked, ask your IT admin to allow custom app upload or to approve it. Install it on your phone by opening Teams on the same account.

## 4. Saving to SharePoint

Pick one:

- **`STORE=graph`**: needs an app registration with permission to write to the site (Sites.Selected is best; Sites.ReadWrite.All works but needs an admin to consent). Fill the `GRAPH_*` and `SHAREPOINT_*` values.
- **`STORE=flow`**: the bot posts JSON to a Power Automate flow, and the flow creates the item and sends the reviewer an approval request. This avoids Graph permissions, but the "When an HTTP request is received" trigger can be a premium connector.
- **`STORE=file`**: local only.

## 5. Voice: record an answer instead of typing it

Each story box has a **microphone** at the right of its question, where a character count usually sits. It opens a dialog: tap the microphone, speak, tap again to stop. The recording is turned into text, the worker reads and fixes the words, then **Add to story** puts them in that box (after anything already typed). What they typed elsewhere on the card stays.

Once recorded, the microphone becomes a small pill with a **play button** and the recording's **length** (for example 1:11). The play button plays it back; tapping the length opens the recorder again. Recording again replaces the last recording, and replaces its words in the box if the worker has not changed them (otherwise the new words are added at the end). In Teams, play opens a small player, because a card cannot play sound itself; the preview plays it in place.

- **Turn it on:** create an **Azure AI Speech** resource in a Canadian region (for example `canadacentral`, so recordings are processed in Canada), then set `SPEECH_KEY` and `SPEECH_REGION` in `.env`. `SPEECH_LOCALES` sets the language (`en-CA` by default; `en-CA,fr-CA` lets people speak either). Without these, Teams shows no microphone buttons.
- **Privacy:** the recording goes from the dialog to this server and on to the speech service once. So the worker can play it back, this server keeps it in `data/recordings` (`/tmp` on Vercel). Recordings are not sent with the story or saved to SharePoint, and there is no automatic clean-up yet, so clear that folder from time to time (as with `data/uploads`). Only the text the worker accepts goes on the card, and the same privacy check runs on it at Review.
- **Limits:** up to 3 minutes per recording. A box still holds its normal maximum (1,500 or 1,000 characters); anything longer is cut off with a message.
- **Teams:** the microphone and play icons load from `PUBLIC_BASE_URL/assets`, like the dialogs. The dialog needs the microphone, so the app package asks for it (`devicePermissions: media`), and Teams asks the worker once. Recording in a dialog has not been tested on Teams mobile; there, the microphone key on the phone keyboard also types into the card's boxes.
- **Preview:** with the same two settings the preview uses the speech service too. Without them it falls back to the browser's own speech recognition (Chrome, Edge, Safari), so the demo works with no account. That sends audio to the browser maker's service, so it is for demos only and is never used in Teams.

## 6. The Foundation's approval gate

Build it in Power Automate: when an item is created, start an approval for whoever reviews stories at the Foundation (a person or a shared group, so it keeps working when staff change), then copy approved stories to the Approved library and set `Status`. The Figma file shows the approval card (screen 6).

## Known limits (prototype)

- Conversations are held in memory, so a restart loses half-finished stories. Swap in `ConversationState` with Azure storage before real use.
- The name check is a light first filter, not a guarantee. The Foundation's review is the real safeguard.
- Up to 10 photos per story (videos and audio are not accepted). The card has an **Add Photos** button in its "Photos (Reading notes, drawings - anonymized)" box. In Teams it opens a dialog with a real file picker (photo library or camera on a phone); picked files upload, the dialog closes and they appear in the box, each with a Remove button. What you had typed stays. This needs `PUBLIC_BASE_URL` set to your tunnel address in `.env`, and the same address in the app package (`PUBLIC_BASE_URL=https://<tunnel> BOT_ID=<id> npm run package`, which fills `validDomains`). Uploads are kept in `data/uploads` until the story is saved. The preview opens the same dialog as a pop-up.
- If the dialog does not open on someone's phone, they can still attach with the + button in the chat; the bot adds those files too.
- Updating a card in place is how Teams behaves in one-to-one chats. If an update ever fails, the bot sends a fresh card instead.
- The Teams, tunnel and SharePoint steps have not been tested against a real tenant. The flow, the card, the privacy check, the Teams message handling (with a simulated adapter) and the SharePoint request shape are covered by the tests.
- Program names are set in `PROGRAMS` (comma separated). The default list includes Season of Sharing and Grocery vouchers from research; confirm the exact names with CFSGE. If `PROGRAMS` is set in your `.env` or in Vercel's environment variables, it replaces the default list, so update it there too.

## Deploy the preview to Vercel

1. Push the repo to GitHub, then in Vercel choose **Add New → Project** and import it.
2. Set **Application Preset** to **Other** (not Express: the Express preset looks for `src/index.ts`, which is the Teams bot, not the preview).
3. Leave build settings as they are and click **Deploy**.

`api/index.ts` and `vercel.json` do the rest. To record through Azure AI Speech on Vercel, add `SPEECH_KEY` and `SPEECH_REGION` under **Settings, Environment Variables** (recordings stay well under Vercel's 4.5 MB request limit). On Vercel, chat state lives in memory and uploads go to `/tmp`, so state can reset between clicks and uploads are capped near 4.5 MB. For a full demo with media use Render, Railway or Fly.
