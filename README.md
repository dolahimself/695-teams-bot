# Bear Friends Desk bot (Microsoft Teams)

A Teams bot that lets a worker share a short, anonymous story in about two minutes, mostly from a phone. Stories go to a private list for Laura to review before anything is used.

**Design:** Figma file "Bear Friends Desk: Teams chat design" (https://www.figma.com/design/nzOCTfqBzR0QQeKisbRWad). Six phone screens and two desktop screens in the Teams look.

## How it works for the worker

1. The first message is the **story card**. The program choice is already on it, so there is no separate step for choosing a program.
2. The worker picks a program and types what happened and what changed. **Everything stays on the card.** Nothing disappears as they go.
3. **Review story** turns the same card into a read-only summary. They can attach a photo, audio or video in the chat, and it shows on the card. **Edit** goes back with everything kept.
4. **Send for review** saves the story. The card stays on screen and says Laura will review it.
5. If the privacy check finds a name, phone number, email, link or postal code, the same card shows a message and keeps everything typed. Nothing is saved.

The card is one Adaptive Card that the bot updates in place. Teams uses the full width on phone and desktop.

## What is in here

| Path | What it does |
|---|---|
| `src/flow.ts` | The conversation as a state machine. No Teams or SharePoint code. Edit the wording here. |
| `src/card.ts` | Builds the Adaptive Card for each state (form, review, sent). |
| `src/bot.ts` | Connects it to Teams: card taps, typed messages, file attachments, in-place updates. |
| `src/store.ts` | Where stories go: `file` (local test), `graph` (SharePoint list), `flow` (Power Automate). |
| `src/local.ts` | Run the same flow in your terminal. |
| `src/preview.ts`, `src/preview-server.ts`, `public/preview.html` | The browser preview (`npm run preview`). |
| `src/index.ts` | The web server Teams talks to (`/api/messages`). |
| `appPackage/` | Teams app manifest and the bear icon. `npm run package` zips them. |
| `assets/bear.png` | The bear image. |
| `test/` | 11 tests: the flow, the card, the privacy check, the Teams layer, and the SharePoint save. |

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

`Program` (single line), `WhatHappened` (multiple lines), `WhatChanged` (multiple lines), `Status` (choice: Unreviewed, Approved, Not usable), `FileLink` (single line), `DurationSeconds` (number).

`DurationSeconds` is how long the worker took, so you can test the "about two minutes" target.

For photos and audio, add a document library and put its drive id in `SHAREPOINT_DRIVE_ID`. Files over 4 MB are skipped (the story still saves).

## 3. Register the bot and connect it to Teams

1. In Teams, open **Developer Portal for Teams** (dev.teams.microsoft.com), then **Tools, Bot management, New bot**. Copy the bot id and create a client secret. (Azure Bot also works.)
2. Copy `.env.example` to `.env` and fill in `MicrosoftAppId`, `MicrosoftAppTenantId`, `MicrosoftAppPassword`.
3. Run the bot: `npm start` (port 3978).
4. Expose it with a tunnel, for example Visual Studio dev tunnels: `devtunnel host -p 3978 --allow-anonymous`. Set the bot's messaging endpoint to `https://<your-tunnel>/api/messages`.
5. Build the app package and upload it: `BOT_ID=<your bot id> npm run package`, then in Teams: Apps, Manage your apps, Upload an app. If upload is blocked, ask your IT admin to allow custom app upload or to approve it. Install it on your phone by opening Teams on the same account.

## 4. Saving to SharePoint

Pick one:

- **`STORE=graph`**: needs an app registration with permission to write to the site (Sites.Selected is best; Sites.ReadWrite.All works but needs an admin to consent). Fill the `GRAPH_*` and `SHAREPOINT_*` values.
- **`STORE=flow`**: the bot posts JSON to a Power Automate flow, and the flow creates the item and sends Laura her approval. This avoids Graph permissions, but the "When an HTTP request is received" trigger can be a premium connector.
- **`STORE=file`**: local only.

## 5. Laura's approval gate

Build it in Power Automate: when an item is created, start an approval for Laura, then copy approved stories to the Approved library and set `Status`. The Figma file shows the approval card (screen 6).

## Known limits (prototype)

- Conversations are held in memory, so a restart loses half-finished stories. Swap in `ConversationState` with Azure storage before real use.
- The name check is a light first filter, not a guarantee. Laura's review is the real safeguard.
- Files attached with the paperclip work. Photos pasted inline in Teams may not be downloadable.
- Updating a card in place is how Teams behaves in one-to-one chats. If an update ever fails, the bot sends a fresh card instead.
- The Teams, tunnel and SharePoint steps have not been tested against a real tenant. The flow, the card, the privacy check, the Teams message handling (with a simulated adapter) and the SharePoint request shape are covered by the tests.
- Program names in `.env` are placeholders. Replace them with CFSGE's real programs.
# 695-teams-bot
