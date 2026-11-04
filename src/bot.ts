import { ActivityHandler, Attachment, TurnContext } from 'botbuilder';
import { buildCardActivity } from './card';
import { Event, FileRef, Out, Session, newSession, parseText, step, viewOf } from './flow';
import type { StoryStore } from './store';

const TEAMS_FILE = 'application/vnd.microsoft.teams.file.download.info';

function toFiles(attachments: Attachment[] | undefined): FileRef[] {
  const out: FileRef[] = [];
  for (const a of attachments ?? []) {
    if (a.contentType === TEAMS_FILE && a.content?.downloadUrl) {
      out.push({ name: a.name ?? 'file', url: a.content.downloadUrl, contentType: a.content.fileType });
    } else if (a.contentUrl && !a.contentType?.startsWith('text/html')) {
      out.push({ name: a.name ?? 'attachment', url: a.contentUrl, contentType: a.contentType });
    }
  }
  return out;
}

interface Slot {
  session: Session;
  /** Id of the live story card, so file attachments can update it in place. */
  cardId?: string;
}

/**
 * Sessions live in memory: fine for a prototype with a handful of testers.
 * Before real use, swap this Map for ConversationState with Azure Blob or Cosmos storage.
 */
export class BearFriendsBot extends ActivityHandler {
  private slots = new Map<string, Slot>();

  constructor(private store: StoryStore, private programs: string[]) {
    super();

    this.onMembersAdded(async (context, next) => {
      for (const m of context.activity.membersAdded ?? []) {
        if (m.id !== context.activity.recipient.id) {
          await this.run_(context, { type: 'start' }, undefined);
        }
      }
      await next();
    });

    this.onMessage(async (context, next) => {
      const a = context.activity;
      const value = a.value as Record<string, unknown> | undefined;
      const fromCard = typeof value?.action === 'string';
      let event: Event;

      if (fromCard) {
        const act = String(value!.action);
        if (act === 'review') {
          event = {
            type: 'review',
            values: {
              program: String(value!.program ?? ''),
              whatHappened: String(value!.whatHappened ?? ''),
              whatChanged: String(value!.whatChanged ?? ''),
            },
          };
        } else if (act === 'edit' || act === 'send' || act === 'again') {
          event = { type: act };
        } else {
          event = { type: 'text', text: '' };
        }
      } else {
        const files = toFiles(a.attachments);
        event = files.length
          ? { type: 'file', file: files[0] }
          : parseText(TurnContext.removeRecipientMention(a) ?? a.text ?? '');
      }

      await this.run_(context, event, fromCard ? a.replyToId : undefined);
      await next();
    });
  }

  private slot(context: TurnContext): Slot {
    const key = context.activity.conversation.id;
    let s = this.slots.get(key);
    if (!s) {
      s = { session: newSession() };
      this.slots.set(key, s);
    }
    return s;
  }

  private async show(context: TurnContext, slot: Slot, out: Out, cardMessageId?: string) {
    if (!out.view) return;
    const activity = buildCardActivity(out.view);
    const target = cardMessageId ?? slot.cardId;
    if (!out.newMessage && target) {
      try {
        await context.updateActivity({ ...activity, id: target });
        slot.cardId = target;
        return;
      } catch (err) {
        console.warn('could not update the card, sending a new one', err);
      }
    }
    const sent = await context.sendActivity(activity);
    slot.cardId = sent?.id;
  }

  /** Runs one event through the flow, then draws the result. */
  private async run_(context: TurnContext, event: Event, cardMessageId: string | undefined) {
    const slot = this.slot(context);
    const out = step(slot.session, event, this.programs);

    if (out.submit) {
      try {
        await this.store.save(out.submit, Date.now());
      } catch (err) {
        console.error('save failed', err);
        // Put the story back in review so the worker can tap Send again. Nothing is lost.
        slot.session.phase = 'review';
        await this.show(context, slot, { view: viewOf(slot.session, this.programs) }, cardMessageId);
        await context.sendActivity('I could not save that just now. Nothing was lost. Please tap Send again in a minute.');
        return;
      }
    }

    await this.show(context, slot, out, cardMessageId);
    if (out.say) await context.sendActivity(out.say);
  }
}
