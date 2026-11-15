import {
  Attachment,
  TaskModuleRequest,
  TaskModuleResponse,
  TeamsActivityHandler,
  TurnContext,
} from 'botbuilder';
import { buildCardActivity } from './card';
import { CarriedStory, Event, FileRef, Out, Session, TextField, Values, isTextField, newSession, parseText, step, viewOf } from './flow';
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
  /** What was typed on the card when the upload dialog opened, so it is not lost. */
  typed?: Values;
}

/** What the bot needs for voice answers (see voice.ts). */
export interface BotVoice {
  /** Address of the recording page for one story box. `again` replaces the last recording. */
  recordUrl(field: TextField, again?: boolean): string;
  /** Address of a small player for a recording made here, or undefined if the address is not one of ours. */
  listenUrl(recordingUrl: string): string | undefined;
  /** Public address the card's microphone and play icons load from. */
  assetBase: string;
}

/**
 * Sessions live in memory: fine for a prototype with a handful of testers.
 * Before real use, swap this Map for ConversationState with Azure Blob or Cosmos storage.
 */
export class BearFriendsBot extends TeamsActivityHandler {
  private slots = new Map<string, Slot>();

  constructor(
    private store: StoryStore,
    private programs: string[],
    /** Gives the address of the upload page for the dialog. Without it the button explains how to attach in the chat. */
    private attachUrl?: () => string,
    /** Voice answers. Without it the card has no microphone buttons. */
    private voice?: BotVoice,
  ) {
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
        } else if (act === 'removeFile') {
          event = {
            type: 'removeFile',
            index: Number(value!.index),
            values: 'whatHappened' in value!
              ? { program: String(value!.program ?? ''), whatHappened: String(value!.whatHappened ?? ''), whatChanged: String(value!.whatChanged ?? '') }
              : undefined,
          };
        } else if (act === 'addMedia') {
          await context.sendActivity('Use the + button next to the message box to attach photos in this chat. They will show on your card.');
          return next();
        } else if (act === 'record') {
          await context.sendActivity('The recorder did not open here. Please type your answer, or use the microphone key on your phone keyboard.');
          return next();
        } else if (act === 'play') {
          await context.sendActivity('Your recording could not be played here. Its words are in the box on your card.');
          return next();
        } else if (act === 'send') {
          // The card sends the permission box as the text "true" or "false".
          event = { type: 'send', story: value!.story as CarriedStory | undefined, consent: value!.consent === 'true' || value!.consent === true };
        } else if (act === 'edit') {
          event = { type: 'edit', story: value!.story as CarriedStory | undefined }; // checked in flow.ts
        } else if (act === 'again') {
          event = { type: act };
        } else {
          event = { type: 'text', text: '' };
        }
      } else {
        const files = toFiles(a.attachments);
        event = files.length
          ? { type: 'file', files }
          : parseText(TurnContext.removeRecipientMention(a) ?? a.text ?? '');
      }

      await this.run_(context, event, fromCard ? a.replyToId : undefined);
      await next();
    });
  }

  /** The worker tapped "Add Photos", the microphone, or a recording's play button or time: open that dialog. */
  protected async handleTeamsTaskModuleFetch(context: TurnContext, req: TaskModuleRequest): Promise<TaskModuleResponse> {
    const slot = this.slot(context);
    const d = (req.data ?? {}) as Record<string, unknown>;
    if (d.action === 'play') {
      const rec = isTextField(d.field) ? slot.session.recordings[d.field] : undefined;
      const url = rec && this.voice?.listenUrl(rec.url);
      if (!url) return { task: { type: 'message', value: 'That recording is no longer available. Tap its time to record again.' } };
      return { task: { type: 'continue', value: { title: 'Your recording', url, fallbackUrl: url, width: 'small', height: 'small' } } };
    }
    slot.typed = 'whatHappened' in d
      ? { program: String(d.program ?? ''), whatHappened: String(d.whatHappened ?? ''), whatChanged: String(d.whatChanged ?? '') }
      : undefined;
    if (d.action === 'record') {
      if (!this.voice || !isTextField(d.field)) {
        return { task: { type: 'message', value: 'Voice recording is not set up. Please type your answer instead.' } };
      }
      const url = this.voice.recordUrl(d.field, d.again === true);
      return { task: { type: 'continue', value: { title: 'Record your answer', url, fallbackUrl: url, width: 'medium', height: 'medium' } } };
    }
    if (!this.attachUrl) {
      return { task: { type: 'message', value: 'Attach photos in the chat with the + button instead.' } };
    }
    const url = this.attachUrl();
    return { task: { type: 'continue', value: { title: 'Add Photos', url, fallbackUrl: url, width: 'medium', height: 'large' } } };
  }

  /** A dialog closed: put the uploaded photos in the collection box, or the recorded words in their story box. */
  protected async handleTeamsTaskModuleSubmit(context: TurnContext, req: TaskModuleRequest): Promise<TaskModuleResponse> {
    const slot = this.slot(context);
    const data = (req.data ?? {}) as { files?: FileRef[]; voice?: { field?: unknown; text?: unknown; audio?: { url: string; seconds: number } } };
    const voice = data.voice;
    const files = (data.files ?? []).filter((f) => f && f.url && f.name);
    if (voice && isTextField(voice.field) && typeof voice.text === 'string' && voice.text.trim()) {
      // The recording is checked in flow.ts; it is only ever played through listenUrl, which accepts our own recordings.
      await this.run_(context, { type: 'voice', field: voice.field, text: voice.text, audio: voice.audio, values: slot.typed }, slot.cardId);
    } else if (files.length) {
      await this.run_(context, { type: 'file', files, values: slot.typed }, slot.cardId);
    }
    slot.typed = undefined;
    return undefined as unknown as TaskModuleResponse; // an empty reply closes the dialog
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
    const activity = buildCardActivity(out.view, { voice: !!this.voice, assetBase: this.voice?.assetBase });
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
