/**
 * The conversation as a small state machine. No Teams or SharePoint code in here,
 * so the same file runs in Teams, in the local terminal, and in the tests.
 *
 * One live story card: program, what happened and what changed sit on the card
 * together. Nothing the worker has picked or typed ever disappears: the card is
 * updated in place (form -> review -> sent) and keeps showing the story.
 */

export interface FileRef {
  name: string;
  url: string;
  contentType?: string;
}

/** A complete, validated story. This is what gets saved. */
export interface Draft {
  program: string;
  whatHappened: string;
  whatChanged: string;
  file?: FileRef;
  startedAt: number;
}

export interface Values {
  program?: string;
  whatHappened?: string;
  whatChanged?: string;
}

export type Phase = 'idle' | 'form' | 'review' | 'sent';

export interface Session {
  phase: Phase;
  values: Values;
  file?: FileRef;
  startedAt: number;
  error?: string;
  errorField?: keyof Values;
  /** Show the hello line above the very first card. */
  greeting: boolean;
}

/** Everything needed to draw the card. */
export interface View {
  kind: 'form' | 'review' | 'sent';
  values: Values;
  file?: FileRef;
  error?: string;
  errorField?: keyof Values;
  greeting: boolean;
  programs: string[];
}

export type Event =
  | { type: 'start' }
  | { type: 'again' }
  | { type: 'startOver' }
  | { type: 'cancel' }
  | { type: 'help' }
  | { type: 'review'; values: Values }
  | { type: 'edit' }
  | { type: 'send' }
  | { type: 'file'; file: FileRef }
  | { type: 'text'; text: string };

export interface Out {
  /** Draw or update the story card with this. */
  view?: View;
  /** True when a fresh card message should be sent instead of updating the old one. */
  newMessage?: boolean;
  /** A short plain message to send. */
  say?: string;
  /** Set once, when the worker taps Send. The caller saves it. */
  submit?: Draft;
}

export const PRIVACY_NOTE = 'Please leave out names and identifying details.';
export const GREETING = 'Hi, I am the Bear Friends Desk. Share a story in about two minutes. ' + PRIVACY_NOTE;

const REMINDER = 'Your story card is above. Fill it in there, or say "start over".';
const ALREADY_SENT = 'That story was already sent. Tap "Share another story" to start a new one.';

export const newSession = (): Session => ({ phase: 'idle', values: {}, startedAt: 0, greeting: false });

/**
 * A light first check, not a guarantee. It catches the obvious (email, phone,
 * links, postal codes, "my name is ...") so the worker can fix it before it is saved.
 * The real safeguard is the human review of every story before it is used.
 */
export function findIdentifyingDetail(text: string): string | null {
  if (/[\w.+-]+@[\w-]+\.[\w.-]+/.test(text)) return 'an email address';
  if (/(?:\+?\d[\s().-]?){7,}/.test(text)) return 'a phone number';
  if (/https?:\/\/|www\./i.test(text)) return 'a link';
  if (/\b[A-Za-z]\d[A-Za-z][ -]?\d[A-Za-z]\d\b/.test(text)) return 'a postal code';
  if (/\b(?:[Mm]y|[Hh]is|[Hh]er|[Tt]heir|[Oo]ur) name is\s+[A-Z][a-z]+|\b(?:named|called|[Nn]amed|[Cc]alled)\s+[A-Z][a-z]+/.test(text)) return 'a name';
  return null;
}

export function viewOf(s: Session, programs: string[]): View {
  return {
    kind: s.phase === 'review' ? 'review' : s.phase === 'sent' ? 'sent' : 'form',
    values: { ...s.values },
    file: s.file,
    error: s.error,
    errorField: s.errorField,
    greeting: s.greeting && s.phase === 'form',
    programs,
  };
}

function begin(s: Session, now: number, greeting: boolean) {
  s.phase = 'form';
  s.values = {};
  s.file = undefined;
  s.error = undefined;
  s.errorField = undefined;
  s.startedAt = now;
  s.greeting = greeting;
}

const clean = (v?: string) => (v ?? '').trim();

/** Turns a typed message into an event. Tapping the card sends events directly. */
export function parseText(text: string): Event {
  const t = clean(text).toLowerCase().replace(/[.!]+$/, '');
  if (['start over', 'restart', 'new story'].includes(t)) return { type: 'startOver' };
  if (['cancel', 'stop', 'quit', 'exit'].includes(t)) return { type: 'cancel' };
  if (t === 'help') return { type: 'help' };
  return { type: 'text', text };
}

export function step(s: Session, e: Event, programs: string[], now: number = Date.now()): Out {
  const draw = (newMessage = false): Out => ({ view: viewOf(s, programs), newMessage });

  switch (e.type) {
    case 'help':
      return {
        say:
          'I help you share a short story about a child or family, in about two minutes. ' +
          'Nothing is used until Laura has reviewed it. Say "start over" to begin again, or "cancel" to stop.',
      };

    case 'cancel':
      s.phase = 'idle';
      s.values = {};
      s.file = undefined;
      s.error = undefined;
      return { say: 'No problem. Nothing was saved. Message me whenever you are ready.' };

    case 'start':
    case 'text':
      if (s.phase === 'form' || s.phase === 'review') return { say: REMINDER };
      begin(s, now, true);
      return draw(true);

    case 'again':
      begin(s, now, false);
      return draw(true);

    case 'startOver':
      begin(s, now, false);
      return draw(false);

    case 'review': {
      if (s.phase === 'sent') return { say: ALREADY_SENT };
      if (s.phase === 'idle') begin(s, now, false);
      const v = e.values;
      s.values = {
        program: clean(v.program) || undefined,
        whatHappened: clean(v.whatHappened) || undefined,
        whatChanged: clean(v.whatChanged) || undefined,
      };
      s.error = undefined;
      s.errorField = undefined;

      const fail = (msg: string, field: keyof Values): Out => {
        s.phase = 'form';
        s.error = msg;
        s.errorField = field;
        return draw();
      };
      if (!s.values.program || !programs.includes(s.values.program)) return fail('Please choose a program.', 'program');
      if (!s.values.whatHappened) return fail('Please say what happened.', 'whatHappened');
      if (!s.values.whatChanged) return fail('Please say what changed for the child or family.', 'whatChanged');
      for (const field of ['whatHappened', 'whatChanged'] as const) {
        const found = findIdentifyingDetail(s.values[field]!);
        if (found) {
          return fail(
            `That looks like it may include ${found}. Please remove it and try again. Nothing has been saved.`,
            field,
          );
        }
      }
      s.phase = 'review';
      return draw();
    }

    case 'edit':
      if (s.phase === 'sent') return { say: ALREADY_SENT };
      if (s.phase === 'idle') return { say: 'Say hi and I will open a story card for you.' };
      s.phase = 'form';
      s.error = undefined;
      s.errorField = undefined;
      return draw();

    case 'file':
      if (s.phase === 'review') {
        s.file = e.file;
        return draw();
      }
      if (s.phase === 'form') {
        s.file = e.file;
        return { say: 'Got your file. It will show when you review your story.' };
      }
      return { say: 'Open a story card first (say hi), then attach your file.' };

    case 'send': {
      if (s.phase === 'sent') return { say: ALREADY_SENT };
      if (s.phase !== 'review') return { say: 'Please review your story first.' };
      const draft: Draft = {
        program: s.values.program!,
        whatHappened: s.values.whatHappened!,
        whatChanged: s.values.whatChanged!,
        file: s.file,
        startedAt: s.startedAt,
      };
      s.phase = 'sent';
      return { ...draw(), submit: draft };
    }
  }
}
