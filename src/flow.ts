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

/** Most photos and videos one story can carry. */
export const MAX_FILES = 10;

export type FileKind = 'photo' | 'video' | 'audio' | 'file';

/** What to call an attachment on the card. */
export function fileKind(f: FileRef): FileKind {
  const t = (f.contentType ?? '').toLowerCase();
  const n = f.name.toLowerCase();
  if (t.startsWith('image/') || /\.(jpe?g|png|gif|heic|heif|webp|bmp)$/.test(n) || /^(jpe?g|png|gif|heic|webp)$/.test(t)) return 'photo';
  if (t.startsWith('video/') || /\.(mp4|mov|m4v|avi|mkv|webm)$/.test(n) || /^(mp4|mov|m4v|webm)$/.test(t)) return 'video';
  if (t.startsWith('audio/') || /\.(mp3|m4a|wav|aac|ogg)$/.test(n) || /^(mp3|m4a|wav)$/.test(t)) return 'audio';
  return 'file';
}

/** A complete, validated story. This is what gets saved. */
export interface Draft {
  program: string;
  whatHappened: string;
  whatChanged: string;
  files: FileRef[];
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
  files: FileRef[];
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
  files: FileRef[];
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
  /** One or more attachments. `values` is what is typed on the card right now, when the caller can read it. */
  | { type: 'file'; files: FileRef[]; values?: Values }
  | { type: 'removeFile'; index: number; values?: Values }
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

export const newSession = (): Session => ({ phase: 'idle', values: {}, files: [], startedAt: 0, greeting: false });

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
    files: [...s.files],
    error: s.error,
    errorField: s.errorField,
    greeting: s.greeting && s.phase === 'form',
    programs,
  };
}

function begin(s: Session, now: number, greeting: boolean) {
  s.phase = 'form';
  s.values = {};
  s.files = [];
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

/** Keeps what the worker has typed so far, without checking it. */
function keepTyped(s: Session, v: Values) {
  s.values = {
    program: clean(v.program) || undefined,
    whatHappened: v.whatHappened?.trim() ? v.whatHappened : undefined,
    whatChanged: v.whatChanged?.trim() ? v.whatChanged : undefined,
  };
  s.error = undefined;
  s.errorField = undefined;
}

export function step(s: Session, e: Event, programs: string[], now: number = Date.now()): Out {
  const draw = (newMessage = false): Out => ({ view: viewOf(s, programs), newMessage });

  switch (e.type) {
    case 'help':
      return {
        say:
          'I help you share a short story about a child or family, in about two minutes. ' +
          'Nothing is used until the Foundation team has reviewed it. Say "start over" to begin again, or "cancel" to stop.',
      };

    case 'cancel':
      s.phase = 'idle';
      s.values = {};
      s.files = [];
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

    case 'file': {
      if (s.phase !== 'form' && s.phase !== 'review') return { say: 'Open a story card first (say hi), then attach your photos or videos.' };
      if (e.values && s.phase === 'form') keepTyped(s, e.values);
      const have = new Set(s.files.map((f) => f.url));
      let added = 0;
      let skipped = 0;
      for (const f of e.files) {
        if (have.has(f.url)) continue;
        if (s.files.length >= MAX_FILES) { skipped++; continue; }
        s.files.push(f);
        have.add(f.url);
        added++;
      }
      const note = skipped ? `Only ${MAX_FILES} photos or videos fit on one story, so ${skipped} were left out.` : undefined;
      // The card can only be redrawn without losing typed text when we know what was typed.
      if (s.phase === 'review' || e.values) return { ...draw(), say: note };
      return {
        say: note ?? `Added ${added || 'your'} file${added === 1 ? '' : 's'}: ${s.files.length} so far. They will show on your card when you tap Review story.`,
      };
    }

    case 'removeFile': {
      if (s.phase !== 'form' && s.phase !== 'review') return {};
      if (e.values && s.phase === 'form') keepTyped(s, e.values);
      if (e.index >= 0 && e.index < s.files.length) s.files.splice(e.index, 1);
      return draw();
    }

    case 'send': {
      if (s.phase === 'sent') return { say: ALREADY_SENT };
      if (s.phase !== 'review') return { say: 'Please review your story first.' };
      const draft: Draft = {
        program: s.values.program!,
        whatHappened: s.values.whatHappened!,
        whatChanged: s.values.whatChanged!,
        files: [...s.files],
        startedAt: s.startedAt,
      };
      s.phase = 'sent';
      return { ...draw(), submit: draft };
    }
  }
}
