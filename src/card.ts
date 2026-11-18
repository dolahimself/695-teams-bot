import { CardFactory, MessageFactory } from 'botbuilder';
import {
  CONSENT_LABEL, CarriedStory, FileKind, GREETING, MAX_FILES, OTHER_LABEL, PRIVACY_NOTE, Recording, TEXT_FIELDS, TextField, View, fileKind,
} from './flow';

/**
 * Builds the Adaptive Card for each state. One card does the whole job, so
 * everything stays on screen: the program, the story and the review.
 * "msteams.width = Full" lets Teams use the whole width on phones and desktop.
 */

const text = (t: string, extra: Record<string, unknown> = {}) => ({ type: 'TextBlock', text: t, wrap: true, ...extra });

function fact(label: string, value: string) {
  return {
    type: 'Container',
    spacing: 'Small',
    items: [
      text(label.toUpperCase(), { size: 'Small', weight: 'Bolder', isSubtle: true }),
      text(value, { spacing: 'None' }),
    ],
  };
}

function banner(message: string, style: 'attention' | 'good' | 'accent') {
  return { type: 'Container', style, bleed: false, items: [text(message)] };
}

const KIND_LABEL: Record<FileKind, string> = { photo: 'Photo', video: 'Video', audio: 'Audio', file: 'File' };

/** The collection box: every photo attached to this story, with a way to take one off. */
function attachments(view: View, editable: boolean) {
  const n = view.files.length;
  const items: unknown[] = [
    text(n ? `Photos (Reading notes, drawings - anonymized) (${n})` : 'Photos (Reading notes, drawings - anonymized)', { size: 'Small', weight: 'Bolder', isSubtle: true }),
  ];
  view.files.forEach((f, i) => {
    items.push({
      type: 'ColumnSet',
      spacing: 'Small',
      columns: [
        { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [text(KIND_LABEL[fileKind(f)], { size: 'Small', color: 'Accent', weight: 'Bolder' })] },
        { type: 'Column', width: 'stretch', verticalContentAlignment: 'Center', items: [text(f.name)] },
        ...(editable
          ? [{
              type: 'Column',
              width: 'auto',
              items: [{ type: 'ActionSet', actions: [{ type: 'Action.Submit', title: 'Remove', data: { action: 'removeFile', index: i } }] }],
            }]
          : []),
      ],
    });
  });
  if (!n) items.push(text('None yet', { spacing: 'None', isSubtle: true }));
  if (editable && n < MAX_FILES) {
    items.push({
      type: 'ActionSet',
      spacing: 'Small',
      actions: [
        {
          type: 'Action.Submit',
          title: n ? 'Add more photos' : 'Add Photos',
          // Teams opens a dialog with a real file picker (see attach.ts).
          data: { action: 'addMedia', msteams: { type: 'task/fetch' } },
        },
      ],
    });
    items.push(text(`Up to ${MAX_FILES} per story.`, { size: 'Small', isSubtle: true, spacing: 'None' }));
  }
  return { type: 'Container', style: 'emphasis', spacing: 'Medium', items };
}

export interface CardOptions {
  /** Show a microphone beside each story box's question, to speak the answer instead of typing it. */
  voice?: boolean;
  /** Address the card's icons load from (the bot's public address; empty in the preview). */
  assetBase?: string;
  /** Who reviews stories, for the Sent message ("The Foundation team" unless a pilot says otherwise). */
  reviewTeam?: string;
}

export const fmtSeconds = (s: number) => `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;

/** Opens the recording dialog for one story box. Teams shows it as a dialog (see voice.ts). */
const recordAction = (field: TextField, again: boolean) => ({
  type: 'Action.Submit',
  title: again ? 'Record again' : 'Record your answer',
  data: { action: 'record', field, ...(again ? { again: true } : {}), msteams: { type: 'task/fetch' } },
});

/** Before recording: a microphone button where a character count usually sits. */
function micButton(field: TextField, assets: string) {
  return {
    type: 'Image',
    id: `mic-${field}`,
    url: `${assets}/assets/mic.png`,
    width: '32px',
    height: '32px',
    altText: 'Record your answer instead of typing',
    selectAction: recordAction(field, false),
  };
}

/** After recording: the play button plays it; the time opens the recorder to record again. */
function recordingPill(field: TextField, rec: Recording, assets: string) {
  const time = fmtSeconds(rec.seconds);
  return {
    type: 'Container',
    id: `rec-${field}`,
    showBorder: true, // Teams draws the pill outline
    roundedCorners: true,
    items: [{
      type: 'ColumnSet',
      columns: [
        {
          type: 'Column',
          width: 'auto',
          verticalContentAlignment: 'Center',
          items: [{
            type: 'Image',
            url: `${assets}/assets/play.png`,
            width: '24px',
            height: '24px',
            altText: `Play your recording (${time})`,
            // Teams cannot play sound inside a card, so it opens a small player; the preview plays it in place.
            selectAction: { type: 'Action.Submit', title: `Play your recording (${time})`, data: { action: 'play', field, msteams: { type: 'task/fetch' } } },
          }],
        },
        {
          type: 'Column',
          width: 'auto',
          spacing: 'Small',
          verticalContentAlignment: 'Center',
          selectAction: recordAction(field, true),
          items: [text(time, { weight: 'Bolder' })],
        },
      ],
    }],
  };
}

/** One story box. With voice on, its question shares a row with the microphone or the recording. */
function storyBox(field: TextField, value: string | undefined, rec: Recording | undefined, opts: CardOptions): unknown[] {
  const f = TEXT_FIELDS[field];
  const input = { type: 'Input.Text', id: field, placeholder: f.placeholder, isMultiline: true, maxLength: f.max, value: value ?? '' };
  if (!opts.voice) return [{ ...input, label: f.label }];
  const assets = opts.assetBase ?? '';
  return [
    {
      type: 'ColumnSet',
      spacing: 'Medium',
      columns: [
        { type: 'Column', width: 'stretch', verticalContentAlignment: 'Center', items: [text(f.label, { weight: 'Bolder' })] },
        { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [rec ? recordingPill(field, rec, assets) : micButton(field, assets)] },
      ],
    },
    { ...input, spacing: 'Small' },
  ];
}

export function buildCardJson(view: View, opts: CardOptions = {}): Record<string, unknown> {
  const body: unknown[] = [];
  const actions: unknown[] = [];
  const v = view.values;

  if (view.kind === 'form') {
    body.push(text('Share a story', { size: 'Large', weight: 'Bolder' }));
    if (!view.greeting) body.push(text('About two minutes. ' + PRIVACY_NOTE, { isSubtle: true, spacing: 'Small' }));
    if (view.error) body.push(banner(view.error, 'attention'));
    body.push({
      type: 'Input.ChoiceSet',
      id: 'program',
      label: 'Which program is this story from?',
      // A dropdown: the list is long and keeps growing.
      style: 'compact',
      placeholder: 'Choose a program',
      value: v.program ?? '',
      choices: view.programs.map((p) => ({ title: p === 'Other' ? OTHER_LABEL : p, value: p })),
    });
    body.push(...storyBox('whatHappened', v.whatHappened, view.recordings.whatHappened, opts));
    body.push(...storyBox('whatChanged', v.whatChanged, view.recordings.whatChanged, opts));
    body.push(attachments(view, true));
    actions.push({ type: 'Action.Submit', title: 'Review story', style: 'positive', data: { action: 'review' } });
  } else {
    body.push(text(view.kind === 'sent' ? 'Your story' : 'Check your story', { size: 'Large', weight: 'Bolder' }));
    body.push(
      view.kind === 'sent'
        ? banner(`Sent. ${opts.reviewTeam || 'The Foundation team'} will review it before anything is used.`, 'good')
        : banner('Nothing is sent until you tap Submit.', 'accent'),
    );
    body.push(fact('Program', v.program ?? ''));
    body.push(fact('What happened', v.whatHappened ?? ''));
    body.push(fact('What changed', v.whatChanged ?? ''));
    body.push(attachments(view, view.kind === 'review'));
    if (view.kind === 'review') {
      // Shown next to the box, where the worker is looking after tapping Submit.
      if (view.error) body.push(banner(view.error, 'attention'));
      body.push({
        type: 'Input.Toggle',
        id: 'consent',
        label: 'Permission to share',
        title: CONSENT_LABEL,
        value: view.consent ? 'true' : 'false',
        valueOn: 'true',
        valueOff: 'false',
        wrap: true,
        spacing: 'Medium',
      });
      // The buttons carry the story, so they still work if the server forgot the session (see CarriedStory).
      const story: CarriedStory = { values: v, files: view.files, startedAt: view.startedAt, recordings: view.recordings };
      actions.push({ type: 'Action.Submit', title: 'Submit', style: 'positive', data: { action: 'send', story } });
      actions.push({ type: 'Action.Submit', title: 'Edit', data: { action: 'edit', story } });
    } else {
      body.push(fact('Permission to share', 'Confirmed'));
      actions.push({ type: 'Action.Submit', title: 'Share another story', data: { action: 'again' } });
    }
  }

  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    msteams: { width: 'Full' },
    body,
    actions,
  };
}

/** The message that carries the card. The hello line rides along on the first card only. */
export function buildCardActivity(view: View, opts: CardOptions = {}) {
  const card = CardFactory.adaptiveCard(buildCardJson(view, opts));
  return MessageFactory.attachment(card, view.greeting ? GREETING : undefined);
}
