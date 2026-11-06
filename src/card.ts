import { CardFactory, MessageFactory } from 'botbuilder';
import { FileKind, GREETING, MAX_FILES, OTHER_LABEL, PRIVACY_NOTE, View, fileKind } from './flow';

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

export function buildCardJson(view: View): Record<string, unknown> {
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
      style: 'expanded',
      value: v.program ?? '',
      choices: view.programs.map((p) => ({ title: p === 'Other' ? OTHER_LABEL : p, value: p })),
    });
    body.push({
      type: 'Input.Text',
      id: 'whatHappened',
      label: 'What happened?',
      placeholder: 'Tell us about the impact moment',
      isMultiline: true,
      maxLength: 1500,
      value: v.whatHappened ?? '',
    });
    body.push({
      type: 'Input.Text',
      id: 'whatChanged',
      label: 'What changed for the child or family?',
      placeholder: 'Tell us a story (4-5 sentences work great).',
      isMultiline: true,
      maxLength: 1000,
      value: v.whatChanged ?? '',
    });
    body.push(attachments(view, true));
    actions.push({ type: 'Action.Submit', title: 'Review story', style: 'positive', data: { action: 'review' } });
  } else {
    body.push(text(view.kind === 'sent' ? 'Your story' : 'Check your story', { size: 'Large', weight: 'Bolder' }));
    body.push(
      view.kind === 'sent'
        ? banner('Sent. The Foundation team will review it before anything is used.', 'good')
        : banner('Nothing is sent until you tap Submit.', 'accent'),
    );
    body.push(fact('Program', v.program ?? ''));
    body.push(fact('What happened', v.whatHappened ?? ''));
    body.push(fact('What changed', v.whatChanged ?? ''));
    body.push(attachments(view, view.kind === 'review'));
    if (view.kind === 'review') {
      actions.push({ type: 'Action.Submit', title: 'Submit', style: 'positive', data: { action: 'send' } });
      actions.push({ type: 'Action.Submit', title: 'Edit', data: { action: 'edit' } });
    } else {
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
export function buildCardActivity(view: View) {
  const card = CardFactory.adaptiveCard(buildCardJson(view));
  return MessageFactory.attachment(card, view.greeting ? GREETING : undefined);
}
