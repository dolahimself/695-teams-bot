import { CardFactory, MessageFactory } from 'botbuilder';
import { GREETING, PRIVACY_NOTE, View } from './flow';

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
      choices: view.programs.map((p) => ({ title: p, value: p })),
    });
    body.push({
      type: 'Input.Text',
      id: 'whatHappened',
      label: 'What happened?',
      placeholder: 'A few sentences is plenty',
      isMultiline: true,
      maxLength: 1500,
      value: v.whatHappened ?? '',
    });
    body.push({
      type: 'Input.Text',
      id: 'whatChanged',
      label: 'What changed for the child or family?',
      placeholder: 'One or two sentences',
      isMultiline: true,
      maxLength: 1000,
      value: v.whatChanged ?? '',
    });
    actions.push({ type: 'Action.Submit', title: 'Review story', style: 'positive', data: { action: 'review' } });
  } else {
    body.push(text(view.kind === 'sent' ? 'Your story' : 'Check your story', { size: 'Large', weight: 'Bolder' }));
    body.push(
      view.kind === 'sent'
        ? banner('Sent. Laura will review it before anything is used.', 'good')
        : banner('Nothing is sent until you tap Send for review.', 'accent'),
    );
    body.push(fact('Program', v.program ?? ''));
    body.push(fact('What happened', v.whatHappened ?? ''));
    body.push(fact('What changed', v.whatChanged ?? ''));
    body.push(fact('File', view.file ? view.file.name : 'none'));
    if (view.kind === 'review') {
      body.push(
        text('Want to add a photo, audio or video? Attach it in this chat before you send.', {
          isSubtle: true,
          spacing: 'Medium',
        }),
      );
      actions.push({ type: 'Action.Submit', title: 'Send for review', style: 'positive', data: { action: 'send' } });
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
