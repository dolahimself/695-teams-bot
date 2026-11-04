/**
 * Try the whole conversation in your terminal, no Teams or Azure needed.
 *   npm run local
 * It plays the same flow the Teams card uses. Stories are written to data/stories.json.
 *
 * Commands:  hi | 1..5 or a program name | happened <text> | changed <text>
 *            review | edit | send | attach <file name> | again | start over | cancel
 */
import * as readline from 'readline';
import { PROGRAMS } from './config';
import { Event, Out, Values, newSession, parseText, step } from './flow';
import { FileStore } from './store';

const store = new FileStore();
const session = newSession();
const values: Values = {};

function render(out: Out) {
  if (out.say) console.log(`\nBear: ${out.say}`);
  const v = out.view;
  if (!v) return;
  if (v.greeting) console.log('\nBear: Hi, I am the Bear Friends Desk. Share a story in about two minutes. Please leave out names and identifying details.');
  console.log('\n+--------------------------------------------+');
  console.log(v.kind === 'form' ? '| Share a story' : v.kind === 'review' ? '| Check your story' : '| Your story');
  if (v.error) console.log(`| !! ${v.error}`);
  if (v.kind === 'form') {
    v.programs.forEach((p, i) => console.log(`| ${v.values.program === p ? '(o)' : '( )'} ${i + 1}. ${p}`));
  } else {
    console.log(`| Program: ${v.values.program}`);
  }
  console.log(`| What happened: ${v.values.whatHappened ?? ''}`);
  console.log(`| What changed:  ${v.values.whatChanged ?? ''}`);
  if (v.kind !== 'form') console.log(`| File: ${v.file ? v.file.name : 'none'}`);
  console.log(v.kind === 'form' ? '| [review]' : v.kind === 'review' ? '| [send] [edit]' : '| Sent. [again]');
  console.log('+--------------------------------------------+');
}

console.log('Bear Friends Desk (local). Say hi to begin. Ctrl+C to quit.');
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
rl.on('line', async (line) => {
  const [cmd, ...rest] = line.trim().split(/\s+/);
  const arg = rest.join(' ');
  const lower = cmd.toLowerCase();
  let event: Event;

  if (/^\d+$/.test(lower)) {
    values.program = PROGRAMS[Number(lower) - 1];
    return console.log(`(program set to ${values.program})`);
  } else if (PROGRAMS.map((p) => p.toLowerCase()).includes(line.trim().toLowerCase())) {
    values.program = PROGRAMS.find((p) => p.toLowerCase() === line.trim().toLowerCase());
    return console.log(`(program set to ${values.program})`);
  } else if (lower === 'happened') {
    values.whatHappened = arg;
    return console.log('(noted)');
  } else if (lower === 'changed') {
    values.whatChanged = arg;
    return console.log('(noted)');
  } else if (lower === 'review') event = { type: 'review', values: { ...values } };
  else if (lower === 'edit') event = { type: 'edit' };
  else if (lower === 'send') event = { type: 'send' };
  else if (lower === 'again') { Object.keys(values).forEach((k) => delete (values as any)[k]); event = { type: 'again' }; }
  else if (lower === 'attach') event = { type: 'file', file: { name: arg || 'photo.jpg', url: 'file://local' } };
  else event = parseText(line);

  const out = step(session, event, PROGRAMS);
  if (out.submit) console.log(`(saved as ${await store.save(out.submit, Date.now())})`);
  if (event.type === 'startOver') Object.keys(values).forEach((k) => delete (values as any)[k]);
  render(out);
});
