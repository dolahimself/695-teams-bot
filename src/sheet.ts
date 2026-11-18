import type { Draft } from './flow';
import { TEXT_FIELDS, TextField } from './flow';
import type { LocalFile, ReadFile, StoryStore } from './store';

/**
 * Saves stories to a Google Sheet, for a pilot run by the project team.
 *
 * The Sheet has a small Google Apps Script (scripts/google-sheet-collector.gs) deployed as a web app.
 * Each story becomes a row; its photos and voice recordings go into a Drive folder for that story,
 * linked from the row. Every request carries a shared secret the script checks.
 *
 * The story is sent first, then each file on its own, so one large photo cannot sink the whole story.
 * A file that is missing, too big or fails is skipped and logged; the story is still saved.
 */

/** Apps Script accepts large posts, but base64 makes files a third bigger. Phone photos are far smaller. */
export const MAX_FILE_BYTES = 25 * 1024 * 1024;

export interface SheetConfig {
  /** The web app address from Deploy > Web app in Apps Script (ends in /exec). */
  url: string;
  /** Must match the SECRET script property. */
  secret: string;
  /** Reads a photo or recording this server kept (see attach.ts and voice.ts). */
  readLocal?: ReadFile;
}

const secondsTaken = (d: Draft, at: number) => Math.max(0, Math.round((at - d.startedAt) / 1000));
const toBase64 = (b: Buffer) => b.toString('base64');

export class SheetStore implements StoryStore {
  constructor(private cfg: SheetConfig, private fetchFn: typeof fetch = fetch) {}

  /** Posts one request to the script. Apps Script answers with a redirect, which fetch follows to the JSON reply. */
  private async call(body: Record<string, unknown>): Promise<Record<string, unknown>> {
    const res = await this.fetchFn(this.cfg.url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ secret: this.cfg.secret, ...body }),
    });
    if (!res.ok) throw new Error(`Google Sheet call failed: ${res.status}`);
    const json = (await res.json().catch(() => ({}))) as Record<string, unknown>;
    if (json.ok !== true) throw new Error(`Google Sheet refused it: ${String(json.error ?? 'no reply')}`);
    return json;
  }

  /** Our own kept files first; anything else (such as a Teams chat attachment) is fetched. */
  private async read(url: string): Promise<LocalFile | undefined> {
    const local = await this.cfg.readLocal?.(url);
    if (local) return local;
    if (!/^https?:\/\//.test(url)) return undefined;
    try {
      const res = await this.fetchFn(url);
      if (!res.ok) return undefined;
      const name = decodeURIComponent(url.split('?')[0].split('/').pop() || 'file');
      return { bytes: Buffer.from(await res.arrayBuffer()), type: res.headers.get('content-type') ?? 'application/octet-stream', name };
    } catch {
      return undefined;
    }
  }

  private async sendFile(storyId: string, url: string, extra: Record<string, unknown>) {
    try {
      const file = await this.read(url);
      if (!file) return console.warn('story file not found, skipped:', url);
      if (file.bytes.length > MAX_FILE_BYTES) return console.warn('story file over 25 MB, skipped:', file.name);
      await this.call({ action: 'file', storyId, name: file.name, type: file.type, base64: toBase64(file.bytes), ...extra });
    } catch (err) {
      console.warn('story file could not be saved, skipped:', url, err); // the story itself is already in the sheet
    }
  }

  async save(d: Draft, at: number): Promise<string> {
    const { id } = await this.call({
      action: 'story',
      story: {
        submittedAt: new Date(at).toISOString(),
        program: d.program,
        whatHappened: d.whatHappened,
        whatChanged: d.whatChanged,
        durationSeconds: secondsTaken(d, at),
        permissionConfirmed: true, // the card does not let a story through without it
      },
    });
    const storyId = String(id);
    for (const f of d.files) await this.sendFile(storyId, f.url, { kind: 'photo' });
    for (const field of Object.keys(TEXT_FIELDS) as TextField[]) {
      const rec = d.recordings?.[field];
      if (rec) await this.sendFile(storyId, rec.url, { kind: 'recording', field: TEXT_FIELDS[field].label, seconds: rec.seconds });
    }
    return `sheet:${storyId}`;
  }
}
