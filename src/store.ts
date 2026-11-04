import { promises as fs } from 'fs';
import * as path from 'path';
import type { Draft } from './flow';

export interface StoryStore {
  /** Saves one submitted story. Returns an id or link for logs. */
  save(draft: Draft, submittedAt: number): Promise<string>;
}

const titleOf = (d: Draft, at: number) =>
  `${d.program} story, ${new Date(at).toISOString().slice(0, 10)}`;

const secondsTaken = (d: Draft, at: number) => Math.max(0, Math.round((at - d.startedAt) / 1000));

/** Local development: appends to data/stories.json. */
export class FileStore implements StoryStore {
  constructor(private file = path.join(process.cwd(), 'data', 'stories.json')) {}

  async save(d: Draft, at: number): Promise<string> {
    await fs.mkdir(path.dirname(this.file), { recursive: true });
    let all: unknown[] = [];
    try {
      all = JSON.parse(await fs.readFile(this.file, 'utf8'));
    } catch {
      /* first story */
    }
    const id = String(all.length + 1);
    all.push({
      id,
      title: titleOf(d, at),
      program: d.program,
      whatHappened: d.whatHappened,
      whatChanged: d.whatChanged,
      file: d.file?.name ?? null,
      status: 'Unreviewed',
      durationSeconds: secondsTaken(d, at),
    });
    await fs.writeFile(this.file, JSON.stringify(all, null, 2));
    return `file:${id}`;
  }
}

export interface GraphConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  siteId: string;
  listId: string;
  /** Optional drive (document library) for photos, audio and video. */
  driveId?: string;
}

type FetchFn = typeof fetch;
const MAX_UPLOAD_BYTES = 4 * 1024 * 1024; // simple upload limit in Graph

/**
 * Writes to the SharePoint list through Microsoft Graph.
 * List columns (internal names, create them without spaces, then rename the display name):
 *   Program, WhatHappened, WhatChanged, Status, FileLink, DurationSeconds
 */
export class GraphStore implements StoryStore {
  private token?: { value: string; expires: number };

  constructor(private cfg: GraphConfig, private fetchFn: FetchFn = fetch) {}

  private async getToken(): Promise<string> {
    if (this.token && this.token.expires > Date.now() + 60_000) return this.token.value;
    const res = await this.fetchFn(
      `https://login.microsoftonline.com/${this.cfg.tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: this.cfg.clientId,
          client_secret: this.cfg.clientSecret,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      },
    );
    if (!res.ok) throw new Error(`Graph token request failed: ${res.status}`);
    const json = (await res.json()) as { access_token: string; expires_in: number };
    this.token = { value: json.access_token, expires: Date.now() + json.expires_in * 1000 };
    return this.token.value;
  }

  private async uploadFile(d: Draft, at: number): Promise<string | undefined> {
    if (!d.file || !this.cfg.driveId) return undefined;
    try {
      const src = await this.fetchFn(d.file.url);
      if (!src.ok) return undefined;
      const bytes = Buffer.from(await src.arrayBuffer());
      if (bytes.length > MAX_UPLOAD_BYTES) return undefined;
      const safe = d.file.name.replace(/[^\w.\- ]+/g, '_');
      const token = await this.getToken();
      const put = await this.fetchFn(
        `https://graph.microsoft.com/v1.0/drives/${this.cfg.driveId}/root:/Story files/${at}-${safe}:/content`,
        { method: 'PUT', headers: { Authorization: `Bearer ${token}` }, body: bytes },
      );
      if (!put.ok) return undefined;
      return ((await put.json()) as { webUrl?: string }).webUrl;
    } catch {
      return undefined; // the story is still saved without the file
    }
  }

  async save(d: Draft, at: number): Promise<string> {
    const fileLink = await this.uploadFile(d, at);
    const token = await this.getToken();
    const res = await this.fetchFn(
      `https://graph.microsoft.com/v1.0/sites/${this.cfg.siteId}/lists/${this.cfg.listId}/items`,
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fields: {
            Title: titleOf(d, at),
            Program: d.program,
            WhatHappened: d.whatHappened,
            WhatChanged: d.whatChanged,
            Status: 'Unreviewed',
            FileLink: fileLink ?? '',
            DurationSeconds: secondsTaken(d, at),
          },
        }),
      },
    );
    if (!res.ok) throw new Error(`SharePoint save failed: ${res.status} ${await res.text()}`);
    return ((await res.json()) as { id: string }).id;
  }
}

/** Posts the story as JSON to a Power Automate flow, which then creates the list item. */
export class FlowStore implements StoryStore {
  constructor(private url: string, private fetchFn: FetchFn = fetch) {}

  async save(d: Draft, at: number): Promise<string> {
    const res = await this.fetchFn(this.url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        title: titleOf(d, at),
        program: d.program,
        whatHappened: d.whatHappened,
        whatChanged: d.whatChanged,
        fileName: d.file?.name ?? '',
        fileUrl: d.file?.url ?? '',
        durationSeconds: secondsTaken(d, at),
      }),
    });
    if (!res.ok) throw new Error(`Flow call failed: ${res.status}`);
    return 'flow:accepted';
  }
}

export function createStore(env: NodeJS.ProcessEnv = process.env): StoryStore {
  const kind = (env.STORE ?? 'file').toLowerCase();
  if (kind === 'graph') {
    const need = (k: string) => {
      const v = env[k];
      if (!v) throw new Error(`Missing ${k} in .env (STORE=graph)`);
      return v;
    };
    return new GraphStore({
      tenantId: need('GRAPH_TENANT_ID'),
      clientId: need('GRAPH_CLIENT_ID'),
      clientSecret: need('GRAPH_CLIENT_SECRET'),
      siteId: need('SHAREPOINT_SITE_ID'),
      listId: need('SHAREPOINT_LIST_ID'),
      driveId: env.SHAREPOINT_DRIVE_ID || undefined,
    });
  }
  if (kind === 'flow') {
    if (!env.FLOW_URL) throw new Error('Missing FLOW_URL in .env (STORE=flow)');
    return new FlowStore(env.FLOW_URL);
  }
  return new FileStore();
}
