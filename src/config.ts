import 'dotenv/config';

export const PROGRAMS: string[] = (
  process.env.PROGRAMS ?? 'Camp,Family support,Youth programs,Community events,My own story as a Social Worker,Other'
)
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

export const PORT = Number(process.env.PORT ?? 3978);

/** Public https address of this server (your dev tunnel). The photo dialog opens a page here. */
export const PUBLIC_BASE_URL = (process.env.PUBLIC_BASE_URL ?? `http://localhost:${PORT}`).replace(/\/$/, '');
