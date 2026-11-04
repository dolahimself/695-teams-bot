import 'dotenv/config';

export const PROGRAMS: string[] = (
  process.env.PROGRAMS ?? 'Camp,Family support,Youth programs,Community events,Other'
)
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

export const PORT = Number(process.env.PORT ?? 3978);
