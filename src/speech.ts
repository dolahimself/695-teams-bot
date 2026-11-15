/**
 * Turns a voice recording into text for the story boxes.
 *
 * Uses Azure AI Speech (fast transcription), which fits the agency's Microsoft setup.
 * Create the Speech resource in a Canadian region (for example canadacentral) so recordings
 * are processed in Canada. This app sends each recording once and never stores it.
 */
export interface Transcriber {
  transcribe(audio: Buffer, contentType: string): Promise<string>;
}

export interface AzureSpeechConfig {
  key: string;
  region: string;
  /** Languages people may speak, such as en-CA. More than one turns on language detection. */
  locales: string[];
}

/** File extension the service expects for each recording format browsers produce. */
const EXT: Record<string, string> = {
  webm: 'webm', ogg: 'ogg', mp4: 'm4a', 'x-m4a': 'm4a', m4a: 'm4a', aac: 'aac', mpeg: 'mp3', wav: 'wav', 'x-wav': 'wav', wave: 'wav',
};

export class AzureSpeech implements Transcriber {
  constructor(private cfg: AzureSpeechConfig, private fetchFn: typeof fetch = fetch) {}

  async transcribe(audio: Buffer, contentType: string): Promise<string> {
    const sub = (contentType.split(';')[0].split('/')[1] ?? '').trim().toLowerCase();
    const form = new FormData();
    form.append('audio', new Blob([new Uint8Array(audio)], { type: contentType }), `recording.${EXT[sub] ?? 'webm'}`);
    form.append('definition', JSON.stringify({ locales: this.cfg.locales }));
    const res = await this.fetchFn(
      `https://${this.cfg.region}.api.cognitive.microsoft.com/speechtotext/transcriptions:transcribe?api-version=2024-11-15`,
      { method: 'POST', headers: { 'Ocp-Apim-Subscription-Key': this.cfg.key }, body: form },
    );
    if (!res.ok) throw new Error(`Speech service failed: ${res.status} ${await res.text()}`);
    const json = (await res.json()) as { combinedPhrases?: { text?: string }[] };
    return (json.combinedPhrases ?? []).map((p) => p.text ?? '').join(' ').replace(/\s+/g, ' ').trim();
  }
}

/** Azure AI Speech when SPEECH_KEY and SPEECH_REGION are set; otherwise no server-side voice typing. */
export function createTranscriber(env: NodeJS.ProcessEnv = process.env): Transcriber | undefined {
  if (!env.SPEECH_KEY || !env.SPEECH_REGION) return undefined;
  const locales = (env.SPEECH_LOCALES || 'en-CA').split(',').map((s) => s.trim()).filter(Boolean);
  return new AzureSpeech({ key: env.SPEECH_KEY, region: env.SPEECH_REGION.trim(), locales });
}
