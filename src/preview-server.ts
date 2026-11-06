import { createPreviewApp } from './preview';

const port = Number(process.env.PREVIEW_PORT ?? 3979);
createPreviewApp().listen(port, () => {
  console.log(`\nBear Friends Stories preview:  http://localhost:${port}\n(plays the same flow and card the Teams bot uses; nothing is saved)\n`);
});
