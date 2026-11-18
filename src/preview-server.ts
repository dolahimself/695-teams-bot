import { createPreviewApp } from './preview';

// Render gives the port to listen on in PORT. Locally PORT is the Teams bot's port (3978), so only use it on Render.
const port = Number(process.env.PREVIEW_PORT || (process.env.RENDER ? process.env.PORT : '') || 3979);
const app = createPreviewApp();
app.listen(port, () => {
  console.log(`\nBear Friends Stories preview:  http://localhost:${port}`);
  console.log(app.locals.collecting
    ? `(pilot: submitted stories are saved, STORE=${process.env.STORE})\n`
    : '(plays the same flow and card the Teams bot uses; nothing is saved)\n');
});
