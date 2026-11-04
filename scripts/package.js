// Builds appPackage/bear-friends-desk.zip with your bot id filled in.
// Usage: BOT_ID=<MicrosoftAppId> node scripts/package.js
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
require('dotenv').config();

const id = process.env.BOT_ID || process.env.MicrosoftAppId;
if (!id) { console.error('Set BOT_ID (or MicrosoftAppId in .env) first.'); process.exit(1); }

const dir = path.join(__dirname, '..', 'appPackage');
const manifest = fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8').replace(/\{\{BOT_ID\}\}/g, id);
const zip = new AdmZip();
zip.addFile('manifest.json', Buffer.from(manifest));
zip.addLocalFile(path.join(dir, 'color.png'));
zip.addLocalFile(path.join(dir, 'outline.png'));
const out = path.join(dir, 'bear-friends-desk.zip');
zip.writeZip(out);
console.log('Wrote', out);
