// Builds appPackage/bear-friends-desk.zip with your bot id filled in.
// Usage: BOT_ID=<MicrosoftAppId> PUBLIC_BASE_URL=https://<tunnel> node scripts/package.js
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
require('dotenv').config();

const id = process.env.BOT_ID || process.env.MicrosoftAppId;
if (!id) { console.error('Set BOT_ID (or MicrosoftAppId in .env) first.'); process.exit(1); }

const base = process.env.PUBLIC_BASE_URL || '';
const domain = base.replace(/^https?:\/\//, '').replace(/\/.*$/, '');
if (!domain) { console.error('Set PUBLIC_BASE_URL (your tunnel address) first. The photo and video dialog needs it.'); process.exit(1); }

const dir = path.join(__dirname, '..', 'appPackage');
const manifest = fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8').replace(/\{\{BOT_ID\}\}/g, id).replace(/\{\{DOMAIN\}\}/g, domain);
const zip = new AdmZip();
zip.addFile('manifest.json', Buffer.from(manifest));
zip.addLocalFile(path.join(dir, 'color.png'));
zip.addLocalFile(path.join(dir, 'outline.png'));
const out = path.join(dir, 'bear-friends-desk.zip');
zip.writeZip(out);
console.log('Wrote', out);
