const fs = require('fs');
const path = require('path');

const prodUrl = process.env.PROD_URL || process.argv[2];
if (!prodUrl) {
  console.error('Usage: node scripts/update-manifest.js <PROD_URL>  or set PROD_URL env var');
  process.exit(1);
}

const manifestPath = path.resolve(__dirname, '..', 'manifest.prod.xml');
let xml = fs.readFileSync(manifestPath, 'utf8');

xml = xml.replace(/https:\/\/[^\"'<>]+\/taskpane\/taskpane.html/g, `${prodUrl.replace(/\/$/, '')}/taskpane/taskpane.html`);
xml = xml.replace(/https:\/\/[^\"'<>]+\/assets\//g, `${prodUrl.replace(/\/$/, '')}/assets/`);

fs.writeFileSync(manifestPath, xml, 'utf8');
console.log('manifest.prod.xml actualizado con:', prodUrl);
