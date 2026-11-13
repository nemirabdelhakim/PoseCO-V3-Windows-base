const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..');
const files = [
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'libs/xlsx.full.min.js'],
  ['node_modules/html5-qrcode/html5-qrcode.min.js', 'libs/html5-qrcode.min.js'],
  ['node_modules/jsbarcode/dist/JsBarcode.all.min.js', 'libs/JsBarcode.all.min.js']
];
for (const [src, dest] of files) {
  const from = path.join(root, src), to = path.join(root, dest);
  if (fs.existsSync(from)) { fs.mkdirSync(path.dirname(to), { recursive: true }); fs.copyFileSync(from, to); }
}
