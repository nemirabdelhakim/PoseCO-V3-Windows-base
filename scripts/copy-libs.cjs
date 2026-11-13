
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');

const files = [
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'libs/xlsx.full.min.js'],
  ['node_modules/html5-qrcode/html5-qrcode.min.js', 'libs/html5-qrcode.min.js'],
  ['node_modules/jsbarcode/dist/JsBarcode.all.min.js', 'libs/JsBarcode.all.min.js'],
  ['node_modules/@sec-ant/zxing-wasm/dist/iife/reader/index.js', 'libs/zxing_reader.js'],
  ['node_modules/@sec-ant/zxing-wasm/dist/reader/zxing_reader.wasm', 'libs/zxing_reader.wasm']
];

fs.mkdirSync(path.join(root, 'libs'), { recursive: true });

for (const [src, dest] of files) {
  const from = path.join(root, src);
  const to = path.join(root, dest);

  if (!fs.existsSync(from)) {
    throw new Error('Fichier bibliothèque introuvable : ' + src);
  }

  fs.copyFileSync(from, to);
  console.log('Copié : ' + dest);
}

console.log('Toutes les bibliothèques Pose CO sont installées.');
