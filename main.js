
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

let win;

const KEY_FILES = {
  poseCO_records_v1: 'PoseCO_Enregistrements.xlsx',
  poseCO_settings_v1: 'PoseCO_Parametres.xlsx',
  poseCO_list_v1: 'PoseCO_Liste_Abonnes.xlsx',
  poseCO_mag_v1: 'PoseCO_Magasin.xlsx'
};

const largeSheet = 'BaseComplete';

// Cache des classeurs Excel en mémoire
const bookCache = new Map();

function dataDir() {
  const portableDir = process.env.PORTABLE_EXECUTABLE_DIR;

  if (portableDir) {
    return path.join(portableDir, 'Donnees');
  }

  return path.join(path.dirname(app.getPath('exe')), 'Donnees');
}

function ensureDir() {
  fs.mkdirSync(dataDir(), { recursive: true });
}

function fileFor(key) {
  const name = KEY_FILES[key];

  if (!name) {
    throw new Error('Clé de stockage inconnue : ' + key);
  }

  ensureDir();
  return path.join(dataDir(), name);
}

console.log('DOSSIER DONNEES :', dataDir());

// ==================================================
// LECTURE DES CLASSEURS EXCEL AVEC CACHE
// ==================================================

function readBook(file) {
  if (!fs.existsSync(file)) {
    const cached = bookCache.get(file);

    if (cached && !cached.exists) {
      return cached.wb;
    }

    const wb = XLSX.utils.book_new();

    bookCache.set(file, {
      exists: false,
      wb: wb
    });

    return wb;
  }

  const stat = fs.statSync(file);
  const cached = bookCache.get(file);

  if (
    cached &&
    cached.exists &&
    cached.mtimeMs === stat.mtimeMs &&
    cached.size === stat.size
  ) {
    return cached.wb;
  }

  const wb = XLSX.readFile(file, {
    cellDates: true
  });

  bookCache.set(file, {
    exists: true,
    mtimeMs: stat.mtimeMs,
    size: stat.size,
    wb: wb
  });

  return wb;
}

// ==================================================
// ECRITURE EXCEL SECURISEE
// ==================================================

function writeBookAtomic(file, wb) {
  ensureDir();

  const temp = file + '.tmp.xlsx';
  const backup = file + '.bak';

  XLSX.writeFile(wb, temp, {
    bookType: 'xlsx',
    compression: true
  });

  let movedOld = false;

  try {
    try {
      if (fs.existsSync(backup + '.replace')) {
        fs.unlinkSync(backup + '.replace');
      }
    } catch (_) {}

    if (fs.existsSync(file)) {
      try {
        fs.copyFileSync(file, backup);
      } catch (_) {}

      fs.renameSync(file, backup + '.replace');
      movedOld = true;
    }

    fs.renameSync(temp, file);

    // Actualiser le cache après écriture
    const stat = fs.statSync(file);

    bookCache.set(file, {
      exists: true,
      mtimeMs: stat.mtimeMs,
      size: stat.size,
      wb: wb
    });

    if (
      movedOld &&
      fs.existsSync(backup + '.replace')
    ) {
      fs.unlinkSync(backup + '.replace');
    }

  } catch (err) {
    try {
      if (
        movedOld &&
        !fs.existsSync(file) &&
        fs.existsSync(backup + '.replace')
      ) {
        fs.renameSync(backup + '.replace', file);
      }
    } catch (_) {}

    try {
      if (fs.existsSync(temp)) {
        fs.unlinkSync(temp);
      }
    } catch (_) {}

    throw err;
  }
}

// ==================================================
// CONVERSION JSON / EXCEL
// ==================================================

function jsonSheet(value) {
  if (Array.isArray(value)) {
    return value.map((v, i) => ({
      Ordre: i + 1,
      Donnee_JSON: JSON.stringify(v)
    }));
  }

  return [{
    Cle: 'valeur',
    Donnee_JSON: JSON.stringify(value)
  }];
}

function readJsonSheet(ws, isArray) {
  if (!ws) {
    return isArray ? [] : null;
  }

  const rows = XLSX.utils.sheet_to_json(ws, {
    defval: ''
  });

  if (isArray) {
    return rows
      .map(r => {
        try {
          return JSON.parse(r.Donnee_JSON);
        } catch (_) {
          return null;
        }
      })
      .filter(v => v !== null);
  }

  if (!rows.length) {
    return null;
  }

  try {
    return JSON.parse(rows[0].Donnee_JSON);
  } catch (_) {
    return null;
  }
}

// ==================================================
// LECTURE DES DONNEES
// ==================================================

function getItem(key) {
  const file = fileFor(key);

  if (!fs.existsSync(file)) {
    return null;
  }

  const wb = readBook(file);

  if (key === 'poseCO_records_v1') {
    const records = readJsonSheet(
      wb.Sheets.Enregistrements,
      true
    );

    console.log('Enregistrements chargés :', records.length);

    return JSON.stringify(records);
  }

  if (key === 'poseCO_mag_v1') {
    return JSON.stringify({
      entrees: readJsonSheet(wb.Sheets.Entrees, true),
      sorties: readJsonSheet(wb.Sheets.Sorties, true)
    });
  }

  if (key === 'poseCO_list_v1') {
    const meta = readJsonSheet(
      wb.Sheets.ListeMeta,
      false
    ) || {};

    const refs = readJsonSheet(
      wb.Sheets.References,
      true
    );

    return JSON.stringify({
      name: meta.name || '',
      date: meta.date || '',
      refs: refs
    });
  }

  if (key === 'poseCO_settings_v1') {
    return JSON.stringify(
      readJsonSheet(wb.Sheets.Parametres, false) || {}
    );
  }

  throw new Error('Clé de stockage inconnue : ' + key);
}

// ==================================================
// ENREGISTREMENT DES DONNEES
// ==================================================

function setItem(key, raw) {
  const file = fileFor(key);

  let value;

  try {
    value = JSON.parse(raw);
  } catch (_) {
    throw new Error('Données JSON invalides pour ' + key);
  }

  const wb = readBook(file);

  if (key === 'poseCO_records_v1') {

    if (!Array.isArray(value)) {
      throw new Error('Liste des enregistrements invalide');
    }

    wb.Sheets.Enregistrements =
      XLSX.utils.json_to_sheet(jsonSheet(value));

    if (!wb.SheetNames.includes('Enregistrements')) {
      wb.SheetNames.push('Enregistrements');
    }

  } else if (key === 'poseCO_mag_v1') {

    wb.Sheets.Entrees = XLSX.utils.json_to_sheet(
      jsonSheet(
        value && Array.isArray(value.entrees)
          ? value.entrees
          : []
      )
    );

    wb.Sheets.Sorties = XLSX.utils.json_to_sheet(
      jsonSheet(
        value && Array.isArray(value.sorties)
          ? value.sorties
          : []
      )
    );

    for (const n of ['Entrees', 'Sorties']) {
      if (!wb.SheetNames.includes(n)) {
        wb.SheetNames.push(n);
      }
    }

  } else if (key === 'poseCO_list_v1') {

    wb.Sheets.ListeMeta = XLSX.utils.json_to_sheet([{
      Cle: 'valeur',
      Donnee_JSON: JSON.stringify({
        name: value.name || '',
        date: value.date || ''
      })
    }]);

    wb.Sheets.References = XLSX.utils.json_to_sheet(
      jsonSheet(
        Array.isArray(value.refs) ? value.refs : []
      )
    );

    for (const n of ['ListeMeta', 'References']) {
      if (!wb.SheetNames.includes(n)) {
        wb.SheetNames.push(n);
      }
    }

  } else if (key === 'poseCO_settings_v1') {

    wb.Sheets.Parametres = XLSX.utils.json_to_sheet([{
      Cle: 'valeur',
      Donnee_JSON: JSON.stringify(value || {})
    }]);

    if (!wb.SheetNames.includes('Parametres')) {
      wb.SheetNames.push('Parametres');
    }

  } else {
    throw new Error('Clé de stockage inconnue : ' + key);
  }

  writeBookAtomic(file, wb);

  console.log('Enregistrement Excel effectué :', key);
}

// ==================================================
// SUPPRESSION DES DONNEES
// ==================================================

function removeItem(key) {
  const file = fileFor(key);

  if (fs.existsSync(file)) {
    fs.unlinkSync(file);
  }

  bookCache.delete(file);
}

// ==================================================
// LECTURE DE LA BASE COMPLETE DES ABONNES
// ==================================================

function getLarge(key) {
  if (key !== 'listFull') {
    return null;
  }

  const file = fileFor('poseCO_list_v1');

  if (!fs.existsSync(file)) {
    return null;
  }

  const wb = readBook(file);
  const ws = wb.Sheets[largeSheet];

  if (!ws) {
    return null;
  }

  const data = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: ''
  });

  if (
    !data.length ||
    data[0][0] !== '__POSECO_META__'
  ) {
    return null;
  }

  let meta = {};

  try {
    meta = JSON.parse(data[0][1] || '{}');
  } catch (_) {}

  if (meta.isNull) {
    return null;
  }

  meta = meta || {};

  return {
    name: meta.name || '',
    date: meta.date || '',
    header: meta.header || [],
    col: meta.col || 0,
    rows: data.slice(1)
  };
}

// ==================================================
// ENREGISTREMENT DE LA BASE COMPLETE
// ==================================================

function setLarge(key, value) {
  if (key !== 'listFull') {
    return;
  }

  const file = fileFor('poseCO_list_v1');
  const wb = readBook(file);

  const meta = {
    isNull: !value,
    name: value && value.name || '',
    date: value && value.date || '',
    header: value && value.header || [],
    col: value && value.col || 0
  };

  const data = [
    ['__POSECO_META__', JSON.stringify(meta)]
  ].concat(
    value && Array.isArray(value.rows)
      ? value.rows
      : []
  );

  wb.Sheets[largeSheet] = XLSX.utils.aoa_to_sheet(data);

  if (!wb.SheetNames.includes(largeSheet)) {
    wb.SheetNames.push(largeSheet);
  }

  writeBookAtomic(file, wb);
}

// ==================================================
// COMMUNICATION AVEC PRELOAD.JS
// ==================================================

ipcMain.on('poseco:get', (event, key) => {
  try {
    event.returnValue = getItem(key);
  } catch (e) {
    console.error('Erreur lecture :', e);
    event.returnValue = null;
  }
});

ipcMain.on('poseco:set', (event, key, value) => {
  try {
    setItem(key, value);
    event.returnValue = true;
  } catch (e) {
    console.error('Erreur écriture :', e);
    event.returnValue = false;
  }
});

ipcMain.on('poseco:remove', (event, key) => {
  try {
    removeItem(key);
    event.returnValue = true;
  } catch (e) {
    console.error('Erreur suppression :', e);
    event.returnValue = false;
  }
});

ipcMain.handle('poseco:get-large', async (_event, key) => {
  return getLarge(key);
});

ipcMain.handle('poseco:set-large', async (_event, key, value) => {
  return setLarge(key, value);
});

// ==================================================
// FENETRE PRINCIPALE
// ==================================================

function createWindow() {
  win = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 1000,
    minHeight: 700,

    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  win.loadFile(
    path.join(__dirname, 'src', 'index.html')
  );

  win.on('closed', () => {
    win = null;
  });
}

// ==================================================
// DEMARRAGE DE L'APPLICATION
// ==================================================

app.whenReady().then(() => {
  ensureDir();
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
