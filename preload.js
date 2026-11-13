const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('poseCOExcel', {
  getItem: key => ipcRenderer.sendSync('poseco:get', key),

  setItem: (key, value) => {
    const ok = ipcRenderer.sendSync('poseco:set', key, value);
    if (!ok) throw new Error('Écriture Excel impossible');
  },

  removeItem: key => {
    const ok = ipcRenderer.sendSync('poseco:remove', key);
    if (!ok) throw new Error('Suppression Excel impossible');
  },

  getLarge: key => ipcRenderer.invoke('poseco:get-large', key),

  setLarge: (key, value) => ipcRenderer.invoke('poseco:set-large', key),

  getRecords: () => ipcRenderer.invoke('poseco:get-records'),

  getList: () => ipcRenderer.invoke('poseco:get-list')
});
