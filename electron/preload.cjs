const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('wordDesktop', {
  platform: 'windows',
  isDesktop: true,

  openFile: async () => {
    return await ipcRenderer.invoke('dialog:openFile');
  },

  saveFile: async (options) => {
    return await ipcRenderer.invoke('dialog:saveFile', options);
  },

  exportPdf: async (options) => {
    return await ipcRenderer.invoke('dialog:exportPdf', options);
  },

  setWindowState: (state) => {
    ipcRenderer.send('window:setState', state);
  },

  onMenuCommand: (callback) => {
    const handler = (event, command, payload) => callback(command, payload);
    ipcRenderer.on('menu:command', handler);
    return () => ipcRenderer.removeListener('menu:command', handler);
  }
});
