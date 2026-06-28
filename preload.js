const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  getData: () => ipcRenderer.invoke('get-data'),
  saveData: (db) => ipcRenderer.invoke('save-data', db),
  printPDF: (monthName) => ipcRenderer.invoke('print-pdf', monthName),
  backup: () => ipcRenderer.invoke('backup'),
  backupExcel: () => ipcRenderer.invoke('backup-excel'),
  exportExcel: (monthName) => ipcRenderer.invoke('export-excel', monthName),
  exportAllExcel: () => ipcRenderer.invoke('export-all-excel'),
  exportWord: (monthName, blockFilter) => ipcRenderer.invoke('export-word', monthName, blockFilter),
  checkUpdate: () => ipcRenderer.invoke('check-update'),
  getVersion: () => ipcRenderer.invoke('get-version'),
  onUpdateStatus: (cb) => ipcRenderer.on('update-status', (_, msg) => cb(msg))
});
