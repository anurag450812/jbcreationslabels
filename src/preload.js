const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktopAPI', {
    isAvailable: () => ipcRenderer.invoke('desktop:is-available'),
    selectDirectory: options => ipcRenderer.invoke('desktop:select-directory', options),
    selectFiles: options => ipcRenderer.invoke('desktop:select-files', options),
    selectSaveFile: options => ipcRenderer.invoke('desktop:select-save-file', options),
    scanSourceFiles: payload => ipcRenderer.invoke('desktop:scan-source-files', payload),
    getSettings: () => ipcRenderer.invoke('desktop:get-settings'),
    saveSettings: settings => ipcRenderer.invoke('desktop:save-settings', settings),
    exportSettings: () => ipcRenderer.invoke('desktop:export-settings'),
    importSettings: () => ipcRenderer.invoke('desktop:import-settings'),
    importServiceAccount: () => ipcRenderer.invoke('desktop:import-service-account'),
    runSkuAutomation: payload => ipcRenderer.invoke('desktop:run-sku-automation', payload),
});