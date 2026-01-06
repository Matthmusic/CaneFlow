const { contextBridge, ipcRenderer, webUtils } = require('electron')

contextBridge.exposeInMainWorld('api', {
  pickInputFile: () => ipcRenderer.invoke('pick-input-file'),
  pickOutputFile: (inputPath) => ipcRenderer.invoke('pick-output-file', inputPath),
  previewRows: (payload) => ipcRenderer.invoke('preview-rows', payload),
  convertFile: (payload) => ipcRenderer.invoke('convert-file', payload),
  getFilePathFromDrop: (file) => {
    try {
      return webUtils.getPathForFile(file)
    } catch (err) {
      console.error('Failed to get file path:', err)
      return null
    }
  },
  onAppLog: (callback) => {
    const listener = (_event, data) => callback(data)
    ipcRenderer.on('app-log', listener)
    return () => ipcRenderer.removeListener('app-log', listener)
  },
  revealPath: (targetPath) => ipcRenderer.invoke('reveal-path', targetPath),
  windowClose: () => ipcRenderer.invoke('window-close'),
  windowMinimize: () => ipcRenderer.invoke('window-minimize'),
  windowToggleMaximize: () => ipcRenderer.invoke('window-toggle-maximize'),
  checkUpdates: () => ipcRenderer.invoke('check-updates'),
  downloadUpdate: () => ipcRenderer.invoke('download-update'),
  installUpdate: () => ipcRenderer.invoke('install-update'),
  onUpdateEvent: (callback) => {
    const listener = (_event, data) => callback(data)
    ipcRenderer.on('update-event', listener)
    return () => ipcRenderer.removeListener('update-event', listener)
  },
})
