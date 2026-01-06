const { app, BrowserWindow, ipcMain, dialog, shell, screen, Menu, globalShortcut } = require('electron')
const fs = require('fs')
const path = require('path')
const { autoUpdater } = require('electron-updater')
const { buildPreviewRows, mapSheetRows } = require('./transform.cjs')
const { readSheetRows, writeSheetRows } = require('./excel.cjs')

const LOG_PREFIX = '[CaneFlow]'

const isDev = !!process.env.VITE_DEV_SERVER_URL
let mainWindow = null

if (process.platform === 'win32') {
  app.setAppUserModelId('com.matthmusic.caneflow')
}

function canSendToMainWindow() {
  return !!(mainWindow && !mainWindow.isDestroyed() && mainWindow.webContents)
}

function sendToMainWindow(channel, payload) {
  if (canSendToMainWindow()) {
    mainWindow.webContents.send(channel, payload)
  }
}

function sendLog(level, message, data = {}) {
  console[level](`${LOG_PREFIX} ${message}`, data)
  sendToMainWindow('app-log', { level, message, data, timestamp: new Date().toISOString() })
}

function buildDefaultOutputPath(inputPath) {
  const parsed = path.parse(inputPath)
  return path.join(parsed.dir, `${parsed.name} - MULTIDOC.xlsx`)
}

function normalizeOutputPath(filePath) {
  if (!filePath) return filePath
  const parsed = path.parse(filePath)
  if (!parsed.ext) {
    return path.join(parsed.dir, `${parsed.name}.xlsx`)
  }
  if (parsed.ext.toLowerCase() !== '.xlsx') {
    return path.join(parsed.dir, `${parsed.name}.xlsx`)
  }
  return filePath
}

function resolveIconPath() {
  if (app.isPackaged) {
    const packagedCandidates = [
      path.join(process.resourcesPath, 'assets', 'caneflow.ico'),
      path.join(process.resourcesPath, 'electron', 'caneflow.ico'),
    ]
    return packagedCandidates.find((candidate) => fs.existsSync(candidate)) || packagedCandidates[0]
  }

  const devCandidates = [
    path.join(__dirname, '..', 'electron', 'caneflow.ico'),
    path.join(__dirname, '..', 'img', 'CaneFlow.ico'),
    path.join(__dirname, '..', 'img', 'CaneFlow.png'),
  ]
  return devCandidates.find((candidate) => fs.existsSync(candidate)) || devCandidates[0]
}

function createWindow() {
  const { width: displayW, height: displayH } = screen.getPrimaryDisplay().workAreaSize
  const targetW = Math.max(1100, Math.min(1400, Math.round(displayW * 0.9)))
  const targetH = Math.max(760, Math.min(1100, Math.round(displayH * 0.8)))

  mainWindow = new BrowserWindow({
    width: targetW,
    height: targetH,
    minWidth: 1040,
    minHeight: 720,
    backgroundColor: '#070b17',
    icon: resolveIconPath(),
    title: 'CaneFlow',
    frame: false,
    titleBarStyle: 'hidden',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  if (isDev) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL)
  } else {
    const indexPath = path.join(__dirname, '../dist/index.html')
    mainWindow.loadFile(indexPath)
  }

  mainWindow.on('closed', () => {
    mainWindow = null
  })

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url)
    return { action: 'deny' }
  })
}

function wireAutoUpdater() {
  if (isDev) return
  autoUpdater.autoDownload = false
  autoUpdater.autoInstallOnAppQuit = true

  autoUpdater.on('update-available', (info) => {
    sendToMainWindow('update-event', { type: 'available', info })
  })
  autoUpdater.on('update-not-available', () => {
    sendToMainWindow('update-event', { type: 'not-available' })
  })
  autoUpdater.on('error', (error) => {
    sendToMainWindow('update-event', { type: 'error', message: error.message })
  })
  autoUpdater.on('download-progress', (progress) => {
    sendToMainWindow('update-event', { type: 'progress', progress })
  })
  autoUpdater.on('update-downloaded', (info) => {
    sendToMainWindow('update-event', { type: 'downloaded', info })
  })
}

ipcMain.handle('pick-input-file', async () => {
  const res = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    title: 'Selectionner un carnet de cables',
    filters: [
      { name: 'Excel (.xls/.xlsx)', extensions: ['xls', 'xlsx'] },
      { name: 'All files', extensions: ['*'] },
    ],
  })
  if (res.canceled || res.filePaths.length === 0) return null
  return res.filePaths[0]
})

ipcMain.handle('pick-output-file', async (_event, inputPath) => {
  let dialogOptions = {
    title: 'Enregistrer le fichier Multidoc',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  }

  if (inputPath) {
    const defaultPath = buildDefaultOutputPath(inputPath)
    dialogOptions.defaultPath = defaultPath
  }

  const res = await dialog.showSaveDialog(mainWindow, dialogOptions)
  if (res.canceled || !res.filePath) return null
  return res.filePath
})

ipcMain.handle('preview-rows', async (_event, payload) => {
  const inputPath = payload?.inputPath
  if (!inputPath) {
    throw new Error('inputPath is required')
  }
  sendLog('log', 'preview-rows start', { inputPath, sheetName: payload?.sheetName || '' })
  const sheetRows = await readSheetRows(inputPath, payload?.sheetName)
  const preview = buildPreviewRows(sheetRows)
  sendLog('log', 'preview-rows done', { count: preview.length })
  return preview
})

ipcMain.handle('convert-file', async (_event, payload) => {
  const inputPath = payload?.inputPath
  if (!inputPath) {
    throw new Error('inputPath is required')
  }

  sendLog('log', 'convert-file start', {
    inputPath,
    outputPath: payload?.outputPath || '',
    unitPricesByType: payload?.unitPricesByType ? Object.keys(payload.unitPricesByType).length : 0,
    unitPrices: Array.isArray(payload?.unitPrices) ? payload.unitPrices.length : 0,
    unitPriceDefault: payload?.unitPrice ?? '',
    tva: payload?.tva ?? '',
    includeHeaders: payload?.includeHeaders !== false,
  })

  const sheetName = payload?.sheetName
  const sheetRows = await readSheetRows(inputPath, sheetName)
  const includeHeaders = payload?.includeHeaders !== false
  const outputRows = mapSheetRows(sheetRows, {
    defaultUnitPrice: payload?.unitPrice ?? '',
    defaultTva: payload?.tva ?? '',
    includeHeaders,
    unitPrices: Array.isArray(payload?.unitPrices) ? payload.unitPrices : null,
    unitPricesByType: payload?.unitPricesByType ?? null,
  })

  const requestedPath = payload?.outputPath && String(payload.outputPath).trim()
    ? payload.outputPath
    : buildDefaultOutputPath(inputPath)
  const outputPath = normalizeOutputPath(requestedPath)

  await writeSheetRows(outputPath, 'Multidoc', outputRows)
  const dataRowCount = includeHeaders ? Math.max(outputRows.length - 1, 0) : outputRows.length
  sendLog('log', 'convert-file done', { outputPath, rowCount: dataRowCount })
  return { outputPath, rowCount: dataRowCount }
})

ipcMain.handle('reveal-path', async (_event, targetPath) => {
  if (!targetPath) return
  shell.showItemInFolder(targetPath)
})

ipcMain.handle('check-updates', async () => {
  if (isDev) return { status: 'dev' }
  try {
    const result = await autoUpdater.checkForUpdates()
    if (result && result.updateInfo && result.updateInfo.version !== app.getVersion()) {
      return { status: 'available', version: result.updateInfo.version }
    }
    return { status: 'up-to-date', version: app.getVersion() }
  } catch (err) {
    return { status: 'error', message: String(err) }
  }
})

ipcMain.handle('download-update', async () => {
  if (isDev) return
  await autoUpdater.downloadUpdate()
})

ipcMain.handle('install-update', async () => {
  if (isDev) return
  autoUpdater.quitAndInstall()
})

ipcMain.handle('window-close', () => {
  if (mainWindow) {
    mainWindow.close()
  }
})

ipcMain.handle('window-minimize', () => {
  if (mainWindow) {
    mainWindow.minimize()
  }
})

ipcMain.handle('window-toggle-maximize', () => {
  if (!mainWindow) return
  if (mainWindow.isMaximized()) {
    mainWindow.unmaximize()
  } else {
    mainWindow.maximize()
  }
})

app.whenReady().then(() => {
  Menu.setApplicationMenu(null)
  createWindow()
  wireAutoUpdater()
  if (!isDev) {
    autoUpdater.checkForUpdates()
  }
  if (isDev) {
    globalShortcut.register('Control+Shift+I', () => {
      const win = BrowserWindow.getFocusedWindow()
      if (win) {
        win.webContents.toggleDevTools({ mode: 'detach' })
      }
    })
    globalShortcut.register('F12', () => {
      const win = BrowserWindow.getFocusedWindow()
      if (win) {
        win.webContents.toggleDevTools({ mode: 'detach' })
      }
    })
  }
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('will-quit', () => {
  globalShortcut.unregisterAll()
})
