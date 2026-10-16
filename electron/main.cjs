const { app, BrowserWindow, Menu, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow = null;
let isDirty = false;

function buildApplicationMenu() {
  const isMac = process.platform === 'darwin';

  const template = [
    {
      label: '&File',
      submenu: [
        {
          label: '&New Document',
          accelerator: 'CmdOrCtrl+N',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'file:new')
        },
        {
          label: '&Open...',
          accelerator: 'CmdOrCtrl+O',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'file:open')
        },
        { type: 'separator' },
        {
          label: '&Save',
          accelerator: 'CmdOrCtrl+S',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'file:save')
        },
        {
          label: 'Save &As...',
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'file:save-as')
        },
        {
          label: 'Export to &PDF...',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'file:export-pdf')
        },
        { type: 'separator' },
        {
          label: '&Print...',
          accelerator: 'CmdOrCtrl+P',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'print-doc')
        },
        { type: 'separator' },
        {
          label: 'E&xit',
          accelerator: isMac ? 'Cmd+Q' : 'Alt+F4',
          click: () => app.quit()
        }
      ]
    },
    {
      label: '&Edit',
      submenu: [
        {
          label: '&Undo',
          accelerator: 'CmdOrCtrl+Z',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'undo')
        },
        {
          label: '&Redo',
          accelerator: 'CmdOrCtrl+Y',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'redo')
        },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' },
        { type: 'separator' },
        {
          label: '&Find...',
          accelerator: 'CmdOrCtrl+F',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-find-dialog')
        },
        {
          label: 'R&eplace...',
          accelerator: 'CmdOrCtrl+H',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-replace-dialog')
        }
      ]
    },
    {
      label: '&View',
      submenu: [
        {
          label: '&Print Layout',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'set-view', { view: 'print' })
        },
        {
          label: '&Read Mode',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'set-view', { view: 'read' })
        },
        {
          label: '&Web Layout',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'set-view', { view: 'web' })
        },
        { type: 'separator' },
        {
          label: '&Navigation Pane',
          accelerator: 'CmdOrCtrl+Alt+N',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'toggle-nav-pane')
        },
        {
          label: '&Dark Mode',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'toggle-dark-mode')
        },
        {
          label: '&Focus Mode',
          accelerator: 'F11',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'toggle-focus')
        },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' }
      ]
    },
    {
      label: '&Insert',
      submenu: [
        {
          label: '&Page Break',
          accelerator: 'CmdOrCtrl+Enter',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'insert-page-break')
        },
        {
          label: '&Table...',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-table-dialog')
        },
        {
          label: '&Picture...',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'trigger-image-upload')
        },
        {
          label: '&Hyperlink...',
          accelerator: 'CmdOrCtrl+K',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-hyperlink-dialog')
        },
        {
          label: '&Equation...',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-equation-dialog')
        },
        {
          label: '&Symbol...',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-symbol-dialog')
        }
      ]
    },
    {
      label: 'F&ormat',
      submenu: [
        {
          label: '&Font Dialog...',
          accelerator: 'CmdOrCtrl+D',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-font-dialog')
        },
        {
          label: '&Paragraph Dialog...',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-paragraph-dialog')
        },
        {
          label: 'Page &Setup...',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-page-setup-dialog')
        }
      ]
    },
    {
      label: '&Tools',
      submenu: [
        {
          label: '&Word Count...',
          accelerator: 'CmdOrCtrl+Shift+G',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'open-word-count-dialog')
        },
        {
          label: '&Tell Me (Command Search)...',
          accelerator: 'Alt+Q',
          click: () => mainWindow && mainWindow.webContents.send('menu:command', 'tell-me')
        },
        { type: 'separator' },
        { role: 'toggleDevTools' }
      ]
    },
    {
      label: '&Help',
      submenu: [
        {
          label: 'About &Word...',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: 'About Microsoft Word Alternative',
              message: 'Word (Windows Desktop Edition)',
              detail: 'Built independently from the installed Microsoft Word reference application (WINWORD.EXE 16.0.20430.20146).\n\nFeatures OpenXML (.docx) fidelity, Fluent ribbon interface, multi-page layout, and offline desktop editing.',
              buttons: ['OK']
            });
          }
        }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

function createWindow() {
  const iconPath = path.join(__dirname, '../public/word-icon.svg');

  mainWindow = new BrowserWindow({
    width: 1280,
    height: 840,
    minWidth: 960,
    minHeight: 600,
    title: 'Document1 - Word',
    icon: fs.existsSync(iconPath) ? iconPath : undefined,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  });

  buildApplicationMenu();

  const isDev = process.env.NODE_ENV === 'development';
  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }

  mainWindow.on('close', (e) => {
    if (isDirty) {
      const choice = dialog.showMessageBoxSync(mainWindow, {
        type: 'warning',
        buttons: ['Save', "Don't Save", 'Cancel'],
        defaultId: 0,
        cancelId: 2,
        title: 'Microsoft Word',
        message: 'Want to save your changes to this document before closing?'
      });

      if (choice === 0) {
        e.preventDefault();
        mainWindow.webContents.send('menu:command', 'file:save');
      } else if (choice === 2) {
        e.preventDefault();
      }
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // Handle opening file passed as command line argument
  const fileArg = process.argv.slice(1).find(arg => !arg.startsWith('--') && !arg.endsWith('.js') && !arg.endsWith('.cjs') && fs.existsSync(arg));
  if (fileArg) {
    mainWindow.webContents.on('did-finish-load', () => {
      try {
        const fileBytes = fs.readFileSync(fileArg);
        mainWindow.webContents.send('menu:command', 'file:open-buffer', {
          filePath: fileArg,
          fileName: path.basename(fileArg),
          data: Array.from(fileBytes)
        });
      } catch (err) {
        console.error('Error opening file arg:', err);
      }
    });
  }
}

// IPC Handlers for Native Dialogs and File I/O
ipcMain.handle('dialog:openFile', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Open Document',
    filters: [
      { name: 'Word Documents', extensions: ['docx'] },
      { name: 'Web Pages', extensions: ['htm', 'html'] },
      { name: 'Rich Text Format', extensions: ['rtf'] },
      { name: 'Plain Text Files', extensions: ['txt'] },
      { name: 'All Files', extensions: ['*'] }
    ],
    properties: ['openFile']
  });

  if (result.canceled || result.filePaths.length === 0) return null;

  const targetPath = result.filePaths[0];
  const fileBytes = fs.readFileSync(targetPath);
  return {
    filePath: targetPath,
    fileName: path.basename(targetPath),
    data: Array.from(fileBytes)
  };
});

ipcMain.handle('dialog:saveFile', async (event, options = {}) => {
  let targetPath = options.defaultPath;

  if (options.forceSaveAs || !targetPath || !fs.existsSync(targetPath)) {
    const ext = options.format || 'docx';
    const filterMap = {
      docx: { name: 'Word Document', extensions: ['docx'] },
      html: { name: 'Web Page', extensions: ['html', 'htm'] },
      rtf: { name: 'Rich Text Format', extensions: ['rtf'] },
      txt: { name: 'Plain Text', extensions: ['txt'] }
    };

    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'Save As',
      defaultPath: targetPath || `Document1.${ext}`,
      filters: [
        filterMap[ext] || { name: 'Word Document', extensions: ['docx'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });

    if (result.canceled || !result.filePath) return null;
    targetPath = result.filePath;
  }

  const buffer = Buffer.from(options.data);
  fs.writeFileSync(targetPath, buffer);
  isDirty = false;

  return {
    filePath: targetPath,
    fileName: path.basename(targetPath)
  };
});

ipcMain.handle('dialog:exportPdf', async (event, options = {}) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'Export to PDF',
    defaultPath: options.defaultPath || 'Document1.pdf',
    filters: [{ name: 'PDF Document', extensions: ['pdf'] }]
  });

  if (result.canceled || !result.filePath) return null;

  const pdfData = await mainWindow.webContents.printToPDF({
    pageSize: 'Letter',
    printBackground: true,
    marginsType: 1
  });

  fs.writeFileSync(result.filePath, pdfData);
  return { filePath: result.filePath };
});

ipcMain.on('window:setState', (event, state) => {
  if (state.title && mainWindow) {
    mainWindow.setTitle(state.title);
  }
  if (typeof state.dirty === 'boolean') {
    isDirty = state.dirty;
  }
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
});
