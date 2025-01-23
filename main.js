const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs');
const mammoth = require('mammoth');
const WordExtractor = require("word-extractor"); 
const extractor = new WordExtractor();
const Promise = require('bluebird');
const iconv = require('iconv-lite'); // 引入 iconv-lite 套件（需先安裝）
const OleDoc = require('ole-doc').OleDoc;
const CONCURRENT_LIMIT = 10;
const fileCache = new Map();
const excelCache = new Map();
let mainWindow;


function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  if (process.platform === 'darwin') {
    app.name = 'Word and Excel Search v1.0.0 by Tom ';
  }

  mainWindow.loadFile('index.html');
}

function clearCache() {
  fileCache.clear();
  excelCache.clear();
  return {
    wordCount: fileCache.size,
    excelCount: excelCache.size
  };
}

app.whenReady().then(createWindow);

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

ipcMain.handle('select-directory', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openDirectory']
  });
  fileCache.clear();
  excelCache.clear();
  return result.filePaths[0];
});

ipcMain.handle('clear-cache', () => {
  return clearCache();
});

ipcMain.handle('get-cache-info', () => {
  return {
    wordCount: fileCache.size,
    excelCount: excelCache.size
  };
});

ipcMain.handle('search-files', async (event, { directory, searchTexts, caseSensitive, fileTypes }) => {
  const results = [];
  const keywords = searchTexts.split(',').map(keyword => keyword.trim()).filter(keyword => keyword);
  
  function getTotalFiles(dir, types) {
    let total = 0;
    try {
      const items = fs.readdirSync(dir);
      for (const item of items) {
        const fullPath = path.join(dir, item);
        const fileName = path.basename(item);
        if (fileName.startsWith('~$')) continue;
        
        if (fs.statSync(fullPath).isDirectory()) {
          total += getTotalFiles(fullPath, types);
        } else {
          const ext = path.extname(fullPath).toLowerCase();
          if ((types.includes('excel') && (ext === '.xlsx' || ext === '.xls')) ||
              (types.includes('word') && (ext === '.docx' || ext === '.doc'))) {
            total++;
          }
        }
      }
    } catch (error) {
      console.error(`計算檔案數量錯誤 ${dir}: ${error}`);
    }
    return total;
  }

  const totalFiles = getTotalFiles(directory, fileTypes);
  let processedFiles = 0;

  async function processFile(filePath, ext, fileTypes) {
    if (fileTypes.includes('excel') && (ext === '.xlsx' || ext === '.xls')) {
      try {
        let workbook;
        if (excelCache.has(filePath)) {
          workbook = excelCache.get(filePath);
        } else {
          workbook = XLSX.readFile(filePath, {type: 'binary'});
          excelCache.set(filePath, workbook);
        }
        
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
          
          data.forEach((row, rowIndex) => {
            if (row) {
              row.forEach((cell, colIndex) => {
                if (cell) {
                  const cellText = cell.toString();
                  keywords.forEach(keyword => {
                    const searchFor = caseSensitive ? keyword.trim() : keyword.trim().toLowerCase();
                    const searchIn = caseSensitive ? cellText.trim() : cellText.trim().toLowerCase();
                    if (searchIn.includes(searchFor)) {
                      results.push({
                        file: filePath,
                        type: 'Excel',
                        sheet: sheetName,
                        row: rowIndex + 1,
                        column: colIndex + 1,
                        content: cellText,
                        keyword: keyword
                      });
                    }
                  });
                }
              });
            }
          });
        });
      } catch (error) {
        console.error(`讀取Excel檔案錯誤 ${filePath}: ${error}`);
      }
    }
    
    if (fileTypes.includes('word') && (ext === '.docx' || ext === '.doc')) {
      try {
        let text;
        if (fileCache.has(filePath)) {
          text = fileCache.get(filePath);
        } else {
          if (ext === '.docx') {
            const buffer = fs.readFileSync(filePath);
            const result = await mammoth.extractRawText({buffer});
            text = result.value;
          } else {
            // 使用 word-extractor 處理 .doc
            text = await extractor.extract(filePath)
              .then(doc => doc.getBody());
          }
          fileCache.set(filePath, text);
        }
        
        if (text) {
          keywords.forEach(keyword => {
            const searchFor = caseSensitive ? keyword : keyword.toLowerCase();
            const searchIn = caseSensitive ? text : text.toLowerCase();
            let startIndex = 0;
            let index;
            while ((index = searchIn.indexOf(searchFor, startIndex)) !== -1) {
              const start = Math.max(0, index - 50);
              const end = Math.min(text.length, index + keyword.length + 50);
              results.push({
                file: filePath,
                type: 'Word',
                content: text.substring(start, end),
                keyword: keyword,
                position: index
              });
              startIndex = index + 1;
            }
          });
        }
      } catch (error) {
        const errorMessage = `read word error -> ${filePath}: ${error}`;
        const utf8ErrorMessage = iconv.encode(errorMessage, 'utf-8');
        console.error(utf8ErrorMessage.toString('utf-8'));
      }
     }

    processedFiles++;
    if (processedFiles <= totalFiles) {
      event.sender.send('search-progress', {
        current: processedFiles,
        total: totalFiles,
        percentage: Math.min(100, Math.round((processedFiles / totalFiles) * 100))
      });
    }
  }

  async function searchInDirectory(dir) {
    try {
      const files = fs.readdirSync(dir);
      const filesToProcess = files.filter(file => {
        const fileName = path.basename(file);
        return !fileName.startsWith('~$');
      });

      await Promise.map(filesToProcess, async (file) => {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);
        
        if (stat.isDirectory()) {
          await searchInDirectory(filePath);
        } else {
          const ext = path.extname(file).toLowerCase();
          await processFile(filePath, ext, fileTypes);
        }
      }, {concurrency: CONCURRENT_LIMIT});
      
    } catch (error) {
      console.error(`讀取目錄錯誤 ${dir}: ${error}`);
    }
  }
  
  await searchInDirectory(directory);
  return results;
});