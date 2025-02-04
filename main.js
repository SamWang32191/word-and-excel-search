const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs').promises;
const mammoth = require('mammoth');
const WordExtractor = require('word-extractor');
const extractor = new WordExtractor();
const Promise = require('bluebird');
const iconv = require('iconv-lite');

// 配置
const CONCURRENT_LIMIT = 5; // 並發限制
const MAX_CACHE_SIZE = 1000; // 緩存最大容量

// 緩存管理器
class CacheManager {
  constructor(maxSize = MAX_CACHE_SIZE) {
    this.cache = new Map();
    this.maxSize = maxSize;
  }

  get(key) {
    const item = this.cache.get(key);
    if (item) item.lastAccessed = Date.now();
    return item?.data;
  }

  set(key, data) {
    if (this.cache.size >= this.maxSize) {
      const lruKey = [...this.cache.entries()].reduce((a, b) =>
        a[1].lastAccessed < b[1].lastAccessed ? a : b
      )[0];
      this.cache.delete(lruKey);
    }
    this.cache.set(key, { data, lastAccessed: Date.now() });
  }

  clear() {
    this.cache.clear();
  }
}

const fileCache = new CacheManager();
const excelCache = new CacheManager();
let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 900,
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  if (process.platform === 'darwin') {
    app.name = 'Word and Excel Search v1.0.1 by Tom';
  }

  mainWindow.loadFile('index.html');
}

// 清空緩存
function clearCache() {
  fileCache.clear();
  excelCache.clear();
  return {
    wordCount: fileCache.cache.size,
    excelCount: excelCache.cache.size,
  };
}

// 啟動應用
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

// 選擇目錄
ipcMain.handle('select-directory', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openDirectory'],
  });
  fileCache.clear();
  excelCache.clear();
  return result.filePaths[0];
});

// 清空緩存
ipcMain.handle('clear-cache', () => {
  return clearCache();
});

// 獲取緩存信息
ipcMain.handle('get-cache-info', () => {
  return {
    wordCount: fileCache.cache.size,
    excelCount: excelCache.cache.size,
  };
});

// 忽略隱藏文件和臨時文件
function shouldIgnoreFile(fileName) {
  return fileName.startsWith('~$') || fileName.startsWith('.');
}

// 搜索文件
ipcMain.handle('search-files', async (event, { directory, searchTexts, caseSensitive, fileTypes, enableCache }) => {
  const results = [];
  const keywords = searchTexts
    .split(',')
    .map((keyword) => keyword.trim())
    .filter((keyword) => keyword);

  let totalFiles = 0;
  let processedFiles = 0;

  // 處理 Excel 文件
  async function processExcelFile(filePath) {
    try {
      let workbook;
      if (enableCache && excelCache.get(filePath)) {
        workbook = excelCache.get(filePath);
      } else {
        const buffer = await fs.readFile(filePath);
        workbook = XLSX.read(buffer, { type: 'buffer' });
        if (enableCache) excelCache.set(filePath, workbook);
      }

      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        data.forEach((row, rowIndex) => {
          if (!row) return; // 跳過空行
          row.forEach((cell, colIndex) => {
            if (cell) {
              const cellText = cell.toString();
              keywords.forEach((keyword) => {
                const searchFor = caseSensitive ? keyword : keyword.toLowerCase();
                const searchIn = caseSensitive ? cellText : cellText.toLowerCase();
                if (searchIn.includes(searchFor)) {
                  results.push({
                    file: filePath,
                    type: 'Excel',
                    sheet: sheetName,
                    row: rowIndex + 1,
                    column: colIndex + 1,
                    content: cellText,
                    keyword: keyword,
                  });
                }
              });
            }
          });
        });
      });
    } catch (error) {
      event.sender.send('search-error', `讀取Excel檔案錯誤 ${filePath}: ${error}`);
    }
  }

  // 處理 Word 文件
  async function processWordFile(filePath, ext) {
    try {
      let text;
      if (enableCache && fileCache.get(filePath)) {
        text = fileCache.get(filePath);
      } else {
        if (ext === '.docx') {
          const buffer = await fs.readFile(filePath);
          const result = await mammoth.extractRawText({ buffer });
          text = result.value;
        } else {
          try {
            text = await extractor.extract(filePath).then((doc) => doc.getBody());
          } catch (extractError) {
            event.sender.send('search-error', `檔案 ${path.basename(filePath)} 讀取失敗，可能是舊版本Word檔案或檔案已損壞: ${extractError.message}`);
            return;
          }
        }
        if (enableCache) fileCache.set(filePath, text);
      }

      if (text) {
        keywords.forEach((keyword) => {
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
              position: index,
            });
            startIndex = index + 1;
          }
        });
      }
    } catch (error) {
      event.sender.send('search-error', `讀取Word檔案錯誤 ${filePath}: ${error}`);
    }
  }

  // 計算總文件數
  async function countFiles(dir) {
    try {
      const files = await fs.readdir(dir);
      for (const file of files) {
        if (shouldIgnoreFile(file)) continue;
        
        const filePath = path.join(dir, file);
        const stat = await fs.stat(filePath);

        if (stat.isDirectory()) {
          await countFiles(filePath);
        } else {
          const ext = path.extname(file).toLowerCase();
          const isValidFile = (fileTypes.includes('excel') && (ext === '.xlsx' || ext === '.xls')) ||
                            (fileTypes.includes('word') && (ext === '.docx' || ext === '.doc'));
          
          if (isValidFile) {
            totalFiles++;
          }
        }
      }
    } catch (error) {
      event.sender.send('search-error', `計算檔案數量錯誤 ${dir}: ${error}`);
    }
  }

  // 處理目錄
  async function processDirectory(dir) {
    try {
      const files = await fs.readdir(dir);
      await Promise.map(
        files,
        async (file) => {
          if (shouldIgnoreFile(file)) return;
          
          const filePath = path.join(dir, file);
          const stat = await fs.stat(filePath);

          if (stat.isDirectory()) {
            await processDirectory(filePath);
          } else {
            const ext = path.extname(file).toLowerCase();
            if (fileTypes.includes('excel') && (ext === '.xlsx' || ext === '.xls')) {
              await processExcelFile(filePath);
              processedFiles++;
            } else if (fileTypes.includes('word') && (ext === '.docx' || ext === '.doc')) {
              await processWordFile(filePath, ext);
              processedFiles++;
            }

            // 更新進度
            const percentage = Math.min(100, Math.round((processedFiles / totalFiles) * 100));
            event.sender.send('search-progress', {
              current: processedFiles,
              total: totalFiles,
              percentage,
            });
          }
        },
        { concurrency: CONCURRENT_LIMIT }
      );
    } catch (error) {
      event.sender.send('search-error', `讀取目錄錯誤 ${dir}: ${error}`);
    }
  }

  // 執行搜索
  await countFiles(directory);  // 先計算總文件數
  await processDirectory(directory);  // 然後處理文件
  
  return results;
});