const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs');
const StreamZip = require('node-stream-zip');

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

  // 在 macOS 上設定應用程式名稱
  if (process.platform === 'darwin') {
    app.name = 'Word and Excel Search v1.0.0 by Tom ';
  }

  mainWindow.loadFile('index.html');
}

app.whenReady().then(createWindow);

// 針對 macOS 的視窗管理
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
  return result.filePaths[0];
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
        
        if (fileName.startsWith('~$')) {
          continue;
        }
        
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

  // 檢查是否包含 XML 標記的特徵
  function isXMLContent(text) {
    const xmlPatterns = [
      'w:', 
      'xml:', 
      '/>', 
      '</',
      '">',
      'w:val=',
      'w:sz=',
      'w:space=',
      'w:color=',
      'xmlns:',
      'office/word/',
      'officeDocument/',
      'wordprocessing',
      'http://www.',
      'Data="http',
      'mso-',      // 加入 Word 樣式標記
      ';mso-',
      'distance-',
      'position-',
      ':absolute',
      ':pt'
    ];
  
    // 檢查內容是否包含任何 XML 或樣式相關的文字
    return xmlPatterns.some(pattern => text.includes(pattern)) ||
           // 檢查內容是否看起來像完整的 URL 或 XML 路徑
           /^[a-zA-Z]+:\/\//.test(text) ||
           // 檢查是否包含多個技術性標記
           text.includes('/') && text.split('/').length > 2 ||
           // 檢查是否包含 CSS 樣式類的內容
           text.includes(':') && text.split(';').length > 1;
  }

  async function searchInDirectory(dir) {
    try {
      const files = fs.readdirSync(dir);
      
      for (const file of files) {
        const filePath = path.join(dir, file);
        const fileName = path.basename(file);
        
        if (fileName.startsWith('~$')) {
          continue;
        }

        const stat = fs.statSync(filePath);
        
        if (stat.isDirectory()) {
          await searchInDirectory(filePath);
        } else {
          const ext = path.extname(file).toLowerCase();
          
          if (fileTypes.includes('excel') && (ext === '.xlsx' || ext === '.xls')) {
            try {
              const workbook = XLSX.readFile(filePath, {type: 'binary'});
              
              workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
                
                data.forEach((row, rowIndex) => {
                  if (row) {
                    row.forEach((cell, colIndex) => {
                      if (cell) {
                        const cellText = cell.toString();
                        
                        keywords.forEach(keyword => {
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
              let text = '';
              
              if (ext === '.doc') {
                const content = fs.readFileSync(filePath);
                try {
                  text = content.toString('utf16le');
                } catch {
                  try {
                    text = content.toString('latin1');
                  } catch {
                    text = content.toString('ascii');
                  }
                }
              } else {
                try {
                  const zip = new StreamZip.async({ file: filePath });
                  const data = await zip.entryData('word/document.xml');
                  text = data.toString('utf8');
                  await zip.close();
                } catch (error) {
                  console.error(`無法讀取 DOCX 檔案 ${filePath}: ${error}`);
                }
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
                    const content = text.substring(start, end);

                    // 只有不是 XML 內容時才加入結果
                    if (!isXMLContent(content)) {
                      results.push({
                        file: filePath,
                        type: 'Word',
                        content: content,
                        keyword: keyword,
                        position: index
                      });
                    }

                    startIndex = index + 1;
                  }
                });
              }
            } catch (error) {
              console.error(`讀取Word檔案錯誤 ${filePath}: ${error}`);
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
      }
    } catch (error) {
      console.error(`讀取目錄錯誤 ${dir}: ${error}`);
    }
  }
  
  await searchInDirectory(directory);
  return results;
});