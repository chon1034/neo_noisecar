const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const multer = require('multer');

const app = express();
const PORT = 4000;

// 設定上傳檔案存放目錄
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

// 每次啟動 app.js 時清空 uploads 資料夾（移除之前上傳的檔案）
fs.readdir(uploadDir, (err, files) => {
  if (err) {
    console.error('讀取 uploads 資料夾時發生錯誤:', err);
  } else {
    files.forEach(file => {
      const filePath = path.join(uploadDir, file);
      fs.unlink(filePath, err => {
        if (err) {
          console.error(`刪除檔案 ${file} 失敗:`, err);
        } else {
          console.log(`已刪除檔案: ${file}`);
        }
      });
    });
  }
});

// 設定 multer 存檔策略，處理上傳檔案並轉換檔名編碼以避免中文亂碼
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    let originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, originalName);
  }
});
const upload = multer({ storage });

// 提供前端靜態頁面（將前端 HTML、CSS、JS 放在 public 資料夾）
app.use(express.static(path.join(__dirname, 'public')));

// 上傳檔案 API，預期上傳欄位名稱為 excelFile 與 wordFile
app.post('/upload', upload.fields([
  { name: 'excelFile', maxCount: 1 },
  { name: 'wordFile', maxCount: 1 }
]), (req, res) => {
  try {
    // 取得上傳的 Excel 與 Word 模板檔案
    const excelFile = req.files['excelFile'][0];
    const wordFile = req.files['wordFile'][0];

    // 建立輸出目錄
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }
    const outputDocx = path.join(outputDir, 'merged.docx');

    // 1. 讀取 Excel 資料
    // 假設工作表名稱為 sheet1，讀取 Excel 轉換成 JSON 陣列，每一筆資料代表一筆記錄
    const workbook = XLSX.readFile(excelFile.path);
    const sheetName = 'sheet1';
    const sheet = workbook.Sheets[sheetName];
    const records = XLSX.utils.sheet_to_json(sheet);
    if (!records || records.length === 0) {
      return res.status(400).send('Excel 沒有資料');
    }
    console.log('Excel 資料筆數:', records.length);

    // 2. 讀取 Word 模板並進行資料合併
    // 請先修改你的 Word 模板，使需要重複的區段包在 {#records} 與 {/records} 之間，
    // 並在每筆資料區塊的結尾插入分頁符號，確保每筆資料各佔一頁。
    const content = fs.readFileSync(wordFile.path, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    
    // 傳入所有 Excel 資料，並將其對應到模板中的 records 迴圈區塊
    doc.render({ records: records });
    
    // 產生合併後的 DOCX 檔案
    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputDocx, buf);
    console.log('DOCX 合併完成:', outputDocx);

    // 直接下載合併後的 Word 檔案
    res.download(outputDocx, 'merged.docx', (err) => {
      if (err) {
        console.error('下載錯誤:', err);
      }
    });
  } catch (err) {
    console.error('伺服器錯誤:', err);
    res.status(500).send('伺服器錯誤');
  }
});

// 啟動伺服器
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
