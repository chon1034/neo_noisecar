const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { exec } = require('child_process');
const multer = require('multer');

const app = express();
const PORT = 4000;

// 設定上傳檔案存放目錄
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

// 清空 uploads 資料夾（每次啟動時清除先前上傳的檔案）
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

// 設定 multer 存檔策略
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    // 轉換編碼，將 file.originalname 從 latin1 轉成 utf8，以避免中文亂碼
    let originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, originalName);
  }
});
const upload = multer({ storage });

// 提供前端靜態頁面（請將前端 HTML 放在 public 資料夾）
app.use(express.static(path.join(__dirname, 'public')));

// 上傳檔案的 API，預期上傳欄位名稱分別為 excelFile 與 wordFile
app.post('/upload', upload.fields([
  { name: 'excelFile', maxCount: 1 },
  { name: 'wordFile', maxCount: 1 }
]), (req, res) => {
  try {
    // 取得上傳的 Excel 與 Word 模板檔案路徑
    const excelFile = req.files['excelFile'][0];
    const wordFile = req.files['wordFile'][0];

    // 建立暫存輸出目錄
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }
    const outputDocx = path.join(outputDir, 'output.docx');
    const outputPdf = path.join(outputDir, 'output.pdf');

    // 1. 讀取 Excel 資料
    // 讀取 Excel 檔案，並將指定工作表（此例中假設工作表名稱為 sheet1）的資料轉換成 JSON 陣列
    const workbook = XLSX.readFile(excelFile.path);
    const sheetName = 'sheet1';
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    if (!jsonData || jsonData.length === 0) {
      return res.status(400).send('Excel 沒有資料');
    }
    console.log('Excel 資料筆數:', jsonData.length);

    // 2. 讀取 Word 模板並進行資料合併
    // 請確認你的 Word 模板中使用迴圈區塊，如下範例：
    // {#records}
    // 姓名：{{姓名}}
    // 性別：{{性別}}
    // 出生年月日：{{出生年月日}}
    // ...其他欄位...
    // ---------------------------
    // {/records}
    const content = fs.readFileSync(wordFile.path, 'binary');
    const zip = new PizZip(content);
    // 建立 docxtemplater 實例，設定 paragraphLoop 與 linebreaks 可改善段落換行問題
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    
    // 傳入 Excel 所有資料（jsonData 陣列）到模板中的 records 迴圈區塊
    doc.render({ records: jsonData });
    
    // 產生合併後的 DOCX 檔案，並寫入 output 目錄
    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputDocx, buf);
    console.log('DOCX 合併完成:', outputDocx);

    // 3. 使用 LibreOffice 將 DOCX 轉換成 PDF
    // 注意：請確認 Linux 系統上已安裝 LibreOffice 並可使用 soffice 指令
    // 若 PATH 有問題，可使用完整路徑（例如 /usr/bin/soffice）
    const sofficePath = '/usr/bin/soffice';
    const cmd = `${sofficePath} --headless --convert-to pdf "${outputDocx}" --outdir "${outputDir}"`;
    exec(cmd, (error, stdout, stderr) => {
      if (error) {
        console.error(`PDF 轉換錯誤: ${error}`);
        return res.status(500).send('PDF 轉換錯誤');
      }
      console.log('PDF 轉換結果:', stdout);
      // 4. 提供下載 PDF 檔案
      res.download(outputPdf, 'merged.pdf', (err) => {
        if (err) {
          console.error('下載錯誤:', err);
        }
      });
    });
  } catch (err) {
    console.error('伺服器錯誤:', err);
    res.status(500).send('伺服器錯誤');
  }
});

// 啟動伺服器，監聽指定 PORT
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
