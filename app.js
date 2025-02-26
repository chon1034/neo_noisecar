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

// 設定 multer 存檔策略
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    // 取原檔名
    cb(null, file.originalname);
  }
});
const upload = multer({ storage });

// 提供前端靜態頁面
app.use(express.static(path.join(__dirname, 'public')));

// 上傳檔案的 API，預期上傳欄位名稱分別為 excelFile 與 wordFile
app.post('/upload', upload.fields([
  { name: 'excelFile', maxCount: 1 },
  { name: 'wordFile', maxCount: 1 }
]), (req, res) => {
  try {
    // 取得上傳的檔案路徑
    const excelFile = req.files['excelFile'][0];
    const wordFile = req.files['wordFile'][0];

    // 建立暫存輸出目錄
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }
    const outputDocx = path.join(outputDir, 'output.docx');
    const outputPdf = path.join(outputDir, 'output.pdf');

    // 1. 讀取 Excel 資料（假設只取第一筆資料）
    const workbook = XLSX.readFile(excelFile.path);
    // 假設工作表名稱為 sheet1（不分大小寫，可自行調整）
    const sheetName = 'sheet1';
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    if (!jsonData || jsonData.length === 0) {
      return res.status(400).send('Excel 沒有資料');
    }
    const data = jsonData[0]; // 取第一筆資料
    console.log('Excel 資料:', data);

    // 2. 讀取 Word 模板並進行資料合併
    const content = fs.readFileSync(wordFile.path, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    doc.render(data);
    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputDocx, buf);
    console.log('DOCX 合併完成:', outputDocx);

    // 3. 使用 LibreOffice 將 DOCX 轉換成 PDF
    const cmd = `soffice --headless --convert-to pdf "${outputDocx}" --outdir "${outputDir}"`;
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

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
