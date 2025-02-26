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

// 設定 multer 存檔策略（處理上傳檔案，並轉換檔名編碼以避免中文亂碼）
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    // 將 file.originalname 從 latin1 轉成 utf8
    let originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, originalName);
  }
});
const upload = multer({ storage });

// 提供前端靜態頁面（public 資料夾內放置前端 HTML、CSS、JS）
app.use(express.static(path.join(__dirname, 'public')));

// 同時提供 output 資料夾作為靜態檔案路由，讓合併後的檔案可下載
app.use('/output', express.static(path.join(__dirname, 'output')));

// 上傳檔案的 API，預期上傳欄位名稱分別為 excelFile 與 wordFile
app.post('/upload', upload.fields([
  { name: 'excelFile', maxCount: 1 },
  { name: 'wordFile', maxCount: 1 }
]), (req, res) => {
  try {
    // 取得上傳的 Excel 與 Word 模板檔案
    const excelFile = req.files['excelFile'][0];
    const wordFile = req.files['wordFile'][0];

    // 建立暫存輸出目錄
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }
    const outputDocx = path.join(outputDir, 'merged.docx');
    const outputPdf = path.join(outputDir, 'merged.pdf');

    // 1. 讀取 Excel 資料
    // 讀取 Excel 檔，並將指定工作表（這裡假設工作表名稱為 sheet1）的資料轉換為 JSON 陣列
    const workbook = XLSX.readFile(excelFile.path);
    const sheetName = 'sheet1';
    const sheet = workbook.Sheets[sheetName];
    const records = XLSX.utils.sheet_to_json(sheet);
    if (!records || records.length === 0) {
      return res.status(400).send('Excel 沒有資料');
    }
    console.log('Excel 資料筆數:', records.length);

    // 2. 讀取 Word 模板並進行資料合併
    // 注意：請先修改你的 Word 模板，將需要重複的區塊包在 {#records} 和 {/records} 中，
    // 並在區塊結尾加入分頁符號，確保每筆資料顯示在單獨一頁。
    const content = fs.readFileSync(wordFile.path, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    
    // 傳入所有 Excel 資料，並將資料對應到模板中的 records 迴圈區塊
    doc.render({ records: records });
    
    // 產生合併後的 DOCX 檔案
    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputDocx, buf);
    console.log('DOCX 合併完成:', outputDocx);

    // 3. 使用 LibreOffice 將 DOCX 轉換成 PDF
    // 請確認 Linux 系統上已安裝 LibreOffice 並且 soffice 的路徑正確（此例中使用 /usr/bin/soffice）
    const sofficePath = '/usr/bin/soffice';
    const cmd = `${sofficePath} --headless --convert-to pdf "${outputDocx}" --outdir "${outputDir}"`;
    exec(cmd, (error, stdout, stderr) => {
      if (error) {
        console.error(`PDF 轉換錯誤: ${error}`);
        return res.status(500).send('PDF 轉換錯誤');
      }
      console.log('PDF 轉換結果:', stdout);

      // 4. 回傳一個簡單的 HTML 頁面，提供合併後的 DOCX 與 PDF 下載連結
      res.send(`
        <h2>合併完成！</h2>
        <p><a href="/output/merged.docx" download>下載合併後的 Word 檔 (.docx)</a></p>
        <p><a href="/output/merged.pdf" download>下載合併後的 PDF 檔 (.pdf)</a></p>
      `);
    });
  } catch (err) {
    console.error('伺服器錯誤:', err);
    res.status(500).send('伺服器錯誤');
  }
});

// 啟動伺服器，監聽指定的 PORT
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
