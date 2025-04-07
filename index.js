const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const archiver = require('archiver');
const cron = require('node-cron');

const app = express();
const PORT = 3000;
const storage = multer.diskStorage({
  destination: 'uploads/',
  filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});

// cron.schedule('59 23 * * *', () => {
//     console.log('[Scheduler] Running daily upload cleanup...');
  
//     fs.readdir(uploadDir, (err, files) => {
//       if (err) {
//         return console.error('[Scheduler] Failed to read upload directory:', err);
//       }
  
//       for (const file of files) {
//         fs.unlink(path.join(uploadDir, file), err => {
//           if (err) {
//             console.error(`[Scheduler] Error deleting file: ${file}`, err);
//           } else {
//             console.log(`[Scheduler] Deleted: ${file}`);
//           }
//         });
//       }
//     });
//   });
const uploadDir = path.join(__dirname, 'uploads');

cron.schedule('* * * * *', () => {
  console.log('[Scheduler] Running cleanup every minute...');

  fs.readdir(uploadDir, (err, files) => {
    if (err) {
      return console.error('[Scheduler] Failed to read upload directory:', err);
    }

    files.forEach(file => {
      const filePath = path.join(uploadDir, file);
      fs.unlink(filePath, err => {
        if (err) {
          console.error(`[Scheduler] Failed to delete ${file}:`, err);
        } else {
          console.log(`[Scheduler] Deleted: ${file}`);
        }
      });
    });
  });
});

  
const upload = multer({ storage });

app.use(express.static('public'));

app.post('/upload', upload.single('excel'), async (req, res) => {
  const filePath = req.file.path;

  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
console.log("sheetData",sheetData);
  if (!sheetData.length) {
    return res.status(400).send('Excel file is empty or invalid.');
  }

  const pdfPaths = [];
  const excelDateToJSDate = (serial) => {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    return date_info.toLocaleDateString('en-IN');
  };
  
  for (let i = 0; i < sheetData.length; i++) {
    const row = sheetData[i];
    const doc = new PDFDocument({ margin: 50 });
    const fileName = `invoice-${i + 1}.pdf`;
    const filePath = path.join(__dirname, 'uploads', fileName);
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);
  
    // Header
    doc.fontSize(20).text('Suraj Travels', { align: 'center' });
    doc.fontSize(10).text('123, Main Road, Mumbai – 400001', { align: 'center' });
    doc.text('Phone: +91 9876543210 | Email: surajtravels@example.com', { align: 'center' });
    doc.text('GSTIN: 27ABCDE1234F1Z5', { align: 'center' });
    doc.moveDown(2);
  
    // Invoice Info
    doc.fontSize(12).text(`Invoice No: INV-${i + 1}`, { align: 'left' });
    doc.text(`Date: ${excelDateToJSDate(row.DATE)}`);
    doc.moveDown();
  
    // Travel Details Table Style
    doc.font('Helvetica-Bold').text('Guest & Travel Details', { underline: true });
    doc.font('Helvetica');
    doc.moveDown();
  
    const details = [
      ['Guest Name', row['GUEST NAME']],
      ['Extra KM', row['EX.KM']],
      ['Extra Hours (DA)', row['EX.HR /DA']],
      ['Extra KM Charges', `${row['EX.KM RS']}`],
      ['Extra Hour Charges', `${row['EX.HR RS']}`],
      ['Toll/Parking', `${row['TOLL /PARKING ']}`],
      ['Base Rate', `${row['RATE ']}`]
    ];
  
    details.forEach(([label, value]) => {
      doc.text(`${label}: ${value}`);
    });
  
    doc.moveDown();
    doc.font('Helvetica-Bold').text(`Total Amount: ${row.TOTAL?.toFixed(2)}`, { align: 'right' });
    doc.moveDown(2);
  
    // Signature
    doc.text('Suraj Travels', { align: 'left' });
    doc.text('Authorized Signatory', { align: 'left' });
  
    doc.end();
    await new Promise(resolve => stream.on('finish', resolve));
    pdfPaths.push(filePath);
  }

  // Create zip
  const zipName = `invoices-${Date.now()}.zip`;
  const zipPath = path.join(__dirname, 'uploads', zipName);
  const output = fs.createWriteStream(zipPath);
  const archive = archiver('zip', { zlib: { level: 9 } });

  output.on('close', () => {
    // ✅ Only trigger download after ZIP is finalized
    res.download(zipPath, zipName, (err) => {
      if (err) {
        console.error('Download error:', err);
        res.status(500).send('Failed to download ZIP.');
      }
    });
  });

  archive.on('error', err => {
    console.error('Archiving failed:', err);
    res.status(500).send('Could not create ZIP file.');
  });

  archive.pipe(output);

  // Add PDFs to archive
  pdfPaths.forEach(file => {
    archive.file(file, { name: path.basename(file) });
  });

  // Finalize ZIP (very important!)
  await archive.finalize();
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
