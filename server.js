const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { processWord } = require('./utils/wordProcessor');

const app = express();
const PORT = process.env.PORT || 3000;

// Ensure uploads and output dirs exist
const uploadsDir = path.join(os.tmpdir(), 'uploads');
const outputDir = path.join(os.tmpdir(), 'output');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

// Multer config
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  },
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.docx') {
      return cb(new Error('Hanya file .docx yang diperbolehkan!'));
    }
    cb(null, true);
  },
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB max
});

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

// Upload and process endpoint
app.post('/api/process', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Tidak ada file yang diupload.' });
    }

    const inputPath = req.file.path;
    const originalName = req.file.originalname.replace('.docx', '');
    const outputFilename = `${originalName}_rapi.docx`;
    const outputPath = path.join(outputDir, outputFilename);

    // Get options from request
    const options = {
      coverType: req.body.coverType || '2026',
      fontFamily: req.body.fontFamily || 'Times New Roman',
      fontSize: parseInt(req.body.fontSize) || 11,
      lineSpacing: parseFloat(req.body.lineSpacing) || 1.5,
      marginTop: parseFloat(req.body.marginTop) || 3,
      marginBottom: parseFloat(req.body.marginBottom) || 3,
      marginLeft: parseFloat(req.body.marginLeft) || 4,
      marginRight: parseFloat(req.body.marginRight) || 3,
      removeImages: req.body.removeImages === 'true',
      fixTables: req.body.fixTables !== 'false',
      fixHeadings: req.body.fixHeadings !== 'false',
    };

    // Process the document
    const result = await processWord(inputPath, outputPath, options);

    // Clean up uploaded file
    fs.unlinkSync(inputPath);

    res.json({
      success: true,
      filename: outputFilename,
      downloadUrl: `/api/download/${encodeURIComponent(outputFilename)}`,
      report: result.report,
    });
  } catch (err) {
    console.error('Processing error:', err);
    // Clean up uploaded file on error
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).json({ error: err.message || 'Terjadi kesalahan saat memproses dokumen.' });
  }
});

// Download endpoint
app.get('/api/download/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(outputDir, filename);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File tidak ditemukan.' });
  }

  res.download(filePath, filename, (err) => {
    if (err) {
      console.error('Download error:', err);
    }
    // Clean up output file after download
    try {
      fs.unlinkSync(filePath);
    } catch (e) {
      // ignore
    }
  });
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`🚀 Rapikan Word server running at http://localhost:${PORT}`);
  });
}

module.exports = app;
