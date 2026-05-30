import 'dotenv/config';
import express, { Request, Response, NextFunction } from 'express';
import multer from 'multer';
import path from 'path';
import { extractFromImages } from './services/imageExtractor';
import { fillAndDownload } from './services/excelHandler';
import { normaliseClaimData, normaliseparts } from './services/dataMapper';
import { ExtractionResult } from './types/index';

const app = express();
const PORT = parseInt(process.env.PORT ?? '3000', 10);

// Memory storage — images never written to disk
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 MB per file
  fileFilter: (_req, file, cb) => {
    if (file.mimetype.startsWith('image/')) {
      cb(null, true);
    } else {
      cb(new Error('Only image files are accepted'));
    }
  },
});

// Serve frontend
app.use(express.static(path.join(__dirname, '..', 'public')));

// ─── POST /api/process ───────────────────────────────────────────────────────
app.post(
  '/api/process',
  upload.array('images', 20),
  async (req: Request, res: Response, next: NextFunction) => {
    try {
      const files = req.files as Express.Multer.File[];
      if (!files || files.length === 0) {
        res.status(400).json({ error: 'No images uploaded.' });
        return;
      }

      console.log(`Processing ${files.length} image(s)...`);

      // Extract data from all images in parallel
      const rawResult: ExtractionResult = await extractFromImages(
        files.map((f) => ({ buffer: f.buffer, mimetype: f.mimetype }))
      );

      // Post-process
      const claimData = normaliseClaimData(rawResult.claimData);
      const parts = normaliseparts(rawResult.parts);

      console.log('Extracted claimData:', JSON.stringify(claimData, null, 2));
      console.log(`Extracted ${parts.length} parts`);

      // Fill Excel and stream back
      const excelBuffer = await fillAndDownload(claimData, parts);

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="Format_filled.xlsx"');
      res.setHeader('Content-Length', excelBuffer.length);
      res.send(excelBuffer);
    } catch (err) {
      next(err);
    }
  }
);

// ─── GET / → frontend ────────────────────────────────────────────────────────
app.get('/', (_req: Request, res: Response) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

// ─── Error handler ───────────────────────────────────────────────────────────
app.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  console.error('Server error (full):', err);
  res.status(500).json({ error: err.message || 'Internal server error' });
});

app.listen(PORT, () => {
  console.log(`\n✅ Claim Excel Filler running at http://localhost:${PORT}\n`);
});
