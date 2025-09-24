import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import multer from 'multer';
import sharp from 'sharp';
import { promises as fs } from 'fs';
import path from 'path';
import os from 'os';
import { execFile } from 'child_process';
import { promisify } from 'util';

const execFileAsync = promisify(execFile);

const app = express();
const PORT = Number(process.env.PORT || 3000);

const RAW_EXTENSIONS = new Set([
  'cr2','cr3','crw','nef','arw','dng','rw2','pef','orf','raf','x3f','raw','sr2','nrw','k25','kdc','dcr'
]);

const isRawFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  return RAW_EXTENSIONS.has(ext) || mimetype.includes('raw') || mimetype.includes('x-');
};

const prepareRawBuffer = async (file: Express.Multer.File): Promise<Buffer> => {
  if (!isRawFile(file)) {
    return file.buffer;
  }

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-'));
  const inputPath = path.join(tmpDir, `input${path.extname(file.originalname) || '.raw'}`);
  const outputPath = path.join(tmpDir, 'output.tiff');

  try {
    await fs.writeFile(inputPath, file.buffer);
    await execFileAsync('dcraw', ['-T', '-6', '-O', outputPath, inputPath]);
    return await fs.readFile(outputPath);
  } catch (error) {
    console.error('dcraw conversion failed:', error);
    throw new Error('Failed to decode RAW image');
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

app.use(helmet());
app.use(cors({
  origin: process.env.FRONTEND_URL || 'https://morphy-1-ulvv.onrender.com',
  credentials: true,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 60,
  message: 'Too many requests from this IP, please try again later.'
});
app.use('/api/', limiter);

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 200 * 1024 * 1024,
    files: 1
  }
});

app.get('/health', (_req, res) => {
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    memory: process.memoryUsage(),
    version: process.env.npm_package_version || '1.0.0'
  });
});

app.get('/api/status', (_req, res) => {
  res.json({ 
    status: 'healthy',
    service: 'morphy-backend',
    timestamp: new Date().toISOString()
  });
});

app.post('/api/convert', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    const {
      quality = 'high',
      lossless = 'false',
      format = 'webp',
      width,
      height,
      iconSize = '16'
    } = req.body;

    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log(`Processing ${file.originalname} (${file.size} bytes)`);

    const inputBuffer = await prepareRawBuffer(file);

    const qualityValue = quality === 'high' ? 95 : quality === 'medium' ? 80 : 60;
    const isLossless = lossless === 'true';

    let pipeline = sharp(inputBuffer, {
      failOn: 'truncated',
      unlimited: true
    });

    const metadata = await pipeline.metadata();
    console.log(`Metadata => ${metadata.width}x${metadata.height}, format: ${metadata.format}`);

    let contentType: string;
    let fileExtension: string;

    switch (format.toLowerCase()) {
      case 'webp':
        pipeline = pipeline.webp({ quality: qualityValue, lossless: isLossless });
        contentType = 'image/webp';
        fileExtension = 'webp';
        break;
      case 'png':
        pipeline = pipeline.png({ compressionLevel: 9 });
        contentType = 'image/png';
        fileExtension = 'png';
        break;
      case 'jpeg':
      case 'jpg':
        pipeline = pipeline.jpeg({ quality: qualityValue, progressive: true });
        contentType = 'image/jpeg';
        fileExtension = 'jpg';
        break;
      case 'ico':
        pipeline = pipeline
          .resize(parseInt(iconSize) || 16, parseInt(iconSize) || 16, {
            fit: 'contain',
            background: { r: 0, g: 0, b: 0, alpha: 0 }
          })
          .png();
        contentType = 'image/x-icon';
        fileExtension = 'ico';
        break;
      default:
        return res.status(400).json({ error: 'Unsupported output format' });
    }

    if (width || height) {
      pipeline = pipeline.resize(
        width ? parseInt(width) : undefined,
        height ? parseInt(height) : undefined,
        {
          fit: 'inside',
          withoutEnlargement: true
        }
      );
    }

    const outputBuffer = await pipeline.toBuffer();
    const outputName = `${file.originalname.replace(/\.[^.]+$/, '')}.${fileExtension}`;

    res.set({
      'Content-Type': contentType,
      'Content-Disposition': `attachment; filename="${outputName}"`,
      'Content-Length': outputBuffer.length.toString(),
      'Cache-Control': 'no-cache'
    });

    res.send(outputBuffer);
  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({
      error: 'Conversion failed',
      details: process.env.NODE_ENV === 'development' && error instanceof Error ? error.message : undefined
    });
  }
});

app.use((err: any, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File too large. Maximum size is 200MB.' });
    }
    return res.status(400).json({ error: err.message });
  }

  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error' });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Morpy backend running on port ${PORT}`);
});

export default app;

