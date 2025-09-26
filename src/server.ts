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
import { randomUUID } from 'crypto';
import Papa from 'papaparse';

const BATCH_OUTPUT_DIR = path.join(os.tmpdir(), 'morphy-batch-outputs');
fs.mkdir(BATCH_OUTPUT_DIR, { recursive: true }).catch(() => undefined);

const batchFileMetadata = new Map<string, { downloadName: string; mime: string }>();
const batchCleanupTimers = new Map<string, NodeJS.Timeout>();

const scheduleBatchFileCleanup = (storedFilename: string) => {
  const existingTimer = batchCleanupTimers.get(storedFilename);
  if (existingTimer) {
    clearTimeout(existingTimer);
  }

  const timeout = setTimeout(() => {
    fs.rm(path.join(BATCH_OUTPUT_DIR, storedFilename), { force: true }).catch(() => undefined);
    batchCleanupTimers.delete(storedFilename);
    batchFileMetadata.delete(storedFilename);
  }, 5 * 60 * 1000);

  batchCleanupTimers.set(storedFilename, timeout);
};

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

const LIBREOFFICE_CONVERSIONS: Record<string, {
  convertTo: string;
  extension: string;
  mime: string;
}> = {
  doc: {
    convertTo: 'doc:MS Word 97',
    extension: 'doc',
    mime: 'application/msword'
  },
  docx: {
    convertTo: 'docx',
    extension: 'docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  },
  pdf: {
    convertTo: 'pdf',
    extension: 'pdf',
    mime: 'application/pdf'
  },
  rtf: {
    convertTo: 'rtf',
    extension: 'rtf',
    mime: 'application/rtf'
  },
  odt: {
    convertTo: 'odt',
    extension: 'odt',
    mime: 'application/vnd.oasis.opendocument.text'
  },
  html: {
    convertTo: 'html',
    extension: 'html',
    mime: 'text/html; charset=utf-8'
  },
  txt: {
    convertTo: 'txt:Text (encoded):UTF8',
    extension: 'txt',
    mime: 'text/plain; charset=utf-8'
  },
  odp: {
    convertTo: 'odp',
    extension: 'odp',
    mime: 'application/vnd.oasis.opendocument.presentation'
  },
  ppt: {
    convertTo: 'ppt',
    extension: 'ppt',
    mime: 'application/vnd.ms-powerpoint'
  },
  pptx: {
    convertTo: 'pptx',
    extension: 'pptx',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  },
  xls: {
    convertTo: 'xls',
    extension: 'xls',
    mime: 'application/vnd.ms-excel'
  },
  xlsx: {
    convertTo: 'xlsx',
    extension: 'xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  },
  ods: {
    convertTo: 'ods',
    extension: 'ods',
    mime: 'application/vnd.oasis.opendocument.spreadsheet'
  }
};

const LIBREOFFICE_CANDIDATES = [
  process.env.LIBREOFFICE_PATH,
  'libreoffice',
  'soffice'
].filter((value): value is string => Boolean(value));

const isCsvFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  return ext === 'csv' || mimetype.includes('csv') || mimetype.includes('text/plain');
};

const sanitizeFilename = (name: string) =>
  name.replace(/[^a-zA-Z0-9-_]+/g, '_').replace(/^_+|_+$/g, '') || 'file';

interface NormalizedCsvResult {
  normalizedCsv: string;
  rowCount: number;
}

const normalizeCsvBuffer = (buffer: Buffer): NormalizedCsvResult => {
  const csvText = buffer.toString('utf8');

  const dialects = [
    { delimiter: ',', description: 'comma separated' },
    { delimiter: ';', description: 'semicolon separated' },
    { delimiter: '\t', description: 'tab separated' }
  ];

  let parsedRows: string[][] | null = null;

  for (const dialect of dialects) {
    const parsed = Papa.parse<string[]>(csvText, {
      delimiter: dialect.delimiter,
      skipEmptyLines: 'greedy',
      dynamicTyping: false,
      header: false
    });

    if (parsed.errors && parsed.errors.length > 0) {
      console.warn(`CSV parse warnings (${dialect.description}):`, parsed.errors.slice(0, 3));
    }

    const rows = (parsed.data || []).filter((row): row is string[] => Array.isArray(row) && row.length > 0);

    if (rows.length === 0) {
      continue;
    }

    const majorityMap = rows.reduce<Record<number, number>>((acc, row) => {
      const len = row.length;
      acc[len] = (acc[len] || 0) + 1;
      return acc;
    }, {});

    const majorityEntries = Object.entries(majorityMap)
      .map(([columns, count]) => ({ columns: Number(columns), count: Number(count) }))
      .sort((a, b) => b.count - a.count);

    const best = majorityEntries.length > 0 ? majorityEntries[0] : { columns: 0, count: 0 };
    const majorityRatio = rows.length > 0 ? best.count / rows.length : 0;

    if (best.columns > 1 && majorityRatio >= 0.6) {
      parsedRows = rows.map(row => row.slice(0, best.columns));
      break;
    }

    if (!parsedRows || rows.length > parsedRows.length) {
      parsedRows = rows;
    }
  }

  if (!parsedRows || parsedRows.length === 0) {
    throw new Error('CSV appears to be empty or malformed');
  }

  const sanitizedRows = parsedRows.map(row =>
    row.map(value => {
      if (value === null || value === undefined) {
        return '';
      }
      if (typeof value === 'string') {
        return value;
      }
      return String(value);
    })
  );

  let normalizedCsv = Papa.unparse(sanitizedRows, {
    delimiter: ',',
    newline: '\n',
    quotes: true
  });

  if (!normalizedCsv.endsWith('\n')) {
    normalizedCsv += '\n';
  }

  return {
    normalizedCsv,
    rowCount: sanitizedRows.length
  };
};

const CALIBRE_CONVERSIONS: Record<string, {
  extension: string;
  mime: string;
}> = {
  mobi: {
    extension: 'mobi',
    mime: 'application/x-mobipocket-ebook'
  },
  epub: {
    extension: 'epub',
    mime: 'application/epub+zip'
  }
};

const CALIBRE_CANDIDATES = [
  process.env.CALIBRE_PATH,
  process.env.EBOOK_CONVERT_PATH,
  'ebook-convert',
  'ebook-convert.exe'
].filter((value): value is string => Boolean(value));

interface CommandResult {
  stdout: string;
  stderr: string;
}

const execLibreOffice = async (args: string[]): Promise<CommandResult> => {
  let lastError: unknown;
  for (const binary of LIBREOFFICE_CANDIDATES) {
    try {
      const result = await execFileAsync(binary, args, {
        env: {
          ...process.env,
          HOME: process.env.HOME || os.homedir(),
          USERPROFILE: process.env.USERPROFILE || os.homedir()
        }
      });
      return result;
    } catch (error: any) {
      lastError = error;
      if (error?.code === 'ENOENT') {
        continue;
      }
      const stderr = typeof error?.stderr === 'string' && error.stderr.trim().length > 0
        ? ` | stderr: ${error.stderr.trim()}`
        : '';
      const stdout = typeof error?.stdout === 'string' && error.stdout.trim().length > 0
        ? ` | stdout: ${error.stdout.trim()}`
        : '';
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`LibreOffice execution failed using "${binary}": ${message}${stderr}${stdout}`);
    }
  }

  throw new Error(
    'LibreOffice binary not found. Please ensure LibreOffice is installed and available on the PATH or set LIBREOFFICE_PATH.' +
      (lastError instanceof Error ? ` (${lastError.message})` : '')
  );
};

const execCalibre = async (args: string[]): Promise<CommandResult> => {
  let lastError: unknown;
  for (const binary of CALIBRE_CANDIDATES) {
    try {
      const result = await execFileAsync(binary, args, {
        env: {
          ...process.env,
          HOME: process.env.HOME || os.homedir(),
          USERPROFILE: process.env.USERPROFILE || os.homedir()
        }
      });
      return result;
    } catch (error: any) {
      lastError = error;
      if (error?.code === 'ENOENT') {
        continue;
      }
      const stderr = typeof error?.stderr === 'string' && error.stderr.trim().length > 0
        ? ` | stderr: ${error.stderr.trim()}`
        : '';
      const stdout = typeof error?.stdout === 'string' && error.stdout.trim().length > 0
        ? ` | stdout: ${error.stdout.trim()}`
        : '';
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`Calibre execution failed using "${binary}": ${message}${stderr}${stdout}`);
    }
  }

  throw new Error(
    'Calibre ebook-convert binary not found. Please ensure Calibre is installed and available on the PATH or set CALIBRE_PATH/EBOOK_CONVERT_PATH.' +
      (lastError instanceof Error ? ` (${lastError.message})` : '')
  );
};

const convertCsvWithLibreOffice = async (
  file: Express.Multer.File,
  targetFormat: string
): Promise<{ buffer: Buffer; filename: string; mime: string }> => {
  const conversion = LIBREOFFICE_CONVERSIONS[targetFormat];
  if (!conversion) {
    throw new Error('Unsupported LibreOffice target format');
  }

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-lo-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const safeBase = `${sanitizeFilename(originalBase)}_${randomUUID()}`;
  const normalized = normalizeCsvBuffer(file.buffer);
  const inputFilename = `${safeBase}.csv`;
  const inputPath = path.join(tmpDir, inputFilename);

  const findOutputFile = async (): Promise<string | null> => {
    const files = await fs.readdir(tmpDir);
    const targetExt = `.${conversion.extension.toLowerCase()}`;
    const directMatch = files.find(name => name.toLowerCase() === `${safeBase.toLowerCase()}${targetExt}`);
    if (directMatch) {
      return path.join(tmpDir, directMatch);
    }

    const fallbackMatch = files.find(name => name.toLowerCase().endsWith(targetExt));
    return fallbackMatch ? path.join(tmpDir, fallbackMatch) : null;
  };

  const commandVariants: string[][] = [
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--calc'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--writer'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:44,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--calc'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:59,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--calc'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:9,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--calc'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:44,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--writer'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:59,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--writer'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:9,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--writer']
  ];

  let lastError: unknown;

  try {
    await fs.writeFile(inputPath, normalized.normalizedCsv, 'utf8');

    for (const args of commandVariants) {
      try {
        const { stdout, stderr } = await execLibreOffice(args);
        if (stdout.trim().length > 0) {
          console.log('LibreOffice stdout:', stdout.trim());
        }
        if (stderr.trim().length > 0) {
          console.warn('LibreOffice stderr:', stderr.trim());
        }

        const outputPath = await findOutputFile();
        if (!outputPath) {
          throw new Error(`LibreOffice did not produce an output file for args: ${args.join(' ')}`);
        }

        const outputBuffer = await fs.readFile(outputPath);
        const downloadName = `${sanitizeFilename(originalBase)}.${conversion.extension}`;

        return {
          buffer: outputBuffer,
          filename: downloadName,
          mime: conversion.mime
        };
      } catch (commandError) {
        lastError = commandError;
        console.error('LibreOffice conversion attempt failed:', commandError);
      }
    }

    if (lastError) {
      throw lastError;
    }

    throw new Error('LibreOffice conversion failed for unknown reasons.');

  } catch (error) {
    console.error('LibreOffice conversion failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown LibreOffice error';
    throw new Error(`Failed to convert CSV with LibreOffice: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const convertCsvWithCalibre = async (
  file: Express.Multer.File,
  targetFormat: string
): Promise<{ buffer: Buffer; filename: string; mime: string }> => {
  const conversion = CALIBRE_CONVERSIONS[targetFormat];
  if (!conversion) {
    throw new Error('Unsupported Calibre target format');
  }

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-calibre-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const safeBase = `${sanitizeFilename(originalBase)}_${randomUUID()}`;
  const normalized = normalizeCsvBuffer(file.buffer);
  const inputPath = path.join(tmpDir, `${safeBase}.html`);
  const outputPath = path.join(tmpDir, `${safeBase}.${conversion.extension}`);

  const htmlTableRows = normalized.normalizedCsv
    .split('\n')
    .filter(line => line.trim().length > 0)
    .map(line => {
      const parsed = Papa.parse<string[]>(line, {
        delimiter: ',',
        dynamicTyping: false,
        header: false
      });
      const values = (parsed.data[0] || []).map(value =>
        String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      );
      const cells = values.map(value => `<td>${value}</td>`).join('');
      return `<tr>${cells}</tr>`;
    })
    .join('\n');

  const htmlContent = `<!DOCTYPE html>\n<html lang="en">\n<head>\n  <meta charset="UTF-8" />\n  <title>${sanitizeFilename(originalBase)}</title>\n  <style>\n    body { font-family: Arial, sans-serif; padding: 24px; }\n    table { border-collapse: collapse; width: 100%; }\n    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }\n  </style>\n</head>\n<body>\n  <table>\n    <tbody>\n      ${htmlTableRows}\n    </tbody>\n  </table>\n</body>\n</html>`;

  try {
    await fs.writeFile(inputPath, htmlContent, 'utf8');

    const { stdout, stderr } = await execCalibre([
      inputPath,
      outputPath,
      '--change-justification', 'left',
      '--pretty-print',
      '--disable-font-rescaling',
      '--input-encoding', 'utf-8'
    ]);

    if (stdout.trim().length > 0) {
      console.log('Calibre stdout:', stdout.trim());
    }
    if (stderr.trim().length > 0) {
      console.warn('Calibre stderr:', stderr.trim());
    }

    const outputBuffer = await fs.readFile(outputPath);
    const downloadName = `${sanitizeFilename(originalBase)}.${conversion.extension}`;

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: conversion.mime
    };
  } catch (error) {
    console.error('Calibre conversion failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown Calibre error';
    throw new Error(`Failed to convert CSV with Calibre: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

app.use(helmet());
app.use(cors({
  origin: process.env.FRONTEND_URL || '*',
  credentials: true
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
    fileSize: 100 * 1024 * 1024,
    files: 1
  }
});

const uploadBatch = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024,
    files: 20
  }
});

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
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

    const targetFormat = String(format).toLowerCase();

    if (isCsvFile(file)) {
      if (LIBREOFFICE_CONVERSIONS[targetFormat]) {
        const { buffer, filename, mime } = await convertCsvWithLibreOffice(file, targetFormat);

        res.set({
          'Content-Type': mime,
          'Content-Disposition': `attachment; filename="${filename}"`,
          'Content-Length': buffer.length.toString(),
          'Cache-Control': 'no-cache'
        });

        return res.send(buffer);
      }

      if (CALIBRE_CONVERSIONS[targetFormat]) {
        const { buffer, filename, mime } = await convertCsvWithCalibre(file, targetFormat);

        res.set({
          'Content-Type': mime,
          'Content-Disposition': `attachment; filename="${filename}"`,
          'Content-Length': buffer.length.toString(),
          'Cache-Control': 'no-cache'
        });

        return res.send(buffer);
      }
    }

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

    switch (targetFormat) {
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
    const errorMessage = error instanceof Error ? error.message : 'Unknown conversion error';
    res.status(500).json({
      error: 'Conversion failed',
      details: errorMessage
    });
  }
});

app.post('/api/convert/batch', uploadBatch.array('files'), async (req, res) => {
  const files = req.files as Express.Multer.File[] | undefined;
  const format = String(req.body?.format ?? '').toLowerCase() || 'webp';

  if (!files || files.length === 0) {
    return res.status(400).json({
      success: false,
      processed: 0,
      results: [],
      error: 'No files uploaded'
    });
  }

  if (files.length > 20) {
    return res.status(400).json({
      success: false,
      processed: 0,
      results: [],
      error: 'Too many files uploaded. Maximum is 20.'
    });
  }

  const results: Array<{
    originalName: string;
    outputFilename?: string;
    size?: number;
    success: boolean;
    error?: string;
    downloadPath?: string;
    storedFilename?: string;
  }> = [];

  let processed = 0;

  for (const file of files) {
    try {
      if (!isCsvFile(file)) {
        throw new Error('Only CSV files are supported for batch conversion');
      }

      let output;
      if (LIBREOFFICE_CONVERSIONS[format]) {
        output = await convertCsvWithLibreOffice(file, format);
      } else if (CALIBRE_CONVERSIONS[format]) {
        output = await convertCsvWithCalibre(file, format);
      } else {
        throw new Error('Unsupported CSV target format for batch conversion');
      }

      const storedFilename = `${Date.now()}_${randomUUID()}_${output.filename}`;
      const storedFilePath = path.join(BATCH_OUTPUT_DIR, storedFilename);

      await fs.writeFile(storedFilePath, output.buffer);
      batchFileMetadata.set(storedFilename, {
        downloadName: output.filename,
        mime: output.mime
      });
      scheduleBatchFileCleanup(storedFilename);

      results.push({
        originalName: file.originalname,
        outputFilename: output.filename,
        size: output.buffer.length,
        success: true,
        downloadPath: `/download/${encodeURIComponent(storedFilename)}`,
        storedFilename
      });

      processed += 1;
    } catch (error) {
      console.error(`Batch conversion failed for ${file.originalname}:`, error);
      results.push({
        originalName: file.originalname,
        success: false,
        error: error instanceof Error ? error.message : 'Unknown conversion error'
      });
    }
  }

  res.json({
    success: results.every(result => result.success),
    processed,
    results
  });
});

app.get('/download/:filename', async (req, res) => {
  try {
    const storedFilename = req.params.filename;
    const metadata = batchFileMetadata.get(storedFilename);

    if (!metadata) {
      return res.status(404).json({ error: 'File not found or expired' });
    }

    const filePath = path.join(BATCH_OUTPUT_DIR, storedFilename);
    const stat = await fs.stat(filePath).catch(() => null);
    if (!stat) {
      batchFileMetadata.delete(storedFilename);
      return res.status(404).json({ error: 'File not found or expired' });
    }

    scheduleBatchFileCleanup(storedFilename);

    res.set({
      'Content-Type': metadata.mime,
      'Content-Disposition': `attachment; filename="${metadata.downloadName}"`,
      'Content-Length': stat.size.toString(),
      'Cache-Control': 'no-cache'
    });

    const stream = (await import('node:fs')).createReadStream(filePath);
    stream.on('error', (error) => {
      console.error('File stream error:', error);
      res.destroy(error);
    });
    stream.pipe(res);
  } catch (error) {
    console.error('Download error:', error);
    res.status(500).json({ error: 'Failed to download file' });
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

