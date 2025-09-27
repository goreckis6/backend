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
  },
  csv: {
    convertTo: 'csv:"Text - txt - csv (StarCalc)"',
    extension: 'csv',
    mime: 'text/csv; charset=utf-8'
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

const isEpubFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  return ext === 'epub' || mimetype.includes('epub');
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
  intermediateExtension?: string;
  postProcessLibreOfficeTarget?: keyof typeof LIBREOFFICE_CONVERSIONS;
}> = {
  mobi: {
    extension: 'mobi',
    mime: 'application/x-mobipocket-ebook'
  },
  doc: {
    extension: 'doc',
    mime: 'application/msword',
    intermediateExtension: 'docx',
    postProcessLibreOfficeTarget: 'doc'
  },
  docx: {
    extension: 'docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  },
  pdf: {
    extension: 'pdf',
    mime: 'application/pdf'
  },
  rtf: {
    extension: 'rtf',
    mime: 'application/rtf'
  },
  odt: {
    extension: 'odt',
    mime: 'application/vnd.oasis.opendocument.text'
  },
  html: {
    extension: 'html',
    mime: 'text/html; charset=utf-8'
  },
  txt: {
    extension: 'txt',
    mime: 'text/plain; charset=utf-8'
  },
  odp: {
    extension: 'odp',
    mime: 'application/vnd.oasis.opendocument.presentation'
  },
  ppt: {
    extension: 'ppt',
    mime: 'application/vnd.ms-powerpoint'
  },
  pptx: {
    extension: 'pptx',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  },
  xlsx: {
    extension: 'xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    intermediateExtension: 'ods',
    postProcessLibreOfficeTarget: 'xlsx'
  },
  csv: {
    extension: 'csv',
    mime: 'text/csv; charset=utf-8',
    intermediateExtension: 'ods',
    postProcessLibreOfficeTarget: 'csv'
  },
  md: {
    extension: 'md',
    mime: 'text/markdown; charset=utf-8'
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

interface CalibreConversionResult {
  buffer: Buffer;
  filename: string;
  mime: string;
  storedFilename?: string;
}

const convertWithCalibre = async (
  file: Express.Multer.File,
  targetFormat: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<CalibreConversionResult> => {
  const conversion = CALIBRE_CONVERSIONS[targetFormat];
  if (!conversion) {
    throw new Error('Unsupported Calibre target format');
  }

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-calibre-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  const intermediateExtension = conversion.intermediateExtension ?? conversion.extension;
  try {
    const inputPath = path.join(tmpDir, `${safeBase}${path.extname(file.originalname) || '.epub'}`);
    await fs.writeFile(inputPath, file.buffer);
    const outputPath = path.join(tmpDir, `${safeBase}.${intermediateExtension}`);

    const args = buildCalibreArgs(inputPath, outputPath, options);

    const { stdout, stderr } = await execCalibre(args);

    if (stdout.trim().length > 0) {
      console.log('Calibre stdout:', stdout.trim());
    }
    if (stderr.trim().length > 0) {
      console.warn('Calibre stderr:', stderr.trim());
    }

    const outputBuffer = await fs.readFile(outputPath);

    if (conversion.postProcessLibreOfficeTarget) {
      const result = await convertBufferWithLibreOffice(
        outputBuffer,
        `.${intermediateExtension}`,
        originalBase,
        conversion.postProcessLibreOfficeTarget,
        options
      );

      if (persistToDisk && result.buffer.length > 0) {
        const storedFilename = `${Date.now()}_${randomUUID()}_${result.filename}`;
        const storedFilePath = path.join(BATCH_OUTPUT_DIR, storedFilename);
        await fs.writeFile(storedFilePath, result.buffer);
        batchFileMetadata.set(storedFilename, {
          downloadName: result.filename,
          mime: result.mime
        });
        scheduleBatchFileCleanup(storedFilename);

        return {
          ...result,
          storedFilename
        };
      }

      return result;
    }

    const downloadName = `${sanitizedBase}.${conversion.extension}`;

    if (persistToDisk) {
      const storedFilename = `${Date.now()}_${randomUUID()}_${downloadName}`;
      const storedFilePath = path.join(BATCH_OUTPUT_DIR, storedFilename);
      await fs.writeFile(storedFilePath, outputBuffer);
      batchFileMetadata.set(storedFilename, {
        downloadName,
        mime: conversion.mime
      });
      scheduleBatchFileCleanup(storedFilename);

      return {
        buffer: outputBuffer,
        filename: downloadName,
        mime: conversion.mime,
        storedFilename
      };
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: conversion.mime
    };
  } catch (error) {
    console.error('Calibre conversion failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown Calibre error';
    throw new Error(`Failed to convert with Calibre: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const convertBufferWithLibreOffice = async (
  buffer: Buffer,
  inputExtension: string,
  originalBase: string,
  targetFormat: keyof typeof LIBREOFFICE_CONVERSIONS,
  options: Record<string, string | undefined> = {}
): Promise<CalibreConversionResult> => {
  const conversion = LIBREOFFICE_CONVERSIONS[targetFormat];
  if (!conversion) {
    throw new Error('Unsupported LibreOffice target format');
  }

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-lo-post-'));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  const normalizedExtension = inputExtension.startsWith('.') ? inputExtension : `.${inputExtension}`;
  const inputFilename = `${safeBase}${normalizedExtension}`;
  const inputPath = path.join(tmpDir, inputFilename);

  try {
    await fs.writeFile(inputPath, buffer);

    const args = [
      '--headless',
      '--nolockcheck',
      '--nodefault',
      '--nologo',
      '--nofirststartwizard',
      ...buildLibreOfficeFilterArgs(options),
      '--convert-to', conversion.convertTo,
      '--outdir', tmpDir,
      inputPath
    ];

    const { stdout, stderr } = await execLibreOffice(args);
    if (stdout.trim().length > 0) {
      console.log('LibreOffice post-process stdout:', stdout.trim());
    }
    if (stderr.trim().length > 0) {
      console.warn('LibreOffice post-process stderr:', stderr.trim());
    }

    const files = await fs.readdir(tmpDir);
    const targetExt = `.${conversion.extension.toLowerCase()}`;
    const outputName = files.find(name => name.toLowerCase().endsWith(targetExt));
    if (!outputName) {
      throw new Error(`LibreOffice did not produce an output file for post-processing to ${conversion.extension}`);
    }

    const outputBuffer = await fs.readFile(path.join(tmpDir, outputName));
    const downloadName = `${sanitizedBase}.${conversion.extension}`;

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: conversion.mime
    };
  } catch (error) {
    console.error('LibreOffice post-processing failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown LibreOffice error';
    throw new Error(`Failed to post-process with LibreOffice: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const buildCalibreArgs = (
  inputPath: string,
  outputPath: string,
  options: Record<string, string | undefined>
): string[] => {
  const args = [
    inputPath,
    outputPath,
    '--change-justification', 'left',
    '--pretty-print',
    '--disable-font-rescaling',
    '--input-encoding', 'utf-8'
  ];

  if (options.bookTitle) args.push('--title', options.bookTitle);
  if (options.author) args.push('--authors', options.author);
  if (options.includeMetadata === 'false') args.push('--no-default-epub-cover');
  if (options.includeImages === 'false') args.push('--disable-dehyphenate');
  if (options.preserveFormatting === 'false') args.push('--disable-font-rescaling');
  if (options.extractTables === 'true') args.push('--extract-tables');
  if (options.delimiter) {
    const delimiterValue = options.delimiter === '\t' ? '\t' : options.delimiter === ';' ? ';' : ',';
    args.push('--csv-input', `delimiter=${delimiterValue}`);
  }
  if (options.pageSize) args.push('--paper-size', options.pageSize);
  if (options.orientation) args.push('--orientation', options.orientation);
  if (options.includeCSS === 'false') args.push('--no-css');
  if (options.responsiveDesign === 'true') args.push('--flow-size', '0');
  if (options.githubCompatible === 'true') args.push('--transform-css', 'github');
  if (options.kindleOptimized === 'true') args.push('--mobi-file-type', 'both');
  if (options.openSourceCompatible === 'true') args.push('--prefer-metadata-author-sort');
  if (options.slideLayout) args.push('--ppt-template', options.slideLayout);
  if (options.enableCollaboration === 'true') args.push('--docx-no-cover');
  if (options.encoding) args.push('--output-encoding', options.encoding);
  if (options.lineEndings) args.push('--line-endings', options.lineEndings);

  return args;
};

const buildLibreOfficeFilterArgs = (
  options: Record<string, string | undefined>
): string[] => {
  if (!options.delimiter) {
    return [];
  }

  const delimiter = options.delimiter === '\t' ? '9' : options.delimiter === ';' ? '59' : '44';
  return [`--infilter=CSV:${delimiter},34,UTF8`];
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
    const requestOptions = { ...(req.body as Record<string, string | undefined>) };

    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log(`Processing ${file.originalname} (${file.size} bytes)`);

    const targetFormat = String(requestOptions.format ?? 'webp').toLowerCase();

    if (isEpubFile(file) && CALIBRE_CONVERSIONS[targetFormat]) {
      const result = await convertWithCalibre(file, targetFormat, requestOptions);

      res.set({
        'Content-Type': result.mime,
        'Content-Disposition': `attachment; filename="${result.filename}"`,
        'Content-Length': result.buffer.length.toString(),
        'Cache-Control': 'no-cache'
      });

      return res.send(result.buffer);
    }

    const quality = requestOptions.quality ?? 'high';
    const lossless = requestOptions.lossless ?? 'false';
    const width = requestOptions.width;
    const height = requestOptions.height;
    const iconSize = requestOptions.iconSize ?? '16';

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
        if (isCsvFile(file) && CALIBRE_CONVERSIONS[targetFormat]) {
          const result = await convertWithCalibre(file, targetFormat, requestOptions);

          res.set({
            'Content-Type': result.mime,
            'Content-Disposition': `attachment; filename="${result.filename}"`,
            'Content-Length': result.buffer.length.toString(),
            'Cache-Control': 'no-cache'
          });

          return res.send(result.buffer);
        }
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
  const requestOptions = { ...(req.body as Record<string, string | undefined>) };
  const format = String(requestOptions.format ?? 'webp').toLowerCase();

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
      let output;
      if (LIBREOFFICE_CONVERSIONS[format] && !isCsvFile(file)) {
        output = await convertCsvWithLibreOffice(file, format);
      } else if (CALIBRE_CONVERSIONS[format] && isEpubFile(file)) {
        output = await convertWithCalibre(file, format, requestOptions, true);
      } else if (CALIBRE_CONVERSIONS[format] && isCsvFile(file)) {
        output = await convertWithCalibre(file, format, requestOptions, true);
      } else if (CALIBRE_CONVERSIONS[format]) {
        output = await convertWithCalibre(file, format, requestOptions, true);
      } else {
        throw new Error('Unsupported target format for batch conversion');
      }

      results.push({
        originalName: file.originalname,
        outputFilename: output.filename,
        size: output.buffer.length,
        success: true,
        downloadPath: output.storedFilename ? `/download/${encodeURIComponent(output.storedFilename)}` : undefined,
        storedFilename: output.storedFilename
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
  const multerModule: any = multer;
  if (multerModule.MulterError && err instanceof multerModule.MulterError) {
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

