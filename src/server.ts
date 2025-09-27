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
import * as XLSX from 'xlsx';

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

const persistOutputBuffer = async (
  buffer: Buffer,
  downloadName: string,
  mime: string
): Promise<ConversionResult> => {
  const storedFilename = `${Date.now()}_${randomUUID()}_${downloadName}`;
  const storedFilePath = path.join(BATCH_OUTPUT_DIR, storedFilename);
  await fs.writeFile(storedFilePath, buffer);
  batchFileMetadata.set(storedFilename, {
    downloadName,
    mime
  });
  scheduleBatchFileCleanup(storedFilename);
  return { buffer, filename: downloadName, mime, storedFilename };
};

const convertEpsFile = async (
  file: Express.Multer.File,
  targetFormat: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-eps-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  const inputFilename = `${safeBase}.eps`;
  const inputPath = path.join(tmpDir, inputFilename);

  try {
    await fs.writeFile(inputPath, file.buffer);

    const quality = options.quality ?? 'high';
    const width = options.width;
    const height = options.height;
    const iconSize = options.iconSize ?? '16';

    const qualityValue = quality === 'high' ? 95 : quality === 'medium' ? 80 : 60;

    let contentType: string;
    let fileExtension: string;
    let outputPath: string;

    switch (targetFormat) {
      case 'webp':
        contentType = 'image/webp';
        fileExtension = 'webp';
        break;
      case 'png':
        contentType = 'image/png';
        fileExtension = 'png';
        break;
      case 'jpeg':
      case 'jpg':
        contentType = 'image/jpeg';
        fileExtension = 'jpg';
        break;
      case 'ico':
        contentType = 'image/x-icon';
        fileExtension = 'ico';
        break;
      default:
        throw new Error('Unsupported output format for EPS conversion');
    }

    outputPath = path.join(tmpDir, `${safeBase}.${fileExtension}`);

    // Use ImageMagick convert command for EPS files
    const convertArgs = [
      inputPath,
      '-density', '300', // High quality conversion
      '-colorspace', 'RGB',
      '-background', 'white',
      '-flatten'
    ];

    // Add quality settings
    if (targetFormat === 'jpeg' || targetFormat === 'jpg') {
      convertArgs.push('-quality', qualityValue.toString());
    }

    // Add size constraints if specified
    if (width || height) {
      const sizeArg = `${width || ''}x${height || ''}`;
      convertArgs.push('-resize', sizeArg);
    }

    // For ICO format, resize to icon size
    if (targetFormat === 'ico') {
      const size = parseInt(iconSize) || 16;
      convertArgs.push('-resize', `${size}x${size}`);
    }

    convertArgs.push(outputPath);

    // Try ImageMagick convert command
    try {
      console.log('Trying ImageMagick convert with args:', convertArgs);
      await execFileAsync('convert', convertArgs);
      console.log('ImageMagick convert successful');
    } catch (convertError) {
      console.warn('ImageMagick convert failed:', convertError);
      // Fallback: try magick command (newer ImageMagick)
      try {
        console.log('Trying magick command with args:', ['convert', ...convertArgs]);
        await execFileAsync('magick', ['convert', ...convertArgs]);
        console.log('Magick command successful');
      } catch (magickError) {
        console.warn('Magick command failed:', magickError);
        // Final fallback: try ghostscript directly
        const gsArgs = [
          '-dNOPAUSE',
          '-dBATCH',
          '-dSAFER',
          '-sDEVICE=png16m',
          `-r${targetFormat === 'ico' ? '72' : '300'}`,
          `-sOutputFile=${outputPath}`,
          inputPath
        ];
        console.log('Trying ghostscript with args:', gsArgs);
        await execFileAsync('gs', gsArgs);
        console.log('Ghostscript successful');
      }
    }

    const outputBuffer = await fs.readFile(outputPath);
    const downloadName = `${sanitizedBase}.${fileExtension}`;

    if (persistToDisk) {
      return persistOutputBuffer(outputBuffer, downloadName, contentType);
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: contentType
    };

  } catch (error) {
    console.error('EPS conversion failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown EPS conversion error';
    throw new Error(`Failed to convert EPS file: ${message}. Please ensure ImageMagick or Ghostscript is installed.`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

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
  // Priority to file extension first, then MIME type, but be more specific about text/plain
  return ext === 'csv' || mimetype.includes('csv') || (mimetype === 'text/plain' && ext !== 'epub');
};

const isEpubFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  // Priority to file extension first, then MIME type
  return ext === 'epub' || mimetype.includes('epub') || mimetype.includes('application/epub');
};

const isEpsFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  return ext === 'eps' || mimetype.includes('eps') || mimetype.includes('postscript');
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
  postProcessExcelTarget?: 'csv' | 'xlsx';
}> = {
  mobi: { extension: 'mobi', mime: 'application/x-mobipocket-ebook' },
  doc: { extension: 'doc', mime: 'application/msword', intermediateExtension: 'docx', postProcessLibreOfficeTarget: 'doc' },
  docx: { extension: 'docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  pdf: { extension: 'pdf', mime: 'application/pdf' },
  rtf: { extension: 'rtf', mime: 'application/rtf' },
  odt: { extension: 'odt', mime: 'application/vnd.oasis.opendocument.text' },
  html: { extension: 'html', mime: 'text/html; charset=utf-8' },
  txt: { extension: 'txt', mime: 'text/plain; charset=utf-8' },
  odp: { extension: 'odp', mime: 'application/vnd.oasis.opendocument.presentation' },
  ppt: { extension: 'ppt', mime: 'application/vnd.ms-powerpoint' },
  pptx: { extension: 'pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
  xlsx: { extension: 'xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', intermediateExtension: 'txt', postProcessExcelTarget: 'xlsx' },
  csv: { extension: 'csv', mime: 'text/csv; charset=utf-8', intermediateExtension: 'txt', postProcessExcelTarget: 'csv' },
  md: { extension: 'md', mime: 'text/markdown; charset=utf-8' }
};

const CALIBRE_CANDIDATES = [
  process.env.CALIBRE_PATH,
  process.env.EBOOK_CONVERT_PATH,
  'ebook-convert',
  'ebook-convert.exe'
].filter((value): value is string => Boolean(value));

interface ConversionResult {
  buffer: Buffer;
  filename: string;
  mime: string;
  storedFilename?: string;
}

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
  targetFormat: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
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

        if (persistToDisk) {
          return persistOutputBuffer(outputBuffer, downloadName, conversion.mime);
        }

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
): Promise<ConversionResult> => {
  console.log(`Calibre conversion requested: ${file.originalname} -> ${targetFormat}`);
  console.log(`Available Calibre conversions:`, Object.keys(CALIBRE_CONVERSIONS));
  
  const conversion = CALIBRE_CONVERSIONS[targetFormat];
  if (!conversion) {
    console.error(`Unsupported Calibre target format: ${targetFormat}`);
    throw new Error(`Unsupported Calibre target format: ${targetFormat}`);
  }
  
  console.log(`Using conversion config:`, conversion);

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
    console.log('Calibre command args:', args);

    try {
      const { stdout, stderr } = await execCalibre(args);

      if (stdout.trim().length > 0) {
        console.log('Calibre stdout:', stdout.trim());
      }
      if (stderr.trim().length > 0) {
        console.warn('Calibre stderr:', stderr.trim());
      }

      const outputBuffer = await fs.readFile(outputPath);
      console.log(`Calibre conversion successful, output size: ${outputBuffer.length} bytes`);
    } catch (calibreError) {
      console.error('Calibre conversion failed:', calibreError);
      
      // Special fallback for HTML conversion
      if (targetFormat === 'html') {
        console.log('Attempting HTML conversion fallback...');
        
        // Try with simpler arguments for HTML
        const fallbackArgs = [
          inputPath,
          outputPath,
          '--no-default-epub-cover',
          '--disable-font-rescaling',
          '--breadth-first'
        ];
        
        console.log('Fallback Calibre args:', fallbackArgs);
        const { stdout: fallbackStdout, stderr: fallbackStderr } = await execCalibre(fallbackArgs);
        
        if (fallbackStdout.trim().length > 0) {
          console.log('Fallback Calibre stdout:', fallbackStdout.trim());
        }
        if (fallbackStderr.trim().length > 0) {
          console.warn('Fallback Calibre stderr:', fallbackStderr.trim());
        }
        
        console.log('HTML fallback conversion successful');
      } else {
        throw calibreError;
      }
    }

    const outputBuffer = await fs.readFile(outputPath);

    if (conversion.postProcessLibreOfficeTarget) {
      console.log(`Post-processing with LibreOffice: ${intermediateExtension} -> ${conversion.postProcessLibreOfficeTarget}, persistToDisk: ${persistToDisk}`);
      const result = await convertBufferWithLibreOffice(
        outputBuffer,
        `.${intermediateExtension}`,
        originalBase,
        conversion.postProcessLibreOfficeTarget,
        options,
        persistToDisk
      );
      console.log(`LibreOffice post-process result: filename=${result.filename}, hasStoredFilename=${!!result.storedFilename}, bufferSize=${result.buffer.length}`);
      return result; // convertBufferWithLibreOffice handles persistToDisk internally
    }

    if (conversion.postProcessExcelTarget) {
      const result = await postProcessToSpreadsheet(
        outputBuffer,
        originalBase,
        conversion.postProcessExcelTarget,
        persistToDisk
      );
      return result;
    }

    const downloadName = `${sanitizedBase}.${conversion.extension}`;

    if (persistToDisk) {
      return persistOutputBuffer(outputBuffer, downloadName, conversion.mime);
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: conversion.mime
    };
  } catch (error) {
    console.error('Calibre conversion failed:', error);
    console.error('File details:', {
      originalName: file.originalname,
      fileSize: file.buffer.length,
      targetFormat,
      hasIntermediateExtension: !!conversion.intermediateExtension,
      hasPostProcess: !!(conversion.postProcessLibreOfficeTarget || conversion.postProcessExcelTarget)
    });
    
    const message = error instanceof Error ? error.message : 'Unknown Calibre error';
    throw new Error(`Failed to convert ${file.originalname} (${targetFormat}): ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const convertBufferWithLibreOffice = async (
  buffer: Buffer,
  inputExtension: string,
  originalBase: string,
  targetFormat: keyof typeof LIBREOFFICE_CONVERSIONS,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
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

    if (persistToDisk) {
      return persistOutputBuffer(outputBuffer, downloadName, conversion.mime);
    }

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

const postProcessToSpreadsheet = async (
  textBuffer: Buffer,
  originalBase: string,
  target: 'csv' | 'xlsx',
  persistToDisk = false
): Promise<ConversionResult> => {
  const sanitizedBase = sanitizeFilename(originalBase);
  const lines = textBuffer.toString('utf8').split(/\r?\n/);
  const data: any[][] = lines.map(line => [line]);
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: target });
  const downloadName = `${sanitizedBase}.${target}`;
  const mime = target === 'xlsx'
    ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    : 'text/csv; charset=utf-8';

  if (persistToDisk) {
    return persistOutputBuffer(buffer as Buffer, downloadName, mime);
  }

  return { buffer: buffer as Buffer, filename: downloadName, mime };
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
  
  // Special handling for HTML conversion to avoid common issues
  if (args[1].endsWith('.html')) {
    args.push('--breadth-first');
    args.push('--max-levels', '5');
    args.push('--max-toc-links', '50');
  }

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
    const isCSV = isCsvFile(file);
    const isEPUB = isEpubFile(file);
    const isEPS = isEpsFile(file);
    console.log(`Single file: isCSV=${isCSV}, isEPUB=${isEPUB}, isEPS=${isEPS}, format=${targetFormat}, mimetype=${file.mimetype}`);

    let result: ConversionResult;

    if (isEPUB && CALIBRE_CONVERSIONS[targetFormat]) {
      console.log('Single: Routing to Calibre (EPUB conversion)');
      result = await convertWithCalibre(file, targetFormat, requestOptions, true);
    } else if (isCSV && LIBREOFFICE_CONVERSIONS[targetFormat]) {
      console.log('Single: Routing to LibreOffice (CSV conversion)');
      result = await convertCsvWithLibreOffice(file, targetFormat, requestOptions, true);
    } else if (isEPS && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(targetFormat)) {
      console.log('Single: Routing to EPS conversion');
      result = await convertEpsFile(file, targetFormat, requestOptions, true);
    } else {
      // Handle Sharp image conversions
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

      result = await persistOutputBuffer(outputBuffer, outputName, contentType);
    }

    // Return download link instead of streaming file
    res.json({
      success: true,
      downloadPath: `/download/${encodeURIComponent(result.storedFilename!)}`,
      filename: result.filename,
      size: result.buffer.length
    });
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
      const isCSV = isCsvFile(file);
      const isEPUB = isEpubFile(file);
      const isEPS = isEpsFile(file);
      console.log(`Processing ${file.originalname}: isCSV=${isCSV}, isEPUB=${isEPUB}, isEPS=${isEPS}, format=${format}, mimetype=${file.mimetype}`);
      
      if (isEPUB && CALIBRE_CONVERSIONS[format]) {
        console.log('Routing to Calibre (EPUB conversion)');
        output = await convertWithCalibre(file, format, requestOptions, true);
      } else if (isCSV && LIBREOFFICE_CONVERSIONS[format]) {
        console.log('Routing to LibreOffice (CSV conversion)');
        output = await convertCsvWithLibreOffice(file, format, requestOptions, true);
      } else if (isEPS && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(format)) {
        console.log('Routing to EPS conversion');
        output = await convertEpsFile(file, format, requestOptions, true);
      } else {
        throw new Error(`Unsupported input file type or target format for batch conversion. File: ${file.originalname}, isCSV: ${isCSV}, isEPUB: ${isEPUB}, isEPS: ${isEPS}, format: ${format}`);
      }

      console.log(`Batch result for ${file.originalname}: filename=${output.filename}, hasStoredFilename=${!!output.storedFilename}, bufferSize=${output.buffer.length}`);
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

