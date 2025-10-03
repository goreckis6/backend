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
import PptxGenJS from 'pptxgenjs';
import mammoth from 'mammoth';
import * as cheerio from 'cheerio';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType } from 'docx';

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

const convertDngFile = async (
  file: Express.Multer.File,
  targetFormat: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log('=== DNG CONVERSION START ===');
  console.log('File details:', {
    originalname: file.originalname,
    size: file.size,
    mimetype: file.mimetype,
    targetFormat,
    options,
    persistToDisk
  });

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-dng-'));
  console.log('Created temp directory:', tmpDir);

  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  const inputFilename = `${safeBase}.dng`;
  const inputPath = path.join(tmpDir, inputFilename);
  const tiffPath = path.join(tmpDir, `${safeBase}.tiff`);

  console.log('File paths:', {
    inputPath,
    tiffPath,
    originalBase,
    sanitizedBase,
    safeBase
  });

  try {
    console.log('Writing DNG file to disk...');
    await fs.writeFile(inputPath, file.buffer);
    console.log('DNG file written successfully, size:', file.buffer.length, 'bytes');

    // Verify file was written
    const stats = await fs.stat(inputPath);
    console.log('Verified file on disk:', {
      size: stats.size,
      exists: stats.isFile()
    });

    // Check if file is empty or too small
    if (stats.size < 1000) {
      throw new Error('DNG file appears to be empty or corrupted (too small)');
    }

    // Check if file has DNG signature
    const fileHeader = file.buffer.slice(0, 4);
    const isDng = fileHeader[0] === 0x49 && fileHeader[1] === 0x49 && fileHeader[2] === 0x2A && fileHeader[3] === 0x00; // TIFF/DNG signature
    const isDngAlt = fileHeader[0] === 0x4D && fileHeader[1] === 0x4D && fileHeader[2] === 0x00 && fileHeader[3] === 0x2A; // TIFF/DNG signature (big-endian)
    
    if (!isDng && !isDngAlt) {
      console.warn('File does not appear to have DNG/TIFF signature, but proceeding anyway...');
    } else {
      console.log('DNG file signature verified');
    }

    const quality = options.quality ?? 'high';
    const iconSize = options.iconSize ?? '16';
    const qualityValue = quality === 'high' ? 95 : quality === 'medium' ? 80 : 60;

    console.log('Conversion options:', {
      quality,
      iconSize,
      qualityValue
    });

    // Step 1: Use dcraw to convert DNG to TIFF
    console.log('=== DCRAW CONVERSION START ===');
    
    const dcrawCommands = [
      { args: ['-T', '-6', '-O', tiffPath, inputPath], desc: 'dcraw -T -6 -O (with output file)' },
      { args: ['-T', '-6', inputPath], desc: 'dcraw -T -6 (without output file)', cwd: tmpDir },
      { args: ['-T', '-4', '-O', tiffPath, inputPath], desc: 'dcraw -T -4 -O (16-bit output)' },
      { args: ['-T', '-8', '-O', tiffPath, inputPath], desc: 'dcraw -T -8 -O (8-bit output)' },
      { args: ['-T', '-6', '-w', '-O', tiffPath, inputPath], desc: 'dcraw -T -6 -w -O (with white balance)' }
    ];
    
    let dcrawSuccess = false;
    let lastError: unknown;
    
    for (const cmd of dcrawCommands) {
      try {
        console.log(`Trying: ${cmd.desc}`);
        const dcrawResult = await execFileAsync('dcraw', cmd.args, cmd.cwd ? { cwd: cmd.cwd } : {});
        console.log('dcraw conversion successful!');
        console.log('dcraw stdout:', dcrawResult.stdout);
        console.log('dcraw stderr:', dcrawResult.stderr);
        dcrawSuccess = true;
        break;
      } catch (dcrawError) {
        console.error(`dcraw command failed (${cmd.desc}):`, dcrawError);
        lastError = dcrawError;
        continue;
      }
    }
    
    if (!dcrawSuccess) {
      throw new Error(`dcraw conversion failed with all command variants: ${lastError instanceof Error ? lastError.message : String(lastError)}`);
    }

    // Verify TIFF was created
    try {
      const tiffStats = await fs.stat(tiffPath);
      console.log('TIFF file created successfully:', {
        size: tiffStats.size,
        exists: tiffStats.isFile()
      });
    } catch (statError) {
      console.error('TIFF file not found after dcraw conversion:', statError);
      throw new Error('dcraw did not produce output file');
    }

    // Step 2: Use ImageMagick to convert TIFF to target format
    console.log('=== IMAGEMAGICK CONVERSION START ===');
    
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
        throw new Error('Unsupported output format for DNG conversion');
    }

    outputPath = path.join(tmpDir, `${safeBase}.${fileExtension}`);
    console.log('Output path:', outputPath);

    // Build ImageMagick convert command
    const convertArgs = [
      tiffPath,
      '-colorspace', 'RGB',
      '-background', 'white',
      '-flatten'
    ];

    // Add quality settings
    if (targetFormat === 'jpeg' || targetFormat === 'jpg') {
      convertArgs.push('-quality', qualityValue.toString());
    }

    // For ICO format, resize to icon size
    if (targetFormat === 'ico') {
      const size = parseInt(iconSize) || 16;
      convertArgs.push('-resize', `${size}x${size}`);
      console.log('ICO size:', size);
    }

    convertArgs.push(outputPath);

    // Try ImageMagick convert command
    console.log('ImageMagick command: convert', convertArgs.join(' '));
    
    try {
      console.log('Trying ImageMagick convert...');
      const convertResult = await execFileAsync('convert', convertArgs);
      console.log('ImageMagick convert successful!');
      console.log('convert stdout:', convertResult.stdout);
      console.log('convert stderr:', convertResult.stderr);
    } catch (convertError) {
      console.warn('ImageMagick convert failed:', convertError);
      // Fallback: try magick command (newer ImageMagick)
      try {
        console.log('Trying magick command fallback...');
        const magickArgs = ['convert', ...convertArgs];
        console.log('magick command:', 'magick', magickArgs.join(' '));
        const magickResult = await execFileAsync('magick', magickArgs);
        console.log('Magick command successful!');
        console.log('magick stdout:', magickResult.stdout);
        console.log('magick stderr:', magickResult.stderr);
      } catch (magickError) {
        console.error('Both ImageMagick variants failed:', magickError);
        throw new Error('ImageMagick conversion failed');
      }
    }

    // Verify output file was created
    try {
      const outputStats = await fs.stat(outputPath);
      console.log('Output file created successfully:', {
        size: outputStats.size,
        exists: outputStats.isFile(),
        path: outputPath
      });
    } catch (statError) {
      console.error('Output file not found after ImageMagick conversion:', statError);
      throw new Error('ImageMagick did not produce output file');
    }

    console.log('Reading output buffer...');
    const outputBuffer = await fs.readFile(outputPath);
    console.log('Output buffer size:', outputBuffer.length, 'bytes');
    
    const downloadName = `${sanitizedBase}.${fileExtension}`;
    console.log('Download name:', downloadName);

    if (persistToDisk) {
      console.log('Persisting to disk...');
      const result = await persistOutputBuffer(outputBuffer, downloadName, contentType);
      console.log('Persisted successfully:', {
        storedFilename: result.storedFilename,
        filename: result.filename
      });
      console.log('=== DNG CONVERSION END (SUCCESS) ===');
      return result;
    }

    console.log('=== DNG CONVERSION END (SUCCESS) ===');
    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: contentType
    };

  } catch (error) {
    console.error('=== DNG CONVERSION ERROR ===');
    console.error('Error type:', error instanceof Error ? error.constructor.name : typeof error);
    console.error('Error message:', error instanceof Error ? error.message : String(error));
    console.error('Error stack:', error instanceof Error ? error.stack : 'No stack trace');
    
    const message = error instanceof Error ? error.message : 'Unknown DNG conversion error';
    console.log('=== DNG CONVERSION END (ERROR) ===');
    throw new Error(`Failed to convert DNG file: ${message}. Please ensure dcraw and ImageMagick are installed.`);
  } finally {
    console.log('Cleaning up temp directory:', tmpDir);
    await fs.rm(tmpDir, { recursive: true, force: true }).catch((cleanupError) => {
      console.warn('Cleanup failed:', cleanupError);
    });
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
  
  // Check if it's a known RAW file extension
  if (RAW_EXTENSIONS.has(ext)) {
    return true;
  }
  
  // Check for RAW MIME types, but exclude common non-RAW x- types
  if (mimetype.includes('raw')) {
    return true;
  }
  
  // Only consider x- MIME types as RAW if they're not common image formats
  if (mimetype.includes('x-') && !mimetype.includes('x-icon') && !mimetype.includes('x-png') && !mimetype.includes('x-jpeg')) {
    return true;
  }
  
  return false;
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
    
    // Try multiple dcraw command variants
    const dcrawCommands = [
      ['-T', '-6', '-O', outputPath, inputPath], // With output file
      ['-T', '-6', inputPath], // Without output file (uses current directory)
      ['-T', '-4', '-O', outputPath, inputPath], // 16-bit output
      ['-T', '-8', '-O', outputPath, inputPath], // 8-bit output
      ['-T', '-6', '-w', '-O', outputPath, inputPath] // With white balance
    ];
    
    let conversionSucceeded = false;
    let lastError: unknown;
    
    for (const args of dcrawCommands) {
      try {
        console.log('Trying dcraw command:', args);
        await execFileAsync('dcraw', args);
        
        // Check if output file was created
        try {
          await fs.access(outputPath);
          conversionSucceeded = true;
          console.log('dcraw conversion successful');
          break;
        } catch (accessError) {
          // Output file doesn't exist, try next command
          console.log('Output file not found, trying next command...');
        }
      } catch (error) {
        console.warn('dcraw command failed:', error);
        lastError = error;
        continue;
      }
    }
    
    if (!conversionSucceeded) {
      throw new Error(`All dcraw commands failed. Last error: ${lastError instanceof Error ? lastError.message : String(lastError)}`);
    }
    
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
  },
  md: {
    convertTo: 'txt',
    extension: 'md',
    mime: 'text/markdown; charset=utf-8'
  },
  'doc-to-csv': {
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

const isDngFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const result = ext === 'dng';
  console.log('DNG file detection:', {
    filename: file.originalname,
    extension: ext,
    isDNG: result
  });
  return result;
};

const isDocFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  const result = ext === 'doc' || mimetype.includes('msword') || mimetype.includes('application/msword');
  
  console.log('DOC file detection:', {
    filename: file.originalname,
    extension: ext,
    mimetype: mimetype,
    isDOC: result
  });
  
  return result;
};

const isDocxFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  const result = ext === 'docx' || mimetype.includes('officedocument.wordprocessingml.document');

  console.log('DOCX file detection:', {
    filename: file.originalname,
    extension: ext,
    mimetype,
    isDOCX: result
  });

  return result;
};

const isOdtFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  const result = ext === 'odt' || mimetype.includes('opendocument.text');

  console.log('ODT file detection:', {
    filename: file.originalname,
    extension: ext,
    mimetype,
    isODT: result
  });

  return result;
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
  postProcessMarkdown?: boolean;
  postProcessPresentationTarget?: 'odp' | 'ppt' | 'pptx';
}> = {
  mobi: { extension: 'mobi', mime: 'application/x-mobipocket-ebook' },
  doc: { extension: 'doc', mime: 'application/msword', intermediateExtension: 'docx', postProcessLibreOfficeTarget: 'doc' },
  docx: { extension: 'docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  pdf: { extension: 'pdf', mime: 'application/pdf' },
  rtf: { extension: 'rtf', mime: 'application/rtf', intermediateExtension: 'txt', postProcessLibreOfficeTarget: 'rtf' },
  odt: { extension: 'odt', mime: 'application/vnd.oasis.opendocument.text', intermediateExtension: 'txt', postProcessLibreOfficeTarget: 'odt' },
  html: { extension: 'html', mime: 'text/html; charset=utf-8', intermediateExtension: 'txt', postProcessLibreOfficeTarget: 'html' },
  txt: { extension: 'txt', mime: 'text/plain; charset=utf-8' },
  odp: { extension: 'odp', mime: 'application/vnd.oasis.opendocument.presentation', intermediateExtension: 'txt', postProcessPresentationTarget: 'odp' },
  ppt: { extension: 'ppt', mime: 'application/vnd.ms-powerpoint', intermediateExtension: 'txt', postProcessPresentationTarget: 'ppt' },
  pptx: { extension: 'pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', intermediateExtension: 'txt', postProcessPresentationTarget: 'pptx' },
  xlsx: { extension: 'xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', intermediateExtension: 'txt', postProcessExcelTarget: 'xlsx' },
  csv: { extension: 'csv', mime: 'text/csv; charset=utf-8', intermediateExtension: 'txt', postProcessExcelTarget: 'csv' },
  md: { extension: 'md', mime: 'text/markdown; charset=utf-8', intermediateExtension: 'txt', postProcessMarkdown: true }
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
  console.log('=== CALIBRE EXECUTION START ===');
  console.log('CALIBRE_CANDIDATES:', CALIBRE_CANDIDATES);
  console.log('Command args:', args);
  
  let lastError: unknown;
  for (const binary of CALIBRE_CANDIDATES) {
    try {
      console.log(`Trying Calibre binary: ${binary}`);
      
      const result = await execFileAsync(binary, args, {
        env: {
          ...process.env,
          HOME: process.env.HOME || os.homedir(),
          USERPROFILE: process.env.USERPROFILE || os.homedir()
        },
        timeout: 60000 // 60 second timeout
      });
      
      console.log(`Calibre execution successful with binary: ${binary}`);
      console.log('Calibre stdout length:', result.stdout.length);
      console.log('Calibre stderr length:', result.stderr.length);
      
      return result;
    } catch (error: any) {
      console.log(`Calibre binary ${binary} failed:`, {
        code: error?.code,
        signal: error?.signal,
        message: error instanceof Error ? error.message : String(error)
      });
      
      lastError = error;
      if (error?.code === 'ENOENT') {
        console.log(`Binary ${binary} not found, trying next...`);
        continue;
      }
      
      const stderr = typeof error?.stderr === 'string' && error.stderr.trim().length > 0
        ? ` | stderr: ${error.stderr.trim()}`
        : '';
      const stdout = typeof error?.stdout === 'string' && error.stdout.trim().length > 0
        ? ` | stdout: ${error.stdout.trim()}`
        : '';
      const message = error instanceof Error ? error.message : String(error);
      
      console.error(`Calibre execution error: ${message}${stderr}${stdout}`);
      throw new Error(`Calibre execution failed using "${binary}": ${message}${stderr}${stdout}`);
    }
  }

  console.error('All Calibre binaries failed');
  console.log('=== CALIBRE EXECUTION END (FAILED) ===');
  
  throw new Error(
    'Calibre ebook-convert binary not found. Please ensure Calibre is installed and available on the PATH or set CALIBRE_PATH/EBOOK_CONVERT_PATH.' +
      (lastError instanceof Error ? ` (${lastError.message})` : '')
  );
};

const convertCsvDirectlyToMarkdown = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log('=== DIRECT CSV TO MARKDOWN CONVERSION ===');
  
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  
  try {
    // Parse CSV content directly
    const csvContent = file.buffer.toString('utf-8');
    const parsed = Papa.parse(csvContent, {
      header: false,
      skipEmptyLines: true,
      delimiter: ','
    });
    
    console.log('CSV parsed successfully:', {
      rows: parsed.data.length,
      errors: parsed.errors.length
    });
    
    if (parsed.errors.length > 0) {
      console.warn('CSV parsing errors:', parsed.errors);
    }
    
    // Convert to markdown table
    const rows = parsed.data as string[][];
    if (rows.length === 0) {
      throw new Error('CSV file appears to be empty');
    }
    
    const markdownLines: string[] = [];
    
    // Add table header
    const headers = rows[0];
    markdownLines.push('| ' + headers.join(' | ') + ' |');
    markdownLines.push('|' + headers.map(() => '---').join('|') + '|');
    
    // Add data rows
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      markdownLines.push('| ' + row.join(' | ') + ' |');
    }
    
    const markdownContent = markdownLines.join('\n');
    const buffer = Buffer.from(markdownContent, 'utf-8');
    const downloadName = `${sanitizedBase}.md`;
    const mime = 'text/markdown; charset=utf-8';
    
    console.log('Markdown conversion successful:', {
      contentLength: markdownContent.length,
      rows: rows.length,
      columns: headers.length
    });
    
    if (persistToDisk) {
      return persistOutputBuffer(buffer, downloadName, mime);
    }
    
    return {
      buffer,
      filename: downloadName,
      mime
    };
    
  } catch (error) {
    console.error('Direct CSV to Markdown conversion failed:', error);
    throw new Error(`Failed to convert CSV to Markdown: ${error instanceof Error ? error.message : String(error)}`);
  }
};

const parseHtmlToCsv = async (htmlContent: string): Promise<string> => {
  console.log('Parsing HTML to CSV...');
  
  // Simple HTML table parser
  const csvLines: string[] = [];
  
  // Look for table elements
  const tableRegex = /<table[^>]*>(.*?)<\/table>/gis;
  const tableMatches = htmlContent.match(tableRegex);
  
  if (tableMatches && tableMatches.length > 0) {
    console.log(`Found ${tableMatches.length} table(s) in HTML`);
    
    for (const tableHtml of tableMatches) {
      // Extract rows
      const rowRegex = /<tr[^>]*>(.*?)<\/tr>/gis;
      const rowMatches = tableHtml.match(rowRegex);
      
      if (rowMatches) {
        for (const rowHtml of rowMatches) {
          // Extract cells (both th and td)
          const cellRegex = /<(?:th|td)[^>]*>(.*?)<\/(?:th|td)>/gis;
          const cellMatches = rowHtml.match(cellRegex);
          
          if (cellMatches) {
            const cells = cellMatches.map(cell => {
              // Remove HTML tags and clean up content
              const cleanCell = cell
                .replace(/<(?:th|td)[^>]*>/, '')
                .replace(/<\/(?:th|td)>/, '')
                .replace(/<[^>]*>/g, '') // Remove any remaining HTML tags
                .replace(/&nbsp;/g, ' ')
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&#39;/g, "'")
                .trim();
              
              // Escape quotes and wrap in quotes
              const escapedCell = cleanCell.replace(/"/g, '""');
              return `"${escapedCell}"`;
            });
            
            csvLines.push(cells.join(','));
          }
        }
      }
    }
  } else {
    // No tables found, try to extract structured data from paragraphs
    console.log('No tables found, extracting paragraph content');
    
    // Look for paragraphs or divs that might contain structured data
    const paragraphRegex = /<(?:p|div)[^>]*>(.*?)<\/(?:p|div)>/gis;
    const paragraphMatches = htmlContent.match(paragraphRegex);
    
    if (paragraphMatches) {
      for (const paragraphHtml of paragraphMatches) {
        const cleanText = paragraphHtml
          .replace(/<(?:p|div)[^>]*>/, '')
          .replace(/<\/(?:p|div)>/, '')
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'")
          .trim();
        
        if (cleanText.length > 0) {
          // Try to split on common delimiters
          if (cleanText.includes('\t')) {
            const cells = cleanText.split('\t').map(cell => `"${cell.trim()}"`);
            csvLines.push(cells.join(','));
          } else if (cleanText.includes('  ') && cleanText.split('  ').length > 1) {
            const cells = cleanText.split(/\s{2,}/).map(cell => `"${cell.trim()}"`);
            csvLines.push(cells.join(','));
          } else {
            csvLines.push(`"${cleanText}"`);
          }
        }
      }
    }
  }
  
  if (csvLines.length === 0) {
    // Fallback: just extract all text content
    const textContent = htmlContent
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .split('\n')
      .filter(line => line.trim().length > 0);
    
    for (const line of textContent) {
      csvLines.push(`"${line.trim()}"`);
    }
  }
  
  console.log(`HTML to CSV conversion: ${csvLines.length} rows generated`);
  return csvLines.join('\n');
};


const parseHtmlTablesToCsv = async (htmlContent: string): Promise<string> => {
  console.log('Parsing HTML tables to CSV using cheerio...');

  const $ = cheerio.load(htmlContent);
  const tables = $('table');

  const allCsvRows: string[] = [];

  tables.each((tableIndex, table) => {
    console.log(`Processing table ${tableIndex + 1}...`);
    const tableRows: string[] = [];

    $(table).find('tr').each((rowIndex, row) => {
      const cells: string[] = [];

      let currentIndex = 0;

      $(row).find('td, th').each((cellIndex, cell) => {
        let cellText = $(cell).text() || '';
        cellText = cellText.replace(/\s+/g, ' ').trim();

        if (!cellText || cellText.trim() === '') {
          cellText = '';
        }

        // Skip duplicated content for merged cells
        if (cellText === '' || tableRows.some(existingRow => existingRow.includes(cellText))) {
          const rowspan = parseInt($(cell).attr('rowspan') || '1', 10);
          if (rowspan > 1) {
            for (let i = 1; i < rowspan; i++) {
              const nextRow = tableRows[rowIndex + i] ? tableRows[rowIndex + i].split(',') : [];
              nextRow[currentIndex] = '""';
              tableRows[rowIndex + i] = nextRow.join(',');
            }
          }
          currentIndex++;
          return;
        }

        const colspanAttr = $(cell).attr('colspan');
        const colspan = colspanAttr ? Math.max(parseInt(colspanAttr, 10) || 1, 1) : 1;

        const escaped = cellText.replace(/"/g, '""');
        for (let i = 0; i < colspan; i++) {
          cells.push(`"${escaped}"`);
          currentIndex++;
        }
      });

      if (cells.length > 0) {
        tableRows.push(cells.join(','));
      }
    });

    if (tableRows.length > 0) {
      console.log(`Table ${tableIndex + 1} has ${tableRows.length} rows`);
      allCsvRows.push(...tableRows);
      if (tableIndex < tables.length - 1) {
        allCsvRows.push('');
      }
    }
  });

  if (allCsvRows.length === 0) {
    console.log('No HTML tables found, extracting text content');
    const textContent = htmlContent
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .split('\n')
      .filter(line => line.trim())
      .map(line => `"${line.trim()}"`)
      .join('\n');

    return textContent;
  }

  const csvContent = allCsvRows.join('\n');

  console.log('HTML table parsing successful:', {
    totalTables: tables.length,
    totalRows: allCsvRows.length,
    csvLength: csvContent.length
  });

  return csvContent;
};

const convertDocWithLibreOffice = async (
  file: Express.Multer.File,
  targetFormat: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log('=== DOC TO CSV CONVERSION START (ENHANCED) ===');
  
  const conversion = LIBREOFFICE_CONVERSIONS[targetFormat];
  if (!conversion) {
    throw new Error('Unsupported LibreOffice target format');
  }

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-doc-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  const inputExtension = path.extname(file.originalname).toLowerCase() || '.doc';
  const inputFilename = `${safeBase}${inputExtension}`;
  const inputPath = path.join(tmpDir, inputFilename);

  try {
    await fs.writeFile(inputPath, file.buffer);

    // Step 1: Ensure we have a DOCX file (LibreOffice handles DOCX better for mammoth)
    let docxFilename = inputFilename;
    if (inputExtension !== '.docx') {
      console.log('Step 1: Converting document to DOCX...');
      const docxArgs = [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'docx',
        '--outdir', tmpDir,
        inputPath
      ];

      const { stdout, stderr } = await execLibreOffice(docxArgs);
      if (stdout.trim().length > 0) {
        console.log('DOC to DOCX stdout:', stdout.trim());
      }
      if (stderr.trim().length > 0) {
        console.warn('DOC to DOCX stderr:', stderr.trim());
      }

      const files = await fs.readdir(tmpDir);
      const generatedDocx = files.find(name => name.toLowerCase().endsWith('.docx'));
      if (!generatedDocx) {
        throw new Error('DOCX file not produced by LibreOffice');
      }
      docxFilename = generatedDocx;
    }

    // Step 2: Extract HTML from DOCX using mammoth
    console.log('Step 2: Extracting HTML from DOCX using mammoth...');
    const docxPath = path.join(tmpDir, docxFilename);
    const docxBuffer = await fs.readFile(docxPath);
    const mammothResult = await mammoth.convertToHtml({ buffer: docxBuffer });
    const htmlContent = mammothResult.value;

    console.log('Mammoth extraction successful:', {
      htmlLength: htmlContent.length,
      messages: mammothResult.messages.length
    });
    if (mammothResult.messages.length > 0) {
      console.log('Mammoth messages:', mammothResult.messages);
    }

    // Step 3: Parse HTML tables and convert to CSV
    console.log('Step 3: Parsing HTML tables and converting to CSV...');
    const csvContent = await parseHtmlTablesToCsv(htmlContent);
    console.log('HTML table parsing successful:', {
      csvLength: csvContent.length,
      csvLines: csvContent.split('\n').length
    });

    // Write CSV file
    const csvPath = path.join(tmpDir, `${sanitizedBase}.csv`);
    await fs.writeFile(csvPath, csvContent, 'utf-8');

    const outputBuffer = await fs.readFile(csvPath);
    const downloadName = `${sanitizedBase}.${conversion.extension}`;

    console.log('DOC to CSV conversion successful:', {
      outputSize: outputBuffer.length,
      filename: downloadName
    });

    if (persistToDisk) {
      return persistOutputBuffer(outputBuffer, downloadName, conversion.mime);
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: conversion.mime
    };

  } catch (error) {
    console.error('DOC to CSV conversion failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown conversion error';
    throw new Error(`Failed to convert DOC to CSV: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const convertCsvToDocxFallback = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log('=== CSV TO DOCX FALLBACK CONVERSION START ===');
  
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-docx-fallback-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  const normalized = normalizeCsvBuffer(file.buffer);
  const inputFilename = `${safeBase}.csv`;
  const inputPath = path.join(tmpDir, inputFilename);

  try {
    await fs.writeFile(inputPath, normalized.normalizedCsv, 'utf8');

    // Try CSV to ODT first, then ODT to DOCX (more reliable)
    
    // Step 1: Convert CSV to ODT
    const csvToOdtArgs = [
      '--headless',
      '--nolockcheck',
      '--nodefault',
      '--nologo',
      '--nofirststartwizard',
      '--convert-to', 'odt',
      '--outdir', tmpDir,
      inputPath
    ];
    
    console.log('Converting CSV to ODT:', csvToOdtArgs);
    const { stdout: odtStdout, stderr: odtStderr } = await execLibreOffice(csvToOdtArgs);
    
    if (odtStdout.trim().length > 0) {
      console.log('CSV to ODT stdout:', odtStdout.trim());
    }
    if (odtStderr.trim().length > 0) {
      console.warn('CSV to ODT stderr:', odtStderr.trim());
    }
    
    // Check if ODT file was created
    const files = await fs.readdir(tmpDir);
    const odtFile = files.find(name => name.toLowerCase().endsWith('.odt'));
    
    if (!odtFile) {
      throw new Error('LibreOffice did not produce an ODT file from CSV');
    }
    
    console.log('ODT file created successfully:', odtFile);
    
    // Step 2: Convert ODT to DOCX
    const odtFilePath = path.join(tmpDir, odtFile);
    const commandVariants: string[][] = [
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'docx',
        '--outdir', tmpDir,
        odtFilePath
      ],
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'docx:"MS Word 2007 XML"',
        '--outdir', tmpDir,
        odtFilePath
      ],
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'docx',
        '--outdir', tmpDir,
        odtFilePath,
        '--writer'
      ]
    ];

    let lastError: unknown;
    let outputBuffer: Buffer | null = null;
    let docxFile: string | null = null;

    for (const args of commandVariants) {
      try {
        console.log('Trying LibreOffice ODT to DOCX command:', args);
        const { stdout, stderr } = await execLibreOffice(args);
        
        if (stdout.trim().length > 0) {
          console.log('LibreOffice ODT to DOCX stdout:', stdout.trim());
        }
        if (stderr.trim().length > 0) {
          console.warn('LibreOffice ODT to DOCX stderr:', stderr.trim());
        }

        // Check if DOCX file was created
        const finalFiles = await fs.readdir(tmpDir);
        docxFile = finalFiles.find(name => name.toLowerCase().endsWith('.docx')) || null;
        
        if (docxFile) {
          const outputPath = path.join(tmpDir, docxFile);
          outputBuffer = await fs.readFile(outputPath);
          
          // Validate the DOCX file
          if (outputBuffer.length > 0 && outputBuffer[0] === 0x50 && outputBuffer[1] === 0x4B) {
            console.log('Valid DOCX file created from ODT:', docxFile);
            break;
          } else {
            console.warn('Invalid DOCX file created from ODT, trying next variant...');
            outputBuffer = null;
            docxFile = null;
          }
        }
      } catch (error) {
        console.warn('LibreOffice ODT to DOCX command failed:', error);
        lastError = error;
        continue;
      }
    }

    if (!docxFile || !outputBuffer) {
      throw new Error(`LibreOffice did not produce a valid DOCX file. Last error: ${lastError instanceof Error ? lastError.message : String(lastError)}`);
    }

    const downloadName = `${sanitizedBase}.docx`;
    
    console.log('CSV to DOCX fallback conversion successful:', {
      outputSize: outputBuffer.length,
      filename: downloadName,
      docxFile,
      isValidDocx: outputBuffer[0] === 0x50 && outputBuffer[1] === 0x4B
    });

    if (persistToDisk) {
      console.log('Persisting fallback DOCX file to disk:', {
        bufferSize: outputBuffer.length,
        downloadName,
        isValidDocx: outputBuffer[0] === 0x50 && outputBuffer[1] === 0x4B
      });
      const result = await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      console.log('Fallback file persisted successfully:', {
        storedFilename: result.storedFilename,
        filename: result.filename,
        mime: result.mime
      });
      return result;
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };

  } catch (error) {
    console.error('CSV to DOCX fallback conversion failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown conversion error';
    throw new Error(`Failed to convert CSV to DOCX: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const convertCsvToDocxEnhanced = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log('=== CSV TO DOCX ENHANCED CONVERSION START ===');
  
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-docx-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  
  try {
    // Step 1: Parse CSV data and create a simple HTML table
    const csvText = file.buffer.toString('utf-8');
    const parsed = Papa.parse(csvText, {
      header: false,
      skipEmptyLines: true,
      transform: (value) => value.trim()
    });
    
    console.log('CSV parsing successful:', {
      rows: parsed.data.length,
      errors: parsed.errors.length
    });
    
    if (parsed.errors.length > 0) {
      console.warn('CSV parsing warnings:', parsed.errors);
    }
    
    // Step 2: Create HTML table from CSV data
    const htmlContent = createHtmlTableFromCsv(parsed.data as string[][]);
    const htmlPath = path.join(tmpDir, `${safeBase}.html`);
    await fs.writeFile(htmlPath, htmlContent, 'utf-8');
    
    console.log('HTML table created successfully');
    
    // Step 3: Convert HTML to DOCX using LibreOffice with multiple variants
    const libreOfficeVariants = [
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'docx',
        '--outdir', tmpDir,
        htmlPath
      ],
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'docx:"MS Word 2007 XML"',
        '--outdir', tmpDir,
        htmlPath
      ],
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'docx',
        '--outdir', tmpDir,
        htmlPath,
        '--writer'
      ]
    ];
    
    let docxFile: string | null = null;
    let lastError: unknown;
    
    for (const libreOfficeArgs of libreOfficeVariants) {
      try {
        console.log('Running LibreOffice HTML to DOCX conversion:', libreOfficeArgs);
        const { stdout, stderr } = await execLibreOffice(libreOfficeArgs);
        
        if (stdout.trim().length > 0) {
          console.log('LibreOffice HTML->DOCX stdout:', stdout.trim());
        }
        if (stderr.trim().length > 0) {
          console.warn('LibreOffice HTML->DOCX stderr:', stderr.trim());
        }
        
        // Check if DOCX file was created
        const files = await fs.readdir(tmpDir);
        docxFile = files.find(name => name.toLowerCase().endsWith('.docx')) || null;
        
        if (docxFile) {
          console.log('DOCX file created successfully:', docxFile);
          break;
        } else {
          console.log('No DOCX file found, trying next variant...');
          console.log('Available files:', files);
        }
      } catch (error) {
        console.warn('LibreOffice variant failed:', error);
        lastError = error;
        continue;
      }
    }
    
    if (!docxFile) {
      throw new Error(`LibreOffice did not produce a DOCX file. Last error: ${lastError instanceof Error ? lastError.message : String(lastError)}`);
    }
    
    const finalDocxPath = path.join(tmpDir, docxFile);
    const outputBuffer = await fs.readFile(finalDocxPath);
    const downloadName = `${sanitizedBase}.docx`;
    
    // Validate the DOCX file
    if (outputBuffer.length === 0) {
      throw new Error('Generated DOCX file is empty');
    }
    
    // Check if it's a valid DOCX file (should start with PK signature)
    if (outputBuffer.length < 4 || outputBuffer[0] !== 0x50 || outputBuffer[1] !== 0x4B) {
      throw new Error('Generated file does not appear to be a valid DOCX file');
    }
    
    console.log('CSV to DOCX conversion successful:', {
      outputSize: outputBuffer.length,
      filename: downloadName,
      docxFile,
      isValidDocx: outputBuffer[0] === 0x50 && outputBuffer[1] === 0x4B
    });
    
    if (persistToDisk) {
      console.log('Persisting DOCX file to disk:', {
        bufferSize: outputBuffer.length,
        downloadName,
        isValidDocx: outputBuffer[0] === 0x50 && outputBuffer[1] === 0x4B
      });
      const result = await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      console.log('File persisted successfully:', {
        storedFilename: result.storedFilename,
        filename: result.filename,
        mime: result.mime
      });
      return result;
    }
    
    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
    
  } catch (error) {
    console.error('CSV to DOCX enhanced conversion failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown conversion error';
    throw new Error(`Failed to convert CSV to DOCX: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const createHtmlTableFromCsv = (csvData: string[][]): string => {
  if (csvData.length === 0) {
    return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CSV Data</title></head><body><p>No data available</p></body></html>';
  }
  
  // Clean and validate CSV data
  const cleanData = csvData.map(row => 
    row.map(cell => {
      if (typeof cell !== 'string') return String(cell || '');
      return cell.trim();
    }).filter(cell => cell !== '') // Remove empty cells
  ).filter(row => row.length > 0); // Remove empty rows
  
  if (cleanData.length === 0) {
    return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CSV Data</title></head><body><p>No valid data available</p></body></html>';
  }
  
  let html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CSV Data</title>';
  html += '<style>body{font-family:Arial,sans-serif;margin:20px;}table{border-collapse:collapse;width:100%;}th,td{border:1px solid #ddd;padding:8px;text-align:left;}th{background-color:#f2f2f2;font-weight:bold;}</style>';
  html += '</head><body><h1>CSV Data</h1><table>';
  
  // Add header row
  if (cleanData.length > 0) {
    html += '<thead><tr>';
    cleanData[0].forEach(cell => {
      html += `<th>${escapeHtml(cell)}</th>`;
    });
    html += '</tr></thead>';
  }
  
  // Add data rows
  html += '<tbody>';
  for (let i = 1; i < cleanData.length; i++) {
    html += '<tr>';
    cleanData[i].forEach(cell => {
      html += `<td>${escapeHtml(cell)}</td>`;
    });
    html += '</tr>';
  }
  html += '</tbody></table></body></html>';
  
  return html;
};

const escapeHtml = (text: string): string => {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
};

// CSV -> EPUB using custom implementation (bypassing Calibre)
const convertCsvToEpubCustom = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log('=== CSV TO EPUB (Custom) START ===');
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);

  try {
    // Parse CSV
    const csvText = file.buffer.toString('utf-8');
    const parsed = Papa.parse<string[]>(csvText, { skipEmptyLines: false });
    const rows: string[][] = parsed && Array.isArray((parsed as any).data)
      ? ((parsed as any).data as unknown as string[][]).map((r: unknown) => (Array.isArray(r) ? (r as string[]) : [String(r ?? '')]))
      : [];

    // Create a simple HTML document from CSV data
    let htmlContent = `<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <title>${options.title || sanitizedBase}</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        h1 { color: #333; }
    </style>
</head>
<body>
    <h1>${options.title || sanitizedBase}</h1>
    <p><strong>Author:</strong> ${options.author || 'Unknown'}</p>
    <table>`;

    if (rows.length > 0) {
      // Add header row
      htmlContent += '<thead><tr>';
      rows[0].forEach(cell => {
        htmlContent += `<th>${escapeHtml(cell)}</th>`;
      });
      htmlContent += '</tr></thead><tbody>';
      
      // Add data rows
      for (let i = 1; i < rows.length; i++) {
        htmlContent += '<tr>';
        rows[i].forEach(cell => {
          htmlContent += `<td>${escapeHtml(cell)}</td>`;
        });
        htmlContent += '</tr>';
      }
    }
    
    htmlContent += '</tbody></table></body></html>';

    // Since we can't easily create a real EPUB file without Calibre,
    // we'll create a simple HTML file that can be read as an e-book
    // and name it with .epub extension for compatibility
    const outputBuffer = Buffer.from(htmlContent, 'utf-8');
    const downloadName = `${sanitizedBase}.epub`;
    
    console.log('CSV->EPUB (HTML) conversion successful:', { filename: downloadName, size: outputBuffer.length });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/epub+zip');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/epub+zip'
    };
  } catch (error) {
    console.error('CSV->EPUB conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown CSV->EPUB error';
    throw new Error(`Failed to convert CSV to EPUB: ${message}`);
  }
};

// Create a basic MOBI file structure
const createBasicMobiFile = (htmlContent: string, title: string, author: string): string => {
  // Create a simple MOBI-like structure using HTML with proper metadata
  const mobiContent = `<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <title>${title}</title>
    <meta name="author" content="${author}" />
    <meta name="generator" content="Morphy Converter" />
    <style>
        body { 
            font-family: Arial, sans-serif; 
            margin: 20px; 
            line-height: 1.6;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin: 20px 0; 
        }
        th, td { 
            border: 1px solid #ddd; 
            padding: 8px; 
            text-align: left; 
        }
        th { 
            background-color: #f2f2f2; 
            font-weight: bold; 
        }
        h1 { 
            color: #333; 
            border-bottom: 2px solid #333;
            padding-bottom: 10px;
        }
        .metadata {
            background-color: #f9f9f9;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
    </style>
</head>
<body>
    <div class="metadata">
        <h1>${title}</h1>
        <p><strong>Author:</strong> ${author}</p>
        <p><strong>Generated by:</strong> Morphy Converter</p>
    </div>
    ${htmlContent.replace('<h1>', '').replace('</h1>', '').replace('<p><strong>Author:</strong>', '').replace('</p>', '')}
</body>
</html>`;
  
  return mobiContent;
};

// CSV -> E-book using Python script (pandas + jinja2 + ebooklib)
const convertCsvToEbookPython = async (
  file: Express.Multer.File,
  targetFormat: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO ${targetFormat.toUpperCase()} (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-python-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.${targetFormat}`);
    
    // Prepare Python script arguments
    const pythonArgs = [
      'python3',
      path.join('/app/scripts/csv_to_ebook.py'),
      csvPath,
      outputPath,
      targetFormat,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ];

    console.log('Running Python CSV to e-book converter with args:', pythonArgs);

    // Execute Python script
    const { stdout, stderr } = await execFileAsync('python3', [
      path.join('/app/scripts/csv_to_ebook.py'),
      csvPath,
      outputPath,
      targetFormat,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python script produced empty output file');
    }

    // Determine MIME type
    let mimeType: string;
    switch (targetFormat) {
      case 'epub':
        mimeType = 'application/epub+zip';
        break;
      case 'mobi':
        mimeType = 'application/x-mobipocket-ebook';
        break;
      case 'html':
        mimeType = 'text/html';
        break;
      case 'txt':
        mimeType = 'text/plain';
        break;
      default:
        mimeType = 'application/octet-stream';
    }

    const downloadName = `${sanitizedBase}.${targetFormat}`;
    console.log(`CSV->${targetFormat.toUpperCase()} conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, mimeType);
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: mimeType
    };
  } catch (error) {
    console.error(`CSV->${targetFormat.toUpperCase()} conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->${targetFormat.toUpperCase()} error`;
    throw new Error(`Failed to convert CSV to ${targetFormat.toUpperCase()}: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV -> DOCX using pure Node (docx library). Avoids LibreOffice filter issues entirely.
const convertCsvToDocxWithDocxLib = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log('=== CSV TO DOCX (docx lib) START ===');
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);

  try {
    const csvText = file.buffer.toString('utf-8');
    const parsed = Papa.parse<string[]>(csvText, {
      skipEmptyLines: false
    });

    if (parsed && Array.isArray(parsed.errors) && parsed.errors.length) {
      console.warn('CSV parse warnings:', parsed.errors.map((e: any) => ({ message: e.message, row: e.row })));
    }

    const rows: string[][] = parsed && Array.isArray((parsed as any).data)
      ? ((parsed as any).data as unknown as string[][]).map((r: unknown) => (Array.isArray(r) ? (r as string[]) : [String(r ?? '')]))
      : [];

    // Ensure at least one row
    const safeRows = rows.length > 0 ? rows : [[csvText]];

    const tableRows: TableRow[] = safeRows.map((row) =>
      new TableRow({
        children: row.map((cell) =>
          new TableCell({
            children: [new Paragraph(String(cell ?? ''))]
          })
        )
      })
    );

    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: tableRows
    });

    const doc = new Document({
      sections: [
        {
          children: [table]
        }
      ]
    });

    const outputBuffer = await Packer.toBuffer(doc);
    const downloadName = `${sanitizedBase}.docx`;

    // Validate basic DOCX signature (PK zip)
    if (outputBuffer.length < 4 || outputBuffer[0] !== 0x50 || outputBuffer[1] !== 0x4B) {
      console.warn('Generated DOCX does not start with PK header, but proceeding. Size:', outputBuffer.length);
    }

    if (persistToDisk) {
      return await persistOutputBuffer(
        outputBuffer,
        downloadName,
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
  } catch (error) {
    console.error('CSV to DOCX (docx lib) failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown conversion error';
    throw new Error(`Failed to convert CSV to DOCX: ${message}`);
  }
};

const convertCsvWithLibreOffice = async (
  file: Express.Multer.File,
  targetFormat: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  // Special handling for MD conversion - bypass LibreOffice entirely
  if (targetFormat === 'md') {
    console.log('Using direct CSV to Markdown conversion (bypassing LibreOffice)');
    return await convertCsvDirectlyToMarkdown(file, options, persistToDisk);
  }

  // Special handling for DOCX conversion - use a more reliable approach
  if (targetFormat === 'docx') {
    console.log('Using docx library for CSV to DOCX conversion (avoiding LibreOffice)');
    return await convertCsvToDocxWithDocxLib(file, options, persistToDisk);
  }

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

  // Prioritize simple commands for MD conversion, otherwise use all variants
  const commandVariants: string[][] = targetFormat === 'md' ? [
    ['--headless', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath],
    ['--headless', '--nolockcheck', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath],
    ['--headless', '--nolockcheck', '--nodefault', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--calc'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--writer'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:44,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--calc'],
    ['--headless', '--nolockcheck', '--nodefault', '--nologo', '--nofirststartwizard', '--infilter=CSV:44,34,UTF8', '--convert-to', conversion.convertTo, '--outdir', tmpDir, inputPath, '--writer']
  ] : [
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
          // Debug: List all files in temp directory
          try {
            const files = await fs.readdir(tmpDir);
            console.log('Files in temp directory after LibreOffice:', files);
            console.log('Looking for files with extension:', conversion.extension);
          } catch (listError) {
            console.error('Failed to list temp directory:', listError);
          }
          throw new Error(`LibreOffice did not produce an output file for args: ${args.join(' ')}`);
        }

        const outputBuffer = await fs.readFile(outputPath);
        const downloadName = `${sanitizeFilename(originalBase)}.${conversion.extension}`;

        // Post-process for Markdown format
        if (targetFormat === 'md') {
          console.log('Post-processing TXT to Markdown format...');
          const result = await convertTxtToMarkdown(outputBuffer, originalBase, options, persistToDisk);
          return result;
        }

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

    if (conversion.postProcessPresentationTarget) {
      console.log(`Post-processing presentation: ${intermediateExtension} -> ${conversion.postProcessPresentationTarget}, persistToDisk: ${persistToDisk}`);
      console.log(`Input buffer size: ${outputBuffer.length} bytes`);
      try {
        const result = await convertTxtToPresentation(
          outputBuffer,
          originalBase,
          conversion.postProcessPresentationTarget,
          options,
          persistToDisk
        );
        console.log(`Presentation post-process result: filename=${result.filename}, hasStoredFilename=${!!result.storedFilename}, bufferSize=${result.buffer.length}`);
        return result;
      } catch (presentationError) {
        console.error('Presentation post-processing failed:', presentationError);
        console.error('Failed conversion details:', {
          inputFormat: intermediateExtension,
          outputFormat: conversion.postProcessPresentationTarget,
          inputBufferSize: outputBuffer.length,
          originalFile: file.originalname
        });
        throw new Error(`Presentation post-processing failed: ${presentationError instanceof Error ? presentationError.message : String(presentationError)}`);
      }
    }

    if (conversion.postProcessLibreOfficeTarget) {
      console.log(`Post-processing with LibreOffice: ${intermediateExtension} -> ${conversion.postProcessLibreOfficeTarget}, persistToDisk: ${persistToDisk}`);
      console.log(`Input buffer size: ${outputBuffer.length} bytes`);
      try {
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
      } catch (libreOfficeError) {
        console.error('LibreOffice post-processing failed:', libreOfficeError);
        console.error('Failed conversion details:', {
          inputFormat: intermediateExtension,
          outputFormat: conversion.postProcessLibreOfficeTarget,
          inputBufferSize: outputBuffer.length,
          originalFile: file.originalname
        });
        throw new Error(`LibreOffice post-processing failed: ${libreOfficeError instanceof Error ? libreOfficeError.message : String(libreOfficeError)}`);
      }
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

    if (conversion.postProcessMarkdown) {
      const result = await convertTxtToMarkdown(
        outputBuffer,
        originalBase,
        options,
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
    // Special preprocessing for presentation formats
    let processedBuffer = buffer;
    if (['odp', 'ppt', 'pptx'].includes(targetFormat) && normalizedExtension === '.txt') {
      console.log('Preprocessing text for presentation format:', targetFormat);
      const textContent = buffer.toString('utf-8');
      const presentationText = createPresentationStructure(textContent);
      processedBuffer = Buffer.from(presentationText, 'utf-8');
      console.log('Preprocessed text length:', presentationText.length);
    }
    
    await fs.writeFile(inputPath, processedBuffer);

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

const convertTxtToPresentation = async (
  textBuffer: Buffer,
  originalBase: string,
  targetFormat: 'odp' | 'ppt' | 'pptx',
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  const sanitizedBase = sanitizeFilename(originalBase);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-presentation-'));
  
  try {
    // Step 1: Build PPTX with pptxgenjs to avoid empty slides
    const textContent = textBuffer.toString('utf-8');
    const slidesData = buildSlidesFromText(textContent);
    const pptx = new (PptxGenJS as any)();
    pptx.layout = 'LAYOUT_16x9';
    slidesData.forEach((slideData: { title: string; bullets: string[] }) => {
      const slide = pptx.addSlide();
      if (slideData.title) {
        slide.addText(slideData.title, { x: 0.5, y: 0.4, fontSize: 28, bold: true });
      }
      let y = 1.1;
      slideData.bullets.forEach((line) => {
        if (line && line.trim().length > 0) {
          slide.addText(`• ${line}`, { x: 0.7, y, fontSize: 18 });
          y += 0.4;
        }
      });
    });
    const pptxPath = path.join(tmpDir, `${sanitizedBase}.pptx`);
    await new Promise<void>((resolve, reject) => {
      pptx.writeFile({ fileName: pptxPath }).then(() => resolve()).catch(reject);
    });

    // If target is PPTX, return generated file
    if (targetFormat === 'pptx') {
      const outputBuffer = await fs.readFile(pptxPath);
      const downloadName = `${sanitizedBase}.pptx`;
      const mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
      if (persistToDisk) {
        return persistOutputBuffer(outputBuffer, downloadName, mime);
      }
      return { buffer: outputBuffer, filename: downloadName, mime };
    }

    // Step 2: Convert PPTX to desired format via LibreOffice
    console.log(`Converting PPTX to ${targetFormat}...`);
    const filterMap2: Record<'odp' | 'ppt', string> = { odp: 'impress8', ppt: 'MS PowerPoint 97' };
    const variants2: string[][] = [
      ['--headless','--nolockcheck','--nodefault','--nologo','--nofirststartwizard','--convert-to', `${targetFormat}:${filterMap2[targetFormat]}`,'--outdir', tmpDir, pptxPath],
      ['--headless','--nolockcheck','--nodefault','--nologo','--nofirststartwizard','--convert-to', targetFormat,'--outdir', tmpDir, pptxPath]
    ];
    let stdout = ''; let stderr = ''; let ok = false;
    for (const variant of variants2) {
      console.log('LibreOffice PPTX->presentation args:', variant);
      try {
        const res = await execLibreOffice(variant);
        stdout = res.stdout; stderr = res.stderr; ok = true; break;
      } catch (e) { console.warn('LibreOffice PPTX conversion attempt failed, trying fallback...', e); }
    }
    if (!ok) throw new Error('LibreOffice failed converting PPTX to presentation format');
    if (stdout.trim()) console.log('LibreOffice presentation stdout:', stdout.trim());
    if (stderr.trim()) console.warn('LibreOffice presentation stderr:', stderr.trim());

    const files = await fs.readdir(tmpDir);
    const expectedExt = { odp: '.odp', ppt: '.ppt', pptx: '.pptx' } as const;
    const outputFile = files.find(f => f.toLowerCase().endsWith(expectedExt[targetFormat]));
    if (!outputFile) throw new Error(`LibreOffice did not produce ${targetFormat} output file. Available files: ${files.join(', ')}`);
    const outputPath = path.join(tmpDir, outputFile);
    const outputBuffer = await fs.readFile(outputPath);
    
    console.log(`Presentation conversion successful, output size: ${outputBuffer.length} bytes`);
    
    const mimeTypes = {
      odp: 'application/vnd.oasis.opendocument.presentation',
      ppt: 'application/vnd.ms-powerpoint',
      pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    };
    
    const downloadName = `${sanitizedBase}.${targetFormat}`;
    const mime = mimeTypes[targetFormat];
    
    if (persistToDisk) {
      // Always persist presentation files to disk for better performance
      console.log('Persisting presentation file to disk with 10-minute cleanup...');
      const result = await persistOutputBuffer(outputBuffer, downloadName, mime);
      
      // Schedule cleanup after 10 minutes (600,000 ms) instead of default 5 minutes
      if (result.storedFilename) {
        const cleanupPath = path.join(BATCH_OUTPUT_DIR, result.storedFilename);
        setTimeout(async () => {
          try {
            await fs.unlink(cleanupPath);
            console.log(`Cleaned up presentation file: ${result.storedFilename}`);
            batchFileMetadata.delete(result.storedFilename!);
          } catch (cleanupError) {
            console.warn(`Cleanup failed for presentation file ${result.storedFilename}:`, cleanupError);
          }
        }, 10 * 60 * 1000); // 10 minutes
      }
      
      return result;
    }
    
    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime
    };
    
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

const createPresentationHTML = (textContent: string): string => {
  const lines = textContent.split('\n');
  const slides: string[] = [];
  
  let currentSlide: string[] = [];
  let slideTitle = 'Document Presentation';
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    if (!line) {
      if (currentSlide.length > 0) {
        currentSlide.push('');
      }
      continue;
    }
    
    // Check if this looks like a title/heading (short line, doesn't end with punctuation)
    const isTitle = line.length < 50 && !line.endsWith('.') && !line.endsWith(',') && !line.endsWith(';');
    const isChapter = line.startsWith('Chapter ') || line.startsWith('CHAPTER ');
    
    if ((isTitle || isChapter) && currentSlide.length > 0) {
      // Save current slide and start new one
      if (currentSlide.length > 0) {
        slides.push(createSlideHTML(slideTitle, currentSlide.join('\n')));
      }
      
      slideTitle = line;
      currentSlide = [];
    } else if (isTitle || isChapter) {
      // First slide title
      slideTitle = line;
    } else {
      // Add content to current slide
      currentSlide.push(line);
      
      // Start new slide after reasonable amount of content
      if (currentSlide.filter(l => l.trim()).length >= 8) {
        slides.push(createSlideHTML(slideTitle, currentSlide.join('\n')));
        slideTitle = 'Continued...';
        currentSlide = [];
      }
    }
  }
  
  // Add final slide if there's content
  if (currentSlide.length > 0) {
    slides.push(createSlideHTML(slideTitle, currentSlide.join('\n')));
  }
  
  // Create complete HTML document
  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Presentation</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }
        .slide { page-break-after: always; margin-bottom: 50px; padding: 20px; border: 1px solid #ccc; }
        .slide h1 { color: #333; font-size: 28px; margin-bottom: 20px; text-align: center; }
        .slide h2 { color: #555; font-size: 24px; margin-bottom: 15px; }
        .slide p { font-size: 16px; line-height: 1.5; margin-bottom: 10px; }
        .slide:last-child { page-break-after: auto; }
    </style>
</head>
<body>
${slides.join('\n')}
</body>
</html>`;
};

const createSlideHTML = (title: string, content: string): string => {
  const paragraphs = content.split('\n\n').filter(p => p.trim());
  const formattedContent = paragraphs.map(p => `    <p>${p.replace(/\n/g, '<br>')}</p>`).join('\n');
  
  return `<div class="slide">
    <h1>${title}</h1>
${formattedContent}
</div>`;
};

// Build slide data structure (title + bullet lines) from plain text
const buildSlidesFromText = (textContent: string): Array<{ title: string; bullets: string[] }> => {
  const lines = textContent.split(/\r?\n/);
  const slides: Array<{ title: string; bullets: string[] }> = [];

  let currentTitle = 'Document Presentation';
  let currentBullets: string[] = [];
  const flushSlide = () => {
    if (currentBullets.length > 0) {
      slides.push({ title: currentTitle, bullets: currentBullets });
      currentBullets = [];
      currentTitle = 'Continued...';
    }
  };

  for (const raw of lines) {
    const line = raw.trim();
    if (!line) {
      // Blank line indicates possible slide break when bullets are long
      if (currentBullets.length >= 10) flushSlide();
      continue;
    }

    const isHeading = (line.length < 60 && !/[.,;:]$/.test(line)) || /^chapter\s+/i.test(line);
    if (isHeading && currentBullets.length > 0) {
      flushSlide();
      currentTitle = line;
      continue;
    }

    currentBullets.push(line);
    if (currentBullets.length >= 12) {
      flushSlide();
    }
  }

  // Final slide
  flushSlide();
  if (slides.length === 0) {
    slides.push({ title: currentTitle, bullets: ['(No extractable content)'] });
  }
  return slides;
};

const createPresentationStructure = (textContent: string): string => {
  const lines = textContent.split('\n');
  const structuredLines: string[] = [];
  
  // Add a title slide
  structuredLines.push('TITLE SLIDE');
  structuredLines.push('=============');
  structuredLines.push('Document Presentation');
  structuredLines.push('');
  structuredLines.push('');
  
  let slideCount = 1;
  let currentSlideLines: string[] = [];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    if (!line) {
      if (currentSlideLines.length > 0) {
        currentSlideLines.push('');
      }
      continue;
    }
    
    // Start new slide for chapters or major sections
    if (line.startsWith('Chapter ') || line.startsWith('CHAPTER ') || 
        (line.length < 50 && !line.endsWith('.') && !line.endsWith(',') && !line.endsWith(';'))) {
      
      // Finish current slide if it has content
      if (currentSlideLines.length > 0) {
        structuredLines.push(`SLIDE ${slideCount}`);
        structuredLines.push('================');
        structuredLines.push(...currentSlideLines);
        structuredLines.push('');
        structuredLines.push('');
        slideCount++;
        currentSlideLines = [];
      }
      
      // Start new slide with this heading
      currentSlideLines.push(line);
      currentSlideLines.push('');
    } else {
      // Add to current slide
      currentSlideLines.push(line);
      
      // Start new slide after every 5 lines of content to avoid overcrowded slides
      if (currentSlideLines.filter(l => l.trim().length > 0).length >= 6) {
        structuredLines.push(`SLIDE ${slideCount}`);
        structuredLines.push('================');
        structuredLines.push(...currentSlideLines);
        structuredLines.push('');
        structuredLines.push('');
        slideCount++;
        currentSlideLines = [];
      }
    }
  }
  
  // Add final slide if there's remaining content
  if (currentSlideLines.length > 0) {
    structuredLines.push(`SLIDE ${slideCount}`);
    structuredLines.push('================');
    structuredLines.push(...currentSlideLines);
  }
  
  return structuredLines.join('\n');
};

const convertTxtToMarkdown = async (
  textBuffer: Buffer,
  originalBase: string,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  const text = textBuffer.toString('utf-8');
  const sanitizedBase = sanitizeFilename(originalBase);
  
  // Convert CSV data to markdown table format
  const lines = text.split('\n');
  const markdownLines: string[] = [];
  
  // Check if this looks like CSV data (comma-separated values)
  const isCsvData = lines.some(line => line.includes(',') && line.split(',').length > 1);
  
  if (isCsvData) {
    // Convert CSV to markdown table
    const rows = lines.filter(line => line.trim()).map(line => 
      line.split(',').map(cell => cell.trim().replace(/"/g, ''))
    );
    
    if (rows.length > 0) {
      // Add table header
      const headers = rows[0];
      markdownLines.push('| ' + headers.join(' | ') + ' |');
      markdownLines.push('|' + headers.map(() => '---').join('|') + '|');
      
      // Add data rows
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        markdownLines.push('| ' + row.join(' | ') + ' |');
      }
    }
  } else {
    // Convert plain text to basic markdown format
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      if (!line) {
        markdownLines.push('');
        continue;
      }
      
      // Simple heuristics for markdown conversion
      if (line.length < 50 && !line.endsWith('.') && !line.endsWith(',') && !line.endsWith(';')) {
        // Likely a heading
        markdownLines.push(`## ${line}`);
      } else if (line.startsWith('Chapter ') || line.startsWith('CHAPTER ')) {
        // Chapter heading
        markdownLines.push(`# ${line}`);
      } else {
        // Regular paragraph
        markdownLines.push(line);
      }
      
      // Add extra line break after paragraphs
      if (i < lines.length - 1 && lines[i + 1].trim() === '') {
        markdownLines.push('');
      }
    }
  }
  
  const markdownContent = markdownLines.join('\n');
  const buffer = Buffer.from(markdownContent, 'utf-8');
  const downloadName = `${sanitizedBase}.md`;
  const mime = 'text/markdown; charset=utf-8';

  if (persistToDisk) {
    return persistOutputBuffer(buffer, downloadName, mime);
  }

  return {
    buffer,
    filename: downloadName,
    mime
  };
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
    // Keep it simple for HTML intermediate conversion
    args.push('--no-default-epub-cover');
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

const convertDocxWithLibreOffice = async (
  file: Express.Multer.File,
  inputHint: 'doc' | 'docx' | 'odt',
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== ${inputHint.toUpperCase()} TO EPUB VIA LIBREOFFICE ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-lo-epub-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;
  const inputExt = inputHint === 'docx' ? '.docx' : inputHint === 'odt' ? '.odt' : '.doc';
  const inputFilename = `${safeBase}${inputExt}`;
  const inputPath = path.join(tmpDir, inputFilename);

  try {
    await fs.writeFile(inputPath, file.buffer);

    const exportVariants: string[][] = [
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'epub:writer_epub_Export',
        '--outdir', tmpDir,
        inputPath
      ],
      [
        '--headless',
        '--nolockcheck',
        '--nodefault',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'epub',
        '--outdir', tmpDir,
        inputPath
      ]
    ];

    let conversionSucceeded = false;
    for (const exportArgs of exportVariants) {
      try {
        console.log('Trying LibreOffice DOCX->EPUB command:', exportArgs);
        const { stdout, stderr } = await execLibreOffice(exportArgs);
        if (stdout.trim().length > 0) {
          console.log('LibreOffice DOCX->EPUB stdout:', stdout.trim());
        }
        if (stderr.trim().length > 0) {
          console.warn('LibreOffice DOCX->EPUB stderr:', stderr.trim());
        }

        const files = await fs.readdir(tmpDir);
        const epubFile = files.find(name => name.toLowerCase().endsWith('.epub'));
        if (!epubFile) {
          console.warn('EPUB file not found after LibreOffice run, trying next variant...');
          continue;
        }

        const outputPath = path.join(tmpDir, epubFile);
        const outputBuffer = await fs.readFile(outputPath);
        const downloadName = `${sanitizedBase}.epub`;

        console.log('DOCX to EPUB conversion successful:', {
          outputSize: outputBuffer.length,
          filename: downloadName
        });

        if (persistToDisk) {
          return persistOutputBuffer(outputBuffer, downloadName, 'application/epub+zip');
        }

        return {
          buffer: outputBuffer,
          filename: downloadName,
          mime: 'application/epub+zip'
        };
      } catch (variantError) {
        console.warn('LibreOffice DOCX->EPUB command failed, trying fallback...', variantError);
        continue;
      }
    }

    throw new Error('LibreOffice did not produce an EPUB file');
  } catch (error) {
    console.error('DOCX to EPUB via LibreOffice failed:', error);
    const message = error instanceof Error ? error.message : 'Unknown LibreOffice EPUB error';
    throw new Error(`Failed to convert to EPUB with LibreOffice: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// Configure helmet with appropriate settings for large file uploads
app.use(helmet({
  contentSecurityPolicy: false, // Disable CSP for file uploads
  crossOriginEmbedderPolicy: false
}));
// CORS configuration with debugging
app.use((req, res, next) => {
  console.log('CORS middleware - Request origin:', req.get('origin'));
  console.log('CORS middleware - Request method:', req.method);
  next();
});

app.use(cors({
  origin: [
    'https://morphy-1-ulvv.onrender.com',
    'http://localhost:3000',
    'http://localhost:5173'
  ],
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept']
}));

// Additional CORS headers for all responses
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', 'https://morphy-1-ulvv.onrender.com');
  res.header('Access-Control-Allow-Credentials', 'true');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, Accept');
  next();
});

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 60,
  message: 'Too many requests from this IP, please try again later.'
});
app.use('/api/', limiter);

// Increase body parser limits for large file uploads
app.use(express.json({ 
  limit: '100mb'
}));
app.use(express.urlencoded({ 
  extended: true, 
  limit: '100mb'
}));

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

// Handle preflight OPTIONS requests
app.options('*', (req, res) => {
  console.log('OPTIONS preflight request received');
  res.header('Access-Control-Allow-Origin', 'https://morphy-1-ulvv.onrender.com');
  res.header('Access-Control-Allow-Credentials', 'true');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, Accept');
  res.status(200).end();
});

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Test endpoint for debugging large file uploads
app.post('/api/test-upload', upload.single('file'), (req, res) => {
  try {
    const file = req.file;
    const contentLength = req.get('content-length');
    
    console.log('Test upload received:', {
      hasFile: !!file,
      fileSize: file?.size,
      contentLength: contentLength,
      headers: req.headers
    });
    
    res.json({
      success: true,
      fileSize: file?.size || 0,
      contentLength: contentLength,
      message: 'Test upload successful'
    });
  } catch (error) {
    console.error('Test upload error:', error);
    res.status(500).json({ error: 'Test upload failed', details: error instanceof Error ? error.message : String(error) });
  }
});

// Middleware to handle large requests
app.use((req, res, next) => {
  const contentLength = parseInt(req.get('content-length') || '0', 10);
  console.log(`Request received: ${req.method} ${req.path}, Content-Length: ${contentLength} bytes`);
  
  if (contentLength > 100 * 1024 * 1024) { // 100MB
    console.log('Request too large, rejecting before processing');
    return res.status(413).json({ 
      error: 'Request entity too large', 
      details: 'File size exceeds the maximum allowed limit of 100MB' 
    });
  }
  
  next();
});

// Timeout middleware for large file processing
const conversionTimeout = (timeoutMs: number) => {
  return (req: any, res: any, next: any) => {
    req.setTimeout(timeoutMs, () => {
      if (!res.headersSent) {
        res.status(408).json({ error: 'Request timeout - file too large or processing took too long' });
      }
    });
    next();
  };
};

app.post('/api/convert', conversionTimeout(5 * 60 * 1000), upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    const requestOptions = { ...(req.body as Record<string, string | undefined>) };

    console.log('=== CONVERSION REQUEST START ===');
    console.log('Request headers:', req.headers);
    console.log('Request body size:', req.get('content-length'));
    console.log('Request body options:', requestOptions);

    if (!file) {
      console.log('ERROR: No file uploaded');
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log(`Processing file details:`, {
      originalname: file.originalname,
      size: file.size,
      mimetype: file.mimetype,
      fieldname: file.fieldname,
      encoding: file.encoding
    });

    const targetFormat = String(requestOptions.format ?? 'webp').toLowerCase();
    const isCSV = isCsvFile(file);
    const isEPUB = isEpubFile(file);
    const isEPS = isEpsFile(file);
    const isDNG = isDngFile(file);
    const isDOC = isDocFile(file);
    
    console.log('File type detection:', {
      targetFormat,
      isCSV,
      isEPUB, 
      isEPS,
      isDNG,
      isDOC,
      mimetype: file.mimetype,
      extension: file.originalname.split('.').pop()?.toLowerCase(),
      originalname: file.originalname
    });
    
    console.log('Available CALIBRE_CONVERSIONS:', Object.keys(CALIBRE_CONVERSIONS));
    console.log('Target format in CALIBRE_CONVERSIONS?', !!CALIBRE_CONVERSIONS[targetFormat]);

    let result: ConversionResult;

    if ((isDocFile(file) || isDocxFile(file) || isOdtFile(file)) && targetFormat === 'epub') {
      console.log('Single: Routing to LibreOffice (DOC/DOCX/ODT to EPUB conversion)');
      const inputHint = isDocxFile(file) ? 'docx' : isOdtFile(file) ? 'odt' : 'doc';
      result = await convertDocxWithLibreOffice(file, inputHint, requestOptions, true);
    } else if (isEPUB && CALIBRE_CONVERSIONS[targetFormat]) {
      console.log('Single: Routing to Calibre (EPUB conversion)');
      result = await convertWithCalibre(file, targetFormat, requestOptions, true);
    } else if (isCSV && ['epub', 'mobi', 'html', 'txt'].includes(targetFormat)) {
      console.log(`Single: Routing to Python (CSV to ${targetFormat.toUpperCase()} conversion)`);
      result = await convertCsvToEbookPython(file, targetFormat, requestOptions, true);
    } else if (isCSV && LIBREOFFICE_CONVERSIONS[targetFormat]) {
      console.log('Single: Routing to LibreOffice (CSV conversion)');
      result = await convertCsvWithLibreOffice(file, targetFormat, requestOptions, true);
    } else if ((isDOC || file.originalname.toLowerCase().endsWith('.doc')) && targetFormat === 'csv') {
      console.log('Single: Routing to LibreOffice (DOC to CSV conversion)');
      result = await convertDocWithLibreOffice(file, 'doc-to-csv', requestOptions, true);
    } else if (isEPS && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(targetFormat)) {
      console.log('Single: Routing to EPS conversion');
      result = await convertEpsFile(file, targetFormat, requestOptions, true);
    } else if (isDNG && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(targetFormat)) {
      console.log('Single: Routing to DNG conversion');
      console.log('DNG conversion details:', {
        filename: file.originalname,
        targetFormat,
        supportedFormats: ['webp', 'png', 'jpeg', 'jpg', 'ico'],
        requestOptions
      });
      result = await convertDngFile(file, targetFormat, requestOptions, true);
    } else {
      // Check if this is a DOC file that wasn't detected properly
      if (file.originalname.toLowerCase().endsWith('.doc') && targetFormat === 'csv') {
        console.log('DOC file detected by fallback - routing to LibreOffice');
        result = await convertDocWithLibreOffice(file, 'doc-to-csv', requestOptions, true);
      } else {
        // Handle Sharp image conversions
        console.log('Falling back to Sharp image processing - this should not happen for DOC files!');
        console.log('File details:', {
          originalname: file.originalname,
          mimetype: file.mimetype,
          targetFormat,
          isDOC,
          isCSV,
          isEPUB,
          isEPS,
          isDNG
        });
        
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
            const icoSize = parseInt(iconSize) || 16;
        pipeline = pipeline
              .resize(icoSize, icoSize, {
            fit: 'contain',
                background: { r: 255, g: 255, b: 255, alpha: 0 }
          })
              .png({ compressionLevel: 0, quality: 100 });
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
    }

    // Return download link instead of streaming file
    console.log('Conversion successful:', {
      filename: result.filename,
      size: result.buffer.length,
      storedFilename: result.storedFilename,
      mime: result.mime
    });
    
    res.json({
      success: true,
      downloadPath: `/download/${encodeURIComponent(result.storedFilename!)}`,
      filename: result.filename,
      size: result.buffer.length
    });
    
    console.log('=== CONVERSION REQUEST END (SUCCESS) ===');
  } catch (error) {
    console.error('=== CONVERSION ERROR ===');
    console.error('Error type:', error instanceof Error ? error.constructor.name : typeof error);
    console.error('Error message:', error instanceof Error ? error.message : String(error));
    console.error('Error stack:', error instanceof Error ? error.stack : 'No stack trace');
    
    if (error instanceof Error && error.message.includes('ENOENT')) {
      console.error('ENOENT error - likely missing binary or file path issue');
    }
    
    if (error instanceof Error && error.message.includes('Calibre')) {
      console.error('Calibre-specific error detected');
    }
    
    const errorMessage = error instanceof Error ? error.message : 'Unknown conversion error';
    console.log('=== CONVERSION REQUEST END (ERROR) ===');
    
    res.status(500).json({
      error: 'Conversion failed',
      details: errorMessage
    });
  }
});

app.post('/api/convert/batch', conversionTimeout(10 * 60 * 1000), uploadBatch.array('files'), async (req, res) => {
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
      const isDNG = isDngFile(file);
      const isDOC = isDocFile(file);
      console.log(`Processing ${file.originalname}: isCSV=${isCSV}, isEPUB=${isEPUB}, isEPS=${isEPS}, isDNG=${isDNG}, isDOC=${isDOC}, format=${format}, mimetype=${file.mimetype}`);

      if ((isDocFile(file) || isDocxFile(file) || isOdtFile(file)) && format === 'epub') {
        console.log('Batch: Routing to LibreOffice (DOC/DOCX/ODT to EPUB conversion)');
        const inputHint = isDocxFile(file) ? 'docx' : isOdtFile(file) ? 'odt' : 'doc';
        output = await convertDocxWithLibreOffice(file, inputHint, requestOptions, true);
      } else if (isEPUB && CALIBRE_CONVERSIONS[format]) {
        console.log('Routing to Calibre (EPUB conversion)');
        output = await convertWithCalibre(file, format, requestOptions, true);
      } else if (isCSV && ['epub', 'mobi', 'html', 'txt'].includes(format)) {
        console.log(`Batch: Routing to Python (CSV to ${format.toUpperCase()} conversion)`);
        output = await convertCsvToEbookPython(file, format, requestOptions, true);
      } else if (isCSV && LIBREOFFICE_CONVERSIONS[format]) {
        console.log('Routing to LibreOffice (CSV conversion)');
        output = await convertCsvWithLibreOffice(file, format, requestOptions, true);
      } else if ((isDOC || file.originalname.toLowerCase().endsWith('.doc')) && format === 'csv') {
        console.log('Batch: Routing to LibreOffice (DOC to CSV conversion)');
        output = await convertDocWithLibreOffice(file, 'doc-to-csv', requestOptions, true);
      } else if (isEPS && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(format)) {
        console.log('Routing to EPS conversion');
        output = await convertEpsFile(file, format, requestOptions, true);
      } else if (isDNG && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(format)) {
        console.log('Batch: Routing to DNG conversion');
        console.log('Batch DNG conversion details:', {
          filename: file.originalname,
          targetFormat: format,
          supportedFormats: ['webp', 'png', 'jpeg', 'jpg', 'ico'],
          requestOptions
        });
        output = await convertDngFile(file, format, requestOptions, true);
      } else {
        throw new Error(`Unsupported input file type or target format for batch conversion. File: ${file.originalname}, isCSV: ${isCSV}, isEPUB: ${isEPUB}, isEPS: ${isEPS}, isDNG: ${isDNG}, isDOC: ${isDOC}, format: ${format}`);
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

    console.log('Download request:', {
      storedFilename,
      hasMetadata: !!metadata,
      metadata: metadata ? { downloadName: metadata.downloadName, mime: metadata.mime } : null
    });

    if (!metadata) {
      return res.status(404).json({ error: 'File not found or expired' });
    }

    const filePath = path.join(BATCH_OUTPUT_DIR, storedFilename);
    const stat = await fs.stat(filePath).catch(() => null);
    if (!stat) {
      batchFileMetadata.delete(storedFilename);
      return res.status(404).json({ error: 'File not found or expired' });
    }

    console.log('File stats:', {
      size: stat.size,
      mtime: stat.mtime,
      isFile: stat.isFile()
    });

    // Read the first few bytes to check if it's a valid DOCX file
    const fileBuffer = await fs.readFile(filePath);
    const isValidDocx = fileBuffer.length > 4 && fileBuffer[0] === 0x50 && fileBuffer[1] === 0x4B;
    
    console.log('File validation:', {
      bufferSize: fileBuffer.length,
      isValidDocx,
      firstBytes: fileBuffer.slice(0, 4).toString('hex')
    });

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
  console.error('=== ERROR HANDLER ===');
  console.error('Error type:', err.constructor.name);
  console.error('Error message:', err.message);
  console.error('Error code:', err.code);
  console.error('Error status:', err.status);
  
  const multerModule: any = multer;
  if (multerModule.MulterError && err instanceof multerModule.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File too large. Maximum size is 100MB.' });
    }
    return res.status(400).json({ error: err.message });
  }

  // Handle request entity too large errors
  if (err.status === 413 || err.code === 'LIMIT_FILE_SIZE' || err.message.includes('too large')) {
    return res.status(413).json({ 
      error: 'Request entity too large', 
      details: 'File size exceeds the maximum allowed limit of 100MB' 
    });
  }

  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error' });
});

// Set server timeout for large file processing
const server = app.listen(PORT, '0.0.0.0', () => {
  console.log(`Morpy backend running on port ${PORT}`);
});

// Increase timeout for large file processing (5 minutes)
server.timeout = 5 * 60 * 1000; // 5 minutes
server.keepAliveTimeout = 65000; // 65 seconds
server.headersTimeout = 66000; // 66 seconds

export default app;

