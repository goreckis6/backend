import express from 'express';
import multer from 'multer';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import sharp from 'sharp';
import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { exec } from 'child_process';
import { promisify } from 'util';
import os from 'os';
import Papa from 'papaparse';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, TextRun } from 'docx';
import XLSX from 'xlsx';

// ---- CSV conversion helpers ----
type CSVTable = { headers: string[]; rows: string[][] };

const parseCSV = (buffer: Buffer, delimiter: string): CSVTable => {
  const text = buffer.toString('utf8');
  const parsed = Papa.parse<string[]>(text, {
    delimiter: delimiter || ',',
    skipEmptyLines: true
  });
  const data = parsed.data as string[][];
  const headers = data.length ? data[0] : [];
  const rows = data.length > 1 ? data.slice(1) : [];
  return { headers, rows };
};

const buildHTMLFromCSV = ({ headers, rows }: CSVTable): string => {
  const head = `<!doctype html><html><head><meta charset="utf-8"><title>CSV</title><style>body{font-family:Arial,Helvetica,sans-serif}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:8px}th{background:#f3f4f6;text-align:left}</style></head><body>`;
  const tblHead = `<table><thead><tr>${headers.map(h => `<th>${h ?? ''}</th>`).join('')}</tr></thead>`;
  const tblBody = `<tbody>${rows.map(r => `<tr>${r.map(c => `<td>${c ?? ''}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  return `${head}${tblHead}${tblBody}</body></html>`;
};

const buildTXTFromCSV = ({ headers, rows }: CSVTable): string => {
  const lines = [headers.join('\t'), ...rows.map(r => r.join('\t'))];
  return lines.join('\n');
};

const buildMDFromCSV = ({ headers, rows }: CSVTable): string => {
  const header = `| ${headers.join(' | ')} |`;
  const sep = `| ${headers.map(() => '---').join(' | ')} |`;
  const body = rows.map(r => `| ${r.join(' | ')} |`).join('\n');
  return [header, sep, body].join('\n');
};

const buildDOCXFromCSV = async ({ headers, rows }: CSVTable): Promise<Buffer> => {
  const tableRows: TableRow[] = [];
  const hdrCells = headers.map(h => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h ?? '', bold: true })] })] }));
  tableRows.push(new TableRow({ children: hdrCells }));
  for (const row of rows) {
    tableRows.push(new TableRow({ children: row.map(c => new TableCell({ children: [new Paragraph(c ?? '')] })) }));
  }
  const doc = new Document({
    sections: [{ properties: {}, children: [new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } })] }]
  });
  return await Packer.toBuffer(doc);
};

const buildXLSXFromCSV = ({ headers, rows }: CSVTable): Buffer => {
  const aoa = [headers, ...rows];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
};

const writeTempFile = async (content: Buffer | string, ext: string): Promise<string> => {
  const tempDir = os.tmpdir();
  const p = path.join(tempDir, `csvconv_${Date.now()}_${randomUUID()}.${ext}`);
  if (typeof content === 'string') {
    await fs.writeFile(p, Buffer.from(content, 'utf8'));
  } else {
    await fs.writeFile(p, content);
  }
  return p;
};

const sofficeConvert = async (inputPath: string, outExt: string): Promise<Buffer> => {
  const outDir = path.dirname(inputPath);
  const cmd = `soffice --headless --convert-to ${outExt} "${inputPath}" --outdir "${outDir}"`;
  await execAsync(cmd, { timeout: 120000, maxBuffer: 100 * 1024 * 1024 });
  const base = path.basename(inputPath, path.extname(inputPath));
  const outPath = path.join(outDir, `${base}.${outExt}`);
  const buf = await fs.readFile(outPath);
  try { await fs.unlink(outPath); } catch {}
  return buf;
};

const calibreConvert = async (inputPath: string, outExt: string): Promise<Buffer> => {
  const outPath = inputPath.replace(path.extname(inputPath), `.${outExt}`);
  const cmd = `ebook-convert "${inputPath}" "${outPath}"`;
  await execAsync(cmd, { timeout: 120000, maxBuffer: 100 * 1024 * 1024 });
  const buf = await fs.readFile(outPath);
  try { await fs.unlink(outPath); } catch {}
  return buf;
};

const processCSVConversion = async (input: Buffer, targetFormat: string, delimiter: string): Promise<{ buffer: Buffer; ext: string; mime: string }> => {
  const table = parseCSV(input, delimiter);
  const fmt = (targetFormat || '').toLowerCase();
  switch (fmt) {
    case 'txt': {
      const txt = buildTXTFromCSV(table);
      return { buffer: Buffer.from(txt, 'utf8'), ext: 'txt', mime: 'text/plain' };
    }
    case 'md': {
      const md = buildMDFromCSV(table);
      return { buffer: Buffer.from(md, 'utf8'), ext: 'md', mime: 'text/markdown' };
    }
    case 'html': {
      const html = buildHTMLFromCSV(table);
      return { buffer: Buffer.from(html, 'utf8'), ext: 'html', mime: 'text/html' };
    }
    case 'docx': {
      const buf = await buildDOCXFromCSV(table);
      return { buffer: buf, ext: 'docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
    }
    case 'xlsx': {
      const buf = buildXLSXFromCSV(table);
      return { buffer: buf, ext: 'xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
    }
    // Formats via LibreOffice from HTML
    case 'pdf':
    case 'odt':
    case 'rtf':
    case 'doc':
    case 'odp':
    case 'ppt':
    case 'pptx':
    case 'xls': {
      const html = buildHTMLFromCSV(table);
      const htmlPath = await writeTempFile(html, 'html');
      try {
        const outBuffer = await sofficeConvert(htmlPath, fmt);
        const mimeMap: Record<string, string> = {
          pdf: 'application/pdf', odt: 'application/vnd.oasis.opendocument.text', rtf: 'application/rtf', doc: 'application/msword',
          odp: 'application/vnd.oasis.opendocument.presentation', ppt: 'application/vnd.ms-powerpoint', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          xls: 'application/vnd.ms-excel'
        };
        return { buffer: outBuffer, ext: fmt, mime: mimeMap[fmt] || 'application/octet-stream' };
      } finally { try { await fs.unlink(htmlPath); } catch {} }
    }
    // Ebooks via Calibre from HTML
    case 'epub':
    case 'mobi': {
      const html = buildHTMLFromCSV(table);
      const htmlPath = await writeTempFile(html, 'html');
      try {
        const outBuffer = await calibreConvert(htmlPath, fmt);
        const mimeMap: Record<string, string> = { epub: 'application/epub+zip', mobi: 'application/x-mobipocket-ebook' };
        return { buffer: outBuffer, ext: fmt, mime: mimeMap[fmt] || 'application/octet-stream' };
      } finally { try { await fs.unlink(htmlPath); } catch {} }
    }
    default:
      throw new Error('Unsupported CSV target format');
  }
};
import { randomUUID } from 'crypto';

const execAsync = promisify(exec);

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Create converted files directory
const CONVERTED_FILES_DIR = path.join(__dirname, '..', 'converted_files');

// Ensure the converted files directory exists
const ensureConvertedFilesDir = async () => {
  try {
    await fs.access(CONVERTED_FILES_DIR);
  } catch {
    await fs.mkdir(CONVERTED_FILES_DIR, { recursive: true });
    console.log('Created converted_files directory');
  }
};

// Clean up old files (older than 5 minutes)
const cleanupOldFiles = async () => {
  try {
    const files = await fs.readdir(CONVERTED_FILES_DIR);
    const now = Date.now();
    const fiveMinutesAgo = now - (5 * 60 * 1000);

    for (const file of files) {
      const filePath = path.join(CONVERTED_FILES_DIR, file);
      const stats = await fs.stat(filePath);
      
      if (stats.mtime.getTime() < fiveMinutesAgo) {
        await fs.unlink(filePath);
        console.log(`Cleaned up old file: ${file}`);
      }
    }
  } catch (error) {
    console.error('Error cleaning up old files:', error);
  }
};

// Schedule cleanup every 2 minutes
setInterval(cleanupOldFiles, 2 * 60 * 1000);

// Utility function to fix UTF-8 encoding issues
const fixUTF8Encoding = (text: string): string => {
  if (!text) return text;
  
  // Handle the specific pattern we're seeing: â\x80\x94 (which should be —)
  // This is the UTF-8 encoding of em dash being displayed incorrectly
  let fixed = text
    .replace(/â\\x80\\x94/g, '—') // Replace â\x80\x94 with proper em dash
    .replace(/â\x80\x94/g, '—')   // Replace â\x80\x94 with proper em dash
    .replace(/\\x80\\x94/g, '—')  // Replace literal \x80\x94 with em dash
    .replace(/\x80\x94/g, '—')    // Replace actual UTF-8 bytes with em dash
    .replace(/\\x80\\x93/g, '–')  // Replace literal \x80\x93 with en dash
    .replace(/\x80\x93/g, '–')    // Replace actual UTF-8 bytes with en dash
    .replace(/\\x80\\x99/g, '')   // Replace literal \x80\x99 with apostrophe
    .replace(/\x80\x99/g, '')     // Replace actual UTF-8 bytes with apostrophe
    .replace(/\\x80\\x9c/g, '"')  // Replace literal \x80\x9c with left quote
    .replace(/\x80\x9c/g, '"')    // Replace actual UTF-8 bytes with left quote
    .replace(/\\x80\\x9d/g, '"')  // Replace literal \x80\x9d with right quote
    .replace(/\x80\x9d/g, '"');   // Replace actual UTF-8 bytes with right quote
  
  return fixed;
};

// Utility function to sanitize filenames for better compatibility
const sanitizeFilename = (filename: string): string => {
  if (!filename) {
    return 'file';
  }

  let workingName = filename;
  try {
    workingName = decodeURIComponent(filename);
  } catch (error) {
    console.log('Could not decode filename in sanitize:', filename);
  }

  // Normalize to decompose accents, then remove combining marks
  workingName = workingName.normalize('NFKD').replace(/[\u0300-\u036f]/g, '');

  // Replace long dashes with regular hyphens
  workingName = workingName.replace(/[\u2012-\u2015]/g, '-');

  // Replace any remaining unsafe characters with underscores
  workingName = workingName
    .replace(/[<>:"/\\|?*]/g, '_')
    .replace(/[^0-9A-Za-z._-]+/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^-+|-+$/g, '')
    .trim();

  if (!workingName.length) {
    workingName = 'file';
  }

  return workingName;
};

// RAW file extensions
const RAW_EXTENSIONS = ['dng', 'cr2', 'cr3', 'nef', 'arw', 'rw2', 'pef', 'orf', 'raf', 'x3f', 'raw'];
const EPS_EXTENSIONS = ['eps', 'ps'];
const CSV_EXTENSIONS = ['csv'];

// Function to check if file is RAW
const isRAWFile = (filename: string): boolean => {
  const ext = filename.split('.').pop()?.toLowerCase();
  return RAW_EXTENSIONS.includes(ext || '');
};

const isEPSFile = (filename: string): boolean => {
  const ext = filename.split('.').pop()?.toLowerCase();
  return EPS_EXTENSIONS.includes(ext || '');
};

const isCSVFile = (filename: string): boolean => {
  const ext = filename.split('.').pop()?.toLowerCase();
  return CSV_EXTENSIONS.includes(ext || '');
};

const processEPSFile = async (inputBuffer: Buffer, filename: string, size: number): Promise<Buffer> => {
  const tempDir = os.tmpdir();
  const uniqueId = randomUUID();
  const inputPath = path.join(tempDir, `eps_input_${uniqueId}.eps`);
  const outputPngPath = path.join(tempDir, `eps_output_${uniqueId}.png`);
  const outputIcoPath = path.join(tempDir, `eps_output_${uniqueId}.ico`);

  try {
    // Verify Ghostscript availability early with a short timeout
    try {
      await execAsync('gs -version', { timeout: 3000 });
    } catch (gsErr) {
      throw new Error('Ghostscript (gs) is not available on the server runtime');
    }

    await fs.writeFile(inputPath, inputBuffer);

    const gsCommand = `gs -dSAFER -dBATCH -dNOPAUSE -dEPSCrop -r${size * 8} -sDEVICE=pngalpha -sOutputFile="${outputPngPath}" "${inputPath}"`;
    console.log('Running Ghostscript:', gsCommand);
    await execAsync(gsCommand, {
      timeout: 60000,
      maxBuffer: 50 * 1024 * 1024
    });

    // Determine ImageMagick CLI flavor (convert vs magick)
    let useMagickRoot = false;
    try {
      await execAsync('convert -version', { timeout: 2000 });
    } catch {
      // Try magick
      await execAsync('magick -version', { timeout: 2000 });
      useMagickRoot = true;
    }

    // Use ImageMagick to produce ICO from the rasterized PNG
    const convertCmd = useMagickRoot
      ? `magick "${outputPngPath}" -resize ${size}x${size} "${outputIcoPath}"`
      : `convert "${outputPngPath}" -resize ${size}x${size} "${outputIcoPath}"`;
    console.log('Running ImageMagick:', convertCmd);
    const { stdout: imStdout, stderr: imStderr } = await execAsync(convertCmd, { timeout: 30000, maxBuffer: 50 * 1024 * 1024 });
    if (imStderr) {
      console.log('ImageMagick stderr:', imStderr);
    }

    const icoBuffer = await fs.readFile(outputIcoPath);
    return icoBuffer;
  } catch (error) {
    console.error('EPS processing error:', error);
    throw new Error(error instanceof Error ? error.message : 'EPS processing failed');
  } finally {
    const cleanupFiles = [inputPath, outputPngPath, outputIcoPath];
    for (const filePath of cleanupFiles) {
      try {
        await fs.unlink(filePath);
      } catch (cleanupError) {
        // ignore
      }
    }
  }
};

// Rasterize EPS to PNG buffer using Ghostscript. Optionally constrain output size via Sharp.
const processEPSRaster = async (
  inputBuffer: Buffer,
  filename: string,
  targetWidth?: number,
  targetHeight?: number
): Promise<Buffer> => {
  const tempDir = os.tmpdir();
  const uniqueId = randomUUID();
  const inputPath = path.join(tempDir, `eps_r_input_${uniqueId}.eps`);
  const outputPngPath = path.join(tempDir, `eps_r_output_${uniqueId}.png`);

  try {
    // Ensure Ghostscript is present
    await execAsync('gs -version', { timeout: 3000 });

    await fs.writeFile(inputPath, inputBuffer);

    // Use a reasonable DPI; if explicit target sizes provided, still rasterize with decent base quality
    const dpi = 300;
    const gsCommand = `gs -dSAFER -dBATCH -dNOPAUSE -dEPSCrop -r${dpi} -sDEVICE=pngalpha -sOutputFile="${outputPngPath}" "${inputPath}"`;
    console.log('GS raster (web):', gsCommand);
    await execAsync(gsCommand, { timeout: 60000, maxBuffer: 50 * 1024 * 1024 });

    let pngBuffer = await fs.readFile(outputPngPath);

    // Optionally resize with Sharp to requested dimensions
    if (targetWidth || targetHeight) {
      pngBuffer = await sharp(pngBuffer, { failOn: 'truncated', unlimited: true })
        .resize(targetWidth || undefined, targetHeight || undefined, {
          fit: 'inside',
          withoutEnlargement: true,
          background: { r: 0, g: 0, b: 0, alpha: 0 }
        })
        .png()
        .toBuffer();
    }

    return pngBuffer;
  } catch (error) {
    console.error('EPS raster error:', error);
    throw new Error(error instanceof Error ? error.message : 'EPS rasterization failed');
  } finally {
    try { await fs.unlink(inputPath); } catch {}
    try { await fs.unlink(outputPngPath); } catch {}
  }
};

// EPUB conversion temporarily disabled

// Function to process RAW file with dcraw or fallback to Sharp
const processRAWFile = async (inputBuffer: Buffer, filename: string): Promise<Buffer> => {
  const tempDir = os.tmpdir();
  const inputPath = path.join(tempDir, `input_${Date.now()}_${filename}`);
  
  try {
    // Check if dcraw is available
    let dcrawAvailable = false;
    try {
      await execAsync('which dcraw');
      await execAsync('dcraw -h');
      dcrawAvailable = true;
      console.log('dcraw is available, using dcraw for RAW processing');
    } catch (dcrawError) {
      console.log('dcraw not available, trying Sharp fallback');
      dcrawAvailable = false;
    }

    if (dcrawAvailable) {
      // Use dcraw for RAW processing
    await fs.writeFile(inputPath, inputBuffer);
    
      const dcrawCommand = `timeout 30s dcraw -T -w -6 -h -c "${inputPath}"`;
      
      const { stdout } = await execAsync(dcrawCommand, { 
        timeout: 30000,
        maxBuffer: 50 * 1024 * 1024
      });
      
      if (!stdout || stdout.length === 0) {
        throw new Error('dcraw produced no output');
      }
      
      return Buffer.from(stdout, 'binary');
    } else {
      // Fallback: Try to process with Sharp directly (some RAW formats might work)
      console.log('Attempting Sharp fallback for RAW file');
      
      try {
        // Try to extract the embedded JPEG preview first (usually has full color)
        console.log('Attempting to extract JPEG preview from RAW file');
        
        let sharpInstance = sharp(inputBuffer, { 
          failOn: 'truncated',
          unlimited: true
        });
        
        const metadata = await sharpInstance.metadata();
        console.log('Sharp metadata for RAW:', metadata);
        
        // If we get a valid color image, use it
        if (metadata.format && metadata.channels && metadata.channels >= 3) {
          console.log('Using extracted preview with', metadata.channels, 'channels');
          return await sharpInstance.png().toBuffer();
        }
        
        // If no preview, try to process as RAW with proper color handling
        console.log('No preview found, attempting RAW processing with color preservation');
        
        sharpInstance = sharp(inputBuffer, { 
          failOn: 'truncated',
          unlimited: true,
          raw: {
            width: metadata.width || 1000,
            height: metadata.height || 1000,
            channels: 3 // Force RGB
          }
        });
        
        // Ensure we maintain color information
        return await sharpInstance
          .png({ 
            quality: 90,
            compressionLevel: 6,
            adaptiveFiltering: true
          })
          .toBuffer();
          
      } catch (sharpError) {
        console.error('Sharp fallback failed:', sharpError);
        
        // Final fallback: try with different RAW parameters
        try {
          console.log('Attempting final fallback with different RAW settings');
          
          const fallbackInstance = sharp(inputBuffer, { 
            unlimited: true,
            sequentialRead: true
          });
          
          return await fallbackInstance
            .jpeg({ quality: 90, progressive: true })
            .toBuffer();
            
        } catch (finalError) {
          console.error('Final fallback failed:', finalError);
          throw new Error('RAW file format not supported. Please try a different RAW file or convert to JPEG/PNG first.');
        }
      }
    }
  } catch (error) {
    console.error('RAW processing error:', error);
    throw new Error(`RAW processing failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
  } finally {
    // Clean up temporary files
    try {
      await fs.unlink(inputPath);
    } catch (cleanupError) {
      // Ignore cleanup errors for files that might not exist
    }
  }
};

const sendBufferAsDownload = async (
  res: express.Response,
  file: Express.Multer.File,
  outputBuffer: Buffer,
  fileExtension: string,
  contentType: string
) => {
  const rawOriginalName = file.originalname.replace(/\.[^.]+$/, '');
  const fixedOriginalName = fixUTF8Encoding(rawOriginalName);
  const sanitizedName = sanitizeFilename(fixedOriginalName);
  const outputFilename = `${sanitizedName}.${fileExtension}`;

  await ensureConvertedFilesDir();
  const timestamp = Date.now();
  const uniqueFilename = `${timestamp}_${outputFilename}`;
  const filePath = path.join(CONVERTED_FILES_DIR, uniqueFilename);
  await fs.writeFile(filePath, outputBuffer);

  const encodedFilename = encodeURIComponent(outputFilename);
  res.set({
    'Content-Type': contentType,
    'Content-Disposition': `attachment; filename*=UTF-8''${encodedFilename}`,
    'Content-Length': outputBuffer.length.toString(),
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive'
  });

  try {
    res.send(outputBuffer);
  } catch (sendError) {
    console.error('Error sending file:', sendError);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Failed to send converted file', details: String(sendError) });
    }
  }

  setTimeout(async () => {
    try { await fs.unlink(filePath); } catch {}
  }, 5 * 60 * 1000);

  setTimeout(() => {
    if (global.gc) { global.gc(); }
  }, 1000);
};

const app = express();
const PORT = process.env.PORT || 3000;

// Enable trust proxy for Render (required for rate limiting)
app.set('trust proxy', 1);

// Handle process signals gracefully
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT received, shutting down gracefully');
  process.exit(0);
});

// Handle uncaught exceptions
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  // Don't exit immediately, log and continue
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  // Don't exit immediately, log and continue
});

// Monitor memory usage with warnings
setInterval(() => {
  const memUsage = process.memoryUsage();
  const heapUsedMB = Math.round(memUsage.heapUsed / 1024 / 1024);
  const heapTotalMB = Math.round(memUsage.heapTotal / 1024 / 1024);
  
  console.log(`Memory: ${heapUsedMB}MB used, ${heapTotalMB}MB total`);
  
  // Warn if memory usage is getting high
  if (heapUsedMB > 400) { // 400MB warning threshold
    console.warn(`⚠️ High memory usage: ${heapUsedMB}MB. Consider garbage collection.`);
    if (global.gc) {
      global.gc();
      console.log('🧹 Forced garbage collection');
    }
  }
}, 30000); // Log every 30 seconds

// Security middleware
app.use(helmet());
app.use(cors({
  origin: process.env.FRONTEND_URL || 'http://localhost:5173',
  credentials: true
}));

// Rate limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100, // limit each IP to 100 requests per windowMs
  message: 'Too many requests from this IP, please try again later.'
});
app.use('/api/', limiter);

// Body parsing
// Middleware to handle UTF-8 encoding properly
app.use((req, res, next) => {
  // Set proper UTF-8 encoding headers
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  next();
});

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// File filter function (moved inline to multer configurations for UTF-8 handling)

// Configure multer for single file uploads with UTF-8 filename preservation
const uploadSingle = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024, // 100MB limit per file
    files: 1
  },
  fileFilter: (req: any, file: any, cb: any) => {
    // Decode the filename if it's URL encoded before file filter
    try {
      const decodedFilename = decodeURIComponent(file.originalname);
      file.originalname = decodedFilename;
      console.log('Decoded filename:', decodedFilename);
    } catch (e) {
      // If decoding fails, keep original filename
      console.log('Could not decode filename:', file.originalname);
    }
    
    // Apply the original file filter logic
    const allowedMimes = [
      'image/jpeg', 'image/jpg', 'image/png', 'image/bmp', 'image/tiff', 'image/tif',
      'image/webp', 'image/gif', 'image/avif', 'image/heic', 'image/heif',
      'image/x-canon-cr2', 'image/x-canon-crw', 'image/x-nikon-nef', 'image/x-sony-arw',
      'image/x-adobe-dng', 'image/x-panasonic-raw', 'image/x-olympus-orf',
      'image/x-pentax-pef', 'image/x-epson-erf', 'image/x-raw',
      // EPS/PostScript
      'application/postscript', 'application/eps', 'application/x-eps', 'image/eps',
      // CSV
      'text/csv', 'application/csv', 'application/vnd.ms-excel'
    ];
    
    if (allowedMimes.includes(file.mimetype) || file.originalname.match(/\.(cr2|crw|nef|arw|dng|raw|orf|pef|erf|eps|ps|csv)$/i)) {
      cb(null, true);
    } else {
      cb(new Error('Unsupported file type') as any, false);
    }
  }
});

// Configure multer for batch file uploads with UTF-8 filename preservation
const uploadBatch = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024, // 100MB limit per file
    files: 20 // Allow up to 20 files for batch processing
  },
  fileFilter: (req: any, file: any, cb: any) => {
    // Decode the filename if it's URL encoded before file filter
    try {
      const decodedFilename = decodeURIComponent(file.originalname);
      file.originalname = decodedFilename;
      console.log('Decoded batch filename:', decodedFilename);
    } catch (e) {
      // If decoding fails, keep original filename
      console.log('Could not decode batch filename:', file.originalname);
    }
    
    // Apply the original file filter logic
    const allowedMimes = [
      'image/jpeg', 'image/jpg', 'image/png', 'image/bmp', 'image/tiff', 'image/tif',
      'image/webp', 'image/gif', 'image/avif', 'image/heic', 'image/heif',
      'image/x-canon-cr2', 'image/x-canon-crw', 'image/x-nikon-nef', 'image/x-sony-arw',
      'image/x-adobe-dng', 'image/x-panasonic-raw', 'image/x-olympus-orf',
      'image/x-pentax-pef', 'image/x-epson-erf', 'image/x-raw',
      // EPS/PostScript
      'application/postscript', 'application/eps', 'application/x-eps', 'image/eps',
      // CSV
      'text/csv', 'application/csv', 'application/vnd.ms-excel'
    ];
    
    if (allowedMimes.includes(file.mimetype) || file.originalname.match(/\.(cr2|crw|nef|arw|dng|raw|orf|pef|erf|eps|ps|csv)$/i)) {
      cb(null, true);
    } else {
      cb(new Error('Unsupported file type') as any, false);
    }
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Operational status endpoint
app.get('/api/status', async (req, res) => {
  try {
    const gsAvailable = await execAsync('gs -version', { timeout: 2000 })
      .then(() => true)
      .catch(() => false);

    const dcrawAvailable = await execAsync('dcraw -h', { timeout: 2000 })
      .then(() => true)
      .catch(() => false);

    const mem = process.memoryUsage();
    const heapUsedMB = Math.round(mem.heapUsed / 1024 / 1024);
    const heapTotalMB = Math.round(mem.heapTotal / 1024 / 1024);

    res.json({
      ok: true,
      timestamp: new Date().toISOString(),
      uptimeSec: Math.round(process.uptime()),
      versions: {
        node: process.version
      },
      tools: {
        ghostscript: gsAvailable,
        dcraw: dcrawAvailable
      },
      memory: {
        heapUsedMB,
        heapTotalMB
      }
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: 'Status probe failed',
      details: error instanceof Error ? error.message : String(error)
    });
  }
});

// Download endpoint for converted files
app.get('/download/:filename', async (req, res) => {
  try {
    const filename = req.params.filename;
    const filePath = path.join(CONVERTED_FILES_DIR, filename);
    
    // Check if file exists
    try {
      await fs.access(filePath);
    } catch {
      return res.status(404).json({ error: 'File not found or expired' });
    }

    // Get file stats
    const stats = await fs.stat(filePath);
    const ext = path.extname(filename).toLowerCase();
    
    // Determine content type
    let contentType = 'application/octet-stream';
    switch (ext) {
      case '.webp':
        contentType = 'image/webp';
        break;
      case '.png':
        contentType = 'image/png';
        break;
      case '.jpg':
      case '.jpeg':
        contentType = 'image/jpeg';
        break;
      case '.ico':
        contentType = 'image/x-icon';
        break;
    }

    // Extract original filename from the stored filename (remove timestamp prefix)
    const originalFilename = filename.replace(/^\d+_/, '');
    const sanitizedFilename = sanitizeFilename(originalFilename);
    const fixedFilename = fixUTF8Encoding(sanitizedFilename);
    const encodedFilename = encodeURIComponent(fixedFilename);
    
    // Set headers with proper UTF-8 encoding for filenames
    res.set({
      'Content-Type': contentType,
      'Content-Disposition': `attachment; filename*=UTF-8''${encodedFilename}`,
      'Content-Length': stats.size.toString(),
      'Cache-Control': 'no-cache'
    });

    // Stream the file
    const fileStream = await fs.readFile(filePath);
    res.send(fileStream);
    
    console.log(`File downloaded: ${filename}`);
  } catch (error) {
    console.error('Download error:', error);
    res.status(500).json({ error: 'Failed to download file' });
  }
});

// Main conversion endpoint
app.post('/api/convert', uploadSingle.single('file'), async (req, res) => {
  try {
    const file = req.file;
    
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    // Log original filename for debugging
    console.log('Original filename:', file.originalname);
    console.log('Filename encoding check:', Buffer.from(file.originalname, 'utf8').toString('hex'));
    const { 
      quality = 'high', 
      lossless = 'false',
      format = 'webp',
      width,
      height,
      iconSize = '16',
      delimiter = ',',
      includeMetadata = 'true',
      extractTables = 'true'
    } = req.body;

    // Calculate common variables
    const targetFormat = String(format || '').toLowerCase();
    const iconSizeNum = parseInt(iconSize) || 16;
    const qualityValue = quality === 'high' ? 95 : quality === 'medium' ? 80 : 60;
    const isLossless = lossless === 'true';
    const isEPS = isEPSFile(file.originalname) || file.mimetype === 'application/postscript';
    const isEPUB = file.originalname.toLowerCase().endsWith('.epub');

  // Early format-specific handling before Sharp metadata
  try {
    // CSV conversions (early)
    if (isCSVFile(file.originalname)) {
      try {
        const { buffer, ext, mime } = await processCSVConversion(file.buffer, targetFormat, delimiter);
        return await sendBufferAsDownload(res, file, buffer, ext, mime);
      } catch (csvErr) {
        return res.status(400).json({ error: 'CSV conversion failed', details: csvErr instanceof Error ? csvErr.message : String(csvErr) });
      }
    }

    if (isEPS && (targetFormat === 'ico' || targetFormat === 'webp')) {
      let outputBuffer: Buffer;
      let contentType: string;
      let fileExtension: string;

      if (targetFormat === 'ico') {
        outputBuffer = await processEPSFile(file.buffer, file.originalname, iconSizeNum);
        contentType = 'image/x-icon';
        fileExtension = 'ico';
      } else {
        const rasterPng = await processEPSRaster(file.buffer, file.originalname);
        outputBuffer = await sharp(rasterPng, { failOn: 'truncated', unlimited: true })
          .webp({ quality: qualityValue, lossless: lossless === 'true', effort: 6, smartSubsample: true })
          .toBuffer();
        contentType = 'image/webp';
        fileExtension = 'webp';
      }

      return await sendBufferAsDownload(res, file, outputBuffer, fileExtension, contentType);
    }

  } catch (earlyError) {
    console.error('Early format handling error:', earlyError);
    return res.status(500).json({ error: 'Conversion failed', details: earlyError instanceof Error ? earlyError.message : String(earlyError) });
    }

    let sharpInstance = sharp(file.buffer, { 
      failOn: 'truncated',
      unlimited: true // Allow very large images
    });

    // Get image metadata
    const metadata = await sharpInstance.metadata();
    console.log(`Image metadata: ${metadata.width}x${metadata.height}, format: ${metadata.format}`);

    // targetFormat, iconSizeNum, isEPS already computed above

    if (isEPS && targetFormat === 'ico') {
      try {
        const outputBuffer = await processEPSFile(file.buffer, file.originalname, iconSizeNum);
        const contentType = 'image/x-icon';
        const fileExtension = 'ico';

        const rawOriginalName = file.originalname.replace(/\.[^.]+$/, '');
        const fixedOriginalName = fixUTF8Encoding(rawOriginalName);
        const sanitizedName = sanitizeFilename(fixedOriginalName);
        const outputFilename = `${sanitizedName}.${fileExtension}`;

        console.log(`EPS conversion successful: ${outputFilename}, size: ${outputBuffer.length} bytes`);

        await ensureConvertedFilesDir();
        const timestamp = Date.now();
        const uniqueFilename = `${timestamp}_${outputFilename}`;
        const filePath = path.join(CONVERTED_FILES_DIR, uniqueFilename);
        await fs.writeFile(filePath, outputBuffer);
        console.log(`File saved to disk: ${uniqueFilename}`);

        const encodedFilename = encodeURIComponent(outputFilename);
        res.set({
          'Content-Type': contentType,
          'Content-Disposition': `attachment; filename*=UTF-8''${encodedFilename}`,
          'Content-Length': outputBuffer.length.toString(),
          'Cache-Control': 'no-cache',
          'Connection': 'keep-alive'
        });

        try {
          res.send(outputBuffer);
          console.log('EPS file sent successfully');
        } catch (sendError) {
          console.error('Error sending EPS file:', sendError);
          if (!res.headersSent) {
            res.status(500).json({ error: 'Failed to send converted file' });
          }
        }

        setTimeout(async () => {
          try {
            await fs.unlink(filePath);
            console.log(`Cleaned up EPS converted file: ${uniqueFilename}`);
          } catch (cleanupError) {
            console.error('Error cleaning up EPS file:', cleanupError);
          }
        }, 5 * 60 * 1000);

        setTimeout(() => {
          if (global.gc) {
            global.gc();
          }
          console.log('Memory cleanup completed after EPS conversion');
        }, 1000);

        return;
      } catch (epsError) {
        console.error('EPS conversion error:', epsError);
        return res.status(500).json({
          error: 'EPS to ICO conversion failed',
          details: epsError instanceof Error ? epsError.message : 'Unknown EPS processing error'
        });
      }
    }

    // Handle different output formats for standard bitmap flows
    let outputBuffer: Buffer;
    let contentType: string;
    let fileExtension: string;

    switch (targetFormat) {
      case 'webp':
        sharpInstance = sharpInstance.webp({ 
          quality: Number(qualityValue), 
          lossless: isLossless,
          effort: 6, // Higher effort for better quality
          smartSubsample: true // Better color handling
        });
        contentType = 'image/webp';
        fileExtension = 'webp';
        break;

      case 'ico':
        sharpInstance = sharpInstance
          .resize(iconSizeNum, iconSizeNum, { 
            fit: 'contain',
            background: { r: 0, g: 0, b: 0, alpha: 0 }
          })
          .png();
        contentType = 'image/x-icon';
        fileExtension = 'ico';
        break;

      case 'png':
        sharpInstance = sharpInstance.png({ 
          quality: Number(qualityValue),
          compressionLevel: 9
        });
        contentType = 'image/png';
        fileExtension = 'png';
        break;

      case 'jpeg':
      case 'jpg':
        sharpInstance = sharpInstance.jpeg({ 
          quality: Number(qualityValue),
          progressive: true
        });
        contentType = 'image/jpeg';
        fileExtension = 'jpg';
        break;

      default:
        return res.status(400).json({ error: 'Unsupported output format' });
    }

    // Apply resizing if specified
    if (width || height) {
      sharpInstance = sharpInstance.resize(
        width ? parseInt(width) : undefined,
        height ? parseInt(height) : undefined,
        { 
          fit: 'inside',
          withoutEnlargement: true
        }
      );
    }

    // Process the image
    outputBuffer = await sharpInstance.toBuffer();

    // Generate output filename with proper sanitization
    const rawOriginalName = file.originalname.replace(/\.[^.]+$/, '');
    const fixedOriginalName = fixUTF8Encoding(rawOriginalName);
    const sanitizedName = sanitizeFilename(fixedOriginalName);
    const outputFilename = `${sanitizedName}.${fileExtension}`;

    console.log(`Conversion successful: ${outputFilename}, size: ${outputBuffer.length} bytes`);

    // Ensure converted files directory exists
    await ensureConvertedFilesDir();

    // Generate unique filename to avoid conflicts (use original for file system, fixed for display)
    const timestamp = Date.now();
    const uniqueFilename = `${timestamp}_${outputFilename}`;
    const filePath = path.join(CONVERTED_FILES_DIR, uniqueFilename);

    // Save file to disk
    await fs.writeFile(filePath, outputBuffer);
    console.log(`File saved to disk: ${uniqueFilename}`);

    // Set response headers with proper UTF-8 encoding for filenames
    const encodedFilename = encodeURIComponent(outputFilename);
    res.set({
      'Content-Type': contentType,
      'Content-Disposition': `attachment; filename*=UTF-8''${encodedFilename}`,
      'Content-Length': outputBuffer.length.toString(),
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive'
    });

    // Send the converted file
    try {
    res.send(outputBuffer);
      console.log('File sent successfully');
    } catch (sendError) {
      console.error('Error sending file:', sendError);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Failed to send converted file' });
      }
    }

    // Schedule cleanup of the file after 5 minutes
    setTimeout(async () => {
      try {
        await fs.unlink(filePath);
        console.log(`Cleaned up converted file: ${uniqueFilename}`);
      } catch (cleanupError) {
        console.error('Error cleaning up file:', cleanupError);
      }
    }, 5 * 60 * 1000); // 5 minutes

    // Clean up memory after sending
    setTimeout(() => {
      if (global.gc) {
        global.gc();
      }
      console.log('Memory cleanup completed after conversion');
    }, 1000);

  } catch (error) {
    console.error('Conversion error:', error);
    
    // Handle specific Sharp errors
    if (error instanceof Error) {
      if (error.message.includes('Input file is missing')) {
        return res.status(400).json({ error: 'Invalid or corrupted image file' });
      }
      if (error.message.includes('unsupported image format')) {
        return res.status(400).json({ error: 'Unsupported image format' });
      }
    }

    res.status(500).json({ 
      error: 'Conversion failed', 
      details: error instanceof Error ? error.message : String(error)
    });
  }
});

// Batch conversion endpoint with timeout
app.post('/api/convert/batch', uploadBatch.array('files', 20), async (req, res) => {
  // Set timeout for batch processing (2 minutes)
  const timeout = setTimeout(() => {
    if (!res.headersSent) {
      res.status(408).json({ 
        error: 'Batch processing timeout. Please try with fewer files or smaller files.' 
      });
    }
  }, 120000); // 2 minutes

  try {
    const files = req.files as Express.Multer.File[];
    const { 
      quality = 'high', 
      lossless = 'false',
      format = 'webp',
      iconSize = '16',
      delimiter = ',',
      includeMetadata = 'true',
      extractTables = 'true'
    } = req.body;

    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    console.log(`Processing batch: ${files.length} files, total size: ${files.reduce((sum, f) => sum + f.size, 0) / 1024 / 1024}MB`);
    
    // Log filenames for debugging
    files.forEach((file, index) => {
      console.log(`File ${index + 1} filename:`, file.originalname);
      console.log(`File ${index + 1} encoding:`, Buffer.from(file.originalname, 'utf8').toString('hex'));
    });

    // Check total size and reject if too large
    const totalSize = files.reduce((sum, file) => sum + file.size, 0);
    const maxBatchSize = 100 * 1024 * 1024; // 100MB limit for batch processing
    
    if (totalSize > maxBatchSize) {
      return res.status(400).json({ 
        error: `Batch too large. Total size: ${Math.round(totalSize / 1024 / 1024)}MB, maximum allowed: ${Math.round(maxBatchSize / 1024 / 1024)}MB. Please process fewer files at once.` 
      });
    }

    const results = [];
    const qualityValue = quality === 'high' ? 95 : quality === 'medium' ? 80 : 60;
    const isLossless = lossless === 'true';

    // Process files sequentially to avoid memory issues
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      console.log(`Processing file ${i + 1}/${files.length}: ${file.originalname} (${Math.round(file.size / 1024)}KB)`);
      
      try {
        // Check memory usage before processing
        const memUsage = process.memoryUsage();
        console.log(`Memory before processing: ${Math.round(memUsage.heapUsed / 1024 / 1024)}MB`);
        
        // Check if this is a RAW file and process it first
        let imageBuffer = file.buffer;
        if (isRAWFile(file.originalname)) {
          console.log(`Processing RAW file in batch: ${file.originalname}`);
          try {
            imageBuffer = await processRAWFile(file.buffer, file.originalname);
          } catch (rawError) {
            console.error(`RAW processing error for ${file.originalname}:`, rawError);
            results.push({
              originalName: file.originalname,
              success: false,
              error: 'Failed to process RAW file'
            });
            continue;
          }
        }

        const fileIsEPS = isEPSFile(file.originalname) || file.mimetype === 'application/postscript';
        const fileIsCSV = isCSVFile(file.originalname);

        const targetFormat = String(format || '').toLowerCase();
        const iconSizeNum = parseInt(iconSize) || 16;
        let outputBuffer: Buffer;
        let fileExtension: string;

        if (fileIsCSV) {
          try {
            const { buffer, ext, mime } = await processCSVConversion(file.buffer, targetFormat, delimiter);
            outputBuffer = buffer;
            fileExtension = ext;
          } catch (csvErr) {
            throw new Error(csvErr instanceof Error ? csvErr.message : String(csvErr));
          }
        } else if (targetFormat === 'ico') {
          if (fileIsEPS) {
            outputBuffer = await processEPSFile(file.buffer, file.originalname, iconSizeNum);
            fileExtension = 'ico';
          } else {
            const tmpDir = os.tmpdir();
            const uid = randomUUID();
            const tmpPng = path.join(tmpDir, `ico_src_${uid}.png`);
            const tmpIco = path.join(tmpDir, `ico_out_${uid}.ico`);

            await sharp(imageBuffer, { failOn: 'truncated', unlimited: true })
              .resize(iconSizeNum, iconSizeNum, { fit: 'contain', background: { r:0, g:0, b:0, alpha:0 } })
              .png()
              .toFile(tmpPng);

            // Determine ImageMagick binary
            let useMagickRoot = false;
            try { await execAsync('convert -version', { timeout: 2000 }); }
            catch { await execAsync('magick -version', { timeout: 2000 }); useMagickRoot = true; }
            const convertCmd = useMagickRoot
              ? `magick "${tmpPng}" -resize ${iconSizeNum}x${iconSizeNum} "${tmpIco}"`
              : `convert "${tmpPng}" -resize ${iconSizeNum}x${iconSizeNum} "${tmpIco}"`;
            console.log('Batch ImageMagick:', convertCmd);
            await execAsync(convertCmd, { timeout: 30000, maxBuffer: 50 * 1024 * 1024 });

            outputBuffer = await fs.readFile(tmpIco);
            fileExtension = 'ico';

            // Cleanup
            try { await fs.unlink(tmpPng); } catch {}
            try { await fs.unlink(tmpIco); } catch {}
          }
        } else if (targetFormat === 'webp' && fileIsEPS) {
          const rasterPng = await processEPSRaster(file.buffer, file.originalname);
          outputBuffer = await sharp(rasterPng, { failOn: 'truncated', unlimited: true })
            .webp({ quality: Number(qualityValue), lossless: isLossless })
            .toBuffer();
          fileExtension = 'webp';
        } else {
          // Standard bitmap formats handled by Sharp
        let sharpInstance = sharp(imageBuffer, { 
          failOn: 'truncated',
          unlimited: true
        });

          switch (targetFormat) {
          case 'webp':
            if (fileIsEPS) {
              const rasterPng = await processEPSRaster(file.buffer, file.originalname);
              outputBuffer = await sharp(rasterPng, { failOn: 'truncated', unlimited: true })
                .webp({ quality: Number(qualityValue), lossless: isLossless })
                .toBuffer();
            } else {
            outputBuffer = await sharpInstance.webp({ 
              quality: Number(qualityValue), 
              lossless: isLossless 
            }).toBuffer();
            }
            fileExtension = 'webp';
            break;
          case 'png':
            outputBuffer = await sharpInstance.png({ 
              quality: Number(qualityValue)
            }).toBuffer();
            fileExtension = 'png';
            break;
          case 'jpeg':
          case 'jpg':
            outputBuffer = await sharpInstance.jpeg({ 
              quality: Number(qualityValue)
            }).toBuffer();
            fileExtension = 'jpg';
            break;
          default:
            throw new Error('Unsupported format');
        }
        }

        const rawOriginalName = file.originalname.replace(/\.[^.]+$/, '');
        const normalizedOriginalName = fixUTF8Encoding(rawOriginalName);
        const sanitizedName = sanitizeFilename(normalizedOriginalName);
        const outputFilename = `${sanitizedName}.${fileExtension}`;

        // Save file to disk for batch processing
        await ensureConvertedFilesDir();
        const timestamp = Date.now();
        const uniqueFilename = `${timestamp}_${outputFilename}`;
        const filePath = path.join(CONVERTED_FILES_DIR, uniqueFilename);
        await fs.writeFile(filePath, outputBuffer);

        // Schedule cleanup after 5 minutes
        setTimeout(async () => {
          try {
            await fs.unlink(filePath);
            console.log(`Cleaned up batch file: ${uniqueFilename}`);
          } catch (cleanupError) {
            console.error('Error cleaning up batch file:', cleanupError);
          }
        }, 5 * 60 * 1000);

        // Log the filenames for debugging
        console.log('Batch result - originalName:', file.originalname);
        console.log('Batch result - outputFilename:', outputFilename);
        console.log('Batch result - originalName hex:', Buffer.from(file.originalname, 'utf8').toString('hex'));
        console.log('Batch result - outputFilename hex:', Buffer.from(outputFilename, 'utf8').toString('hex'));
        
        // Fix UTF-8 encoding issues before sending to frontend
        const responseOriginalName = sanitizeFilename(fixUTF8Encoding(file.originalname));
        const responseOutputFilename = outputFilename;
        
        console.log('Batch result - response originalName:', responseOriginalName);
        console.log('Batch result - response outputFilename:', responseOutputFilename);
        
        const encodedDownloadPath = `/download/${encodeURIComponent(uniqueFilename)}`;

        results.push({
          originalName: responseOriginalName,
          outputFilename: responseOutputFilename,
          size: outputBuffer.length,
          success: true,
          downloadPath: encodedDownloadPath,
          storedFilename: uniqueFilename
        });

        // Force garbage collection to free memory
        if (global.gc) {
          global.gc();
        }
        
        // Small delay between files to prevent memory buildup
        await new Promise(resolve => setTimeout(resolve, 100));

      } catch (fileError) {
        console.error(`Error processing ${file.originalname}:`, fileError);
        results.push({
          originalName: file.originalname,
          success: false,
          error: fileError instanceof Error ? fileError.message : 'Unknown error'
        });
      }
    }

    clearTimeout(timeout);
    
    // Log final results for debugging
    console.log('Final batch results:', JSON.stringify(results, null, 2));

    res.json({
      success: true,
      processed: results.length,
      results
    });

  } catch (error) {
    clearTimeout(timeout);
    console.error('Batch conversion error:', error);
    res.status(500).json({ 
      error: 'Batch conversion failed',
      details: process.env.NODE_ENV === 'development' ? (error instanceof Error ? error.message : String(error)) : 'Internal server error'
    });
  }
});

// Error handling middleware
app.use((error: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error('Unhandled error:', error);
  
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File too large. Maximum size is 200MB.' });
    }
    if (error.code === 'LIMIT_FILE_COUNT') {
      return res.status(400).json({ error: 'Too many files. Maximum is 10 files.' });
    }
  }

  const details = error instanceof Error ? error.message : String(error);
  res.status(500).json({ 
    error: 'Internal server error',
    details
  });
});

// 404 handler
app.use('*', (req, res) => {
  res.status(404).json({ error: 'Endpoint not found' });
});

// Initialize server
const startServer = async () => {
  try {
    // Ensure converted files directory exists
    await ensureConvertedFilesDir();

    // Start server
    app.listen(Number(PORT), '0.0.0.0', () => {
      console.log(`🚀 Server running on port ${PORT}`);
      console.log(`📁 Health check: http://localhost:${PORT}/health`);
      console.log(`🔧 API status: http://localhost:${PORT}/api/status`);
    });
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
};

// Start the server
startServer();