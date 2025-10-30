import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import multer from 'multer';
import sharp from 'sharp';
import { promises as fs } from 'fs';
import path from 'path';
import os from 'os';
import { execFile, spawn } from 'child_process';
import { promisify } from 'util';
import { randomUUID } from 'crypto';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import mammoth from 'mammoth';
import * as cheerio from 'cheerio';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType } from 'docx';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import dotenv from 'dotenv';

// ES module compatibility: define __dirname
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// No external ICO library needed - we'll create ICO manually

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

// Configure multer for document file uploads (DOCX, RTF, ODT, TXT, etc.)
const uploadDocument = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024, // 100MB limit
    files: 1
  },
  fileFilter: (req, file, cb) => {
    // Allow document and spreadsheet formats
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // DOCX
      'application/msword', // DOC
      'application/rtf', // RTF
      'text/rtf', // RTF alternative
      'application/vnd.oasis.opendocument.text', // ODT
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // XLSX
      'application/vnd.ms-excel', // XLS
      'application/vnd.oasis.opendocument.spreadsheet', // ODS
      // PowerPoint formats
      'application/vnd.openxmlformats-officedocument.presentationml.presentation', // PPTX
      'application/vnd.ms-powerpoint', // PPT
      'application/vnd.openxmlformats-officedocument.presentationml.slideshow', // PPSX
      'application/vnd.ms-powerpoint.presentation.macroEnabled.12', // PPTM
      'application/vnd.ms-powerpoint.slideshow.macroEnabled.12', // PPSM
      'application/vnd.oasis.opendocument.presentation', // ODP
      'application/vnd.oasis.opendocument.presentation-template', // OTP
      'text/plain', // TXT
      'text/markdown', // MD
      'text/x-markdown', // MD alternative
      'application/json', // JSON
      'text/xml', // XML
      'application/xml', // XML alternative
      'text/csv', // CSV
      'application/octet-stream' // Generic fallback
    ];
    
    const extension = file.originalname.split('.').pop()?.toLowerCase();
    const allowedExtensions = [
      // Documents
      'docx', 'doc', 'docm', 'dotx', 'dotm', 'rtf', 'odt', 
      // Spreadsheets
      'xlsx', 'xls', 'xlsm', 'xlsb', 'ods', 
      // Presentations
      'pptx', 'ppt', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'pot', 'pps', 'odp', 'otp', 'sdd', 'sti', 'uop',
      // Text/Code
      'txt', 'log', 'md', 'markdown', 'json', 'xml', 'csv', 'tsv', 'html', 'css', 'js', 'py', 'java', 'c', 'cpp',
      // RAW Image Formats (for viewers)
      'nef', 'cr2', 'dng', 'arw', 'orf', 'raf', 'rw2', 'pef', '3fr', 'dcr', 'kdc', 'mrw', 'nrw', 'sr2', 'srf', 'x3f',
      // Standard Images (for viewers)
      'jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'svg', 'ico', 'heic', 'heif', 'avif',
      // Documents (for viewers)
      'pdf'
    ];
    
    if (allowedMimes.includes(file.mimetype) || allowedExtensions.includes(extension || '')) {
      cb(null, true);
    } else {
      cb(new Error('Unsupported document type'), false);
    }
  }
});

// TypeScript interface for conversion results
interface ConversionResult {
  buffer: Buffer;
  filename: string;
  mime: string;
  storedFilename?: string;
  downloadUrl?: string;
  size?: number;
}

// Utility function to create a basic ICO file
const createBasicICO = (width: number, height: number): Buffer => {
  // Create a simple 32x32 ICO file with a basic structure
  // This is a minimal ICO file format implementation
  const icoHeader = Buffer.alloc(6);
  icoHeader.writeUInt16LE(0, 0); // Reserved, must be 0
  icoHeader.writeUInt16LE(1, 2); // Type: 1 = ICO
  icoHeader.writeUInt16LE(1, 4); // Number of images
  
  // Image directory entry
  const imageDir = Buffer.alloc(16);
  imageDir.writeUInt8(width, 0); // Width
  imageDir.writeUInt8(height, 1); // Height
  imageDir.writeUInt8(0, 2); // Color palette
  imageDir.writeUInt8(0, 3); // Reserved
  imageDir.writeUInt16LE(1, 4); // Color planes
  imageDir.writeUInt16LE(32, 6); // Bits per pixel
  imageDir.writeUInt32LE(0, 8); // Image size (0 for uncompressed)
  imageDir.writeUInt32LE(22, 12); // Offset to image data
  
  // Create a simple 32x32 RGBA image (32x32x4 = 4096 bytes)
  const imageData = Buffer.alloc(4096);
  // Fill with a simple pattern (light gray background)
  for (let i = 0; i < 4096; i += 4) {
    imageData.writeUInt8(200, i); // R
    imageData.writeUInt8(200, i + 1); // G
    imageData.writeUInt8(200, i + 2); // B
    imageData.writeUInt8(255, i + 3); // A
  }
  
  // Combine all parts
  const icoFile = Buffer.concat([icoHeader, imageDir, imageData]);
  return icoFile;
};

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
  return { 
    buffer, 
    filename: downloadName, 
    mime, 
    storedFilename,
    downloadUrl: `/batch-download/${storedFilename}`,
    size: buffer.length
  };
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

// DNG to WebP converter using Python
const convertDngToWebpPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== DNG TO WEBP (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-dng-webp-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write DNG file to temp directory
    const dngPath = path.join(tmpDir, `${safeBase}.dng`);
    await fs.writeFile(dngPath, file.buffer);
    
    // Verify the DNG file was written successfully
    const dngExists = await fs.access(dngPath).then(() => true).catch(() => false);
    if (!dngExists) {
      throw new Error('Failed to write DNG file to temporary directory');
    }
    console.log('DNG file written successfully:', dngPath);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.webp`);

    // Use Python script for WebP
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'dng_to_webp.py');
    
    // Check if Python script exists
    console.log('Checking for Python script at:', scriptPath);
    const scriptExists = await fs.access(scriptPath).then(() => true).catch((err) => {
      console.error('Script access check failed:', err);
      return false;
    });
    
    if (!scriptExists) {
      console.error(`!!! Python script not found at: ${scriptPath} !!!`);
      console.error('Current directory (__dirname):', __dirname);
      console.error('Expected script path:', scriptPath);
      
      // Try to list the scripts directory
      try {
        const scriptsDir = path.join(__dirname, '..', 'scripts');
        console.log('Attempting to list scripts directory:', scriptsDir);
        const files = await fs.readdir(scriptsDir);
        console.log('Files in scripts directory:', files);
      } catch (listError) {
        console.error('Could not list scripts directory:', listError);
      }
      
      throw new Error('DNG to WebP conversion script not found. Please check deployment.');
    }
    
    console.log('✅ Python script found, proceeding with conversion...');

    // Check if Python is available
    try {
      console.log('Checking Python availability...');
      const pythonCheck = await execFileAsync(pythonPath, ['--version']);
      console.log('Python version:', pythonCheck.stdout.trim());
    } catch (pythonError) {
      console.error('!!! Python is not available !!!');
      console.error('Python path tried:', pythonPath);
      console.error('Error:', pythonError);
      throw new Error('Python is not available on the system');
    }

    // Parse options
    // Convert quality from 0-1 range to 1-100 range for WebP
    let qualityValue = parseFloat(options.quality || '0.95');
    if (qualityValue <= 1) {
      // Convert from 0-1 scale to 1-100 scale
      qualityValue = Math.round(qualityValue * 100);
    }
    const quality = Math.max(1, Math.min(100, qualityValue)); // Ensure between 1-100
    const lossless = options.lossless === 'true';
    const width = options.width ? parseInt(options.width) : undefined;
    const height = options.height ? parseInt(options.height) : undefined;

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      dngPath,
      outputPath,
      quality,
      lossless,
      width,
      height,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      dngPath,
      outputPath,
      '--quality', quality.toString(),
    ];

    if (lossless) {
      args.push('--lossless');
    }

    if (width) {
      args.push('--width', width.toString());
    }

    if (height) {
      args.push('--height', height.toString());
    }

    let stdout, stderr;
    let pythonFailed = false;
    try {
      console.log('Executing Python script:', pythonPath, args.join(' '));
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
      console.log('Python script execution completed successfully');
    } catch (execError: any) {
      pythonFailed = true;
      console.error('!!! Python script execution FAILED !!!');
      console.error('Error type:', typeof execError);
      console.error('Error details:', execError);
      if (execError instanceof Error) {
        console.error('Error message:', execError.message);
        console.error('Error stack:', execError.stack);
      }
      // When execFileAsync fails, the error object contains stdout and stderr
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
      console.error('Python stdout:', stdout);
      console.error('Python stderr:', stderr);
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for specific error patterns
    if (stderr.includes('Unsupported DNG file format')) {
      throw new Error('Unsupported DNG file format. Please ensure the file is a valid DNG image.');
    }
    if (stderr.includes('DNG file I/O error')) {
      throw new Error('DNG file I/O error. Please check that the file is not corrupted.');
    }
    if (stderr.includes('rawpy not available') || stderr.includes('ImportError')) {
      throw new Error('Required Python library (rawpy) is not available. Please install it.');
    }
    if (stderr.includes('Traceback') || stderr.includes('SyntaxError')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      console.error('!!! Python script did not produce output file !!!');
      console.error('Expected output path:', outputPath);
      console.error('Full Python stdout:', stdout);
      console.error('Full Python stderr:', stderr);
      
      // Try to provide more helpful error message based on stderr
      let errorMessage = `Python WebP script did not produce output file.`;
      
      if (stderr.includes('ModuleNotFoundError') || stderr.includes('No module named')) {
        errorMessage = 'Missing Python dependencies (rawpy or Pillow). Please install them.';
      } else if (stderr.includes('ERROR:') || stderr.includes('Failed')) {
        errorMessage = `Python error: ${stderr}`;
      } else if (stdout.includes('ERROR:')) {
        errorMessage = `Python error: ${stdout}`;
      } else {
        errorMessage = `Conversion failed. Check logs for details. stderr: ${stderr}, stdout: ${stdout}`;
      }
      
      throw new Error(errorMessage);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python WebP script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.webp`;
    console.log(`DNG->WebP conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'image/webp');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'image/webp'
    };
  } catch (error) {
    console.error(`DNG->WebP conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown DNG->WebP error`;
    throw new Error(`Failed to convert DNG to WebP: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// DNG to ICO converter using Python
const convertDngToIcoPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== DNG TO ICO (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-dng-ico-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write DNG file to temp directory
    const dngPath = path.join(tmpDir, `${safeBase}.dng`);
    await fs.writeFile(dngPath, file.buffer);
    
    // Verify the DNG file was written successfully
    const dngExists = await fs.access(dngPath).then(() => true).catch(() => false);
    if (!dngExists) {
      throw new Error('Failed to write DNG file to temporary directory');
    }
    console.log('DNG file written successfully:', dngPath);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.ico`);

    // Use Python script for ICO with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'dng_to_ico.py');
    
    // Check if Python script exists
    console.log('Checking for Python script at:', scriptPath);
    const scriptExists = await fs.access(scriptPath).then(() => true).catch((err) => {
      console.error('Script access check failed:', err);
      return false;
    });
    
    if (!scriptExists) {
      console.error(`!!! Python script not found at: ${scriptPath} !!!`);
      console.error('Current directory (__dirname):', __dirname);
      console.error('Expected script path:', scriptPath);
      
      // Try to list the scripts directory
      try {
        const scriptsDir = path.join(__dirname, '..', 'scripts');
        console.log('Attempting to list scripts directory:', scriptsDir);
        const files = await fs.readdir(scriptsDir);
        console.log('Files in scripts directory:', files);
      } catch (listError) {
        console.error('Could not list scripts directory:', listError);
      }
      
      throw new Error('DNG to ICO conversion script not found. Please check deployment.');
    }
    
    console.log('✅ Python script found, proceeding with conversion...');

    // Check if Python is available
    try {
      console.log('Checking Python availability...');
      const pythonCheck = await execFileAsync(pythonPath, ['--version']);
      console.log('Python version:', pythonCheck.stdout.trim());
    } catch (pythonError) {
      console.error('!!! Python is not available !!!');
      console.error('Python path tried:', pythonPath);
      console.error('Error:', pythonError);
      throw new Error('Python is not available on the system');
    }

    // Parse options
    const sizes = options.sizes 
      ? options.sizes.split(',').map(s => parseInt(s.trim())).filter(s => !isNaN(s) && s >= 16 && s <= 256)
      : [16, 32, 48, 64, 128, 256];
    const qualityLevel = options.quality || 'high';

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      dngPath,
      outputPath,
      sizes,
      qualityLevel,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      dngPath,
      outputPath,
      '--sizes', ...sizes.map(s => s.toString()),
      '--quality', qualityLevel
    ];

    let stdout, stderr;
    let pythonFailed = false;
    try {
      console.log('Executing Python script:', pythonPath, args.join(' '));
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
      console.log('Python script execution completed successfully');
    } catch (execError: any) {
      pythonFailed = true;
      console.error('!!! Python script execution FAILED !!!');
      console.error('Error type:', typeof execError);
      console.error('Error details:', execError);
      if (execError instanceof Error) {
        console.error('Error message:', execError.message);
        console.error('Error stack:', execError.stack);
      }
      // When execFileAsync fails, the error object contains stdout and stderr
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
      console.error('Python stdout:', stdout);
      console.error('Python stderr:', stderr);
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for specific error patterns
    if (stderr.includes('Unsupported DNG file format')) {
      throw new Error('Unsupported DNG file format. Please ensure the file is a valid DNG image.');
    }
    if (stderr.includes('DNG file I/O error')) {
      throw new Error('DNG file I/O error. Please check that the file is not corrupted.');
    }
    if (stderr.includes('rawpy not available') || stderr.includes('ImportError')) {
      throw new Error('Required Python library (rawpy) is not available. Please install it.');
    }
    if (stderr.includes('Traceback') || stderr.includes('SyntaxError')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      console.error('!!! Python script did not produce output file !!!');
      console.error('Expected output path:', outputPath);
      console.error('Full Python stdout:', stdout);
      console.error('Full Python stderr:', stderr);
      
      // Try to provide more helpful error message based on stderr
      let errorMessage = `Python ICO script did not produce output file: ${outputPath}.`;
      
      if (stderr.includes('ModuleNotFoundError') || stderr.includes('No module named')) {
        errorMessage += ' Missing Python dependencies (rawpy or Pillow). Please install them.';
      } else if (stderr.includes('ERROR:') || stderr.includes('Failed')) {
        errorMessage += ` Python error: ${stderr.substring(0, 200)}`;
      } else if (stdout.includes('ERROR:') || stdout.includes('Failed')) {
        errorMessage += ` Python error: ${stdout.substring(0, 200)}`;
      } else {
        errorMessage += ' Check Python environment and dependencies.';
      }
      
      throw new Error(errorMessage);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python ICO script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.ico`;
    console.log(`DNG->ICO conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'image/x-icon');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'image/x-icon'
    };
  } catch (error) {
    console.error(`DNG->ICO conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown DNG->ICO error`;
    throw new Error(`Failed to convert DNG to ICO: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// EPS to ICO converter using Python
const convertEpsToIcoPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== EPS TO ICO (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-eps-ico-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write EPS file to temp directory
    const epsPath = path.join(tmpDir, `${safeBase}.eps`);
    await fs.writeFile(epsPath, file.buffer);
    
    // Verify the EPS file was written successfully
    const epsExists = await fs.access(epsPath).then(() => true).catch(() => false);
    if (!epsExists) {
      throw new Error('Failed to write EPS file to temporary directory');
    }
    console.log('EPS file written successfully:', epsPath);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.ico`);

    // Get script path
    const scriptPath = path.join(__dirname, '..', 'scripts', 'eps_to_ico.py');
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      throw new Error(`Python script not found: ${scriptPath}`);
    }

    // Get Python path
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    
    // Parse options
    const sizes = options.sizes ? options.sizes.split(',').map(s => parseInt(s.trim())) : [16, 32, 48, 64, 128, 256];
    const quality = options.quality || 'high';

    console.log('Conversion parameters:', {
      pythonPath,
      scriptPath,
      epsPath,
      outputPath,
      sizes,
      quality,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      epsPath,
      outputPath,
      '--sizes', ...sizes.map(s => s.toString()),
      '--quality', quality
    ];

    let stdout, stderr;
    let pythonFailed = false;
    try {
      console.log('Executing Python script:', pythonPath, args.join(' '));
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
      console.log('Python script execution completed successfully');
    } catch (execError: any) {
      pythonFailed = true;
      console.error('!!! Python script execution FAILED !!!');
      console.error('Error type:', typeof execError);
      console.error('Error details:', execError);
      console.error('stdout:', stdout);
      console.error('stderr:', stderr);
      throw new Error(`Python ICO script execution failed: ${execError.message}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      console.error('Python ICO script did not produce output file:', outputPath);
      console.error('Python stdout:', stdout);
      console.error('Python stderr:', stderr);
      throw new Error(`Python ICO script did not produce output file: ${outputPath}. Python error: ${stderr}`);
    }

    console.log('ICO file created successfully:', outputPath);
    
    // Read the output file
    const outputBuffer = await fs.readFile(outputPath);
    const downloadName = `${originalBase}.ico`;

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'image/x-icon');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'image/x-icon'
    };
  } catch (error) {
    console.error(`EPS->ICO conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown EPS->ICO error`;
    throw new Error(`Failed to convert EPS to ICO: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// EPS to WebP converter using Python
const convertEpsToWebpPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== EPS TO WEBP (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-eps-webp-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write EPS file to temp directory
    const epsPath = path.join(tmpDir, `${safeBase}.eps`);
    await fs.writeFile(epsPath, file.buffer);
    
    // Verify the EPS file was written successfully
    const epsExists = await fs.access(epsPath).then(() => true).catch(() => false);
    if (!epsExists) {
      throw new Error('Failed to write EPS file to temporary directory');
    }
    console.log('EPS file written successfully:', epsPath);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.webp`);

    // Get script path
    const scriptPath = path.join(__dirname, '..', 'scripts', 'eps_to_webp.py');
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      throw new Error(`Python script not found: ${scriptPath}`);
    }

    // Get Python path
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    
    // Parse options
    const qualityValue = options.quality || 'high';
    let quality = 80; // Default quality
    
    // Handle string quality values
    if (typeof qualityValue === 'string') {
      switch (qualityValue.toLowerCase()) {
        case 'high':
          quality = 90;
          break;
        case 'medium':
          quality = 70;
          break;
        case 'low':
          quality = 50;
          break;
        default:
          // Try to parse as number
          const parsedQuality = parseInt(qualityValue);
          if (!isNaN(parsedQuality) && parsedQuality >= 0 && parsedQuality <= 100) {
            quality = parsedQuality;
          } else {
            console.warn(`Invalid quality value: ${qualityValue}, using default 80`);
            quality = 80;
          }
      }
    } else {
      // Handle numeric quality values
      const parsedQuality = parseInt(String(qualityValue));
      if (!isNaN(parsedQuality) && parsedQuality >= 0 && parsedQuality <= 100) {
        quality = parsedQuality;
      } else {
        console.warn(`Invalid quality value: ${qualityValue}, using default 80`);
        quality = 80;
      }
    }
    
    const lossless = options.lossless === 'true';

    console.log('Conversion parameters:', {
      pythonPath,
      scriptPath,
      epsPath,
      outputPath,
      quality,
      lossless,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      epsPath,
      outputPath,
      '--quality', quality.toString()
    ];

    if (lossless) {
      args.push('--lossless');
    }

    let stdout, stderr;
    let pythonFailed = false;
    try {
      console.log('Executing Python script:', pythonPath, args.join(' '));
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
      console.log('Python script execution completed successfully');
    } catch (execError: any) {
      pythonFailed = true;
      console.error('!!! Python script execution FAILED !!!');
      console.error('Error type:', typeof execError);
      console.error('Error details:', execError);
      console.error('stdout:', stdout);
      console.error('stderr:', stderr);
      throw new Error(`Python WebP script execution failed: ${execError.message}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      console.error('Python WebP script did not produce output file:', outputPath);
      console.error('Python stdout:', stdout);
      console.error('Python stderr:', stderr);
      throw new Error(`Python WebP script did not produce output file: ${outputPath}. Python error: ${stderr}`);
    }

    console.log('WebP file created successfully:', outputPath);
    
    // Read the output file
    const outputBuffer = await fs.readFile(outputPath);
    const downloadName = `${originalBase}.webp`;

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'image/webp');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'image/webp'
    };
  } catch (error) {
    console.error(`EPS->WebP conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown EPS->WebP error`;
    throw new Error(`Failed to convert EPS to WebP: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// GIF to ICO converter using Python
const convertGifToIcoPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== GIF TO ICO (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-gif-ico-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write GIF file to temp directory
    const gifPath = path.join(tmpDir, `${safeBase}.gif`);
    await fs.writeFile(gifPath, file.buffer);
    
    // Verify the GIF file was written successfully
    const gifExists = await fs.access(gifPath).then(() => true).catch(() => false);
    if (!gifExists) {
      throw new Error('Failed to write GIF file to temporary directory');
    }
    console.log('GIF file written successfully:', gifPath);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.ico`);

    // Get script path
    const scriptPath = path.join(__dirname, '..', 'scripts', 'gif_to_ico.py');
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      throw new Error(`Python script not found: ${scriptPath}`);
    }

    // Get Python path
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    
    // Parse options
    const sizes = options.sizes ? options.sizes.split(',').map(s => parseInt(s.trim())) : [16, 32, 48, 64, 128, 256];
    const quality = options.quality || 'high';

    console.log('Conversion parameters:', {
      pythonPath,
      scriptPath,
      gifPath,
      outputPath,
      sizes,
      quality,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      gifPath,
      outputPath,
      '--sizes', ...sizes.map(s => s.toString()),
      '--quality', quality
    ];

    let stdout, stderr;
    let pythonFailed = false;
    try {
      console.log('Executing Python script:', pythonPath, args.join(' '));
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
      console.log('Python script execution completed successfully');
    } catch (execError: any) {
      pythonFailed = true;
      console.error('!!! Python script execution FAILED !!!');
      console.error('Error type:', typeof execError);
      console.error('Error details:', execError);
      console.error('stdout:', stdout);
      console.error('stderr:', stderr);
      throw new Error(`Python ICO script execution failed: ${execError.message}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      console.error('Python ICO script did not produce output file:', outputPath);
      console.error('Python stdout:', stdout);
      console.error('Python stderr:', stderr);
      throw new Error(`Python ICO script did not produce output file: ${outputPath}. Python error: ${stderr}`);
    }

    console.log('ICO file created successfully:', outputPath);
    
    // Read the output file
    const outputBuffer = await fs.readFile(outputPath);
    const downloadName = `${originalBase}.ico`;

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'image/x-icon');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'image/x-icon'
    };
  } catch (error) {
    console.error(`GIF->ICO conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown GIF->ICO error`;
    throw new Error(`Failed to convert GIF to ICO: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to JSON converter using Python
const convertCsvToJsonPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO JSON (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-json-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);
    
    console.log('CSV file written successfully:', csvPath);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.json`);

    // Use Python script for JSON with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_json.py');
    
    // Check if Python script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      throw new Error('CSV to JSON conversion script not found');
    }

    // Parse options
    const orient = options.orient || 'records';
    const indent = options.indent ? parseInt(options.indent) : 2;
    const dateFormat = options.dateFormat || 'iso';

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      orient,
      indent,
      dateFormat,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--orient', orient,
      '--indent', indent.toString(),
      '--date-format', dateFormat
    ];

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

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

    const downloadName = `${sanitizedBase}.json`;
    console.log(`CSV->JSON conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/json');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/json'
    };
  } catch (error) {
    console.error(`CSV->JSON conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->JSON error`;
    throw new Error(`Failed to convert CSV to JSON: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to NDJSON converter using Python
const convertCsvToNdjsonPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO NDJSON (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-ndjson-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);
    
    console.log('CSV file written successfully:', csvPath);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.ndjson`);

    // Use Python script for NDJSON with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_ndjson.py');
    
    // Check if Python script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      throw new Error('CSV to NDJSON conversion script not found');
    }

    // Parse options
    const includeHeaders = options.includeHeaders || 'true';

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      includeHeaders,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--include-headers', includeHeaders
    ];

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

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

    const downloadName = `${sanitizedBase}.ndjson`;
    console.log(`CSV->NDJSON conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/x-ndjson');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/x-ndjson'
    };
  } catch (error) {
    console.error(`CSV->NDJSON conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->NDJSON error`;
    throw new Error(`Failed to convert CSV to NDJSON: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

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
  if (!file) {
    console.log('DNG file detection: No file provided');
    return false;
  }
  
  // Try to detect DNG by filename first
  if (file.originalname) {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const result = ext === 'dng';
    console.log('DNG file detection (by filename):', {
    filename: file.originalname,
    extension: ext,
    isDNG: result
  });
  return result;
  }
  
  // Fallback: try to detect DNG by MIME type or file content
  if (file.mimetype === 'application/octet-stream' || file.mimetype === 'image/x-adobe-dng') {
    console.log('DNG file detection (by MIME type):', {
      mimetype: file.mimetype,
      isDNG: true
    });
    return true;
  }
  
  console.log('DNG file detection: No filename or MIME type match');
  return false;
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

const isBmpFile = (file: Express.Multer.File) => {
  const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
  const mimetype = file.mimetype?.toLowerCase() ?? '';
  const result = ext === 'bmp' || mimetype.includes('bmp') || mimetype.includes('image/bmp');
  
  console.log('BMP file detection:', {
    filename: file.originalname,
    extension: ext,
    mimetype,
    isBMP: result
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
        timeout: 300000 // 5 minute timeout for large files
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
      
      // Check if it's a timeout error
      if (error?.code === 'TIMEOUT' || error?.signal === 'SIGTERM') {
        console.log(`Calibre conversion timed out after 5 minutes for binary: ${binary}`);
        lastError = new Error(`Calibre conversion timed out after 5 minutes. Large files may need more time to process.`);
        continue;
      }
      
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

// Fallback function to create simple HTML from CSV
const createSimpleHtmlFromCsv = async (
  file: Express.Multer.File,
  title: string,
  author: string
): Promise<string> => {
  try {
    // Parse CSV using Papa Parse
    const csvText = file.buffer.toString('utf-8');
    const parsed = Papa.parse<string[]>(csvText, { 
      skipEmptyLines: true,
      transform: (value) => {
        if (typeof value !== 'string') return String(value || '');
        return value.trim();
      }
    });
    
    const rows: string[][] = parsed && Array.isArray((parsed as any).data)
      ? ((parsed as any).data as unknown as string[][]).map((r: unknown) => {
          if (Array.isArray(r)) {
            return r.map(cell => String(cell || '').trim());
          }
          return [String(r || '').trim()];
        }).filter(row => row.some(cell => cell.length > 0))
      : [];

    // Create simple HTML
    let htmlContent = `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>${title}</title>
    <meta name="author" content="${author}">
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        h1 { color: #333; border-bottom: 2px solid #333; padding-bottom: 10px; }
    </style>
</head>
<body>
    <h1>${title}</h1>
    <p><strong>Author:</strong> ${author}</p>
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
    return htmlContent;
  } catch (error) {
    console.error('Fallback HTML generation failed:', error);
    return `<!DOCTYPE html><html><head><title>Error</title></head><body><h1>Conversion Error</h1><p>Failed to convert CSV file.</p></body></html>`;
  }
};

// Simple CSV to MOBI converter using fallback approach
const convertCsvToMobiSimple = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO MOBI (Simple) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-mobi-simple-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Create simple HTML from CSV
    const htmlContent = await createSimpleHtmlFromCsv(file, options.title || sanitizedBase, options.author || 'Unknown');
    
    // Write HTML file
    const htmlPath = path.join(tmpDir, `${safeBase}.html`);
    await fs.writeFile(htmlPath, htmlContent, 'utf-8');
    
    console.log(`Generated HTML: ${htmlPath} (${htmlContent.length} characters)`);
    
    // Convert HTML to MOBI using Calibre with simple command
    const outputPath = path.join(tmpDir, `${safeBase}.mobi`);
    
    const calibreArgs = [
      htmlPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--authors', options.author || 'Unknown'
    ];
    
    console.log('Converting HTML to MOBI with Calibre:', calibreArgs.join(' '));
    console.log(`HTML file size: ${htmlContent.length} characters`);
    console.log(`Starting Calibre conversion for file: ${file.originalname} (${file.buffer.length} bytes)`);
    
    const { stdout, stderr } = await execCalibre(calibreArgs);
    
    if (stdout.trim().length > 0) console.log('Calibre stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Calibre stderr:', stderr.trim());
    
    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Calibre did not produce output file: ${outputPath}`);
    }
    
    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Calibre produced empty output file');
    }
    
    const downloadName = `${sanitizedBase}.mobi`;
    console.log(`CSV->MOBI (simple) conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/x-mobipocket-ebook');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/x-mobipocket-ebook'
    };
  } catch (error) {
    console.error(`CSV->MOBI (simple) conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->MOBI error`;
    throw new Error(`Failed to convert CSV to MOBI (simple): ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV -> E-book using Python script (pandas + jinja2 + ebooklib)
// Optimized CSV to MOBI converter for large files
const convertCsvToMobiOptimized = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO MOBI (Optimized Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-mobi-opt-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.mobi`);
    
    // Use optimized Python script for MOBI
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_mobi_optimized.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown',
      '--rows-per-chapter', '1000' // Optimize for large files
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) {
      console.warn('Python stderr:', stderr.trim());
      // If there's stderr, it might indicate an error
      if (stderr.includes('ERROR') || stderr.includes('FATAL')) {
        throw new Error(`Python script error: ${stderr.trim()}`);
      }
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Optimized Python script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Optimized Python script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.mobi`;
    console.log(`CSV->MOBI (optimized) conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/x-mobipocket-ebook');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/x-mobipocket-ebook'
    };
  } catch (error) {
    console.error(`CSV->MOBI (optimized) conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->MOBI error`;
    throw new Error(`Failed to convert CSV to MOBI (optimized): ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};


// CSV to ODT converter using Python
const convertCsvToOdtPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO ODT (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-odt-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.odt`);
    
    // Use Python script for ODT
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_odt.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python ODT script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python ODT script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.odt`;
    console.log(`CSV->ODT conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.oasis.opendocument.text');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.oasis.opendocument.text'
    };
  } catch (error) {
    console.error(`CSV->ODT conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->ODT error`;
    throw new Error(`Failed to convert CSV to ODT: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to PDF converter using Python
const convertCsvToPdfPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO PDF (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-pdf-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.pdf`);
    
    // Use Python script for PDF
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_pdf.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python PDF script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python PDF script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.pdf`;
    console.log(`CSV->PDF conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/pdf');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/pdf'
    };
  } catch (error) {
    console.error(`CSV->PDF conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->PDF error`;
    throw new Error(`Failed to convert CSV to PDF: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to DOCX converter using Python
const convertCsvToDocxPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO DOCX (Python) START ===`);
  const startTime = Date.now();
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-docx-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.docx`);
    
    // Use existing Python script for DOCX
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_docx.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ], {
      timeout: 300000, // 5 minutes timeout for large files
      maxBuffer: 10 * 1024 * 1024 // 10MB buffer for large outputs
    });

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python DOCX script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python DOCX script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.docx`;
    const processingTime = Date.now() - startTime;
    console.log(`CSV->DOCX conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length,
      processingTimeMs: processingTime,
      processingTimeSec: (processingTime / 1000).toFixed(2)
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
  } catch (error) {
    console.error(`CSV->DOCX conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->DOCX error`;
    throw new Error(`Failed to convert CSV to DOCX: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};


// CSV to PPT converter using Python
const convertCsvToPptPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO PPT (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-ppt-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.pptx`);
    
    // Use Python script for PPT
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_ppt.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown',
      '--max-rows-per-slide', '50' // Optimize for presentations
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python PPT script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python PPT script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.pptx`;
    console.log(`CSV->PPT conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    };
  } catch (error) {
    console.error(`CSV->PPT conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->PPT error`;
    throw new Error(`Failed to convert CSV to PPT: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to PPTX converter using Python
const convertCsvToPptxPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO PPTX (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-pptx-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.pptx`);
    
    // Use Python script for PPTX
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_pptx.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown',
      '--max-rows-per-slide', '50' // Optimize for presentations
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python PPTX script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python PPTX script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.pptx`;
    console.log(`CSV->PPTX conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    };
  } catch (error) {
    console.error(`CSV->PPTX conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->PPTX error`;
    throw new Error(`Failed to convert CSV to PPTX: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to RTF converter using Python
const convertCsvToRtfPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO RTF (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-rtf-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.rtf`);
    
    // Use Python script for RTF
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_rtf.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python RTF script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python RTF script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.rtf`;
    console.log(`CSV->RTF conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/rtf');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/rtf'
    };
  } catch (error) {
    console.error(`CSV->RTF conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->RTF error`;
    throw new Error(`Failed to convert CSV to RTF: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to TXT converter using Python
const convertCsvToTxtPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO TXT (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-txt-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.txt`);

    // Use Python script for TXT
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_txt.py');

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python TXT script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python TXT script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.txt`;
    console.log(`CSV->TXT conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'text/plain');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'text/plain'
    };
  } catch (error) {
    console.error(`CSV->TXT conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->TXT error`;
    throw new Error(`Failed to convert CSV to TXT: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to XLS converter using Python
const convertCsvToXlsPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO XLS (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-xls-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.xls`);

    // Use Python script for XLS
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_xls.py');

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python XLS script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python XLS script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.xls`;
    console.log(`CSV->XLS conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.ms-excel');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.ms-excel'
    };
  } catch (error) {
    console.error(`CSV->XLS conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->XLS error`;
    throw new Error(`Failed to convert CSV to XLS: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to XLSX converter using Python
const convertCsvToXlsxPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO XLSX (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-xlsx-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.xlsx`);

    // Use Python script for XLSX
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_xlsx.py');

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      fileSize: file.buffer.length
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ]);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python XLSX script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python XLSX script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.xlsx`;
    console.log(`CSV->XLSX conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };
  } catch (error) {
    console.error(`CSV->XLSX conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->XLSX error`;
    throw new Error(`Failed to convert CSV to XLSX: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

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
      path.join(__dirname, '..', 'scripts', 'csv_to_ebook.py'),
      csvPath,
      outputPath,
      targetFormat,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ];

    console.log('Running Python CSV to e-book converter with args:', pythonArgs);

    // Execute Python script using virtual environment
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_ebook.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      targetFormat,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown'
    });

    try {
      const { stdout, stderr } = await execFileAsync(pythonPath, [
        scriptPath,
        csvPath,
        outputPath,
        targetFormat,
        '--title', options.title || sanitizedBase,
        '--author', options.author || 'Unknown'
      ]);

      if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
      if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    } catch (pythonError) {
      console.error('Python execution failed:', pythonError);
      console.log('Falling back to simple conversion...');
      
      // For MOBI format, create HTML and convert with Calibre
      if (targetFormat === 'mobi') {
        const htmlPath = path.join(tmpDir, `${safeBase}.html`);
        const fallbackHtml = await createSimpleHtmlFromCsv(file, options.title || sanitizedBase, options.author || 'Unknown');
        await fs.writeFile(htmlPath, fallbackHtml, 'utf-8');
        
        // Convert HTML to MOBI using Calibre
        try {
          const calibreArgs = [
            htmlPath, 
            outputPath,
            '--output-profile=kindle',
            '--disable-font-rescaling'
          ];
          
          if (options.title) calibreArgs.push('--title', String(options.title));
          if (options.author) calibreArgs.push('--authors', String(options.author));
          
          const { stdout, stderr } = await execCalibre(calibreArgs);
          if (stdout.trim().length > 0) console.log('Calibre stdout:', stdout.trim());
          if (stderr.trim().length > 0) console.warn('Calibre stderr:', stderr.trim());
        } catch (calibreError) {
          console.error('Calibre fallback failed:', calibreError);
          throw new Error('Both Python and Calibre conversion failed');
        }
      } else {
        // For other formats, create appropriate file
        const fallbackContent = await createSimpleHtmlFromCsv(file, options.title || sanitizedBase, options.author || 'Unknown');
        await fs.writeFile(outputPath, fallbackContent, 'utf-8');
      }
    }

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
    console.error('Error details:', {
      errorType: error instanceof Error ? error.constructor.name : typeof error,
      errorMessage: error instanceof Error ? error.message : String(error),
      errorStack: error instanceof Error ? error.stack : undefined,
      targetFormat,
      fileSize: file.buffer.length,
      fileName: file.originalname
    });
    
    const message = error instanceof Error ? error.message : `Unknown CSV->${targetFormat.toUpperCase()} error`;
    throw new Error(`Failed to convert CSV to ${targetFormat.toUpperCase()}: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to HTML converter using Python
const convertCsvToHtmlPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO HTML (Python) START ===`);
  const startTime = Date.now();
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-html-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.html`);
    
    // Use Python script for HTML
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_html.py');
    
    // Determine chunk size based on file size for optimal performance
    const fileSizeMB = file.buffer.length / (1024 * 1024);
    const chunkSize = fileSizeMB > 10 ? 2000 : fileSizeMB > 5 ? 1500 : 1000;
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      tableClass: options.tableClass || 'simple',
      includeHeaders: options.includeHeaders !== 'false',
      fileSize: file.buffer.length,
      fileSizeMB: fileSizeMB.toFixed(2),
      chunkSize
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--table-class', options.tableClass || 'simple',
      '--chunk-size', chunkSize.toString()
    ].concat(options.includeHeaders === 'false' ? ['--no-headers'] : []), {
      timeout: 300000, // 5 minutes timeout for large files
      maxBuffer: 10 * 1024 * 1024 // 10MB buffer for large outputs
    });

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python HTML script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python HTML script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.html`;
    const processingTime = Date.now() - startTime;
    console.log(`CSV->HTML conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length,
      processingTimeMs: processingTime,
      processingTimeSec: (processingTime / 1000).toFixed(2)
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'text/html');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'text/html'
    };
  } catch (error) {
    console.error(`CSV->HTML conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->HTML error`;
    throw new Error(`Failed to convert CSV to HTML: ${message}`);
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

// CSV to Markdown converter using Python
const convertCsvToMdPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO MARKDOWN (Python) START ===`);
  const startTime = Date.now();
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-md-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.md`);

    // Use Python script for Markdown
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_md.py');

    // Determine chunk size based on file size for optimal performance
    const fileSizeMB = file.buffer.length / (1024 * 1024);
    const chunkSize = fileSizeMB > 10 ? 2000 : fileSizeMB > 5 ? 1500 : 1000;

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      tableAlignment: options.tableAlignment || 'left',
      includeHeaders: options.includeHeaders !== 'false',
      fileSize: file.buffer.length,
      fileSizeMB: fileSizeMB.toFixed(2),
      chunkSize
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--table-alignment', options.tableAlignment || 'left',
      '--chunk-size', chunkSize.toString()
    ].concat(options.includeHeaders === 'false' ? ['--no-headers'] : []), {
      timeout: 300000, // 5 minutes timeout for large files
      maxBuffer: 10 * 1024 * 1024 // 10MB buffer for large outputs
    });

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python Markdown script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python Markdown script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.md`;
    const processingTime = Date.now() - startTime;
    console.log(`CSV->Markdown conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length,
      processingTimeMs: processingTime,
      processingTimeSec: (processingTime / 1000).toFixed(2)
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'text/markdown');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'text/markdown'
    };
  } catch (error) {
    console.error(`CSV->Markdown conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->Markdown error`;
    throw new Error(`Failed to convert CSV to Markdown: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to MOBI converter using Python
const convertCsvToMobiPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO MOBI (Python) START ===`);
  const startTime = Date.now();
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-mobi-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.mobi`);

    // Use Python script for MOBI
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_mobi.py');

    // Determine chunk size based on file size for optimal performance
    const fileSizeMB = file.buffer.length / (1024 * 1024);
    const chunkSize = fileSizeMB > 10 ? 2000 : fileSizeMB > 5 ? 1500 : 1000;

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      bookTitle: options.bookTitle || sanitizedBase,
      author: options.author || 'CSV Converter',
      includeHeaders: options.includeHeaders !== 'false',
      fileSize: file.buffer.length,
      fileSizeMB: fileSizeMB.toFixed(2),
      chunkSize
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.bookTitle || sanitizedBase,
      '--author', options.author || 'CSV Converter',
      '--chunk-size', chunkSize.toString()
    ].concat(options.includeHeaders === 'false' ? ['--no-headers'] : []), {
      timeout: 300000, // 5 minutes timeout for large files
      maxBuffer: 10 * 1024 * 1024 // 10MB buffer for large outputs
    });

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python MOBI script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python MOBI script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.mobi`;
    const processingTime = Date.now() - startTime;
    console.log(`CSV->MOBI conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length,
      processingTimeMs: processingTime,
      processingTimeSec: (processingTime / 1000).toFixed(2)
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/x-mobipocket-ebook');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/x-mobipocket-ebook'
    };
  } catch (error) {
    console.error(`CSV->MOBI conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->MOBI error`;
    throw new Error(`Failed to convert CSV to MOBI: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to ODP converter using Python
const convertCsvToOdpPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO ODP (Python) START ===`);
  const startTime = Date.now();
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-odp-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.odp`);

    // Use Python script for ODP
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_odp.py');

    // Determine chunk size based on file size for optimal performance
    const fileSizeMB = file.buffer.length / (1024 * 1024);
    const chunkSize = fileSizeMB > 10 ? 2000 : fileSizeMB > 5 ? 1500 : 1000;

    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'CSV Converter',
      slideLayout: options.slideLayout || 'table',
      includeHeaders: options.includeHeaders !== 'false',
      fileSize: file.buffer.length,
      fileSizeMB: fileSizeMB.toFixed(2),
      chunkSize
    });

    const { stdout, stderr } = await execFileAsync(pythonPath, [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'CSV Converter',
      '--slide-layout', options.slideLayout || 'table',
      '--chunk-size', chunkSize.toString()
    ].concat(options.includeHeaders === 'false' ? ['--no-headers'] : []), {
      timeout: 300000, // 5 minutes timeout for large files
      maxBuffer: 10 * 1024 * 1024 // 10MB buffer for large outputs
    });

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python ODP script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python ODP script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.odp`;
    const processingTime = Date.now() - startTime;
    console.log(`CSV->ODP conversion successful:`, {
      filename: downloadName,
      size: outputBuffer.length,
      processingTimeMs: processingTime,
      processingTimeSec: (processingTime / 1000).toFixed(2)
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/vnd.oasis.opendocument.presentation');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/vnd.oasis.opendocument.presentation'
    };
  } catch (error) {
    console.error(`CSV->ODP conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->ODP error`;
    throw new Error(`Failed to convert CSV to ODP: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// Initialize Express app
const app = express();
const PORT = parseInt(process.env.PORT || '3000', 10);

// Trust proxy - Trust only the first proxy (Traefik in Docker network)
// This is more secure than 'true' which would allow anyone to spoof IPs
app.set('trust proxy', 1);

// Rate limiting to prevent abuse
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100, // Limit each IP to 100 requests per 15 minutes
  message: 'Too many requests from this IP, please try again after 15 minutes',
  // Validate that the configuration is correct
  validate: { trustProxy: false }, // Disable validation since we're explicitly setting trust proxy
});

// Configure helmet with appropriate settings for large file uploads
app.use(helmet({
  contentSecurityPolicy: false, // Disable CSP for file uploads
  crossOriginEmbedderPolicy: false
}));
// CORS configuration is handled later in the file

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
  },
  fileFilter: (req, file, cb) => {
    console.log('Multer file filter:', {
      fieldname: file.fieldname,
      originalname: file.originalname,
      mimetype: file.mimetype,
      encoding: file.encoding
    });
    cb(null, true);
  }
});

const uploadSingle = upload.single('file');

const uploadBatchMulter = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024,
    files: 20
  }
});

const uploadBatch = uploadBatchMulter.array('files', 20);

// Handle preflight OPTIONS requests
app.options('*', (req, res) => {
  console.log('OPTIONS preflight request received');
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Credentials', 'true');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, Accept');
  res.status(200).end();
});

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Route: Batch Download
app.get('/batch-download/:filename', async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  const { filename } = req.params;
  console.log(`Batch download request for: ${filename}`);
  
  try {
    const filePath = path.join(BATCH_OUTPUT_DIR, filename);
    const metadata = batchFileMetadata.get(filename);
    
    if (!metadata) {
      console.error(`Metadata not found for: ${filename}`);
      return res.status(404).json({ error: 'File not found or expired' });
    }
    
    // Check if file exists
    const fileExists = await fs.access(filePath).then(() => true).catch(() => false);
    if (!fileExists) {
      console.error(`File not found on disk: ${filePath}`);
      batchFileMetadata.delete(filename);
      return res.status(404).json({ error: 'File not found or expired' });
    }
    
    // Read and send file
    const fileBuffer = await fs.readFile(filePath);
    
    res.set({
      'Content-Type': metadata.mime,
      'Content-Disposition': `attachment; filename="${metadata.downloadName}"`,
      'Content-Length': fileBuffer.length
    });
    
    res.send(fileBuffer);
    console.log(`Successfully sent file: ${metadata.downloadName}`);
  } catch (error) {
    console.error('Batch download error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: `Download failed: ${message}` });
  }
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

// Debug middleware for file uploads
app.use('/api/convert', (req, res, next) => {
  console.log('=== CONVERSION REQUEST DEBUG ===');
  console.log('Request method:', req.method);
  console.log('Request headers:', req.headers);
  console.log('Content-Type:', req.get('content-type'));
  console.log('Content-Length:', req.get('content-length'));
  console.log('Request body keys:', Object.keys(req.body || {}));
  next();
});

app.post('/api/convert', conversionTimeout(5 * 60 * 1000), upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    let file = req.file;
    const requestOptions = { ...(req.body as Record<string, string | undefined>) };

    console.log('=== CONVERSION REQUEST START ===');
    console.log('Request headers:', req.headers);
    console.log('Request body size:', req.get('content-length'));
    console.log('Request body options:', requestOptions);

    if (!file) {
      console.log('ERROR: No file uploaded');
      console.log('Available files in request:', req.files);
      console.log('Request body:', req.body);
      
      // Check if file was sent with a different field name
      if (req.files && Array.isArray(req.files) && req.files.length > 0) {
        console.log('Found files in array format, using first file');
        const firstFile = req.files[0];
        req.file = firstFile;
        file = firstFile;
      } else if (req.files && typeof req.files === 'object') {
        console.log('Found files in object format, checking for common field names');
        const commonFieldNames = ['file', 'upload', 'image', 'document', 'attachment'];
        for (const fieldName of commonFieldNames) {
          if ((req.files as any)[fieldName]) {
            console.log(`Found file with field name: ${fieldName}`);
            req.file = (req.files as any)[fieldName];
            file = (req.files as any)[fieldName];
            break;
          }
        }
      }
      
      if (!file) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No file uploaded' });
      }
    }

    // Validate file object
    if (!file.originalname) {
      console.log('ERROR: File missing originalname property');
      console.log('File object details:', {
        fieldname: file.fieldname,
        mimetype: file.mimetype,
        size: file.size,
        buffer: file.buffer ? `Buffer(${file.buffer.length})` : 'undefined'
      });
      
      // Try to create a fallback filename based on content type
      if (file.mimetype) {
        let ext = file.mimetype.split('/')[1] || 'bin';
        
        // Handle special cases
        if (file.mimetype === 'application/octet-stream') {
          // This could be a DNG or BMP file, try to detect from buffer
          if (file.buffer && file.buffer.length > 0) {
            // Check for BMP signature
            if (file.buffer[0] === 0x42 && file.buffer[1] === 0x4D) { // BM
              ext = 'bmp';
              console.log('Detected BMP file from signature');
            }
            // Check for DNG signature (starts with TIFF header)
            else if (file.buffer[0] === 0x49 && file.buffer[1] === 0x49) { // II (Intel byte order)
              ext = 'dng';
              console.log('Detected DNG file from signature');
            } else if (file.buffer[0] === 0x4D && file.buffer[1] === 0x4D) { // MM (Motorola byte order)
              ext = 'dng';
              console.log('Detected DNG file from signature');
            }
          }
        }
        
        file.originalname = `upload.${ext}`;
        console.log('Created fallback filename:', file.originalname);
      } else {
        // If no mimetype, try to detect from buffer
        if (file.buffer && file.buffer.length > 0) {
          if (file.buffer[0] === 0x42 && file.buffer[1] === 0x4D) { // BM
            file.originalname = 'upload.bmp';
            console.log('Detected BMP file from buffer signature');
          } else {
            file.originalname = 'upload.bin';
            console.log('Created generic fallback filename');
          }
        } else {
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          return res.status(400).json({ error: 'Invalid file upload - missing filename and content' });
        }
      }
    }

    if (!file.buffer || file.buffer.length === 0) {
      console.log('ERROR: File has no content');
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'Invalid file upload - empty file' });
    }

    console.log(`Processing file details:`, {
      originalname: file.originalname,
      size: file.size,
      mimetype: file.mimetype,
      fieldname: file.fieldname,
      encoding: file.encoding
    });

    const targetFormat = String(requestOptions.format ?? 'webp').toLowerCase();
    
    console.log('=== TARGET FORMAT AND BUFFER CHECK ===');
    console.log('Target format:', targetFormat);
    console.log('Has buffer:', !!file.buffer);
    console.log('Buffer length:', file.buffer ? file.buffer.length : 0);
    if (file.buffer && file.buffer.length > 10) {
      console.log('First 10 bytes:', Array.from(file.buffer.slice(0, 10)).map((b: number) => '0x' + b.toString(16).padStart(2, '0')).join(' '));
      console.log('First 2 bytes check - [0]:', file.buffer[0], 'expected: 0x42 (66)');
      console.log('First 2 bytes check - [1]:', file.buffer[1], 'expected: 0x4D (77)');
      console.log('Is BMP signature?:', file.buffer[0] === 0x42 && file.buffer[1] === 0x4D);
    }
    
  
  // EMERGENCY CHECK: If requesting WebP and file has DNG/TIFF signature, route immediately
  if (targetFormat === 'webp' && file.buffer && file.buffer.length > 2) {
    // Check for TIFF/DNG signature
    const isDngSignature = (file.buffer[0] === 0x49 && file.buffer[1] === 0x49) || // II (Intel)
                          (file.buffer[0] === 0x4D && file.buffer[1] === 0x4D);    // MM (Motorola)
    
    console.log('WebP target format - checking for DNG signature:', {
      isDngSignature,
      byte0: file.buffer[0],
      byte1: file.buffer[1],
      expectedII: [0x49, 0x49],
      expectedMM: [0x4D, 0x4D]
    });
    
    if (isDngSignature) {
      console.log('!!! EMERGENCY: DNG to WebP detected at entry point !!!');
      console.log('!!! Routing directly to Python script !!!');
      
      try {
        const result = await convertDngToWebpPython(file, requestOptions, true);
        
        res.set({
          'Content-Type': result.mime,
          'Content-Disposition': `attachment; filename="${result.filename}"`,
          'Content-Length': result.buffer.length.toString(),
          'Cache-Control': 'no-cache'
        });
        res.send(result.buffer);
        console.log('=== CONVERSION REQUEST END (SUCCESS via emergency DNG entry) ===');
        return;
      } catch (emergencyError) {
        console.error('!!! Emergency DNG to WebP conversion failed:', emergencyError);
        console.error('Error details:', {
          message: emergencyError instanceof Error ? emergencyError.message : String(emergencyError),
          stack: emergencyError instanceof Error ? emergencyError.stack : undefined
        });
        return res.status(500).json({ 
          error: 'DNG to WebP conversion failed', 
          details: emergencyError instanceof Error ? emergencyError.message : String(emergencyError)
        });
      }
    }
  }
  
    const isCSV = isCsvFile(file);
    const isEPUB = isEpubFile(file);
    const isEPS = isEpsFile(file);
    const isDNG = isDngFile(file);
    const isDOC = isDocFile(file);
    let isBMP = isBmpFile(file);
    
    console.log('File type detection:', {
      targetFormat,
      isCSV,
      isEPUB, 
      isEPS,
      isDNG,
      isDOC,
      isBMP,
      mimetype: file.mimetype,
      extension: file.originalname ? file.originalname.split('.').pop()?.toLowerCase() : 'undefined',
      originalname: file.originalname,
      fileSize: file.size,
      hasBuffer: !!file.buffer
    });

    // Additional check: if file has BMP signature but isBMP is false, force it
    if (!isBMP && file.buffer && file.buffer.length > 2) {
      if (file.buffer[0] === 0x42 && file.buffer[1] === 0x4D) { // BM signature
        console.log('!!! BMP signature detected in buffer, forcing isBMP to true !!!');
        isBMP = true;
      }
    }

    // CRITICAL CHECK: If target format is ICO and we have any indication of BMP, force routing to Python
    if (targetFormat === 'ico') {
      console.log('!!! ICO conversion requested - checking if this should be BMP to ICO !!!');
      
      // Check if it's a BMP file by any means
      if (!isBMP && file.buffer && file.buffer.length > 2) {
        if (file.buffer[0] === 0x42 && file.buffer[1] === 0x4D) {
          console.log('!!! FORCING BMP detection for ICO conversion !!!');
          isBMP = true;
          if (!file.originalname || !file.originalname.endsWith('.bmp')) {
            file.originalname = 'upload.bmp';
          }
        }
      }
      
      console.log('After ICO check - isBMP:', isBMP, 'originalname:', file.originalname);
    }
    
    console.log('Available CALIBRE_CONVERSIONS:', Object.keys(CALIBRE_CONVERSIONS));
    console.log('Target format in CALIBRE_CONVERSIONS?', !!CALIBRE_CONVERSIONS[targetFormat]);

    let result: ConversionResult | null = null;

    
    if (!result && (isDocFile(file) || isDocxFile(file) || isOdtFile(file)) && targetFormat === 'epub') {
      console.log('Single: Routing to LibreOffice (DOC/DOCX/ODT to EPUB conversion)');
      const inputHint = isDocxFile(file) ? 'docx' : isOdtFile(file) ? 'odt' : 'doc';
      result = await convertDocxWithLibreOffice(file, inputHint, requestOptions, true);
    }
    
    if (!result && isEPUB && CALIBRE_CONVERSIONS[targetFormat]) {
      console.log('Single: Routing to Calibre (EPUB conversion)');
      result = await convertWithCalibre(file, targetFormat, requestOptions, true);
    } else if (!result && isCSV && targetFormat === 'mobi') {
      console.log('Single: Routing to Simple CSV to MOBI conversion');
      result = await convertCsvToMobiSimple(file, requestOptions, true);
    } else if (!result && isCSV && targetFormat === 'odp') {
      console.log('Single: Routing to Python (CSV to ODP conversion)');
      result = await convertCsvToOdpPython(file, requestOptions, true);
    } else if (!result && isCSV && targetFormat === 'odt') {
      console.log('Single: Routing to Python (CSV to ODT conversion)');
      result = await convertCsvToOdtPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'pdf') {
        console.log('Single: Routing to Python (CSV to PDF conversion)');
        result = await convertCsvToPdfPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'ppt') {
        console.log('Single: Routing to Python (CSV to PPT conversion)');
        result = await convertCsvToPptPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'pptx') {
        console.log('Single: Routing to Python (CSV to PPTX conversion)');
        result = await convertCsvToPptxPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'rtf') {
        console.log('Single: Routing to Python (CSV to RTF conversion)');
        result = await convertCsvToRtfPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'txt') {
        console.log('Single: Routing to Python (CSV to TXT conversion)');
        result = await convertCsvToTxtPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'xls') {
        console.log('Single: Routing to Python (CSV to XLS conversion)');
        result = await convertCsvToXlsPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'xlsx') {
        console.log('Single: Routing to Python (CSV to XLSX conversion)');
        result = await convertCsvToXlsxPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'json') {
        console.log('Single: Routing to Python (CSV to JSON conversion)');
        result = await convertCsvToJsonPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'ndjson') {
        console.log('Single: Routing to Python (CSV to NDJSON conversion)');
        result = await convertCsvToNdjsonPython(file, requestOptions, true);
      } else if (!result && isCSV && targetFormat === 'docx') {
        console.log('Single: Routing to Python (CSV to DOCX conversion)');
        result = await convertCsvToDocxPython(file, requestOptions, true);
      } else if (!result && isCSV && ['epub', 'html'].includes(targetFormat)) {
      console.log(`Single: Routing to Python (CSV to ${targetFormat.toUpperCase()} conversion)`);
      result = await convertCsvToEbookPython(file, targetFormat, requestOptions, true);
    } else if (!result && isCSV && LIBREOFFICE_CONVERSIONS[targetFormat] && targetFormat !== 'doc') {
      console.log('Single: Routing to LibreOffice (CSV conversion)');
      result = await convertCsvWithLibreOffice(file, targetFormat, requestOptions, true);
    } else if (!result && (isDOC || file.originalname.toLowerCase().endsWith('.doc')) && targetFormat === 'csv') {
      console.log('Single: Routing to LibreOffice (DOC to CSV conversion)');
      result = await convertDocWithLibreOffice(file, 'doc-to-csv', requestOptions, true);
    } else if (!result && isEPS && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(targetFormat)) {
      console.log('Single: Routing to EPS conversion');
      result = await convertEpsFile(file, targetFormat, requestOptions, true);
    } else if (!result && isDNG && targetFormat === 'webp') {
      console.log('Single: Routing to Python (DNG to WebP conversion)');
      result = await convertDngToWebpPython(file, requestOptions, true);
    } else if (!result && isDNG && targetFormat === 'ico') {
      console.log('Single: Routing to Python (DNG to ICO conversion)');
      result = await convertDngToIcoPython(file, requestOptions, true);
    } else if (!result && isDNG && ['png', 'jpeg', 'jpg'].includes(targetFormat)) {
      console.log('Single: Routing to DNG conversion (legacy)');
      console.log('DNG conversion details:', {
        filename: file.originalname,
        targetFormat,
        supportedFormats: ['png', 'jpeg', 'jpg'],
        requestOptions
      });
      result = await convertDngFile(file, targetFormat, requestOptions, true);
    } else if (!result) {
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
          isDNG,
          isBMP
        });
        
        const quality = requestOptions.quality ?? 'high';
        const lossless = requestOptions.lossless ?? 'false';
        const width = requestOptions.width;
        const height = requestOptions.height;
        const iconSize = requestOptions.iconSize ?? '16';

    const inputBuffer = await prepareRawBuffer(file);

    const qualityValue = quality === 'high' ? 95 : quality === 'medium' ? 80 : 60;
    const isLossless = lossless === 'true';

    let pipeline;
    let metadata;
    
    try {
      pipeline = sharp(inputBuffer, {
      failOn: 'truncated',
      unlimited: true
    });

      metadata = await pipeline.metadata();
    console.log(`Metadata => ${metadata.width}x${metadata.height}, format: ${metadata.format}`);
    } catch (sharpError) {
      console.error('Sharp cannot process this file format:', sharpError);
      
      // Check if this is a DNG file that Sharp can't handle
      if (isDNG || (file.originalname && file.originalname.toLowerCase().endsWith('.dng'))) {
        throw new Error('DNG files cannot be processed directly by Sharp. Please use the Python conversion method.');
      }
      
      // Check if this is a BMP file that Sharp can't handle
      if (isBMP || (file.originalname && file.originalname.toLowerCase().endsWith('.bmp'))) {
        throw new Error('BMP files cannot be processed directly by Sharp. Please use ImageMagick or another conversion method.');
      }
      
      // For other unsupported formats, throw a generic error
      const filename = file.originalname || 'unknown';
      throw new Error(`Unsupported image format: ${filename}. Sharp cannot process this file type.`);
    }

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
    
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
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
    
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({
      error: 'Conversion failed',
      details: errorMessage
    });
  }
});

app.post('/api/convert/batch', conversionTimeout(10 * 60 * 1000), uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  const files = req.files as Express.Multer.File[] | undefined;
  const requestOptions = { ...(req.body as Record<string, string | undefined>) };
  const format = String(requestOptions.format ?? 'webp').toLowerCase();

  if (!files || files.length === 0) {
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    return res.status(400).json({
      success: false,
      processed: 0,
      results: [],
      error: 'No files uploaded'
    });
  }

  if (files.length > 20) {
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
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
      const isBMP = isBmpFile(file);
      console.log(`Processing ${file.originalname}: isCSV=${isCSV}, isEPUB=${isEPUB}, isEPS=${isEPS}, isDNG=${isDNG}, isDOC=${isDOC}, isBMP=${isBMP}, format=${format}, mimetype=${file.mimetype}`);

      if ((isDocFile(file) || isDocxFile(file) || isOdtFile(file)) && format === 'epub') {
        console.log('Batch: Routing to LibreOffice (DOC/DOCX/ODT to EPUB conversion)');
        const inputHint = isDocxFile(file) ? 'docx' : isOdtFile(file) ? 'odt' : 'doc';
        output = await convertDocxWithLibreOffice(file, inputHint, requestOptions, true);
      } else if (isEPUB && CALIBRE_CONVERSIONS[format]) {
        console.log('Routing to Calibre (EPUB conversion)');
        output = await convertWithCalibre(file, format, requestOptions, true);
      } else if (isCSV && format === 'mobi') {
        console.log('Batch: Routing to Simple CSV to MOBI conversion');
        output = await convertCsvToMobiSimple(file, requestOptions, true);
      } else if (isCSV && format === 'odp') {
        console.log('Batch: Routing to Python (CSV to ODP conversion)');
        output = await convertCsvToOdpPython(file, requestOptions, true);
      } else if (isCSV && format === 'odt') {
        console.log('Batch: Routing to Python (CSV to ODT conversion)');
        output = await convertCsvToOdtPython(file, requestOptions, true);
      } else if (isCSV && format === 'pdf') {
        console.log('Batch: Routing to Python (CSV to PDF conversion)');
        output = await convertCsvToPdfPython(file, requestOptions, true);
      } else if (isCSV && format === 'ppt') {
        console.log('Batch: Routing to Python (CSV to PPT conversion)');
        output = await convertCsvToPptPython(file, requestOptions, true);
      } else if (isCSV && format === 'pptx') {
        console.log('Batch: Routing to Python (CSV to PPTX conversion)');
        output = await convertCsvToPptxPython(file, requestOptions, true);
      } else if (isCSV && format === 'rtf') {
        console.log('Batch: Routing to Python (CSV to RTF conversion)');
        output = await convertCsvToRtfPython(file, requestOptions, true);
      } else if (isCSV && format === 'txt') {
        console.log('Batch: Routing to Python (CSV to TXT conversion)');
        output = await convertCsvToTxtPython(file, requestOptions, true);
      } else if (isCSV && format === 'xls') {
        console.log('Batch: Routing to Python (CSV to XLS conversion)');
        output = await convertCsvToXlsPython(file, requestOptions, true);
      } else if (isCSV && format === 'xlsx') {
        console.log('Batch: Routing to Python (CSV to XLSX conversion)');
        output = await convertCsvToXlsxPython(file, requestOptions, true);
      } else if (isCSV && format === 'json') {
        console.log('Batch: Routing to Python (CSV to JSON conversion)');
        output = await convertCsvToJsonPython(file, requestOptions, true);
      } else if (isCSV && format === 'ndjson') {
        console.log('Batch: Routing to Python (CSV to NDJSON conversion)');
        output = await convertCsvToNdjsonPython(file, requestOptions, true);
      } else if (isCSV && format === 'docx') {
        console.log('Batch: Routing to Python (CSV to DOCX conversion)');
        output = await convertCsvToDocxPython(file, requestOptions, true);
      } else if (isCSV && ['epub', 'html'].includes(format)) {
        console.log(`Batch: Routing to Python (CSV to ${format.toUpperCase()} conversion)`);
        output = await convertCsvToEbookPython(file, format, requestOptions, true);
      } else if (isCSV && LIBREOFFICE_CONVERSIONS[format] && format !== 'doc' && format !== 'docx') {
        console.log('Routing to LibreOffice (CSV conversion)');
        output = await convertCsvWithLibreOffice(file, format, requestOptions, true);
      } else if ((isDOC || file.originalname.toLowerCase().endsWith('.doc')) && format === 'csv') {
        console.log('Batch: Routing to LibreOffice (DOC to CSV conversion)');
        output = await convertDocWithLibreOffice(file, 'doc-to-csv', requestOptions, true);
      } else if (isEPS && ['webp', 'png', 'jpeg', 'jpg', 'ico'].includes(format)) {
        console.log('Routing to EPS conversion');
        output = await convertEpsFile(file, format, requestOptions, true);
      } else if (isDNG && format === 'webp') {
        console.log('Batch: Routing to Python (DNG to WebP conversion)');
        output = await convertDngToWebpPython(file, requestOptions, true);
      } else if (isDNG && format === 'ico') {
        console.log('Batch: Routing to Python (DNG to ICO conversion)');
        output = await convertDngToIcoPython(file, requestOptions, true);
      } else if (isDNG && ['png', 'jpeg', 'jpg'].includes(format)) {
        console.log('Batch: Routing to DNG conversion (legacy)');
        console.log('Batch DNG conversion details:', {
          filename: file.originalname,
          targetFormat: format,
          supportedFormats: ['png', 'jpeg', 'jpg'],
          requestOptions
        });
        output = await convertDngFile(file, format, requestOptions, true);
      } else {
        throw new Error(`Unsupported input file type or target format for batch conversion. File: ${file.originalname}, isCSV: ${isCSV}, isEPUB: ${isEPUB}, isEPS: ${isEPS}, isDNG: ${isDNG}, isDOC: ${isDOC}, isBMP: ${isBMP}, format: ${format}`);
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

  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
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
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(404).json({ error: 'File not found or expired' });
    }

    const filePath = path.join(BATCH_OUTPUT_DIR, storedFilename);
    const stat = await fs.stat(filePath).catch(() => null);
    if (!stat) {
      batchFileMetadata.delete(storedFilename);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
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
      'Cache-Control': 'no-cache',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });

    const stream = (await import('node:fs')).createReadStream(filePath);
    stream.on('error', (error) => {
      console.error('File stream error:', error);
      res.destroy(error);
    });
    stream.pipe(res);
  } catch (error) {
    console.error('Download error:', error);
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
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

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    timestamp: new Date().toISOString(),
    python: {
      available: true,
      path: '/opt/venv/bin/python3'
    },
    calibre: {
      available: true,
      path: 'ebook-convert'
    }
  });
});

// Test Python script endpoint
app.get('/test-python', async (req, res) => {
  try {
    const { execFile } = await import('child_process');
    const { promisify } = await import('util');
    const execFileAsync = promisify(execFile);
    const { stdout, stderr } = await execFileAsync('/opt/venv/bin/python3', ['--version']);
    res.json({ 
      success: true, 
      pythonVersion: stdout.trim(),
      stderr: stderr.trim()
    });
  } catch (error) {
    res.status(500).json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error)
    });
  }
});

// HEIC Preview endpoint - OPTIONS for CORS preflight
app.options('/api/preview/heic', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// DOCX Preview endpoint - OPTIONS for CORS preflight
app.options('/api/preview/docx', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// DOCX Preview endpoint - ensure CORS even on errors
app.post('/api/preview/docx', upload.single('file'), async (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('=== DOCX PREVIEW REQUEST ===');
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-docx-preview-'));
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.(docx)$/i, '.html'));
    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_html.py');
    const { execFile } = await import('child_process');
    const { promisify } = await import('util');
    const execFileAsync = promisify(execFile);

    let usedPandocFallback = false;
    try {
      await fs.access(scriptPath);
      // Fast preview flags (restore no-images for stability/speed)
      const args = [
        scriptPath,
        inputPath,
        outputPath,
        '--no-images',
        '--max-paragraphs', '200',
        '--max-chars', '200000',
        '--no-prettify'
      ];
      await execFileAsync('/opt/venv/bin/python', args);
    } catch {
      // Fallback to pandoc if script missing or fails access
      usedPandocFallback = true;
      try {
        await execFileAsync('pandoc', [inputPath, '-f', 'docx', '-t', 'html', '-o', outputPath]);
      } catch (e: any) {
        return res.status(500).json({ error: 'Preview generation failed', details: e?.stderr || e?.message || (usedPandocFallback ? 'Pandoc fallback failed' : '') });
      }
    }

    // Read and return HTML
    try {
      const html = await fs.readFile(outputPath);
      res.set({ 'Content-Type': 'text/html; charset=utf-8' });
      return res.send(html);
    } catch {
      return res.status(500).json({ error: 'Preview output not found' });
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    return res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// RTF Preview endpoint - OPTIONS for CORS preflight
app.options('/api/preview/rtf', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// RTF Preview endpoint - converts RTF to HTML for web viewing
app.post('/api/preview/rtf', upload.single('file'), async (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('=== RTF PREVIEW REQUEST ===');
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-rtf-preview-'));
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.(rtf)$/i, '.html'));
    await fs.writeFile(inputPath, file.buffer);

    // Use pandoc to convert RTF -> HTML for fast preview
    try {
      const { execFile } = await import('child_process');
      const { promisify } = await import('util');
      const execFileAsync = promisify(execFile);
      await execFileAsync('pandoc', [inputPath, '-f', 'rtf', '-t', 'html', '-o', outputPath]);
    } catch (e: any) {
      return res.status(500).json({ error: 'RTF preview conversion failed', details: e?.stderr || e?.message || '' });
    }

    try {
      const html = await fs.readFile(outputPath);
      res.set({ 'Content-Type': 'text/html; charset=utf-8' });
      return res.send(html);
    } catch {
      return res.status(500).json({ error: 'Preview output not found' });
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    return res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// HEIC Preview endpoint - converts HEIC to PNG for web viewing
app.post('/api/preview/heic', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('=== HEIC PREVIEW REQUEST ===');
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-heic-preview-'));

  try {
    const file = req.file;
    if (!file) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('HEIC preview request:', {
      filename: file.originalname,
      size: file.size,
      mimetype: file.mimetype
    });

    // Write HEIC file to temp directory
    const heicPath = path.join(tmpDir, `input.heic`);
    await fs.writeFile(heicPath, file.buffer);
    console.log('HEIC file written:', heicPath);

    // Prepare output PNG file
    const pngPath = path.join(tmpDir, `preview.png`);

    // Use Python script for HEIC to PNG conversion
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_preview.py');
    
    // Check if Python script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'HEIC preview script not found' });
    }

    const args = [
      scriptPath,
      heicPath,
      pngPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const pngExists = await fs.access(pngPath).then(() => true).catch(() => false);
    if (!pngExists) {
      throw new Error(`Python script did not produce PNG preview: ${pngPath}`);
    }

    // Read PNG file and send as response
    const pngBuffer = await fs.readFile(pngPath);
    if (!pngBuffer || pngBuffer.length === 0) {
      throw new Error('Python script produced empty PNG file');
    }

    console.log('HEIC preview successful:', {
      inputSize: file.size,
      outputSize: pngBuffer.length
    });

    // Send PNG as response
    res.set({
      'Content-Type': 'image/png',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.send(pngBuffer);

  } catch (error) {
    console.error('HEIC preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown HEIC preview error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: `Failed to generate HEIC preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: HEIC to PNG (Single) - OPTIONS for CORS preflight
app.options('/convert/heic-to-png/single', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to PNG (Single)
app.post('/convert/heic-to-png/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('HEIC->PNG single conversion request');

  const tmpDir = path.join(os.tmpdir(), `heic-png-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No file uploaded' });
    }

    try {
      await fs.mkdir(tmpDir, { recursive: true });
    } catch (mkdirError) {
      console.error('HEIC to PNG: Failed to create temp directory:', mkdirError);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(500).json({ error: 'Failed to create temporary directory', details: mkdirError instanceof Error ? mkdirError.message : 'Unknown error' });
    }

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.png'));

    try {
      await fs.writeFile(inputPath, file.buffer);
    } catch (writeError) {
      console.error('HEIC to PNG: Failed to write input file:', writeError);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      return res.status(500).json({ error: 'Failed to write input file', details: writeError instanceof Error ? writeError.message : 'Unknown error' });
    }

    const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_png.py');
    console.log('HEIC to PNG: Executing Python script:', scriptPath);
    console.log('HEIC to PNG: Input file:', inputPath);
    console.log('HEIC to PNG: Output file:', outputPath);

    try {
      await fs.access(scriptPath);
      console.log('HEIC to PNG: Script exists');
    } catch (error) {
      console.error('HEIC to PNG: Script does not exist:', scriptPath);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Optional quality param (0-100) for PNG compression level mapping
    const quality = parseInt(req.body.quality) || 95;
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', String(quality), '--max-dimension', String(maxDimension)];

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.on('error', async (error: Error) => {
      console.error('HEIC to PNG: Failed to start Python process:', error);
      if (!res.headersSent) {
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Failed to start conversion process', details: error.message });
      }
      await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
    });

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('HEIC to PNG stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('HEIC to PNG stderr:', data.toString());
    });

    const timeout = setTimeout(async () => {
      console.error('HEIC to PNG: Conversion timeout after 5 minutes');
      python.kill();
      if (!res.headersSent) {
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion timeout. The file may be too large or complex.' });
      }
      await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
    }, 5 * 60 * 1000);

    python.on('close', async (code: number) => {
      clearTimeout(timeout);
      console.log('HEIC to PNG: Python script finished with code:', code);
      console.log('HEIC to PNG: stdout:', stdout);
      console.log('HEIC to PNG: stderr:', stderr);

      try {
        if (!res.headersSent) {
          if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
            const outputBuffer = await fs.readFile(outputPath);
            console.log('HEIC to PNG: Output file size:', outputBuffer.length);
            res.set({
              'Content-Type': 'image/png',
              'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
            });
            res.send(outputBuffer);
          } else {
            console.error('HEIC to PNG conversion failed. Code:', code, 'Stderr:', stderr);
            res.set({
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
            });
            res.status(500).json({ error: 'Conversion failed', details: stderr });
          }
        }
      } catch (error) {
        console.error('Error handling conversion result (PNG):', error);
        if (!res.headersSent) {
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
        }
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('HEIC to PNG conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: HEIC to PNG (Batch) - OPTIONS for CORS preflight
app.options('/convert/heic-to-png/batch', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to PNG (Batch)
app.post('/convert/heic-to-png/batch', uploadBatch, async (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('HEIC->PNG batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `heic-png-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No files uploaded' });
    }

    await fs.mkdir(tmpDir, { recursive: true });

    const results: any[] = [];

    const quality = parseInt(req.body.quality) || 95;
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.png'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_png.py');
        try {
          await fs.access(scriptPath);
        } catch {
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', String(quality), '--max-dimension', String(maxDimension)];
        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (d: Buffer) => { stdout += d.toString(); });
        python.stderr.on('data', (d: Buffer) => { stderr += d.toString(); });

        await new Promise<void>((resolve) => {
          python.on('close', async (code: number) => {
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:image/png;base64,${outputBuffer.toString('base64')}`
                });
              } else {
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
              }
            } catch (err) {
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: err instanceof Error ? err.message : 'Unknown error'
              });
            } finally {
              resolve();
            }
          });
        });
      } catch (error) {
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.json({ success: true, results });

  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: HEIC to EPS (Single) - OPTIONS for CORS preflight
app.options('/convert/heic-to-eps/single', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to EPS (Single)
app.post('/convert/heic-to-eps/single', upload.single('file'), async (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('HEIC->EPS single conversion request');
  const tmpDir = path.join(os.tmpdir(), `heic-eps-${Date.now()}`);
  try {
    const file = req.file;
    if (!file) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });
    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.eps'));
    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_eps.py');
    try { await fs.access(scriptPath); } catch { return res.status(500).json({ error: 'Conversion script not found' }); }

    const maxDimension = parseInt(req.body.maxDimension) || 4096;
    const pythonArgs = [scriptPath, inputPath, outputPath, '--max-dimension', String(maxDimension)];
    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '', stderr = '';
    python.on('error', async (err: Error) => {
      console.error('HEIC->WEBP single: failed to start python:', err);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      if (!res.headersSent) res.status(500).json({ error: 'Failed to start conversion process', details: err.message });
      await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
    });
    python.stdout.on('data', (d: Buffer) => { stdout += d.toString(); });
    python.stderr.on('data', (d: Buffer) => { stderr += d.toString(); });
    python.on('close', async (code: number) => {
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          res.set({
            'Content-Type': 'application/postscript',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
        } else {
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: HEIC to EPS (Batch) - OPTIONS for CORS preflight
app.options('/convert/heic-to-eps/batch', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to EPS (Batch)
app.post('/convert/heic-to-eps/batch', uploadBatch, async (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('HEIC->EPS batch conversion request');
  const tmpDir = path.join(os.tmpdir(), `heic-eps-batch-${Date.now()}`);
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });
    const results: any[] = [];
    const maxDimension = parseInt(req.body.maxDimension) || 4096;
    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.eps'));
        await fs.writeFile(inputPath, file.buffer);
        const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_eps.py');
        try { await fs.access(scriptPath); } catch { results.push({ originalName: file.originalname, outputFilename: '', size: 0, success: false, error: 'Conversion script not found' }); continue; }
        const pythonArgs = [scriptPath, inputPath, outputPath, '--max-dimension', String(maxDimension)];
        const python = spawn('/opt/venv/bin/python', pythonArgs);
        let stdout = '', stderr = '';
        python.on('error', (err: Error) => {
          stderr += ` spawn error: ${err.message}`;
        });
        python.stdout.on('data', (d: Buffer) => { stdout += d.toString(); });
        python.stderr.on('data', (d: Buffer) => { stderr += d.toString(); });
        await new Promise<void>((resolve) => {
          python.on('close', async (code: number) => {
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                results.push({ originalName: file.originalname, outputFilename: path.basename(outputPath), size: outputBuffer.length, success: true, downloadPath: `data:application/postscript;base64,${outputBuffer.toString('base64')}` });
              } else {
                results.push({ originalName: file.originalname, outputFilename: '', size: 0, success: false, error: stderr || `Conversion failed with code ${code}` });
              }
            } finally { resolve(); }
          });
        });
      } catch (err) {
        results.push({ originalName: file.originalname, outputFilename: '', size: 0, success: false, error: err instanceof Error ? err.message : 'Unknown error' });
      }
    }
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.json({ success: true, results });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// TIFF Preview endpoint - converts TIFF to PNG for web viewing
app.post('/api/preview/tiff', upload.single('file'), async (req, res) => {
  console.log('=== TIFF PREVIEW REQUEST ===');
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-tiff-preview-'));

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('TIFF preview request:', {
      filename: file.originalname,
      size: file.size,
      mimetype: file.mimetype
    });

    // Write TIFF file to temp directory
    const tiffPath = path.join(tmpDir, `input.tiff`);
    await fs.writeFile(tiffPath, file.buffer);
    console.log('TIFF file written:', tiffPath);

    // Prepare output PNG file
    const pngPath = path.join(tmpDir, `preview.png`);

    // Use Python script for TIFF to PNG conversion
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'tiff_preview.py');
    
    // Check if Python script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'TIFF preview script not found' });
    }

    const args = [
      scriptPath,
      tiffPath,
      pngPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const pngExists = await fs.access(pngPath).then(() => true).catch(() => false);
    if (!pngExists) {
      throw new Error(`Python script did not produce PNG preview: ${pngPath}`);
    }

    // Read PNG file and send as response
    const pngBuffer = await fs.readFile(pngPath);
    if (!pngBuffer || pngBuffer.length === 0) {
      throw new Error('Python script produced empty PNG file');
    }

    console.log('TIFF preview successful:', {
      inputSize: file.size,
      outputSize: pngBuffer.length
    });

    // Send PNG as response
    res.set('Content-Type', 'image/png');
    res.send(pngBuffer);

  } catch (error) {
    console.error('TIFF preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown TIFF preview error';
    res.status(500).json({ error: `Failed to generate TIFF preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// DOCX Preview endpoint - convert DOCX to HTML for web viewing
app.post('/api/preview/docx', uploadDocument.single('file'), async (req, res) => {
  console.log('=== DOCX PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `docx-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('DOCX file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save DOCX file to temp location
    const docxPath = path.join(tmpDir, 'input.docx');
    await fs.writeFile(docxPath, file.buffer);

    // Use Python script with LibreOffice to convert DOCX to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'docx_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'DOCX preview script not found' });
    }

    const args = [
      scriptPath,
      docxPath,
      htmlPath
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read HTML file
    const html = await fs.readFile(htmlPath, 'utf-8');

    console.log('DOCX preview successful:', {
      inputSize: file.size,
      outputLength: html.length
    });

    // Send HTML wrapped in a styled document with A4 page layout
    const styledHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${file.originalname}</title>
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }
          body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #525659;
            color: #333;
            line-height: 1.6;
            padding: 20px 0;
          }
          .toolbar {
            position: sticky;
            top: 0;
            background: #2c3e50;
            color: white;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            z-index: 1000;
          }
          .toolbar button {
            background: #3498db;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 10px;
            font-size: 14px;
          }
          .toolbar button:hover {
            background: #2980b9;
          }
          .page-container {
            max-width: 210mm;
            margin: 20px auto;
            padding: 0 10px;
          }
          .page {
            width: 210mm;
            min-height: 297mm;
            padding: 20mm;
            margin-bottom: 20px;
            background: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.3);
            page-break-after: always;
          }
          .page:last-child {
            margin-bottom: 0;
          }
          h1, h2, h3, h4, h5, h6 {
            margin-top: 1.2em;
            margin-bottom: 0.6em;
            color: #2c3e50;
            page-break-after: avoid;
          }
          h1 { font-size: 2em; }
          h2 { font-size: 1.5em; }
          h3 { font-size: 1.17em; }
          p { 
            margin-bottom: 1em;
            text-align: justify;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
            page-break-inside: avoid;
          }
          table, th, td {
            border: 1px solid #ddd;
          }
          th, td {
            padding: 8px 12px;
            text-align: left;
          }
          th {
            background-color: #f8f9fa;
            font-weight: bold;
          }
          img {
            max-width: 100%;
            height: auto;
            page-break-inside: avoid;
          }
          ul, ol {
            margin-left: 2em;
            margin-bottom: 1em;
          }
          li {
            margin-bottom: 0.5em;
          }
          @media print {
            body {
              background: white;
              padding: 0;
            }
            .toolbar {
              display: none;
            }
            .page-container {
              max-width: 100%;
              margin: 0;
              padding: 0;
            }
            .page {
              margin: 0;
              box-shadow: none;
              page-break-after: always;
            }
          }
        </style>
      </head>
      <body>
        <div class="toolbar">
          <span><strong>📄 ${file.originalname}</strong></span>
          <div>
            <button onclick="window.print()">🖨️ Print</button>
            <button onclick="window.close()">✖️ Close</button>
          </div>
        </div>
        <div class="page-container">
          <div class="page">
            ${html}
          </div>
        </div>
      </body>
      </html>
    `;

    res.set('Content-Type', 'text/html');
    res.send(styledHtml);

  } catch (error) {
    console.error('DOCX preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown DOCX preview error';
    res.status(500).json({ error: `Failed to generate DOCX preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// RTF Preview endpoint - convert RTF to HTML for web viewing
app.post('/api/preview/rtf', uploadDocument.single('file'), async (req, res) => {
  console.log('=== RTF PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `rtf-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('RTF file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save RTF file to temp location
    const rtfPath = path.join(tmpDir, 'input.rtf');
    await fs.writeFile(rtfPath, file.buffer);

    // Use Python script with LibreOffice/Pandoc to convert RTF to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'rtf_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'RTF preview script not found' });
    }

    const args = [
      scriptPath,
      rtfPath,
      htmlPath
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read HTML file
    let htmlContent = await fs.readFile(htmlPath, 'utf-8');
    
    // Wrap HTML in styled template with A4 page layout
    const styledHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${file.originalname}</title>
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }
          body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #525659;
            color: #333;
            line-height: 1.6;
            padding: 20px 0;
          }
          .toolbar {
            position: sticky;
            top: 0;
            background: #d35400;
            color: white;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            z-index: 1000;
          }
          .toolbar button {
            background: #e67e22;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 10px;
            font-size: 14px;
          }
          .toolbar button:hover {
            background: #f39c12;
          }
          .page-container {
            max-width: 210mm;
            margin: 20px auto;
            padding: 0 10px;
          }
          .page {
            width: 210mm;
            min-height: 297mm;
            padding: 20mm;
            margin-bottom: 20px;
            background: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.3);
            page-break-after: always;
          }
          .page:last-child {
            margin-bottom: 0;
          }
          h1, h2, h3, h4, h5, h6 {
            margin-top: 1.2em;
            margin-bottom: 0.6em;
            color: #2c3e50;
            page-break-after: avoid;
          }
          h1 { font-size: 2em; }
          h2 { font-size: 1.5em; }
          h3 { font-size: 1.17em; }
          p { 
            margin-bottom: 1em;
            text-align: justify;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
            page-break-inside: avoid;
          }
          table, th, td {
            border: 1px solid #ddd;
          }
          th, td {
            padding: 8px 12px;
            text-align: left;
          }
          th {
            background-color: #f8f9fa;
            font-weight: bold;
          }
          img {
            max-width: 100%;
            height: auto;
            page-break-inside: avoid;
          }
          ul, ol {
            margin-left: 2em;
            margin-bottom: 1em;
          }
          li {
            margin-bottom: 0.5em;
          }
          @media print {
            body {
              background: white;
              padding: 0;
            }
            .toolbar {
              display: none;
            }
            .page-container {
              max-width: 100%;
              margin: 0;
              padding: 0;
            }
            .page {
              margin: 0;
              box-shadow: none;
              page-break-after: always;
            }
          }
        </style>
      </head>
      <body>
        <div class="toolbar">
          <span><strong>📄 ${file.originalname}</strong></span>
          <div>
            <button onclick="window.print()">🖨️ Print</button>
            <button onclick="window.close()">✖️ Close</button>
          </div>
        </div>
        <div class="page-container">
          <div class="page">
            ${htmlContent}
          </div>
        </div>
      </body>
      </html>
    `;

    console.log('RTF preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(styledHtml);

  } catch (error) {
    console.error('RTF preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown RTF preview error';
    res.status(500).json({ error: `Failed to generate RTF preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// ODT Preview endpoint - convert ODT to HTML for web viewing
app.post('/api/preview/odt', uploadDocument.single('file'), async (req, res) => {
  console.log('=== ODT PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `odt-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('ODT file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save ODT file to temp location
    const odtPath = path.join(tmpDir, 'input.odt');
    await fs.writeFile(odtPath, file.buffer);

    // Use Python script with LibreOffice to convert ODT to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'odt_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'ODT preview script not found' });
    }

    const args = [
      scriptPath,
      odtPath,
      htmlPath
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read HTML file
    let htmlContent = await fs.readFile(htmlPath, 'utf-8');
    
    // Wrap HTML in styled template with A4 page layout
    const styledHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${file.originalname}</title>
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }
          body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #525659;
            color: #333;
            line-height: 1.6;
            padding: 20px 0;
          }
          .toolbar {
            position: sticky;
            top: 0;
            background: #f59e0b;
            color: white;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            z-index: 1000;
          }
          .toolbar button {
            background: #d97706;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 10px;
            font-size: 14px;
          }
          .toolbar button:hover {
            background: #b45309;
          }
          .page-container {
            max-width: 210mm;
            margin: 20px auto;
            padding: 0 10px;
          }
          .page {
            width: 210mm;
            min-height: 297mm;
            padding: 20mm;
            margin-bottom: 20px;
            background: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.3);
            page-break-after: always;
          }
          .page:last-child {
            margin-bottom: 0;
          }
          h1, h2, h3, h4, h5, h6 {
            margin-top: 1.2em;
            margin-bottom: 0.6em;
            color: #2c3e50;
            page-break-after: avoid;
          }
          h1 { font-size: 2em; }
          h2 { font-size: 1.5em; }
          h3 { font-size: 1.17em; }
          p { 
            margin-bottom: 1em;
            text-align: justify;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
            page-break-inside: avoid;
          }
          table, th, td {
            border: 1px solid #ddd;
          }
          th, td {
            padding: 8px 12px;
            text-align: left;
          }
          th {
            background-color: #f8f9fa;
            font-weight: bold;
          }
          img {
            max-width: 100%;
            height: auto;
            page-break-inside: avoid;
          }
          ul, ol {
            margin-left: 2em;
            margin-bottom: 1em;
          }
          li {
            margin-bottom: 0.5em;
          }
          @media print {
            body {
              background: white;
              padding: 0;
            }
            .toolbar {
              display: none;
            }
            .page-container {
              max-width: 100%;
              margin: 0;
              padding: 0;
            }
            .page {
              margin: 0;
              box-shadow: none;
              page-break-after: always;
            }
          }
        </style>
      </head>
      <body>
        <div class="toolbar">
          <span><strong>📄 ${file.originalname}</strong></span>
          <div>
            <button onclick="window.print()">🖨️ Print</button>
            <button onclick="window.close()">✖️ Close</button>
          </div>
        </div>
        <div class="page-container">
          <div class="page">
            ${htmlContent}
          </div>
        </div>
      </body>
      </html>
    `;

    console.log('ODT preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(styledHtml);

  } catch (error) {
    console.error('ODT preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown ODT preview error';
    res.status(500).json({ error: `Failed to generate ODT preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// PDF Preview endpoint - convert PDF to HTML viewer with PDF.js
app.options('/api/preview/pdf', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.status(204).send();
});

app.post('/api/preview/pdf', uploadDocument.single('file'), async (req, res) => {
  // Set CORS headers immediately
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('=== PDF PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `pdf-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('PDF file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save PDF file to temp location
    const pdfPath = path.join(tmpDir, 'input.pdf');
    await fs.writeFile(pdfPath, file.buffer);

    // Use Python script to create HTML viewer
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'pdf_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`PDF script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'PDF preview script not found' });
    }

    const args = [
      scriptPath,
      pdfPath,
      outputPath
    ];

    console.log('Executing PDF script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('PDF script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'PDF execution failed';
    }

    if (stdout.trim().length > 0) console.log('PDF stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('PDF stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`PDF script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`PDF script did not produce preview: ${outputPath}`);
    }

    // Read and send HTML file
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('PDF preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set({
      'Content-Type': 'text/html',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.send(htmlContent);

  } catch (error) {
    console.error('PDF preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown PDF preview error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: `Failed to generate PDF preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// TXT Preview endpoint - convert TXT to HTML for web viewing
app.post('/api/preview/txt', uploadDocument.single('file'), async (req, res) => {
  console.log('=== TXT PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `txt-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('TXT file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save TXT file to temp location
    const txtPath = path.join(tmpDir, 'input.txt');
    await fs.writeFile(txtPath, file.buffer);

    // Use Python script to convert TXT to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'txt_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'TXT preview script not found' });
    }

    const args = [
      scriptPath,
      txtPath,
      htmlPath,
      '--max-lines', '50000'  // Support up to 50k lines
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read HTML file
    let htmlContent = await fs.readFile(htmlPath, 'utf-8');
    
    // Wrap HTML in styled template
    const styledHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${file.originalname}</title>
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }
          body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #1e1e1e;
            color: #d4d4d4;
            line-height: 1.6;
          }
          .toolbar {
            position: sticky;
            top: 0;
            background: #2d2d30;
            color: white;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.3);
            z-index: 1000;
            border-bottom: 1px solid #3e3e42;
          }
          .toolbar-info {
            display: flex;
            align-items: center;
            gap: 15px;
          }
          .file-info {
            display: flex;
            flex-direction: column;
          }
          .file-name {
            font-weight: bold;
            font-size: 14px;
          }
          .file-meta {
            font-size: 12px;
            color: #858585;
          }
          .toolbar button {
            background: #0e639c;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 10px;
            font-size: 14px;
          }
          .toolbar button:hover {
            background: #1177bb;
          }
          .content-container {
            padding: 20px;
            max-width: 1400px;
            margin: 0 auto;
          }
          .text-container {
            background: #1e1e1e;
            border: 1px solid #3e3e42;
            border-radius: 8px;
            padding: 20px;
            overflow-x: auto;
          }
          .lines-container {
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 14px;
            line-height: 1.6;
          }
          .line {
            display: flex;
          }
          .line-number {
            padding-right: 20px;
            margin-right: 20px;
            border-right: 1px solid #3e3e42;
            color: #858585;
            user-select: none;
            text-align: right;
            min-width: 70px;
            flex-shrink: 0;
          }
          .line-content {
            white-space: pre-wrap;
            word-wrap: break-word;
            flex: 1;
          }
          .truncated-warning {
            background: #f59e0b;
            color: white;
            padding: 12px 20px;
            margin: 20px;
            border-radius: 8px;
            text-align: center;
            font-weight: bold;
          }
          @media print {
            body {
              background: white;
              color: black;
            }
            .toolbar {
              display: none;
            }
            .text-container {
              border: none;
              background: white;
            }
            .line-number {
              color: #666;
            }
            .line-content {
              color: black;
            }
          }
        </style>
      </head>
      <body>
        <div class="toolbar">
          <div class="toolbar-info">
            <span>📄</span>
            <div class="file-info">
              <span class="file-name">${file.originalname}</span>
              <span class="file-meta">${(file.size / 1024).toFixed(2)} KB</span>
            </div>
          </div>
          <div>
            <button onclick="window.print()">🖨️ Print</button>
            <button onclick="window.close()">✖️ Close</button>
          </div>
        </div>
        <div class="content-container">
          <div class="text-container">
            ${htmlContent}
          </div>
        </div>
      </body>
      </html>
    `;

    console.log('TXT preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(styledHtml);

  } catch (error) {
    console.error('TXT preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown TXT preview error';
    res.status(500).json({ error: `Failed to generate TXT preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Markdown Preview endpoint - convert Markdown to HTML for web viewing
app.post('/api/preview/md', uploadDocument.single('file'), async (req, res) => {
  console.log('=== MARKDOWN PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `md-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('Markdown file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save Markdown file to temp location
    const mdPath = path.join(tmpDir, 'input.md');
    await fs.writeFile(mdPath, file.buffer);

    // Use Python script to convert Markdown to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'md_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'Markdown preview script not found' });
    }

    const args = [
      scriptPath,
      mdPath,
      htmlPath
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read HTML file
    let htmlContent = await fs.readFile(htmlPath, 'utf-8');
    
    // Wrap HTML in styled template with GitHub styling
    const styledHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>${file.originalname}</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/github-markdown-css/5.2.0/github-markdown.min.css">
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }
          body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
            background: #f6f8fa;
            color: #24292f;
            line-height: 1.6;
          }
          .toolbar {
            position: sticky;
            top: 0;
            background: #24292f;
            color: white;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            z-index: 1000;
          }
          .toolbar-info {
            display: flex;
            align-items: center;
            gap: 15px;
          }
          .file-info {
            display: flex;
            flex-direction: column;
          }
          .file-name {
            font-weight: bold;
            font-size: 14px;
          }
          .file-meta {
            font-size: 12px;
            color: #8b949e;
          }
          .toolbar button {
            background: #238636;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            margin-left: 10px;
            font-size: 14px;
          }
          .toolbar button:hover {
            background: #2ea043;
          }
          .content-container {
            max-width: 980px;
            margin: 40px auto;
            padding: 0 20px;
          }
          .markdown-body {
            background: white;
            border: 1px solid #d0d7de;
            border-radius: 8px;
            padding: 48px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          @media print {
            .toolbar {
              display: none;
            }
            body {
              background: white;
            }
            .content-container {
              max-width: 100%;
              margin: 0;
              padding: 0;
            }
            .markdown-body {
              border: none;
              box-shadow: none;
            }
          }
        </style>
      </head>
      <body>
        <div class="toolbar">
          <div class="toolbar-info">
            <span>📝</span>
            <div class="file-info">
              <span class="file-name">${file.originalname}</span>
              <span class="file-meta">${(file.size / 1024).toFixed(2)} KB • Markdown</span>
            </div>
          </div>
          <div>
            <button onclick="window.print()">🖨️ Print</button>
            <button onclick="window.close()">✖️ Close</button>
          </div>
        </div>
        <div class="content-container">
          <div class="markdown-body">
            ${htmlContent}
          </div>
        </div>
      </body>
      </html>
    `;

    console.log('Markdown preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(styledHtml);

  } catch (error) {
    console.error('Markdown preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown Markdown preview error';
    res.status(500).json({ error: `Failed to generate Markdown preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// RAF Preview endpoint - convert RAF (Fujifilm RAW) to web-viewable image
app.post('/api/preview/raf', uploadDocument.single('file'), async (req, res) => {
  console.log('=== RAF PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `raf-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('RAF file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save RAF file to temp location
    const rafPath = path.join(tmpDir, 'input.raf');
    await fs.writeFile(rafPath, file.buffer);

    // Use Python script to convert RAF to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'raf_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`RAF script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'RAF preview script not found' });
    }

    const args = [
      scriptPath,
      rafPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing RAF script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('RAF script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'RAF execution failed';
    }

    if (stdout.trim().length > 0) console.log('RAF stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('RAF stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`RAF script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`RAF script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('RAF preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('RAF preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown RAF preview error';
    res.status(500).json({ error: `Failed to generate RAF preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// ORF Preview endpoint - convert ORF (Olympus RAW) to web-viewable image
app.post('/api/preview/orf', uploadDocument.single('file'), async (req, res) => {
  console.log('=== ORF PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `orf-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('ORF file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save ORF file to temp location
    const orfPath = path.join(tmpDir, 'input.orf');
    await fs.writeFile(orfPath, file.buffer);

    // Use Python script to convert ORF to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'orf_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`ORF script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'ORF preview script not found' });
    }

    const args = [
      scriptPath,
      orfPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing ORF script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('ORF script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'ORF execution failed';
    }

    if (stdout.trim().length > 0) console.log('ORF stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('ORF stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`ORF script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`ORF script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('ORF preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('ORF preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown ORF preview error';
    res.status(500).json({ error: `Failed to generate ORF preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// DNG Preview endpoint - convert DNG (Adobe Digital Negative) to web-viewable image
app.post('/api/preview/dng', uploadDocument.single('file'), async (req, res) => {
  console.log('=== DNG PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `dng-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('DNG file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save DNG file to temp location
    const dngPath = path.join(tmpDir, 'input.dng');
    await fs.writeFile(dngPath, file.buffer);

    // Use Python script to convert DNG to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'dng_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`DNG script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'DNG preview script not found' });
    }

    const args = [
      scriptPath,
      dngPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing DNG script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('DNG script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'DNG execution failed';
    }

    if (stdout.trim().length > 0) console.log('DNG stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('DNG stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`DNG script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`DNG script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('DNG preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('DNG preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown DNG preview error';
    res.status(500).json({ error: `Failed to generate DNG preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// ARW Preview endpoint - convert ARW (Sony RAW) to web-viewable image
app.post('/api/preview/arw', uploadDocument.single('file'), async (req, res) => {
  console.log('=== ARW PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `arw-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('ARW file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save ARW file to temp location
    const arwPath = path.join(tmpDir, 'input.arw');
    await fs.writeFile(arwPath, file.buffer);

    // Use Python script to convert ARW to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'arw_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`ARW script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'ARW preview script not found' });
    }

    const args = [
      scriptPath,
      arwPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing ARW script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('ARW script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'ARW execution failed';
    }

    if (stdout.trim().length > 0) console.log('ARW stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('ARW stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`ARW script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`ARW script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('ARW preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('ARW preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown ARW preview error';
    res.status(500).json({ error: `Failed to generate ARW preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// X3F Preview endpoint - convert X3F (Sigma RAW) to web-viewable image
app.post('/api/preview/x3f', uploadDocument.single('file'), async (req, res) => {
  console.log('=== X3F PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `x3f-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('X3F file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save X3F file to temp location
    const x3fPath = path.join(tmpDir, 'input.x3f');
    await fs.writeFile(x3fPath, file.buffer);

    // Use Python script to convert X3F to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'x3f_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`X3F script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'X3F preview script not found' });
    }

    const args = [
      scriptPath,
      x3fPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing X3F script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('X3F script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'X3F execution failed';
    }

    if (stdout.trim().length > 0) console.log('X3F stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('X3F stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`X3F script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`X3F script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('X3F preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('X3F preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown X3F preview error';
    res.status(500).json({ error: `Failed to generate X3F preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// DCR Preview endpoint - convert DCR (Kodak RAW) to web-viewable image
app.post('/api/preview/dcr', uploadDocument.single('file'), async (req, res) => {
  console.log('=== DCR PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `dcr-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('DCR file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save DCR file to temp location
    const dcrPath = path.join(tmpDir, 'input.dcr');
    await fs.writeFile(dcrPath, file.buffer);

    // Use Python script to convert DCR to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'dcr_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`DCR script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'DCR preview script not found' });
    }

    const args = [
      scriptPath,
      dcrPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing DCR script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('DCR script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'DCR execution failed';
    }

    if (stdout.trim().length > 0) console.log('DCR stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('DCR stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`DCR script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`DCR script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('DCR preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('DCR preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown DCR preview error';
    res.status(500).json({ error: `Failed to generate DCR preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// CR2 Preview endpoint - convert CR2 (Canon RAW) to web-viewable image
app.post('/api/preview/cr2', uploadDocument.single('file'), async (req, res) => {
  console.log('=== CR2 PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `cr2-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('CR2 file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save CR2 file to temp location
    const cr2Path = path.join(tmpDir, 'input.cr2');
    await fs.writeFile(cr2Path, file.buffer);

    // Use Python script to convert CR2 to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'cr2_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`CR2 script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'CR2 preview script not found' });
    }

    const args = [
      scriptPath,
      cr2Path,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing CR2 script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('CR2 script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'CR2 execution failed';
    }

    if (stdout.trim().length > 0) console.log('CR2 stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('CR2 stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`CR2 script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`CR2 script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('CR2 preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('CR2 preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown CR2 preview error';
    res.status(500).json({ error: `Failed to generate CR2 preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// NEF Preview endpoint - convert NEF (Nikon RAW) to web-viewable image
app.post('/api/preview/nef', uploadDocument.single('file'), async (req, res) => {
  console.log('=== NEF PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `nef-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('NEF file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save NEF file to temp location
    const nefPath = path.join(tmpDir, 'input.nef');
    await fs.writeFile(nefPath, file.buffer);

    // Use Python script to convert NEF to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'nef_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`NEF script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'NEF preview script not found' });
    }

    const args = [
      scriptPath,
      nefPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing NEF script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('NEF script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'NEF execution failed';
    }

    if (stdout.trim().length > 0) console.log('NEF stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('NEF stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`NEF script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`NEF script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('NEF preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('NEF preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown NEF preview error';
    res.status(500).json({ error: `Failed to generate NEF preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Python Preview endpoint - format Python for web viewing
app.post('/api/preview/python', uploadDocument.single('file'), async (req, res) => {
  console.log('=== PYTHON PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `python-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('Python file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save Python file to temp location
    const pyPath = path.join(tmpDir, 'input.py');
    await fs.writeFile(pyPath, file.buffer);

    // Use Python script to format Python code
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'python_to_formatted.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'Python preview script not found' });
    }

    const args = [
      scriptPath,
      pyPath,
      outputPath,
      '--max-size-mb', '10'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python script did not produce preview: ${outputPath}`);
    }

    // Read and send formatted HTML file
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('Python preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('Python preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown Python preview error';
    res.status(500).json({ error: `Failed to generate Python preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// JavaScript Preview endpoint - format JS for web viewing
app.post('/api/preview/js', uploadDocument.single('file'), async (req, res) => {
  console.log('=== JAVASCRIPT PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `js-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('JavaScript file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save JavaScript file to temp location
    const jsPath = path.join(tmpDir, 'input.js');
    await fs.writeFile(jsPath, file.buffer);

    // Use Python script to format JavaScript
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'js_to_formatted.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'JavaScript preview script not found' });
    }

    const args = [
      scriptPath,
      jsPath,
      outputPath,
      '--max-size-mb', '10'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python script did not produce JavaScript preview: ${outputPath}`);
    }

    // Read and send formatted HTML file
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('JavaScript preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('JavaScript preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown JavaScript preview error';
    res.status(500).json({ error: `Failed to generate JavaScript preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// CSS Preview endpoint - format CSS for web viewing
app.post('/api/preview/css', uploadDocument.single('file'), async (req, res) => {
  console.log('=== CSS PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `css-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('CSS file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save CSS file to temp location
    const cssPath = path.join(tmpDir, 'input.css');
    await fs.writeFile(cssPath, file.buffer);

    // Use Python script to format CSS
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'css_to_formatted.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'CSS preview script not found' });
    }

    const args = [
      scriptPath,
      cssPath,
      outputPath,
      '--max-size-mb', '10'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python script did not produce CSS preview: ${outputPath}`);
    }

    // Read and send formatted HTML file
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('CSS preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('CSS preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown CSS preview error';
    res.status(500).json({ error: `Failed to generate CSS preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// HTML Preview endpoint - format HTML for web viewing
app.post('/api/preview/html', uploadDocument.single('file'), async (req, res) => {
  console.log('=== HTML PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `html-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('HTML file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save HTML file to temp location
    const htmlPath = path.join(tmpDir, 'input.html');
    await fs.writeFile(htmlPath, file.buffer);

    // Use Python script to format HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'html_to_formatted.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'HTML preview script not found' });
    }

    const args = [
      scriptPath,
      htmlPath,
      outputPath,
      '--max-size-mb', '10'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python script did not produce HTML preview: ${outputPath}`);
    }

    // Read and send formatted HTML file
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('HTML preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('HTML preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown HTML preview error';
    res.status(500).json({ error: `Failed to generate HTML preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// XML Preview endpoint - convert XML to HTML for web viewing
app.post('/api/preview/xml', uploadDocument.single('file'), async (req, res) => {
  console.log('=== XML PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `xml-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('XML file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save XML file to temp location
    const xmlPath = path.join(tmpDir, 'input.xml');
    await fs.writeFile(xmlPath, file.buffer);

    // Use Python script to convert XML to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'xml_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'XML preview script not found' });
    }

    const args = [
      scriptPath,
      xmlPath,
      htmlPath,
      '--max-size-mb', '10'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read and send HTML file
    const htmlContent = await fs.readFile(htmlPath, 'utf-8');

    console.log('XML preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('XML preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown XML preview error';
    res.status(500).json({ error: `Failed to generate XML preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// JSON Preview endpoint - convert JSON to HTML for web viewing
app.post('/api/preview/json', uploadDocument.single('file'), async (req, res) => {
  console.log('=== JSON PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `json-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('JSON file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save JSON file to temp location
    const jsonPath = path.join(tmpDir, 'input.json');
    await fs.writeFile(jsonPath, file.buffer);

    // Use Python script to convert JSON to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'json_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'JSON preview script not found' });
    }

    const args = [
      scriptPath,
      jsonPath,
      htmlPath,
      '--max-size-mb', '10'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read and send HTML file
    const htmlContent = await fs.readFile(htmlPath, 'utf-8');

    console.log('JSON preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('JSON preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown JSON preview error';
    res.status(500).json({ error: `Failed to generate JSON preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// ODS Preview endpoint - convert ODS to HTML for web viewing
app.post('/api/preview/ods', uploadDocument.single('file'), async (req, res) => {
  console.log('=== ODS PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `ods-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('ODS file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save ODS file to temp location
    const odsPath = path.join(tmpDir, 'input.ods');
    await fs.writeFile(odsPath, file.buffer);

    // Use Python script to convert ODS to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'ods_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'ODS preview script not found' });
    }

    const args = [
      scriptPath,
      odsPath,
      htmlPath,
      '--max-rows', '2000'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read and send HTML file
    const htmlContent = await fs.readFile(htmlPath, 'utf-8');

    console.log('ODS preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('ODS preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown ODS preview error';
    res.status(500).json({ error: `Failed to generate ODS preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// CSV Preview endpoint - convert CSV to HTML for web viewing
app.post('/api/preview/csv', uploadDocument.single('file'), async (req, res) => {
  console.log('=== CSV PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `csv-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('CSV file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save CSV file to temp location
    const csvPath = path.join(tmpDir, 'input.csv');
    await fs.writeFile(csvPath, file.buffer);

    // Use Python script to convert CSV to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'csv_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'CSV preview script not found' });
    }

    const args = [
      scriptPath,
      csvPath,
      htmlPath,
      '--max-rows', '2000'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read and send HTML file
    const htmlContent = await fs.readFile(htmlPath, 'utf-8');

    console.log('CSV preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('CSV preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown CSV preview error';
    res.status(500).json({ error: `Failed to generate CSV preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Excel Preview endpoint - convert Excel to HTML for web viewing
app.post('/api/preview/xlsx', uploadDocument.single('file'), async (req, res) => {
  console.log('=== EXCEL PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `xlsx-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('Excel file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Determine file extension from original filename
    const fileExt = path.extname(file.originalname).toLowerCase() || '.xlsx';
    console.log('File extension:', fileExt);

    // Save Excel file to temp location with correct extension
    const xlsxPath = path.join(tmpDir, `input${fileExt}`);
    await fs.writeFile(xlsxPath, file.buffer);

    // Use Python script to convert Excel to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'xlsx_to_html.py');
    const htmlPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`Python script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'Excel preview script not found' });
    }

    const args = [
      scriptPath,
      xlsxPath,
      htmlPath,
      '--max-rows', '2000'
    ];

    console.log('Executing Python script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args);
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('Python script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'Python execution failed';
    }

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`Python script error: ${stderr}`);
    }

    // Check if output file was created
    const htmlExists = await fs.access(htmlPath).then(() => true).catch(() => false);
    if (!htmlExists) {
      throw new Error(`Python script did not produce HTML preview: ${htmlPath}`);
    }

    // Read and send HTML file
    const htmlContent = await fs.readFile(htmlPath, 'utf-8');

    console.log('Excel preview successful:', {
      inputSize: file.size,
      outputLength: htmlContent.length
    });

    res.set('Content-Type', 'text/html');
    res.send(htmlContent);

  } catch (error) {
    console.error('Excel preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown Excel preview error';
    res.status(500).json({ error: `Failed to generate Excel preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// ODP Preview endpoint - convert ODP (OpenDocument Presentation) to HTML
app.post('/api/preview/odp', uploadDocument.single('file'), async (req, res) => {
  console.log('=== ODP PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `odp-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('ODP file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save ODP file to temp location
    const odpPath = path.join(tmpDir, 'input.odp');
    await fs.writeFile(odpPath, file.buffer);

    // Use Python script to convert ODP to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'odp_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`ODP script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'ODP preview script not found' });
    }

    const args = [scriptPath, odpPath, outputPath];

    console.log('Executing ODP script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('ODP script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'ODP execution failed';
    }

    if (stdout.trim().length > 0) console.log('ODP stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('ODP stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`ODP script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`ODP script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('ODP preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('ODP preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown ODP preview error';
    res.status(500).json({ error: `Failed to generate ODP preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// OTP Preview endpoint - convert OTP (OpenDocument Presentation Template) to HTML
app.post('/api/preview/otp', uploadDocument.single('file'), async (req, res) => {
  console.log('=== OTP PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `otp-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('OTP file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save OTP file to temp location
    const otpPath = path.join(tmpDir, 'input.otp');
    await fs.writeFile(otpPath, file.buffer);

    // Use Python script to convert OTP to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'otp_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`OTP script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'OTP preview script not found' });
    }

    const args = [scriptPath, otpPath, outputPath];

    console.log('Executing OTP script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('OTP script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'OTP execution failed';
    }

    if (stdout.trim().length > 0) console.log('OTP stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('OTP stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`OTP script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`OTP script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('OTP preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('OTP preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown OTP preview error';
    res.status(500).json({ error: `Failed to generate OTP preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// POT Preview endpoint - convert POT (PowerPoint Template) to HTML
app.post('/api/preview/pot', uploadDocument.single('file'), async (req, res) => {
  console.log('=== POT PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `pot-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('POT file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save POT file to temp location
    const potPath = path.join(tmpDir, 'input.pot');
    await fs.writeFile(potPath, file.buffer);

    // Use Python script to convert POT to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'pot_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`POT script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'POT preview script not found' });
    }

    const args = [scriptPath, potPath, outputPath];

    console.log('Executing POT script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('POT script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'POT execution failed';
    }

    if (stdout.trim().length > 0) console.log('POT stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('POT stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`POT script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`POT script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('POT preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('POT preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown POT preview error';
    res.status(500).json({ error: `Failed to generate POT preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// PPS/PPSX Preview endpoint - convert PPS/PPSX (PowerPoint Slide Show) to HTML
app.post('/api/preview/pps', uploadDocument.single('file'), async (req, res) => {
  console.log('=== PPS/PPSX PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `pps-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('PPS/PPSX file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Determine file extension
    const fileExt = file.originalname.split('.').pop()?.toLowerCase() || 'pps';
    const inputFileName = `input.${fileExt}`;

    // Save PPS/PPSX file to temp location
    const ppsPath = path.join(tmpDir, inputFileName);
    await fs.writeFile(ppsPath, file.buffer);

    // Use Python script to convert PPS/PPSX to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'pps_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`PPS/PPSX script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'PPS/PPSX preview script not found' });
    }

    const args = [scriptPath, ppsPath, outputPath];

    console.log('Executing PPS/PPSX script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('PPS/PPSX script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'PPS/PPSX execution failed';
    }

    if (stdout.trim().length > 0) console.log('PPS/PPSX stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('PPS/PPSX stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`PPS/PPSX script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`PPS/PPSX script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('PPS/PPSX preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length,
      format: fileExt.toUpperCase()
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('PPS/PPSX preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown PPS/PPSX preview error';
    res.status(500).json({ error: `Failed to generate PPS/PPSX preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// PPT/PPTX Preview endpoint - convert PPT/PPTX (PowerPoint Presentation) to HTML
app.post('/api/preview/ppt', uploadDocument.single('file'), async (req, res) => {
  console.log('=== PPT/PPTX PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `ppt-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('PPT/PPTX file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Determine file extension
    const fileExt = file.originalname.split('.').pop()?.toLowerCase() || 'ppt';
    const inputFileName = `input.${fileExt}`;

    // Save PPT/PPTX file to temp location
    const pptPath = path.join(tmpDir, inputFileName);
    await fs.writeFile(pptPath, file.buffer);
    
    // Set file permissions to ensure LibreOffice can read it
    await fs.chmod(pptPath, 0o644);
    
    console.log('PPT/PPTX file saved:', {
      path: pptPath,
      size: (await fs.stat(pptPath)).size,
      exists: await fs.access(pptPath).then(() => true).catch(() => false)
    });

    // Use Python script to convert PPT/PPTX to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'ppt_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`PPT/PPTX script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'PPT/PPTX preview script not found' });
    }

    const args = [scriptPath, pptPath, outputPath];

    console.log('Executing PPT/PPTX script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    let timedOut = false;
    try {
      const result = await execFileAsync(pythonPath, args, { 
        timeout: 60000,  // Reduced to 1 minute for faster feedback
        maxBuffer: 10 * 1024 * 1024  // 10MB buffer for large outputs
      });
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('PPT/PPTX script execution failed:', execError);
      if (execError.killed && execError.signal === 'SIGTERM') {
        timedOut = true;
        console.error('Script timed out after 60 seconds');
      }
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'PPT/PPTX execution failed';
    }

    console.log('=== PPT/PPTX PYTHON SCRIPT OUTPUT ===');
    if (stdout && stdout.trim().length > 0) {
      console.log('STDOUT:');
      console.log(stdout.substring(0, 5000)); // Limit output to prevent flooding logs
    }
    if (stderr && stderr.trim().length > 0) {
      console.log('STDERR:');
      console.log(stderr.substring(0, 5000)); // Limit output to prevent flooding logs
    }
    console.log('=== END SCRIPT OUTPUT ===');
    
    if (timedOut) {
      throw new Error('PowerPoint conversion timed out. The file may be too large or complex.');
    }
    
    // Check for errors - be more lenient with warnings
    if (stderr.includes('ERROR:') || stderr.includes('Traceback') || stderr.includes('CONVERSION FAILED')) {
      throw new Error(`PPT/PPTX conversion failed: ${stderr.substring(0, 500)}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      // List directory contents for debugging
      const dirContents = await fs.readdir(tmpDir);
      console.error('Output file not found. Directory contents:', dirContents);
      throw new Error(`PPT/PPTX script did not produce preview. Check logs for details.`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('PPT/PPTX preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length,
      format: fileExt.toUpperCase()
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('PPT/PPTX preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown PPT/PPTX preview error';
    res.status(500).json({ error: `Failed to generate PowerPoint preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// SDD Preview endpoint - convert SDD (StarOffice Presentation) to HTML
app.post('/api/preview/sdd', uploadDocument.single('file'), async (req, res) => {
  console.log('=== SDD PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `sdd-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('SDD file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save SDD file to temp location
    const sddPath = path.join(tmpDir, 'input.sdd');
    await fs.writeFile(sddPath, file.buffer);

    // Use Python script to convert SDD to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'sdd_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`SDD script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'SDD preview script not found' });
    }

    const args = [scriptPath, sddPath, outputPath];

    console.log('Executing SDD script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('SDD script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'SDD execution failed';
    }

    if (stdout.trim().length > 0) console.log('SDD stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('SDD stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`SDD script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`SDD script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('SDD preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('SDD preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown SDD preview error';
    res.status(500).json({ error: `Failed to generate SDD preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// STI Preview endpoint - convert STI (StarOffice Presentation Template) to HTML
app.post('/api/preview/sti', uploadDocument.single('file'), async (req, res) => {
  console.log('=== STI PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `sti-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('STI file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save STI file to temp location
    const stiPath = path.join(tmpDir, 'input.sti');
    await fs.writeFile(stiPath, file.buffer);

    // Use Python script to convert STI to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'sti_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`STI script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'STI preview script not found' });
    }

    const args = [scriptPath, stiPath, outputPath];

    console.log('Executing STI script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('STI script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'STI execution failed';
    }

    if (stdout.trim().length > 0) console.log('STI stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('STI stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`STI script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`STI script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('STI preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('STI preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown STI preview error';
    res.status(500).json({ error: `Failed to generate STI preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// SX Preview endpoint - convert SX (Stat Studio Program) to HTML
app.post('/api/preview/sx', uploadDocument.single('file'), async (req, res) => {
  console.log('=== SX PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `sx-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('SX file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save SX file to temp location
    const sxPath = path.join(tmpDir, 'input.sx');
    await fs.writeFile(sxPath, file.buffer);

    // Use Python script to convert SX to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'sx_to_formatted.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`SX script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'SX preview script not found' });
    }

    const args = [scriptPath, sxPath, outputPath, '--max-lines', '50000'];

    console.log('Executing SX script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('SX script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'SX execution failed';
    }

    if (stdout.trim().length > 0) console.log('SX stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('SX stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`SX script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`SX script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('SX preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('SX preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown SX preview error';
    res.status(500).json({ error: `Failed to generate SX preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// UOP Preview endpoint - convert UOP (Uniform Office Presentation) to HTML
app.post('/api/preview/uop', uploadDocument.single('file'), async (req, res) => {
  console.log('=== UOP PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `uop-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('UOP file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save UOP file to temp location
    const uopPath = path.join(tmpDir, 'input.uop');
    await fs.writeFile(uopPath, file.buffer);

    // Use Python script to convert UOP to HTML
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'uop_to_html.py');
    const outputPath = path.join(tmpDir, 'output.html');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`UOP script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'UOP preview script not found' });
    }

    const args = [scriptPath, uopPath, outputPath];

    console.log('Executing UOP script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('UOP script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'UOP execution failed';
    }

    if (stdout.trim().length > 0) console.log('UOP stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('UOP stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`UOP script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`UOP script did not produce preview: ${outputPath}`);
    }

    // Read HTML content
    const htmlContent = await fs.readFile(outputPath, 'utf-8');

    console.log('UOP preview successful:', {
      inputSize: file.size,
      outputSize: htmlContent.length
    });

    res.json({ htmlContent });

  } catch (error) {
    console.error('UOP preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown UOP preview error';
    res.status(500).json({ error: `Failed to generate UOP preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// PEF Preview endpoint - convert PEF (Pentax RAW) to web-viewable image
app.post('/api/preview/pef', uploadDocument.single('file'), async (req, res) => {
  console.log('=== PEF PREVIEW REQUEST ===');
  const tmpDir = path.join(os.tmpdir(), `pef-preview-${Date.now()}`);
  
  try {
    await fs.mkdir(tmpDir, { recursive: true });
    
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('PEF file received:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });

    // Save PEF file to temp location
    const pefPath = path.join(tmpDir, 'input.pef');
    await fs.writeFile(pefPath, file.buffer);

    // Use Python script to convert PEF to JPEG
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    const scriptPath = path.join(__dirname, '..', 'viewers', 'pef_to_image.py');
    const outputPath = path.join(tmpDir, 'output.jpg');
    const metadataPath = path.join(tmpDir, 'metadata.json');

    // Check if script exists
    const scriptExists = await fs.access(scriptPath).then(() => true).catch(() => false);
    if (!scriptExists) {
      console.error(`PEF script not found: ${scriptPath}`);
      return res.status(500).json({ error: 'PEF preview script not found' });
    }

    const args = [
      scriptPath,
      pefPath,
      outputPath,
      metadataPath,
      '--max-dimension', '2048'
    ];

    console.log('Executing PEF script:', { pythonPath, scriptPath, args });

    let stdout, stderr;
    try {
      const result = await execFileAsync(pythonPath, args, { timeout: 120000 }); // 2 min timeout for RAW processing
      stdout = result.stdout;
      stderr = result.stderr;
    } catch (execError: any) {
      console.error('PEF script execution failed:', execError);
      stdout = execError.stdout || '';
      stderr = execError.stderr || execError.message || 'PEF execution failed';
    }

    if (stdout.trim().length > 0) console.log('PEF stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('PEF stderr:', stderr.trim());
    
    // Check for errors
    if (stderr.includes('ERROR:') || stderr.includes('Traceback')) {
      throw new Error(`PEF script error: ${stderr}`);
    }

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`PEF script did not produce preview: ${outputPath}`);
    }

    // Read image file and metadata
    const imageBuffer = await fs.readFile(outputPath);
    let metadata = {};
    
    try {
      const metadataContent = await fs.readFile(metadataPath, 'utf-8');
      metadata = JSON.parse(metadataContent);
    } catch (error) {
      console.warn('Could not read metadata:', error);
    }

    console.log('PEF preview successful:', {
      inputSize: file.size,
      outputSize: imageBuffer.length,
      metadata
    });

    // Convert image to base64 data URL
    const base64Image = imageBuffer.toString('base64');
    const imageDataUrl = `data:image/jpeg;base64,${base64Image}`;

    res.json({
      imageUrl: imageDataUrl,
      metadata
    });

  } catch (error) {
    console.error('PEF preview error:', error);
    const message = error instanceof Error ? error.message : 'Unknown PEF preview error';
    res.status(500).json({ error: `Failed to generate PEF preview: ${message}` });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: CSV to Parquet (Single)
app.post('/convert/csv-to-parquet/single', upload.single('file'), async (req, res) => {
  console.log('CSV->Parquet single conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToParquetPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->Parquet single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to Parquet (Batch)
app.post('/convert/csv-to-parquet/batch', uploadBatch, async (req, res) => {
  console.log('CSV->Parquet batch conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToParquetPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->Parquet batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to SQL (Single)
app.post('/convert/csv-to-sql/single', upload.single('file'), async (req, res) => {
  console.log('CSV->SQL single conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToSqlPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->SQL single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to SQL (Batch)
app.post('/convert/csv-to-sql/batch', uploadBatch, async (req, res) => {
  console.log('CSV->SQL batch conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToSqlPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->SQL batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to TOML (Single)
app.post('/convert/csv-to-toml/single', upload.single('file'), async (req, res) => {
  console.log('CSV->TOML single conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToTomlPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->TOML single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to TOML (Batch)
app.post('/convert/csv-to-toml/batch', uploadBatch, async (req, res) => {
  console.log('CSV->TOML batch conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToTomlPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->TOML batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// CSV to TOML converter using Python
const convertCsvToTomlPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO TOML (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-toml-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.toml`);
    
    // Use Python script for TOML with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_toml.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      indent: options.indent || '2',
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--indent', options.indent || '2'
    ];

    const { stdout, stderr } = await execFileAsync(pythonPath, args);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python TOML script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python TOML script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.toml`;
    console.log(`CSV->TOML conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/toml');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/toml'
    };
  } catch (error) {
    console.error(`CSV->TOML conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->TOML error`;
    throw new Error(`Failed to convert CSV to TOML: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// Route: CSV to XML (Single)
app.post('/convert/csv-to-xml/single', upload.single('file'), async (req, res) => {
  console.log('CSV->XML single conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToXmlPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->XML single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to XML (Batch)
app.post('/convert/csv-to-xml/batch', uploadBatch, async (req, res) => {
  console.log('CSV->XML batch conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToXmlPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->XML batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// CSV to XML converter using Python
const convertCsvToXmlPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO XML (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-xml-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.xml`);
    
    // Use Python script for XML with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_xml.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      rootElement: options.rootElement || 'data',
      rowElement: options.rowElement || 'row',
      prettyPrint: options.prettyPrint !== 'false',
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--root-element', options.rootElement || 'data',
      '--row-element', options.rowElement || 'row'
    ];

    // Add no-pretty-print flag if prettyPrint is false
    if (options.prettyPrint === 'false') {
      args.push('--no-pretty-print');
    }

    const { stdout, stderr } = await execFileAsync(pythonPath, args);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python XML script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python XML script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.xml`;
    console.log(`CSV->XML conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/xml');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/xml'
    };
  } catch (error) {
    console.error(`CSV->XML conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->XML error`;
    throw new Error(`Failed to convert CSV to XML: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// Route: CSV to YAML (Single)
app.post('/convert/csv-to-yaml/single', upload.single('file'), async (req, res) => {
  console.log('CSV->YAML single conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToYamlPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->YAML single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to YAML (Batch)
app.post('/convert/csv-to-yaml/batch', uploadBatch, async (req, res) => {
  console.log('CSV->YAML batch conversion request');
  
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToYamlPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->YAML batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// CSV to YAML converter using Python
const convertCsvToYamlPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO YAML (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-yaml-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.yaml`);
    
    // Use Python script for YAML with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_yaml.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      indent: options.indent || '2',
      defaultFlowStyle: options.defaultFlowStyle || 'false',
      allowUnicode: options.allowUnicode || 'true',
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--indent', options.indent || '2',
      '--default-flow-style', options.defaultFlowStyle || 'false',
      '--allow-unicode', options.allowUnicode || 'true'
    ];

    const { stdout, stderr } = await execFileAsync(pythonPath, args);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python YAML script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python YAML script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.yaml`;
    console.log(`CSV->YAML conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/x-yaml');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/x-yaml'
    };
  } catch (error) {
    console.error(`CSV->YAML conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->YAML error`;
    throw new Error(`Failed to convert CSV to YAML: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to EPUB converter using Python
const convertCsvToEpubPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO EPUB (Python) START ===`);
  const startTime = Date.now();
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-epub-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.epub`);
    
    // Use specific Python script for EPUB
    const pythonPath = '/opt/venv/bin/python3';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_epub.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      title: options.title || sanitizedBase,
      author: options.author || 'Unknown',
      includeToc: options.includeTableOfContents !== 'false',
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--title', options.title || sanitizedBase,
      '--author', options.author || 'Unknown'
    ];

    // Add --no-toc flag if table of contents is disabled
    if (options.includeTableOfContents === 'false') {
      args.push('--no-toc');
    }

    const { stdout, stderr } = await execFileAsync(pythonPath, args, {
      timeout: 300000, // 5 minutes timeout for large files
      maxBuffer: 10 * 1024 * 1024 // 10MB buffer for large outputs
    });

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python EPUB script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python EPUB script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.epub`;
    const processingTime = Date.now() - startTime;
    console.log(`CSV->EPUB conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length,
      processingTimeMs: processingTime,
      processingTimeSec: (processingTime / 1000).toFixed(2)
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/epub+zip');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/epub+zip'
    };
  } catch (error) {
    console.error(`CSV->EPUB conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->EPUB error`;
    throw new Error(`Failed to convert CSV to EPUB: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to SQL converter using Python
const convertCsvToSqlPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO SQL (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-sql-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.sql`);
    
    // Use Python script for SQL with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_sql.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      tableName: options.tableName || 'data_table',
      dialect: options.dialect || 'mysql',
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--table-name', options.tableName || 'data_table',
      '--dialect', options.dialect || 'mysql'
    ];

    // Add optional flag
    if (options.includeCreateTable === 'false') {
      args.push('--no-create-table');
    }

    const { stdout, stderr } = await execFileAsync(pythonPath, args);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python SQL script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python SQL script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.sql`;
    console.log(`CSV->SQL conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/sql');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/sql'
    };
  } catch (error) {
    console.error(`CSV->SQL conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->SQL error`;
    throw new Error(`Failed to convert CSV to SQL: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// CSV to Parquet converter using Python
const convertCsvToParquetPython = async (
  file: Express.Multer.File,
  options: Record<string, string | undefined> = {},
  persistToDisk = false
): Promise<ConversionResult> => {
  console.log(`=== CSV TO PARQUET (Python) START ===`);
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-csv-parquet-'));
  const originalBase = path.basename(file.originalname, path.extname(file.originalname));
  const sanitizedBase = sanitizeFilename(originalBase);
  const safeBase = `${sanitizedBase}_${randomUUID()}`;

  try {
    // Write CSV file to temp directory
    const csvPath = path.join(tmpDir, `${safeBase}.csv`);
    await fs.writeFile(csvPath, file.buffer);

    // Prepare output file
    const outputPath = path.join(tmpDir, `${safeBase}.parquet`);
    
    // Use Python script for Parquet with virtual environment
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_parquet.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      csvPath,
      outputPath,
      compression: options.compression || 'snappy',
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      csvPath,
      outputPath,
      '--compression', options.compression || 'snappy'
    ];

    const { stdout, stderr } = await execFileAsync(pythonPath, args);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python Parquet script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python Parquet script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.parquet`;
    console.log(`CSV->Parquet conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/octet-stream');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/octet-stream'
    };
  } catch (error) {
    console.error(`CSV->Parquet conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown CSV->Parquet error`;
    throw new Error(`Failed to convert CSV to Parquet: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
};

// ====================================================================
// DOC to EPUB Conversion
// ====================================================================

async function convertDocToEpubPython(
  file: Express.Multer.File,
  persistToDisk: boolean = false
): Promise<ConversionResult> {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'morphy-doc-epub-'));
  
  try {
    const sanitizedBase = sanitizeFilename(path.parse(file.originalname).name);
    const docPath = path.join(tmpDir, `${sanitizedBase}.doc`);
    const outputPath = path.join(tmpDir, `${sanitizedBase}.epub`);

    console.log(`DOC->EPUB conversion request:`, { 
      filename: file.originalname, 
      size: file.buffer.length 
    });

    // Write DOC file
    await fs.writeFile(docPath, file.buffer);

    // Execute Python script
    const pythonPath = '/opt/venv/bin/python';
    const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_epub.py');
    
    console.log('Python execution details:', {
      pythonPath,
      scriptPath,
      docPath,
      outputPath,
      fileSize: file.buffer.length
    });

    const args = [
      scriptPath,
      docPath,
      outputPath
    ];

    const { stdout, stderr } = await execFileAsync(pythonPath, args);

    if (stdout.trim().length > 0) console.log('Python stdout:', stdout.trim());
    if (stderr.trim().length > 0) console.warn('Python stderr:', stderr.trim());

    // Check if output file was created
    const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
    if (!outputExists) {
      throw new Error(`Python EPUB script did not produce output file: ${outputPath}`);
    }

    // Read output file
    const outputBuffer = await fs.readFile(outputPath);
    if (!outputBuffer || outputBuffer.length === 0) {
      throw new Error('Python EPUB script produced empty output file');
    }

    const downloadName = `${sanitizedBase}.epub`;
    console.log(`DOC->EPUB conversion successful:`, { 
      filename: downloadName, 
      size: outputBuffer.length 
    });

    if (persistToDisk) {
      return await persistOutputBuffer(outputBuffer, downloadName, 'application/epub+zip');
    }

    return {
      buffer: outputBuffer,
      filename: downloadName,
      mime: 'application/epub+zip'
    };

  } catch (error) {
    console.error(`DOC->EPUB conversion error:`, error);
    const message = error instanceof Error ? error.message : `Unknown DOC->EPUB error`;
    throw new Error(`Failed to convert DOC to EPUB: ${message}`);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
}

app.post('/convert/doc-to-epub/single', uploadSingle, async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const result = await convertDocToEpubPython(req.file);

    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length
    });
    res.send(result.buffer);
  } catch (error) {
    console.error('DOC->EPUB single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

app.post('/convert/doc-to-epub/batch', uploadBatch, async (req, res) => {
  try {
    if (!req.files || !Array.isArray(req.files) || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const files = req.files as Express.Multer.File[];
    console.log(`DOC->EPUB batch request: ${files.length} files`);

    // Process files individually and handle errors gracefully
    const results = await Promise.allSettled(
      files.map(file => convertDocToEpubPython(file, true))
    );

    // Map results to include success/failure status
    const processedResults = results.map((result, index) => {
      if (result.status === 'fulfilled') {
        return result.value;
      } else {
        console.error(`File ${index} failed:`, result.reason);
        
        // Extract clean error message from Python exceptions
        let errorMessage = 'Conversion failed';
        if (result.reason instanceof Error) {
          const fullError = result.reason.message;
          // Look for the actual exception message (after "Exception: ")
          const exceptionMatch = fullError.match(/Exception: (.+?)(?:\n|$)/);
          if (exceptionMatch) {
            errorMessage = exceptionMatch[1];
          } else {
            errorMessage = fullError.split('\n')[0]; // Use first line if no exception found
          }
        }
        
        return {
          filename: files[index].originalname.replace(/\.doc$/i, '.epub'),
          error: errorMessage,
          downloadUrl: '',
          size: 0
        };
      }
    });

    res.json(processedResults);
  } catch (error) {
    console.error('DOC->EPUB batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to DOCX (Single)
app.post('/convert/csv-to-docx/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->DOCX single conversion request');
  
  // Set longer timeout for large CSV files (15 minutes)
  req.setTimeout(15 * 60 * 1000);
  res.setTimeout(15 * 60 * 1000);
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToDocxPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->DOCX single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CSV to DOCX (Batch)
app.post('/convert/csv-to-docx/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->DOCX batch conversion request');
  
  // Set longer timeout for large CSV files (15 minutes)
  req.setTimeout(15 * 60 * 1000);
  res.setTimeout(15 * 60 * 1000);
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToDocxPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->DOCX batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CSV to XLSX (Batch)
app.post('/convert/csv-to-xlsx/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->XLSX batch conversion request');
  
  // Set longer timeout for large CSV files (15 minutes)
  req.setTimeout(15 * 60 * 1000);
  res.setTimeout(15 * 60 * 1000);
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToXlsxPython(file, options, true);
        results.push({
          success: true,
          originalName: file.originalname,
          outputFilename: result.filename,
          downloadPath: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          originalName: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ 
      success: true,
      processed: results.length,
      results 
    });
  } catch (error) {
    console.error('CSV->XLSX batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CSV to JSON (Batch)
app.post('/convert/csv-to-json/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->JSON batch conversion request');
  
  // Set longer timeout for large CSV files (15 minutes)
  req.setTimeout(15 * 60 * 1000);
  res.setTimeout(15 * 60 * 1000);
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToJsonPython(file, options, true);
        results.push({
          success: true,
          originalName: file.originalname,
          outputFilename: result.filename,
          downloadPath: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          originalName: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ 
      success: true,
      processed: results.length,
      results 
    });
  } catch (error) {
    console.error('CSV->JSON batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CSV to EPUB (Single)
app.post('/convert/csv-to-epub/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->EPUB single conversion request');
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToEpubPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->EPUB single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CSV to EPUB (Batch)
app.post('/convert/csv-to-epub/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->EPUB batch conversion request');
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToEpubPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->EPUB batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CSV to HTML (Single)
app.post('/convert/csv-to-html/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->HTML single conversion request');
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToHtmlPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->HTML single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to HTML (Batch)
app.post('/convert/csv-to-html/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->HTML batch conversion request');
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToHtmlPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->HTML batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to Markdown (Single)
app.post('/convert/csv-to-md/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->Markdown single conversion request');
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToMdPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->Markdown single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to Markdown (Batch)
app.post('/convert/csv-to-md/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->Markdown batch conversion request');
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToMdPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->Markdown batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to EPUB (Single)
app.post('/convert/csv-to-epub/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->EPUB single conversion request');
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToEpubPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->EPUB single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to EPUB (Batch)
app.post('/convert/csv-to-epub/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->EPUB batch conversion request');
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToEpubPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->EPUB batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to MOBI (Single)
app.post('/convert/csv-to-mobi/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->MOBI single conversion request');
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToMobiPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->MOBI single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to MOBI (Batch)
app.post('/convert/csv-to-mobi/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->MOBI batch conversion request');
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToMobiPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.json({ results });
  } catch (error) {
    console.error('CSV->MOBI batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CSV to ODP (Single)
app.post('/convert/csv-to-odp/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->ODP single conversion request');
  
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const options = req.body || {};
    const result = await convertCsvToOdpPython(file, options, false);
    
    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    
    res.send(result.buffer);
  } catch (error) {
    console.error('CSV->ODP single error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Route: CSV to ODP (Batch)
app.post('/convert/csv-to-odp/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CSV->ODP batch conversion request');
  
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files provided' });
    }

    const options = req.body || {};
    const results = [];

    for (const file of files) {
      try {
        const result = await convertCsvToOdpPython(file, options, true);
        results.push({
          success: true,
          filename: result.filename,
          downloadUrl: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          success: false,
          filename: file.originalname,
          error: error instanceof Error ? error.message : 'Conversion failed'
        });
      }
    }

    res.json({ results });
  } catch (error) {
    console.error('CSV->ODP batch error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ 
      error: message,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
  }
});

// Initialize dotenv
dotenv.config();

// CORS configuration
const allowedOrigins = [
  process.env.FRONTEND_URL,
  process.env.FRONTEND_URL_ALT,
  'https://morphyimg.com',
  'https://morphy-1-ulvv.onrender.com',
  'https://morphy-2-n2tb.onrender.com',
  'http://localhost:5173', // Frontend dev server
  'http://localhost:3000', // Backend dev server
].filter(Boolean) as string[];

// Temporary permissive CORS for debugging
app.use(cors({
  origin: true, // Allow all origins temporarily
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept'],
  credentials: true,
}));

// Original CORS configuration (commented out for debugging)
/*
app.use(cors({
  origin: (origin, callback) => {
    console.log('CORS check - Request origin:', origin);
    console.log('CORS check - Allowed origins:', allowedOrigins);
    
    // Allow requests with no origin (like mobile apps or curl requests)
    if (!origin) {
      console.log('CORS check - No origin, allowing request');
      return callback(null, true);
    }
    
    if (allowedOrigins.indexOf(origin) === -1) {
      const msg = `The CORS policy for this site does not allow access from the specified Origin: ${origin}`;
      console.log('CORS check - Origin not allowed:', origin);
      return callback(new Error(msg), false);
    }
    
    console.log('CORS check - Origin allowed:', origin);
    return callback(null, true);
  },
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept'],
  credentials: true,
}));
*/

// Security middleware
app.use(helmet());

// Rate limiting to prevent abuse
app.use(limiter);

// Body parser for JSON and URL-encoded data
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV,
    database: {
      host: process.env.DB_HOST || 'Not configured',
      port: process.env.DB_PORT || 'Not configured',
      database: process.env.DB_NAME || 'Not configured',
      user: process.env.DB_USER || 'Not configured',
      ssl: process.env.DB_SSL === 'true'
    }
  });
});










// ==================== IMAGE CONVERSION ROUTES ====================

// Duplicate BMP routes removed - using the correct ones later in the file

// Route: AVRO to JSON (Single)
app.post('/convert/avro-to-json/single', upload.single('file'), async (req, res) => {
  console.log('AVRO->JSON single conversion request');

  const tmpDir = path.join(os.tmpdir(), `avro-json-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.avro$/i, '.json'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'avro_to_json.py');
    console.log('AVRO to JSON: Executing Python script:', scriptPath);
    console.log('AVRO to JSON: Input file:', inputPath);
    console.log('AVRO to JSON: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('AVRO to JSON: Script exists');
    } catch (error) {
      console.error('AVRO to JSON: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    const python = spawn('/opt/venv/bin/python', [
      scriptPath,
      inputPath,
      outputPath
    ]);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('AVRO to JSON stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('AVRO to JSON stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('AVRO to JSON: Python script finished with code:', code);
      console.log('AVRO to JSON: stdout:', stdout);
      console.log('AVRO to JSON: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('AVRO to JSON: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/json',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`
          });
          res.send(outputBuffer);
        } else {
          console.error('AVRO to JSON conversion failed. Code:', code, 'Stderr:', stderr);
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.status(500).json({ error: 'Conversion failed', details: error.message });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('AVRO to JSON conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: AVRO to JSON (Batch)
app.post('/convert/avro-to-json/batch', uploadBatch, async (req, res) => {
  console.log('AVRO->JSON batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `avro-json-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.avro$/i, '.json'));

        await fs.writeFile(inputPath, file.buffer);

        const python = spawn('/opt/venv/bin/python', [
          path.join(__dirname, '..', 'scripts', 'avro_to_json.py'),
          inputPath,
          outputPath
        ]);

        await new Promise<void>((resolve, reject) => {
          python.on('close', (code: number) => {
            if (code === 0) {
              resolve();
            } else {
              reject(new Error(`Conversion failed for ${file.originalname}`));
            }
          });
        });

        const outputBuffer = await fs.readFile(outputPath);
        results.push({
          originalName: file.originalname,
          outputFilename: path.basename(outputPath),
          size: outputBuffer.length,
          success: true,
          downloadPath: `data:application/json;base64,${outputBuffer.toString('base64')}`
        });
      } catch (error) {
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
  } catch (error) {
    console.error('AVRO to JSON batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: AVRO to NDJSON (Single)
app.post('/convert/avro-to-ndjson/single', upload.single('file'), async (req, res) => {
  console.log('AVRO->NDJSON single conversion request');

  const tmpDir = path.join(os.tmpdir(), `avro-ndjson-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.avro$/i, '.ndjson'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'avro_to_ndjson.py');
    console.log('AVRO to NDJSON: Executing Python script:', scriptPath);
    console.log('AVRO to NDJSON: Input file:', inputPath);
    console.log('AVRO to NDJSON: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('AVRO to NDJSON: Script exists');
    } catch (error) {
      console.error('AVRO to NDJSON: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    const python = spawn('/opt/venv/bin/python', [
      scriptPath,
      inputPath,
      outputPath
    ]);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('AVRO to NDJSON stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('AVRO to NDJSON stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('AVRO to NDJSON: Python script finished with code:', code);
      console.log('AVRO to NDJSON: stdout:', stdout);
      console.log('AVRO to NDJSON: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('AVRO to NDJSON: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/x-ndjson',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`
          });
          res.send(outputBuffer);
        } else {
          console.error('AVRO to NDJSON conversion failed. Code:', code, 'Stderr:', stderr);
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.status(500).json({ error: 'Conversion failed', details: error.message });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('AVRO to NDJSON conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: AVRO to NDJSON (Batch)
app.post('/convert/avro-to-ndjson/batch', uploadBatch, async (req, res) => {
  console.log('AVRO->NDJSON batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `avro-ndjson-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.avro$/i, '.ndjson'));

        await fs.writeFile(inputPath, file.buffer);

        const python = spawn('/opt/venv/bin/python', [
          path.join(__dirname, '..', 'scripts', 'avro_to_ndjson.py'),
          inputPath,
          outputPath
        ]);

        await new Promise<void>((resolve, reject) => {
          python.on('close', (code: number) => {
            if (code === 0) {
              resolve();
            } else {
              reject(new Error(`Conversion failed for ${file.originalname}`));
            }
          });
        });

        const outputBuffer = await fs.readFile(outputPath);
        results.push({
          originalName: file.originalname,
          outputFilename: path.basename(outputPath),
          size: outputBuffer.length,
          success: true,
          downloadPath: `data:application/x-ndjson;base64,${outputBuffer.toString('base64')}`
        });
      } catch (error) {
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
  } catch (error) {
    console.error('AVRO to NDJSON batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: CSV to AVRO (Single)
app.post('/convert/csv-to-avro/single', upload.single('file'), async (req, res) => {
  console.log('CSV->AVRO single conversion request');

  const tmpDir = path.join(os.tmpdir(), `csv-avro-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.csv$/i, '.avro'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'csv_to_avro.py');
    console.log('CSV to AVRO: Executing Python script:', scriptPath);
    console.log('CSV to AVRO: Input file:', inputPath);
    console.log('CSV to AVRO: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('CSV to AVRO: Script exists');
    } catch (error) {
      console.error('CSV to AVRO: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    const python = spawn('/opt/venv/bin/python', [
      scriptPath,
      inputPath,
      outputPath
    ]);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('CSV to AVRO stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('CSV to AVRO stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('CSV to AVRO: Python script finished with code:', code);
      console.log('CSV to AVRO: stdout:', stdout);
      console.log('CSV to AVRO: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('CSV to AVRO: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/avro',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`
          });
          res.send(outputBuffer);
        } else {
          console.error('CSV to AVRO conversion failed. Code:', code, 'Stderr:', stderr);
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.status(500).json({ error: 'Conversion failed', details: error.message });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('CSV to AVRO conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CSV to AVRO (Batch)
app.post('/convert/csv-to-avro/batch', uploadBatch, async (req, res) => {
  console.log('CSV->AVRO batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `csv-avro-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.csv$/i, '.avro'));

        await fs.writeFile(inputPath, file.buffer);

        const python = spawn('/opt/venv/bin/python', [
          path.join(__dirname, '..', 'scripts', 'csv_to_avro.py'),
          inputPath,
          outputPath
        ]);

        await new Promise<void>((resolve, reject) => {
          python.on('close', (code: number) => {
            if (code === 0) {
              resolve();
            } else {
              reject(new Error(`Conversion failed for ${file.originalname}`));
            }
          });
        });

        const outputBuffer = await fs.readFile(outputPath);
        results.push({
          originalName: file.originalname,
          outputFilename: path.basename(outputPath),
          size: outputBuffer.length,
          success: true,
          downloadPath: `data:application/avro;base64,${outputBuffer.toString('base64')}`
        });
      } catch (error) {
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
  } catch (error) {
    console.error('CSV to AVRO batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// ==================== IMAGE CONVERSION ROUTES ====================

// Route: BMP to WebP (Single)
app.post('/convert/bmp-to-webp/single', upload.single('file'), async (req, res) => {
  console.log('BMP->WebP single conversion request');

  const tmpDir = path.join(os.tmpdir(), `bmp-webp-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.bmp$/i, '.webp'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'bmp_to_webp.py');
    console.log('BMP to WebP: Executing Python script:', scriptPath);
    console.log('BMP to WebP: Input file:', inputPath);
    console.log('BMP to WebP: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('BMP to WebP: Script exists');
    } catch (error) {
      console.error('BMP to WebP: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    const python = spawn('/opt/venv/bin/python', [
      scriptPath,
      inputPath,
      outputPath,
      '--quality', '80'
    ]);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('BMP to WebP stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('BMP to WebP stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('BMP to WebP: Python script finished with code:', code);
      console.log('BMP to WebP: stdout:', stdout);
      console.log('BMP to WebP: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('BMP to WebP: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'image/webp',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`
          });
          res.send(outputBuffer);
          
        } else {
          console.error('BMP to WebP conversion failed. Code:', code, 'Stderr:', stderr);
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.status(500).json({ error: 'Conversion failed', details: error.message });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('BMP to WebP conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: BMP to WebP (Batch)
app.post('/convert/bmp-to-webp/batch', uploadBatch, async (req, res) => {
  console.log('BMP->WebP batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `bmp-webp-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.bmp$/i, '.webp'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'bmp_to_webp.py');
        console.log('BMP to WebP batch: Executing Python script:', scriptPath);
        console.log('BMP to WebP batch: Input file:', inputPath);
        console.log('BMP to WebP batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('BMP to WebP batch: Script exists');
        } catch (error) {
          console.error('BMP to WebP batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const python = spawn('/opt/venv/bin/python', [
          scriptPath,
          inputPath,
          outputPath,
          '--quality', '80'
        ]);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('BMP to WebP batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('BMP to WebP batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('BMP to WebP batch: Python script finished with code:', code);
            console.log('BMP to WebP batch: stdout:', stdout);
            console.log('BMP to WebP batch: stderr:', stderr);

            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('BMP to WebP batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:image/webp;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('BMP to WebP batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files even if this one failed
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files even if this one failed
            }
          });
        });
      } catch (error) {
        console.error('BMP to WebP batch conversion error for file:', file.originalname, error);
        // Error already pushed to results in the promise rejection
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('BMP to WebP batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: CR2 to ICO (Single)
app.post('/convert/cr2-to-ico/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CR2->ICO single conversion request');

  // Set longer timeout for CR2 processing (10 minutes)
  req.setTimeout(10 * 60 * 1000);
  res.setTimeout(10 * 60 * 1000);
  
  // Handle timeout gracefully with CORS headers
  const timeoutHandler = () => {
    console.log('CR2 to ICO: Request timeout - sending timeout response');
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    if (!res.headersSent) {
      res.status(408).json({ 
        error: 'Conversion timeout', 
        message: 'CR2 to ICO conversion is taking longer than expected. Please try with a smaller file or contact support.',
        timeout: true
      });
    }
  };
  
  // Set timeout handler
  req.on('timeout', timeoutHandler);
  res.on('timeout', timeoutHandler);

  const tmpDir = path.join(os.tmpdir(), `cr2-ico-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.cr2$/i, '.ico'));

    await fs.writeFile(inputPath, file.buffer);

    // Get icon size from request body (use original size if not specified)
    const iconSize = req.body.iconSize;
    console.log('CR2 to ICO: Icon size received:', iconSize, 'Type:', typeof iconSize);
    console.log('CR2 to ICO: Icon size check (iconSize !== "default"):', iconSize !== 'default');

    const scriptPath = path.join(__dirname, '..', 'scripts', 'cr2_to_ico.py');
    console.log('CR2 to ICO: Executing Python script:', scriptPath);
    console.log('CR2 to ICO: Input file:', inputPath);
    console.log('CR2 to ICO: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('CR2 to ICO: Script exists');
    } catch (error) {
      console.error('CR2 to ICO: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', 'high'];
    console.log('CR2 to ICO: iconSize value:', iconSize, 'type:', typeof iconSize);
    
    // SIMPLIFIED: Only pass --sizes if user selected a specific size
    // If iconSize is 'default' or not provided, don't pass --sizes at all
    // Python script will use original image size by default
    if (iconSize && iconSize !== 'default') {
      pythonArgs.push('--sizes', iconSize.toString());
      console.log('CR2 to ICO: Using custom size:', iconSize);
    } else {
      console.log('CR2 to ICO: Using ORIGINAL image size (no --sizes parameter)');
    }
    
    console.log('CR2 to ICO: Final Python command:', pythonArgs.join(' '));
    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('CR2 to ICO stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('CR2 to ICO stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('CR2 to ICO: Python script finished with code:', code);
      console.log('CR2 to ICO: stdout:', stdout);
      console.log('CR2 to ICO: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('CR2 to ICO: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'image/x-icon',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('CR2 to ICO conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error.message });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('CR2 to ICO conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CR2 to ICO (Batch)
app.post('/convert/cr2-to-ico/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CR2->ICO batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `cr2-ico-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.cr2$/i, '.ico'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'cr2_to_ico.py');
        console.log('CR2 to ICO batch: Executing Python script:', scriptPath);
        console.log('CR2 to ICO batch: Input file:', inputPath);
        console.log('CR2 to ICO batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('CR2 to ICO batch: Script exists');
        } catch (error) {
          console.error('CR2 to ICO batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', 'high'];
        console.log('CR2 to ICO batch: iconSize value:', req.body.iconSize, 'type:', typeof req.body.iconSize);
        
        // SIMPLIFIED: Only pass --sizes if user selected a specific size
        if (req.body.iconSize && req.body.iconSize !== 'default') {
          pythonArgs.push('--sizes', req.body.iconSize.toString());
          console.log('CR2 to ICO batch: Using custom size:', req.body.iconSize);
        } else {
          console.log('CR2 to ICO batch: Using ORIGINAL image size (no --sizes parameter)');
        }
        
        console.log('CR2 to ICO batch: Final Python command:', pythonArgs.join(' '));
        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('CR2 to ICO batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('CR2 to ICO batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('CR2 to ICO batch: Python script finished with code:', code);
            console.log('CR2 to ICO batch: stdout:', stdout);
            console.log('CR2 to ICO batch: stderr:', stderr);

            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('CR2 to ICO batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:image/x-icon;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('CR2 to ICO batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files even if this one failed
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files even if this one failed
            }
          });
        });
      } catch (error) {
        console.error('CR2 to ICO batch conversion error for file:', file.originalname, error);
        // Error already pushed to results in the promise rejection
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('CR2 to ICO batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});


// Route: DNG to WebP (Single)
app.post('/convert/dng-to-webp/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DNG->WebP single conversion request');

  // Set longer timeout for DNG processing (10 minutes)
  req.setTimeout(10 * 60 * 1000);
  res.setTimeout(10 * 60 * 1000);
  
  // Handle timeout gracefully with CORS headers
  const timeoutHandler = () => {
    console.log('DNG to WebP: Request timeout - sending timeout response');
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    if (!res.headersSent) {
      res.status(408).json({ 
        error: 'Conversion timeout', 
        message: 'DNG to WebP conversion is taking longer than expected. Please try with a smaller file or contact support.',
        timeout: true
      });
    }
  };
  
  // Set timeout handler
  req.on('timeout', timeoutHandler);
  res.on('timeout', timeoutHandler);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const options = req.body || {};
    const result = await convertDngToWebpPython(file, options, false);

    res.set({
      'Content-Type': result.mime,
      'Content-Disposition': `attachment; filename="${result.filename}"`,
      'Content-Length': result.buffer.length,
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });

    res.send(result.buffer);
  } catch (error) {
    console.error('DNG->WebP single conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DNG to WebP (Batch)
app.post('/convert/dng-to-webp/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DNG->WebP batch conversion request');

  try {
    const files = req.files as Express.Multer.File[] | undefined;
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const options = req.body || {};
    const results = [];
    
    for (const file of files) {
      try {
        const result = await convertDngToWebpPython(file, options, true);
        results.push({
          originalName: file.originalname,
          outputFilename: result.filename,
          success: true,
          downloadPath: result.downloadUrl,
          size: result.size
        });
      } catch (error) {
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }
    
    res.json({ success: true, results });
  } catch (error) {
    console.error('DNG->WebP batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: CR2 to WebP (Single)
app.post('/convert/cr2-to-webp/single', upload.single('file'), async (req, res) => {
  console.log('CR2->WebP single conversion request');

  // Set longer timeout for CR2 processing (10 minutes)
  req.setTimeout(10 * 60 * 1000);
  res.setTimeout(10 * 60 * 1000);
  
  // Handle timeout gracefully with CORS headers
  const timeoutHandler = () => {
    console.log('CR2 to WebP: Request timeout - sending timeout response');
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    if (!res.headersSent) {
      res.status(408).json({ 
        error: 'Conversion timeout', 
        message: 'CR2 to WebP conversion is taking longer than expected. Please try with a smaller file or contact support.',
        timeout: true
      });
    }
  };
  
  // Set timeout handler
  req.on('timeout', timeoutHandler);
  res.on('timeout', timeoutHandler);

  const tmpDir = path.join(os.tmpdir(), `cr2-webp-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.cr2$/i, '.webp'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'cr2_to_webp.py');
    console.log('CR2 to WebP: Executing Python script:', scriptPath);
    console.log('CR2 to WebP: Input file:', inputPath);
    console.log('CR2 to WebP: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('CR2 to WebP: Script exists');
    } catch (error) {
      console.error('CR2 to WebP: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    const python = spawn('/opt/venv/bin/python', [
      scriptPath,
      inputPath,
      outputPath,
      '--quality', '80'
    ]);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('CR2 to WebP stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('CR2 to WebP stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('CR2 to WebP: Python script finished with code:', code);
      console.log('CR2 to WebP: stdout:', stdout);
      console.log('CR2 to WebP: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('CR2 to WebP: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'image/webp',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('CR2 to WebP conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error.message });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('CR2 to WebP conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: CR2 to WebP (Batch)
app.post('/convert/cr2-to-webp/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('CR2->WebP batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `cr2-webp-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.cr2$/i, '.webp'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'cr2_to_webp.py');
        console.log('CR2 to WebP batch: Executing Python script:', scriptPath);
        console.log('CR2 to WebP batch: Input file:', inputPath);
        console.log('CR2 to WebP batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('CR2 to WebP batch: Script exists');
        } catch (error) {
          console.error('CR2 to WebP batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const python = spawn('/opt/venv/bin/python', [
          scriptPath,
          inputPath,
          outputPath,
          '--quality', '80'
        ]);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('CR2 to WebP batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('CR2 to WebP batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('CR2 to WebP batch: Python script finished with code:', code);
            console.log('CR2 to WebP batch: stdout:', stdout);
            console.log('CR2 to WebP batch: stderr:', stderr);

            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('CR2 to WebP batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:image/webp;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('CR2 to WebP batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files even if this one failed
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files even if this one failed
            }
          });
        });
      } catch (error) {
        console.error('CR2 to WebP batch conversion error for file:', file.originalname, error);
        // Error already pushed to results in the promise rejection
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('CR2 to WebP batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});


// Route: EPS to WebP (Single)
app.post('/convert/eps-to-webp/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('EPS->WebP single conversion request');

  const tmpDir = path.join(os.tmpdir(), `eps-webp-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.eps$/i, '.webp'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'eps_to_webp.py');
    console.log('EPS to WebP: Executing Python script:', scriptPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('EPS to WebP: Script exists');
    } catch (error) {
      console.error('EPS to WebP: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get quality from request body and convert string values to numeric
    let quality = req.body.quality || 80;
    if (typeof quality === 'string') {
      switch (quality.toLowerCase()) {
        case 'high':
          quality = 90;
          break;
        case 'medium':
          quality = 70;
          break;
        case 'low':
          quality = 50;
          break;
        default:
          quality = parseInt(quality) || 80;
      }
    }
    const lossless = req.body.lossless === 'true' || req.body.lossless === true;
    
    const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', quality.toString()];
    if (lossless) {
      pythonArgs.push('--lossless');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('EPS to WebP stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('EPS to WebP stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('EPS to WebP: Python script finished with code:', code);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('EPS to WebP: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'image/webp',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`
          });
          res.send(outputBuffer);
        } else {
          console.error('EPS to WebP conversion failed. Code:', code, 'Stderr:', stderr);
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('EPS to WebP conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  }
});

// Route: EPS to WebP (Batch)
app.post('/convert/eps-to-webp/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('EPS->WebP batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `eps-webp-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get quality and lossless from request body and convert string values to numeric
    let quality = req.body.quality || 80;
    if (typeof quality === 'string') {
      switch (quality.toLowerCase()) {
        case 'high':
          quality = 90;
          break;
        case 'medium':
          quality = 70;
          break;
        case 'low':
          quality = 50;
          break;
        default:
          quality = parseInt(quality) || 80;
      }
    }
    const lossless = req.body.lossless === 'true' || req.body.lossless === true;

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.eps$/i, '.webp'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'eps_to_webp.py');
        console.log('EPS to WebP batch: Executing Python script:', scriptPath);
        
        // Check if script exists
        try {
          await fs.access(scriptPath);
          console.log('EPS to WebP batch: Script exists');
        } catch (error) {
          console.error('EPS to WebP batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', quality.toString()];
        if (lossless) {
          pythonArgs.push('--lossless');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('EPS to WebP batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('EPS to WebP batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('EPS to WebP batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('EPS to WebP batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:image/webp;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('EPS to WebP batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('EPS to WebP batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('EPS to WebP batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: EPUB to CSV (Single)
app.post('/convert/epub-to-csv/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('EPUB->CSV single conversion request');

  const tmpDir = path.join(os.tmpdir(), `epub-csv-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.epub$/i, '.csv'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'epub_to_csv.py');
    console.log('EPUB to CSV: Executing Python script:', scriptPath);
    console.log('EPUB to CSV: Input file:', inputPath);
    console.log('EPUB to CSV: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('EPUB to CSV: Script exists');
    } catch (error) {
      console.error('EPUB to CSV: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const includeMetadata = req.body.includeMetadata !== 'false';
    const delimiter = req.body.delimiter || ',';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!includeMetadata) {
      pythonArgs.push('--no-metadata');
    }
    if (delimiter !== ',') {
      pythonArgs.push('--delimiter', delimiter);
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('EPUB to CSV stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('EPUB to CSV stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('EPUB to CSV: Python script finished with code:', code);
      console.log('EPUB to CSV: stdout:', stdout);
      console.log('EPUB to CSV: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('EPUB to CSV: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'text/csv',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('EPUB to CSV conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('EPUB to CSV conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: EPUB to CSV (Batch)
app.post('/convert/epub-to-csv/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('EPUB->CSV batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `epub-csv-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const includeMetadata = req.body.includeMetadata !== 'false';
    const delimiter = req.body.delimiter || ',';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.epub$/i, '.csv'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'epub_to_csv.py');
        console.log('EPUB to CSV batch: Executing Python script:', scriptPath);
        console.log('EPUB to CSV batch: Input file:', inputPath);
        console.log('EPUB to CSV batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('EPUB to CSV batch: Script exists');
        } catch (error) {
          console.error('EPUB to CSV batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!includeMetadata) {
          pythonArgs.push('--no-metadata');
        }
        if (delimiter !== ',') {
          pythonArgs.push('--delimiter', delimiter);
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('EPUB to CSV batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('EPUB to CSV batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('EPUB to CSV batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('EPUB to CSV batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:text/csv;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('EPUB to CSV batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('EPUB to CSV batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('EPUB to CSV batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOCX to CSV (Single)
app.post('/convert/docx-to-csv/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->CSV single conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-csv-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.csv'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_csv.py');
    console.log('DOCX to CSV: Executing Python script:', scriptPath);
    console.log('DOCX to CSV: Input file:', inputPath);
    console.log('DOCX to CSV: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOCX to CSV: Script exists');
    } catch (error) {
      console.error('DOCX to CSV: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const extractTables = req.body.extractTables !== 'false';
    const includeParagraphs = req.body.includeParagraphs !== 'false';
    const delimiter = req.body.delimiter || ',';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!extractTables) {
      pythonArgs.push('--no-tables');
    }
    if (!includeParagraphs) {
      pythonArgs.push('--no-paragraphs');
    }
    if (delimiter !== ',') {
      pythonArgs.push('--delimiter', delimiter);
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOCX to CSV stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOCX to CSV stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOCX to CSV: Python script finished with code:', code);
      console.log('DOCX to CSV: stdout:', stdout);
      console.log('DOCX to CSV: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOCX to CSV: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'text/csv',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOCX to CSV conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOCX to CSV conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOCX to CSV (Batch)
app.post('/convert/docx-to-csv/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->CSV batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-csv-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const extractTables = req.body.extractTables !== 'false';
    const includeParagraphs = req.body.includeParagraphs !== 'false';
    const delimiter = req.body.delimiter || ',';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.csv'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_csv.py');
        console.log('DOCX to CSV batch: Executing Python script:', scriptPath);
        console.log('DOCX to CSV batch: Input file:', inputPath);
        console.log('DOCX to CSV batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOCX to CSV batch: Script exists');
        } catch (error) {
          console.error('DOCX to CSV batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!extractTables) {
          pythonArgs.push('--no-tables');
        }
        if (!includeParagraphs) {
          pythonArgs.push('--no-paragraphs');
        }
        if (delimiter !== ',') {
          pythonArgs.push('--delimiter', delimiter);
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOCX to CSV batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOCX to CSV batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOCX to CSV batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOCX to CSV batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:text/csv;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOCX to CSV batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOCX to CSV batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOCX to CSV batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOCX to EPUB (Single)
app.post('/convert/docx-to-epub/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->EPUB single conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-epub-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.epub'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_epub.py');
    console.log('DOCX to EPUB: Executing Python script:', scriptPath);
    console.log('DOCX to EPUB: Input file:', inputPath);
    console.log('DOCX to EPUB: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOCX to EPUB: Script exists');
    } catch (error) {
      console.error('DOCX to EPUB: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateTOC = req.body.generateTOC !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!includeImages) {
      pythonArgs.push('--no-images');
    }
    if (!preserveFormatting) {
      pythonArgs.push('--no-formatting');
    }
    if (!generateTOC) {
      pythonArgs.push('--no-toc');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOCX to EPUB stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOCX to EPUB stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOCX to EPUB: Python script finished with code:', code);
      console.log('DOCX to EPUB: stdout:', stdout);
      console.log('DOCX to EPUB: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOCX to EPUB: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/epub+zip',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOCX to EPUB conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOCX to EPUB conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOCX to EPUB (Batch)
app.post('/convert/docx-to-epub/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->EPUB batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-epub-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateTOC = req.body.generateTOC !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.epub'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_epub.py');
        console.log('DOCX to EPUB batch: Executing Python script:', scriptPath);
        console.log('DOCX to EPUB batch: Input file:', inputPath);
        console.log('DOCX to EPUB batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOCX to EPUB batch: Script exists');
        } catch (error) {
          console.error('DOCX to EPUB batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!includeImages) {
          pythonArgs.push('--no-images');
        }
        if (!preserveFormatting) {
          pythonArgs.push('--no-formatting');
        }
        if (!generateTOC) {
          pythonArgs.push('--no-toc');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOCX to EPUB batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOCX to EPUB batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOCX to EPUB batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOCX to EPUB batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:application/epub+zip;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOCX to EPUB batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOCX to EPUB batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOCX to EPUB batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOCX to MOBI (Single)
app.post('/convert/docx-to-mobi/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->MOBI single conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-mobi-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.mobi'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_mobi.py');
    console.log('DOCX to MOBI: Executing Python script:', scriptPath);
    console.log('DOCX to MOBI: Input file:', inputPath);
    console.log('DOCX to MOBI: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOCX to MOBI: Script exists');
    } catch (error) {
      console.error('DOCX to MOBI: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateTOC = req.body.generateTOC !== 'false';
    const kindleOptimized = req.body.kindleOptimized !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!includeImages) {
      pythonArgs.push('--no-images');
    }
    if (!preserveFormatting) {
      pythonArgs.push('--no-formatting');
    }
    if (!generateTOC) {
      pythonArgs.push('--no-toc');
    }
    if (!kindleOptimized) {
      pythonArgs.push('--no-kindle-optimize');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOCX to MOBI stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOCX to MOBI stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOCX to MOBI: Python script finished with code:', code);
      console.log('DOCX to MOBI: stdout:', stdout);
      console.log('DOCX to MOBI: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOCX to MOBI: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/x-mobipocket-ebook',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOCX to MOBI conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOCX to MOBI conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOCX to MOBI (Batch)
app.post('/convert/docx-to-mobi/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->MOBI batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-mobi-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateTOC = req.body.generateTOC !== 'false';
    const kindleOptimized = req.body.kindleOptimized !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.mobi'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_mobi.py');
        console.log('DOCX to MOBI batch: Executing Python script:', scriptPath);
        console.log('DOCX to MOBI batch: Input file:', inputPath);
        console.log('DOCX to MOBI batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOCX to MOBI batch: Script exists');
        } catch (error) {
          console.error('DOCX to MOBI batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!includeImages) {
          pythonArgs.push('--no-images');
        }
        if (!preserveFormatting) {
          pythonArgs.push('--no-formatting');
        }
        if (!generateTOC) {
          pythonArgs.push('--no-toc');
        }
        if (!kindleOptimized) {
          pythonArgs.push('--no-kindle-optimize');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOCX to MOBI batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOCX to MOBI batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOCX to MOBI batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOCX to MOBI batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:application/x-mobipocket-ebook;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOCX to MOBI batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOCX to MOBI batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOCX to MOBI batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOCX to ODT (Single)
app.post('/convert/docx-to-odt/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->ODT single conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-odt-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.odt'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_odt.py');
    console.log('DOCX to ODT: Executing Python script:', scriptPath);
    console.log('DOCX to ODT: Input file:', inputPath);
    console.log('DOCX to ODT: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOCX to ODT: Script exists');
    } catch (error) {
      console.error('DOCX to ODT: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const preserveFormatting = req.body.preserveFormatting !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!preserveFormatting) {
      pythonArgs.push('--no-formatting');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOCX to ODT stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOCX to ODT stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOCX to ODT: Python script finished with code:', code);
      console.log('DOCX to ODT: stdout:', stdout);
      console.log('DOCX to ODT: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOCX to ODT: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/vnd.oasis.opendocument.text',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOCX to ODT conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOCX to ODT conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOCX to ODT (Batch)
app.post('/convert/docx-to-odt/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->ODT batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-odt-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const preserveFormatting = req.body.preserveFormatting !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.odt'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_odt.py');
        console.log('DOCX to ODT batch: Executing Python script:', scriptPath);
        console.log('DOCX to ODT batch: Input file:', inputPath);
        console.log('DOCX to ODT batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOCX to ODT batch: Script exists');
        } catch (error) {
          console.error('DOCX to ODT batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!preserveFormatting) {
          pythonArgs.push('--no-formatting');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOCX to ODT batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOCX to ODT batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOCX to ODT batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOCX to ODT batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:application/vnd.oasis.opendocument.text;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOCX to ODT batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOCX to ODT batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOCX to ODT batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOCX to TXT (Single)
app.post('/convert/docx-to-txt/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->TXT single conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-txt-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.txt'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_txt.py');
    console.log('DOCX to TXT: Executing Python script:', scriptPath);
    console.log('DOCX to TXT: Input file:', inputPath);
    console.log('DOCX to TXT: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOCX to TXT: Script exists');
    } catch (error) {
      console.error('DOCX to TXT: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const preserveLineBreaks = req.body.preserveLineBreaks !== 'false';
    const removeFormatting = req.body.removeFormatting !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!preserveLineBreaks) {
      pythonArgs.push('--no-line-breaks');
    }
    if (!removeFormatting) {
      pythonArgs.push('--keep-formatting');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOCX to TXT stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOCX to TXT stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOCX to TXT: Python script finished with code:', code);
      console.log('DOCX to TXT: stdout:', stdout);
      console.log('DOCX to TXT: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOCX to TXT: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'text/plain; charset=utf-8',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOCX to TXT conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOCX to TXT conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOCX to TXT (Batch)
app.post('/convert/docx-to-txt/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOCX->TXT batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `docx-txt-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const preserveLineBreaks = req.body.preserveLineBreaks !== 'false';
    const removeFormatting = req.body.removeFormatting !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.docx$/i, '.txt'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'docx_to_txt.py');
        console.log('DOCX to TXT batch: Executing Python script:', scriptPath);
        console.log('DOCX to TXT batch: Input file:', inputPath);
        console.log('DOCX to TXT batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOCX to TXT batch: Script exists');
        } catch (error) {
          console.error('DOCX to TXT batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!preserveLineBreaks) {
          pythonArgs.push('--no-line-breaks');
        }
        if (!removeFormatting) {
          pythonArgs.push('--keep-formatting');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOCX to TXT batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOCX to TXT batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOCX to TXT batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOCX to TXT batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:text/plain;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOCX to TXT batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOCX to TXT batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOCX to TXT batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOC to CSV (Single)
app.post('/convert/doc-to-csv/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->CSV single conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-csv-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.csv'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_csv.py');
    console.log('DOC to CSV: Executing Python script:', scriptPath);
    console.log('DOC to CSV: Input file:', inputPath);
    console.log('DOC to CSV: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOC to CSV: Script exists');
    } catch (error) {
      console.error('DOC to CSV: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const delimiter = req.body.delimiter || ',';
    const extractTables = req.body.extractTables !== 'false';
    const includeParagraphs = req.body.includeParagraphs !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath, '--delimiter', delimiter];
    if (!extractTables) {
      pythonArgs.push('--no-tables');
    }
    if (!includeParagraphs) {
      pythonArgs.push('--no-paragraphs');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOC to CSV stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOC to CSV stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOC to CSV: Python script finished with code:', code);
      console.log('DOC to CSV: stdout:', stdout);
      console.log('DOC to CSV: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOC to CSV: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'text/csv; charset=utf-8',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOC to CSV conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOC to CSV conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOC to CSV (Batch)
app.post('/convert/doc-to-csv/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->CSV batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-csv-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const delimiter = req.body.delimiter || ',';
    const extractTables = req.body.extractTables !== 'false';
    const includeParagraphs = req.body.includeParagraphs !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.csv'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_csv.py');
        console.log('DOC to CSV batch: Executing Python script:', scriptPath);
        console.log('DOC to CSV batch: Input file:', inputPath);
        console.log('DOC to CSV batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOC to CSV batch: Script exists');
        } catch (error) {
          console.error('DOC to CSV batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath, '--delimiter', delimiter];
        if (!extractTables) {
          pythonArgs.push('--no-tables');
        }
        if (!includeParagraphs) {
          pythonArgs.push('--no-paragraphs');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOC to CSV batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOC to CSV batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOC to CSV batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOC to CSV batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:text/csv;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOC to CSV batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOC to CSV batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOC to CSV batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOC to EPUB (Single)
app.post('/convert/doc-to-epub/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->EPUB single conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-epub-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.epub'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_epub.py');
    console.log('DOC to EPUB: Executing Python script:', scriptPath);
    console.log('DOC to EPUB: Input file:', inputPath);
    console.log('DOC to EPUB: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOC to EPUB: Script exists');
    } catch (error) {
      console.error('DOC to EPUB: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateToc = req.body.generateToc !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!includeImages) {
      pythonArgs.push('--no-images');
    }
    if (!preserveFormatting) {
      pythonArgs.push('--no-formatting');
    }
    if (!generateToc) {
      pythonArgs.push('--no-toc');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOC to EPUB stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOC to EPUB stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOC to EPUB: Python script finished with code:', code);
      console.log('DOC to EPUB: stdout:', stdout);
      console.log('DOC to EPUB: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOC to EPUB: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/epub+zip',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOC to EPUB conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOC to EPUB conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOC to EPUB (Batch)
app.post('/convert/doc-to-epub/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->EPUB batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-epub-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateToc = req.body.generateToc !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.epub'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_epub.py');
        console.log('DOC to EPUB batch: Executing Python script:', scriptPath);
        console.log('DOC to EPUB batch: Input file:', inputPath);
        console.log('DOC to EPUB batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOC to EPUB batch: Script exists');
        } catch (error) {
          console.error('DOC to EPUB batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!includeImages) {
          pythonArgs.push('--no-images');
        }
        if (!preserveFormatting) {
          pythonArgs.push('--no-formatting');
        }
        if (!generateToc) {
          pythonArgs.push('--no-toc');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOC to EPUB batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOC to EPUB batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOC to EPUB batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOC to EPUB batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:application/epub+zip;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOC to EPUB batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOC to EPUB batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOC to EPUB batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOC to MOBI (Single)
app.post('/convert/doc-to-mobi/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->MOBI single conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-mobi-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.mobi'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_mobi.py');
    console.log('DOC to MOBI: Executing Python script:', scriptPath);
    console.log('DOC to MOBI: Input file:', inputPath);
    console.log('DOC to MOBI: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOC to MOBI: Script exists');
    } catch (error) {
      console.error('DOC to MOBI: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateToc = req.body.generateToc !== 'false';
    const kindleOptimized = req.body.kindleOptimized !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!includeImages) {
      pythonArgs.push('--no-images');
    }
    if (!preserveFormatting) {
      pythonArgs.push('--no-formatting');
    }
    if (!generateToc) {
      pythonArgs.push('--no-toc');
    }
    if (!kindleOptimized) {
      pythonArgs.push('--no-kindle-optimize');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOC to MOBI stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOC to MOBI stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOC to MOBI: Python script finished with code:', code);
      console.log('DOC to MOBI: stdout:', stdout);
      console.log('DOC to MOBI: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOC to MOBI: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'application/x-mobipocket-ebook',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOC to MOBI conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOC to MOBI conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOC to MOBI (Batch)
app.post('/convert/doc-to-mobi/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->MOBI batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-mobi-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const includeImages = req.body.includeImages !== 'false';
    const preserveFormatting = req.body.preserveFormatting !== 'false';
    const generateToc = req.body.generateToc !== 'false';
    const kindleOptimized = req.body.kindleOptimized !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.mobi'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_mobi.py');
        console.log('DOC to MOBI batch: Executing Python script:', scriptPath);
        console.log('DOC to MOBI batch: Input file:', inputPath);
        console.log('DOC to MOBI batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOC to MOBI batch: Script exists');
        } catch (error) {
          console.error('DOC to MOBI batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!includeImages) {
          pythonArgs.push('--no-images');
        }
        if (!preserveFormatting) {
          pythonArgs.push('--no-formatting');
        }
        if (!generateToc) {
          pythonArgs.push('--no-toc');
        }
        if (!kindleOptimized) {
          pythonArgs.push('--no-kindle-optimize');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOC to MOBI batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOC to MOBI batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOC to MOBI batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOC to MOBI batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:application/x-mobipocket-ebook;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOC to MOBI batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOC to MOBI batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOC to MOBI batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: DOC to TXT (Single)
app.post('/convert/doc-to-txt/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->TXT single conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-txt-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.txt'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_txt.py');
    console.log('DOC to TXT: Executing Python script:', scriptPath);
    console.log('DOC to TXT: Input file:', inputPath);
    console.log('DOC to TXT: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('DOC to TXT: Script exists');
    } catch (error) {
      console.error('DOC to TXT: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const preserveLineBreaks = req.body.preserveLineBreaks !== 'false';
    const removeFormatting = req.body.removeFormatting !== 'false';

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (!preserveLineBreaks) {
      pythonArgs.push('--no-line-breaks');
    }
    if (removeFormatting === false) {
      pythonArgs.push('--keep-formatting');
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('DOC to TXT stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('DOC to TXT stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('DOC to TXT: Python script finished with code:', code);
      console.log('DOC to TXT: stdout:', stdout);
      console.log('DOC to TXT: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('DOC to TXT: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'text/plain; charset=utf-8',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('DOC to TXT conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('DOC to TXT conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: DOC to TXT (Batch)
app.post('/convert/doc-to-txt/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('DOC->TXT batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `doc-txt-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const preserveLineBreaks = req.body.preserveLineBreaks !== 'false';
    const removeFormatting = req.body.removeFormatting !== 'false';

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.doc$/i, '.txt'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'doc_to_txt.py');
        console.log('DOC to TXT batch: Executing Python script:', scriptPath);
        console.log('DOC to TXT batch: Input file:', inputPath);
        console.log('DOC to TXT batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('DOC to TXT batch: Script exists');
        } catch (error) {
          console.error('DOC to TXT batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (!preserveLineBreaks) {
          pythonArgs.push('--no-line-breaks');
        }
        if (removeFormatting === false) {
          pythonArgs.push('--keep-formatting');
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('DOC to TXT batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('DOC to TXT batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('DOC to TXT batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('DOC to TXT batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:text/plain;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('DOC to TXT batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('DOC to TXT batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('DOC to TXT batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: HEIC to SVG (Single)
app.post('/convert/heic-to-svg/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('HEIC->SVG single conversion request');

  const tmpDir = path.join(os.tmpdir(), `heic-svg-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.svg'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_svg.py');
    console.log('HEIC to SVG: Executing Python script:', scriptPath);
    console.log('HEIC to SVG: Input file:', inputPath);
    console.log('HEIC to SVG: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('HEIC to SVG: Script exists');
    } catch (error) {
      console.error('HEIC to SVG: Script does not exist:', scriptPath);
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const quality = parseInt(req.body.quality) || 95;
    const preserveTransparency = req.body.preserveTransparency !== 'false';
    // Use 4096 as default max dimension for faster conversion (can be adjusted)
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (quality !== 95) {
      pythonArgs.push('--quality', quality.toString());
    }
    if (!preserveTransparency) {
      pythonArgs.push('--no-transparency');
    }
    if (maxDimension !== 8192) {
      pythonArgs.push('--max-dimension', maxDimension.toString());
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('HEIC to SVG stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('HEIC to SVG stderr:', data.toString());
    });

    python.on('close', async (code: number) => {
      console.log('HEIC to SVG: Python script finished with code:', code);
      console.log('HEIC to SVG: stdout:', stdout);
      console.log('HEIC to SVG: stderr:', stderr);
      
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          console.log('HEIC to SVG: Output file size:', outputBuffer.length);
          res.set({
            'Content-Type': 'image/svg+xml',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
          
        } else {
          console.error('HEIC to SVG conversion failed. Code:', code, 'Stderr:', stderr);
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('HEIC to SVG conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: HEIC to SVG (Batch)
app.post('/convert/heic-to-svg/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('HEIC->SVG batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `heic-svg-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const quality = parseInt(req.body.quality) || 95;
    const preserveTransparency = req.body.preserveTransparency !== 'false';
    // Use 4096 as default max dimension for faster conversion (can be adjusted)
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.svg'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_svg.py');
        console.log('HEIC to SVG batch: Executing Python script:', scriptPath);
        console.log('HEIC to SVG batch: Input file:', inputPath);
        console.log('HEIC to SVG batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('HEIC to SVG batch: Script exists');
        } catch (error) {
          console.error('HEIC to SVG batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (quality !== 95) {
          pythonArgs.push('--quality', quality.toString());
        }
        if (!preserveTransparency) {
          pythonArgs.push('--no-transparency');
        }
        if (maxDimension !== 8192) {
          pythonArgs.push('--max-dimension', maxDimension.toString());
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('HEIC to SVG batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('HEIC to SVG batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('HEIC to SVG batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('HEIC to SVG batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:image/svg+xml;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('HEIC to SVG batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('HEIC to SVG batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.json({ success: true, results });
    
  } catch (error) {
    console.error('HEIC to SVG batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: HEIC to PDF (Single) - OPTIONS for CORS preflight
app.options('/convert/heic-to-pdf/single', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to PDF (Single)
app.post('/convert/heic-to-pdf/single', upload.single('file'), async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('HEIC->PDF single conversion request');

  const tmpDir = path.join(os.tmpdir(), `heic-pdf-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No file uploaded' });
    }
    try {
      await fs.mkdir(tmpDir, { recursive: true });
    } catch (mkdirError) {
      console.error('HEIC to PDF: Failed to create temp directory:', mkdirError);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(500).json({ error: 'Failed to create temporary directory', details: mkdirError instanceof Error ? mkdirError.message : 'Unknown error' });
    }

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.pdf'));

    try {
      await fs.writeFile(inputPath, file.buffer);
    } catch (writeError) {
      console.error('HEIC to PDF: Failed to write input file:', writeError);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      return res.status(500).json({ error: 'Failed to write input file', details: writeError instanceof Error ? writeError.message : 'Unknown error' });
    }

    const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_pdf.py');
    console.log('HEIC to PDF: Executing Python script:', scriptPath);
    console.log('HEIC to PDF: Input file:', inputPath);
    console.log('HEIC to PDF: Output file:', outputPath);
    
    // Check if script exists
    try {
      await fs.access(scriptPath);
      console.log('HEIC to PDF: Script exists');
    } catch (error) {
      console.error('HEIC to PDF: Script does not exist:', scriptPath);
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    // Get options from request body
    const quality = parseInt(req.body.quality) || 95;
    const pageSize = req.body.pageSize || 'auto';
    const fitToPage = req.body.fitToPage !== 'false';
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    const pythonArgs = [scriptPath, inputPath, outputPath];
    if (quality !== 95) {
      pythonArgs.push('--quality', quality.toString());
    }
    if (pageSize !== 'auto') {
      pythonArgs.push('--page-size', pageSize);
    }
    if (!fitToPage) {
      pythonArgs.push('--no-fit-to-page');
    }
    if (maxDimension !== 8192) {
      pythonArgs.push('--max-dimension', maxDimension.toString());
    }

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '';
    let stderr = '';

    // Handle spawn errors (if process fails to start)
    python.on('error', async (error: Error) => {
      console.error('HEIC to PDF: Failed to start Python process:', error);
      if (!res.headersSent) {
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Failed to start conversion process', details: error.message });
      }
      await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
    });

    python.stdout.on('data', (data: Buffer) => {
      stdout += data.toString();
      console.log('HEIC to PDF stdout:', data.toString());
    });

    python.stderr.on('data', (data: Buffer) => {
      stderr += data.toString();
      console.log('HEIC to PDF stderr:', data.toString());
    });

    // Set a timeout for the conversion (5 minutes)
    const timeout = setTimeout(async () => {
      console.error('HEIC to PDF: Conversion timeout after 5 minutes');
      python.kill();
      if (!res.headersSent) {
        res.set({
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
        });
        res.status(500).json({ error: 'Conversion timeout. The file may be too large or complex.' });
      }
      await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
    }, 5 * 60 * 1000);

    python.on('close', async (code: number) => {
      clearTimeout(timeout);
      console.log('HEIC to PDF: Python script finished with code:', code);
      console.log('HEIC to PDF: stdout:', stdout);
      console.log('HEIC to PDF: stderr:', stderr);
      
      try {
        if (!res.headersSent) {
          if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
            const outputBuffer = await fs.readFile(outputPath);
            console.log('HEIC to PDF: Output file size:', outputBuffer.length);
            res.set({
              'Content-Type': 'application/pdf',
              'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
            });
            res.send(outputBuffer);
            
          } else {
            console.error('HEIC to PDF conversion failed. Code:', code, 'Stderr:', stderr);
            res.set({
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
              'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
            });
            res.status(500).json({ error: 'Conversion failed', details: stderr });
          }
        }
      } catch (error) {
        console.error('Error handling conversion result:', error);
        if (!res.headersSent) {
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: error instanceof Error ? error.message : 'Unknown error' });
        }
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    console.error('HEIC to PDF conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: HEIC to PDF (Batch) - OPTIONS for CORS preflight
app.options('/convert/heic-to-pdf/batch', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to PDF (Batch)
app.post('/convert/heic-to-pdf/batch', uploadBatch, async (req, res) => {
  // Set CORS headers
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });
  
  console.log('HEIC->PDF batch conversion request');

  const tmpDir = path.join(os.tmpdir(), `heic-pdf-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No files uploaded' });
    }
    await fs.mkdir(tmpDir, { recursive: true });

    const results = [];

    // Get options from request body
    const quality = parseInt(req.body.quality) || 95;
    const pageSize = req.body.pageSize || 'auto';
    const fitToPage = req.body.fitToPage !== 'false';
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.pdf'));

        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_pdf.py');
        console.log('HEIC to PDF batch: Executing Python script:', scriptPath);
        console.log('HEIC to PDF batch: Input file:', inputPath);
        console.log('HEIC to PDF batch: Output file:', outputPath);

        try {
          await fs.access(scriptPath);
          console.log('HEIC to PDF batch: Script exists');
        } catch (error) {
          console.error('HEIC to PDF batch: Script does not exist:', scriptPath);
          results.push({
            originalName: file.originalname,
            outputFilename: '',
            size: 0,
            success: false,
            error: 'Conversion script not found'
          });
          continue;
        }

        const pythonArgs = [scriptPath, inputPath, outputPath];
        if (quality !== 95) {
          pythonArgs.push('--quality', quality.toString());
        }
        if (pageSize !== 'auto') {
          pythonArgs.push('--page-size', pageSize);
        }
        if (!fitToPage) {
          pythonArgs.push('--no-fit-to-page');
        }
        if (maxDimension !== 8192) {
          pythonArgs.push('--max-dimension', maxDimension.toString());
        }

        const python = spawn('/opt/venv/bin/python', pythonArgs);

        let stdout = '';
        let stderr = '';

        python.stdout.on('data', (data: Buffer) => {
          stdout += data.toString();
          console.log('HEIC to PDF batch stdout:', data.toString());
        });

        python.stderr.on('data', (data: Buffer) => {
          stderr += data.toString();
          console.log('HEIC to PDF batch stderr:', data.toString());
        });

        await new Promise<void>((resolve, reject) => {
          python.on('close', async (code: number) => {
            console.log('HEIC to PDF batch: Python script finished with code:', code);
            
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                console.log('HEIC to PDF batch: Output file size:', outputBuffer.length);
                results.push({
                  originalName: file.originalname,
                  outputFilename: path.basename(outputPath),
                  size: outputBuffer.length,
                  success: true,
                  downloadPath: `data:application/pdf;base64,${outputBuffer.toString('base64')}`
                });
                resolve();
              } else {
                console.error('HEIC to PDF batch conversion failed. Code:', code, 'Stderr:', stderr);
                results.push({
                  originalName: file.originalname,
                  outputFilename: '',
                  size: 0,
                  success: false,
                  error: stderr || `Conversion failed with code ${code}`
                });
                resolve(); // Continue with other files
              }
            } catch (error) {
              console.error('Error handling batch conversion result:', error);
              results.push({
                originalName: file.originalname,
                outputFilename: '',
                size: 0,
                success: false,
                error: error instanceof Error ? error.message : 'Unknown error'
              });
              resolve(); // Continue with other files
            }
          });
        });
      } catch (error) {
        console.error('HEIC to PDF batch conversion error for file:', file.originalname, error);
        results.push({
          originalName: file.originalname,
          outputFilename: '',
          size: 0,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.json({ success: true, results });
    
  } catch (error) {
    console.error('HEIC to PDF batch conversion error:', error);
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Route: HEIC to WEBP (Single) - OPTIONS for CORS preflight
app.options('/convert/heic-to-webp/single', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to WEBP (Single)
app.post('/convert/heic-to-webp/single', upload.single('file'), async (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('HEIC->WEBP single conversion request');

  const tmpDir = path.join(os.tmpdir(), `heic-webp-${Date.now()}`);

  try {
    const file = req.file;
    if (!file) {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(400).json({ error: 'No file uploaded' });
    }

    await fs.mkdir(tmpDir, { recursive: true });

    const inputPath = path.join(tmpDir, file.originalname);
    const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.webp'));

    await fs.writeFile(inputPath, file.buffer);

    const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_webp.py');
    try { await fs.access(scriptPath); } catch {
      res.set({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
      });
      return res.status(500).json({ error: 'Conversion script not found' });
    }

    const quality = parseInt(req.body.quality) || 90;
    const lossless = req.body.lossless === 'true';
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', String(quality), '--max-dimension', String(maxDimension)];
    if (lossless) pythonArgs.push('--lossless');

    const python = spawn('/opt/venv/bin/python', pythonArgs);

    let stdout = '', stderr = '';
    python.stdout.on('data', (d: Buffer) => { stdout += d.toString(); });
    python.stderr.on('data', (d: Buffer) => { stderr += d.toString(); });

    python.on('close', async (code: number) => {
      try {
        if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
          const outputBuffer = await fs.readFile(outputPath);
          res.set({
            'Content-Type': 'image/webp',
            'Content-Disposition': `attachment; filename="${path.basename(outputPath)}"`,
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.send(outputBuffer);
        } else {
          res.set({
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
          });
          res.status(500).json({ error: 'Conversion failed', details: stderr });
        }
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
      }
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  }
});

// Route: HEIC to WEBP (Batch) - OPTIONS for CORS preflight
app.options('/convert/heic-to-webp/batch', (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
    'Access-Control-Max-Age': '86400'
  });
  res.sendStatus(200);
});

// Route: HEIC to WEBP (Batch)
app.post('/convert/heic-to-webp/batch', uploadBatch, async (req, res) => {
  res.set({
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
  });

  console.log('HEIC->WEBP batch conversion request');
  const tmpDir = path.join(os.tmpdir(), `heic-webp-batch-${Date.now()}`);

  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    await fs.mkdir(tmpDir, { recursive: true });

    const results: any[] = [];
    const quality = parseInt(req.body.quality) || 90;
    const lossless = req.body.lossless === 'true';
    const maxDimension = parseInt(req.body.maxDimension) || 4096;

    for (const file of files) {
      try {
        const inputPath = path.join(tmpDir, file.originalname);
        const outputPath = path.join(tmpDir, file.originalname.replace(/\.(heic|heif)$/i, '.webp'));
        await fs.writeFile(inputPath, file.buffer);

        const scriptPath = path.join(__dirname, '..', 'scripts', 'heic_to_webp.py');
        try { await fs.access(scriptPath); } catch { results.push({ originalName: file.originalname, outputFilename: '', size: 0, success: false, error: 'Conversion script not found' }); continue; }

        const pythonArgs = [scriptPath, inputPath, outputPath, '--quality', String(quality), '--max-dimension', String(maxDimension)];
        if (lossless) pythonArgs.push('--lossless');

        const python = spawn('/opt/venv/bin/python', pythonArgs);
        let stdout = '', stderr = '';
        python.stdout.on('data', (d: Buffer) => { stdout += d.toString(); });
        python.stderr.on('data', (d: Buffer) => { stderr += d.toString(); });

        await new Promise<void>((resolve) => {
          python.on('close', async (code: number) => {
            try {
              if (code === 0 && await fs.access(outputPath).then(() => true).catch(() => false)) {
                const outputBuffer = await fs.readFile(outputPath);
                results.push({ originalName: file.originalname, outputFilename: path.basename(outputPath), size: outputBuffer.length, success: true, downloadPath: `data:image/webp;base64,${outputBuffer.toString('base64')}` });
              } else {
                results.push({ originalName: file.originalname, outputFilename: '', size: 0, success: false, error: stderr || `Conversion failed with code ${code}` });
              }
            } finally {
              resolve();
            }
          });
        });
      } catch (err) {
        results.push({ originalName: file.originalname, outputFilename: '', size: 0, success: false, error: err instanceof Error ? err.message : 'Unknown error' });
      }
    }

    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.json({ success: true, results });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    res.set({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept'
    });
    res.status(500).json({ error: message });
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => undefined);
  }
});

// Start server
const server = app.listen(PORT, '0.0.0.0', () => {
  console.log(`Morpy backend running on port ${PORT}`);
  console.log('✅ Server started successfully');
});

// Increase timeout for large file processing (15 minutes for large CSV files)
server.timeout = 15 * 60 * 1000; // 15 minutes
server.keepAliveTimeout = 15 * 60 * 1000; // 15 minutes
server.headersTimeout = 15 * 60 * 1000 + 1000; // 15 minutes + 1 second

// Graceful shutdown
const gracefulShutdown = (signal: string) => {
  console.log(`\n${signal} received. Shutting down gracefully...`);
  
  server.close(() => {
    console.log('✅ HTTP server closed');
    process.exit(0);
  });
};

process.on('SIGINT', () => gracefulShutdown('SIGINT'));
process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));

export default app;

