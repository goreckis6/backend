import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Read the server.ts file
const filePath = path.join(__dirname, 'src', 'server.ts');
let content = fs.readFileSync(filePath, 'utf8');

// Pattern to match res.set() calls with CORS headers
const corsPattern = /res\.set\(\s*\{\s*['"]Access-Control-Allow-Origin['"]:\s*['"][^'"]*['"],\s*['"]Access-Control-Allow-Methods['"]:\s*['"][^'"]*['"],\s*['"]Access-Control-Allow-Headers['"]:\s*['"][^'"]*['"]\s*\}\s*\);/g;

// Remove all CORS res.set() calls
content = content.replace(corsPattern, '');

// Also remove any standalone CORS header lines
const corsHeaderPattern = /^\s*['"]Access-Control-Allow-[^'"]*['"]:\s*['"][^'"]*['"],?\s*$/gm;
content = content.replace(corsHeaderPattern, '');

// Clean up any empty res.set() calls
content = content.replace(/res\.set\(\s*\{\s*\}\s*\);/g, '');

// Clean up any empty res.set() calls with just whitespace
content = content.replace(/res\.set\(\s*\{\s*\n\s*\}\s*\);/g, '');

// Write the cleaned content back
fs.writeFileSync(filePath, content, 'utf8');

console.log('Removed all manual CORS headers. The centralized CORS middleware will handle all CORS requests.');
