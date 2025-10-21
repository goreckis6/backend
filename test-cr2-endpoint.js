#!/usr/bin/env node

import https from 'https';
import fs from 'fs';
import FormData from 'form-data';

console.log('Testing CR2 to ICO endpoint...');

// Create a simple test file (we'll use a dummy file for testing)
const testFile = Buffer.from('dummy cr2 content');
const form = new FormData();
form.append('file', testFile, { filename: 'test.cr2' });

const options = {
  hostname: 'api.morphyimg.com',
  port: 443,
  path: '/convert/cr2-to-ico/single',
  method: 'POST',
  headers: {
    'Origin': 'https://morphyimg.com',
    ...form.getHeaders()
  }
};

const req = https.request(options, (res) => {
  console.log('Status:', res.statusCode);
  console.log('Headers:', res.headers);
  
  let data = '';
  res.on('data', (chunk) => {
    data += chunk;
  });
  
  res.on('end', () => {
    console.log('Response body:', data);
    console.log('CORS Headers:');
    console.log('- Access-Control-Allow-Origin:', res.headers['access-control-allow-origin']);
    console.log('- Access-Control-Allow-Methods:', res.headers['access-control-allow-methods']);
    console.log('- Access-Control-Allow-Headers:', res.headers['access-control-allow-headers']);
  });
});

req.on('error', (error) => {
  console.error('Request error:', error);
});

form.pipe(req);
