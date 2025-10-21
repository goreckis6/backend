#!/usr/bin/env node

import http from 'http';
import https from 'https';

console.log('Testing CORS configuration...');

// Test the CORS endpoint
const options = {
  hostname: 'api.morphyimg.com',
  port: 443,
  path: '/convert/cr2-to-ico/single',
  method: 'OPTIONS',
  headers: {
    'Origin': 'https://morphyimg.com',
    'Access-Control-Request-Method': 'POST',
    'Access-Control-Request-Headers': 'Content-Type'
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

req.end();
