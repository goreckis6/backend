#!/usr/bin/env node

import https from 'https';

console.log('Testing CR2 to ICO batch endpoint...');

// Test OPTIONS request first
const optionsReq = {
  hostname: 'api.morphyimg.com',
  port: 443,
  path: '/convert/cr2-to-ico/batch',
  method: 'OPTIONS',
  headers: {
    'Origin': 'https://morphyimg.com',
    'Access-Control-Request-Method': 'POST',
    'Access-Control-Request-Headers': 'Content-Type'
  }
};

console.log('Testing OPTIONS request...');
const req1 = https.request(optionsReq, (res) => {
  console.log('OPTIONS Status:', res.statusCode);
  console.log('OPTIONS CORS Headers:');
  console.log('- Access-Control-Allow-Origin:', res.headers['access-control-allow-origin']);
  console.log('- Access-Control-Allow-Methods:', res.headers['access-control-allow-methods']);
  console.log('- Access-Control-Allow-Headers:', res.headers['access-control-allow-headers']);
  
  // Now test POST request
  console.log('\nTesting POST request...');
  const postData = 'files=test1&files=test2';
  
  const postOptions = {
    hostname: 'api.morphyimg.com',
    port: 443,
    path: '/convert/cr2-to-ico/batch',
    method: 'POST',
    headers: {
      'Origin': 'https://morphyimg.com',
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(postData)
    }
  };
  
  const req2 = https.request(postOptions, (res) => {
    console.log('POST Status:', res.statusCode);
    console.log('POST CORS Headers:');
    console.log('- Access-Control-Allow-Origin:', res.headers['access-control-allow-origin']);
    console.log('- Access-Control-Allow-Methods:', res.headers['access-control-allow-methods']);
    console.log('- Access-Control-Allow-Headers:', res.headers['access-control-allow-headers']);
    
    let data = '';
    res.on('data', (chunk) => {
      data += chunk;
    });
    
    res.on('end', () => {
      console.log('POST Response body:', data);
    });
  });
  
  req2.on('error', (error) => {
    console.error('POST Request error:', error);
  });
  
  req2.write(postData);
  req2.end();
});

req1.on('error', (error) => {
  console.error('OPTIONS Request error:', error);
});

req1.end();
