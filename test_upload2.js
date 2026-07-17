const http = require('http');
const fs = require('fs');
const path = require('path');

const filePath = 'D:/桌面/7-14凡洋德国陆运包税FBA不带电20箱发票.xlsx';
const fileName = encodeURIComponent(path.basename(filePath));
const fileContent = fs.readFileSync(filePath);

const boundary = '----Boundary' + Math.random().toString(36).slice(2);

const header = [
  '--' + boundary,
  'Content-Disposition: form-data; name="file"; filename="' + fileName + '"',
  'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '', ''
].join('\r\n');

const footer = '\r\n--' + boundary + '--\r\n';

const headerBuf = Buffer.from(header, 'utf8');
const footerBuf = Buffer.from(footer, 'utf8');
const body = Buffer.concat([headerBuf, fileContent, footerBuf]);

console.log('File size:', fileContent.length);
console.log('Body size:', body.length);

const req = http.request({
  hostname: 'localhost', port: 3000, path: '/api/upload-invoice', method: 'POST',
  headers: {
    'Content-Type': 'multipart/form-data; boundary=' + boundary,
    'Content-Length': body.length,
  },
}, (res) => {
  let data = '';
  res.on('data', c => data += c);
  res.on('end', () => {
    try {
      const result = JSON.parse(data);
      console.log('Response:', JSON.stringify(result, null, 2));
    } catch(e) {
      console.log('Raw response (' + data.length + ' chars):');
      console.log(data.substring(0, 2000));
    }
  });
});
req.on('error', e => console.error('Error:', e));
req.write(body);
req.end();
