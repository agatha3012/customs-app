const http = require('http');
const fs = require('fs');
const path = require('path');

const filePath = 'D:/桌面/7-14凡洋德国陆运包税FBA不带电20箱发票.xlsx';
const fileName = path.basename(filePath);
const fileContent = fs.readFileSync(filePath);

const boundary = '----Boundary' + Math.random().toString(36).slice(2);
const parts = [];
parts.push(Buffer.from('--' + boundary + '\r\n'));
parts.push(Buffer.from('Content-Disposition: form-data; name="file"; filename="' + fileName + '"\r\n'));
parts.push(Buffer.from('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n'));
parts.push(fileContent);
parts.push(Buffer.from('\r\n--' + boundary + '--\r\n'));
const body = Buffer.concat(parts);

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
    const result = JSON.parse(data);
    console.log('Success:', result.success);
    console.log('Products:', result.products ? result.products.length : 0);
    console.log('Boxes:', result.boxes ? result.boxes.length : 0);
    console.log('');
    if (result.products) {
      result.products.forEach(p => {
        console.log('  ' + p.sku + ' | ' + p.nameCN + ' | qty=' + p.totalQty + ' | hs=' + p.hsCode + ' | mat=' + p.material);
      });
    }
    if (result.boxes) {
      console.log('\nBoxes:');
      result.boxes.forEach(b => {
        console.log('  Box ' + b.boxSeq + ': ' + b.id + ' | weight=' + b.weight + ' | products=' + b.products.length);
        b.products.forEach(bp => {
          console.log('    - ' + bp.sku + ' | ' + bp.nameCN + ' | qty=' + bp.qtyPerBox + ' | eur=' + bp.unitPriceEUR);
        });
      });
    }
  });
});
req.write(body);
req.end();
