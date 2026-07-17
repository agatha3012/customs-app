const http = require('http');
const fs = require('fs');
const XLSX = require('xlsx');

function apiPost(path, body) {
  return new Promise((resolve, reject) => {
    const json = JSON.stringify(body);
    const req = http.request({
      hostname: 'localhost', port: 3000, path, method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(json) },
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch(e) { reject(new Error('Parse error: ' + data.substring(0,500))); }
      });
    });
    req.on('error', reject);
    req.write(json);
    req.end();
  });
}

async function main() {
  console.log('=== E2E Test: Customs File Generation ===\n');

  // Step 1: Upload was already done in previous test (last_parsed.json exists)
  // Let's read last_parsed.json to get the invoice data
  const parsed = JSON.parse(fs.readFileSync('uploads/last_parsed.json', 'utf8'));
  console.log('1. Invoice products:');
  parsed.products.forEach(p => {
    console.log(`   ${p.sku} | ${p.nameCN} | qty=${p.totalQty} | bestPriceEUR=${p.bestPriceEUR}`);
  });
  console.log(`   Total boxes: ${parsed.totalBoxes}\n`);

  // Step 2: Lookup products in DB
  const skus = parsed.products.map(p => p.sku);
  console.log('2. Looking up ' + skus.length + ' SKUs in product database...');
  const lookup = await apiPost('/api/lookup-products', { skus });
  console.log('   Found:', Object.keys(lookup.found).length);
  console.log('   Not found:', lookup.notFound.join(', ') || 'none');

  // Show results
  Object.entries(lookup.found).forEach(([sku, info]) => {
    console.log(`   ${sku}: maxPrice=¥${info.maxPrice.toFixed(2)}, supplier=${info.supplier}, city=${info.city}, unitPriceUSD=$${info.unitPriceUSD.toFixed(5)}`);
  });
  console.log('');

  // Step 3: Confirm locations (use auto-detected cities)
  const locations = {};
  Object.entries(lookup.found).forEach(([sku, info]) => {
    if (!info.cityUncertain) locations[sku] = info.city;
  });

  // For uncertain ones, use web search results (simulate)
  const uncertain = Object.entries(lookup.found).filter(([,info]) => info.cityUncertain);
  if (uncertain.length > 0) {
    console.log('3. Uncertain locations:');
    uncertain.forEach(([sku, info]) => {
      console.log(`   ${sku}: supplier="${info.rawSupplier||info.supplier}"`);
    });
    // Use raw supplier name as-is
    uncertain.forEach(([sku, info]) => {
      locations[sku] = info.rawSupplier || info.supplier;
    });
  }
  console.log('');

  // Step 4: Generate the file
  console.log('4. Generating customs file...');
  const genResult = await apiPost('/api/generate', { confirmedLocations: locations });
  console.log('   Success:', genResult.success);
  console.log('   File:', genResult.fileName);
  console.log('   URL:', genResult.downloadUrl);
  console.log('');

  // Step 5: Verify generated file
  if (genResult.success) {
    const outputPath = 'output/' + genResult.fileName;
    console.log('5. Verifying generated file: ' + outputPath);
    const wb = XLSX.readFile(outputPath);
    console.log('   Sheets: ' + wb.SheetNames.join(', '));
    wb.SheetNames.forEach(sn => {
      const ws = wb.Sheets[sn];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      console.log(`   ${sn}: ${data.length} rows`);
    });
    console.log('\n=== E2E Test PASSED ===');
  } else {
    console.log('FAILED:', genResult.message);
  }
}

main().catch(e => console.error('E2E Test Error:', e));
