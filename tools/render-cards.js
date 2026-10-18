// Renders scorecard pages (JSON from the Orexis app) to printable A4 PDFs.
const fs = require('fs'), { chromium } = require('playwright');
(async () => {
  const j = JSON.parse(fs.readFileSync(process.argv[2], 'utf8')), out = process.argv[3] || 'out';
  fs.mkdirSync(out, { recursive: true });
  const b = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {}); const pg = await b.newPage();
  const pdfs = [];
  for (const c of j.cards) {
    await pg.setContent(c.html, { waitUntil: 'load' });
    const buf = await pg.pdf({ format: 'A4', printBackground: true, margin: { top: '12mm', bottom: '12mm', left: '12mm', right: '12mm' } });
    fs.writeFileSync(out + '/' + c.name + '.pdf', buf); pdfs.push({ name: c.name, b64: buf.toString('base64') });
  }
  await b.close();
  fs.writeFileSync(out + '/payload.json', JSON.stringify({ tag: j.tag, pdfs }));
  console.log('rendered', pdfs.length, j.tag);
})();
