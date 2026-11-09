// Рендер HTML-схем в PNG: node render_diagrams.js <html_dir> <png_dir>
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const [src, dst] = process.argv.slice(2);
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 1200 }, deviceScaleFactor: 1 });
  for (const n of ['windows_nt', 'linux', 'minix3']) {
    await page.goto('file://' + path.resolve(src, n + '.html'));
    await page.locator('#cap').screenshot({ path: path.resolve(dst, n + '.png') });
  }
  await browser.close();
})();
