// Forhåndsbilledet docs/billeder/social.png (1280×640) til delte links og GitHub › Settings › Social preview: hjemmesidens
// overskrift (med "Gratis" og "Ikke tilknyttet …") og skærmbilledet af Mine skemaer fra Windows. Skærmbillederne af selve
// appen laves af tests/AulaSync.App.Tests/Skaermbilleder.cs. Kræver Node og Playwright med Chromium:
//
//   node tests/site/billeder.js
const http = require("http");
const fs = require("fs");
const path = require("path");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "playwright");

const docs = path.join(__dirname, "..", "..", "docs");
const out = path.join(docs, "billeder");
const types = { ".html": "text/html; charset=utf-8", ".png": "image/png", ".woff2": "font/woff2" };

const server = http.createServer((req, res) => {
  const file = path.join(docs, decodeURIComponent(new URL(req.url, "http://x").pathname));
  if (!file.startsWith(docs) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { "Content-Type": types[path.extname(file)] || "application/octet-stream" });
  fs.createReadStream(file).pipe(res);
});

(async () => {
  await new Promise(r => server.listen(0, "127.0.0.1", r));
  const url = `http://127.0.0.1:${server.address().port}/index.html`;
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ viewport: { width: 1280, height: 640 }, colorScheme: "light", deviceScaleFactor: 1 });
    await page.goto(url, { waitUntil: "networkidle" });
    await page.evaluate(() => document.fonts.ready);
    await page.addStyleTag({ content: `
      .nav, .trust, main > section:not(.hero), footer, .cta, .meta { display: none !important; }
      .hero { height: 640px; border: 0; }
      .hero .wrap { max-width: none; height: 640px; padding: 0 56px; grid-template-columns: 0.92fr 1.08fr; gap: 40px; }
      .hero h1 { font-size: 66px; }
      .hero .lead { font-size: 20px; }
      .social-brand { display: flex; align-items: center; gap: 12px; font: 750 24px/1 var(--display); letter-spacing: -0.02em; }
      .social-brand img { width: 44px; height: 44px; }
      .social-url { font: 600 16px/1 var(--body); color: var(--muted); }` });
    await page.evaluate(() => {
      const copy = document.querySelector(".hero-copy");
      copy.insertAdjacentHTML("afterbegin", '<div class="social-brand"><img src="billeder/ikon.png" alt="">AulaSync</div>');
      copy.insertAdjacentHTML("beforeend", '<p class="social-url">rpaasch.github.io/AulaSync</p>');
    });
    await page.evaluate(() => Promise.all([...document.querySelectorAll(".hero img")].map(i => i.complete ? null : new Promise(r => { i.onload = r; }))));
    await page.screenshot({ path: path.join(out, "social.png"), clip: { x: 0, y: 0, width: 1280, height: 640 } });
    await page.close();
  } finally {
    await browser.close();
    server.close();
  }
})();
