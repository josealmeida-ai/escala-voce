// Gera apresentacao.pdf a partir de apresentacao.html usando o Chrome/Edge
// já instalado na máquina (via puppeteer-core, sem baixar Chromium extra).
// Uso: npm run pdf
const puppeteer = require("puppeteer-core");
const fs = require("fs");
const path = require("path");

const CHROME_CANDIDATES = [
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
];

const executablePath = CHROME_CANDIDATES.find((p) => fs.existsSync(p));
if (!executablePath) {
  console.error("Nenhum navegador Chrome/Edge encontrado nos caminhos padrão.");
  process.exit(1);
}

const HTML_PATH = path.join(__dirname, "..", "apresentacao.html");
const PDF_PATH = path.join(__dirname, "..", "apresentacao.pdf");

(async () => {
  const browser = await puppeteer.launch({ executablePath, headless: "new" });
  const page = await browser.newPage();
  await page.goto("file:///" + HTML_PATH.replace(/\\/g, "/"), { waitUntil: "networkidle0" });
  // dá um tempo extra para o Tailwind CDN (JIT) processar as classes
  await new Promise((r) => setTimeout(r, 1200));

  await page.pdf({
    path: PDF_PATH,
    printBackground: true,
    preferCSSPageSize: true,
  });

  await browser.close();
  console.log("PDF gerado em " + PDF_PATH);
})();
