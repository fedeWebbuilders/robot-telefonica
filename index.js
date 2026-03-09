import { chromium } from 'playwright';
import { createInterface } from 'readline';
import { readFileSync, existsSync, writeFileSync, mkdirSync } from 'fs';
import { resolve, normalize, join, dirname, basename, extname } from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';

const __dirname = dirname(fileURLToPath(import.meta.url));

// ── Helpers ────────────────────────────────────────────────────────────────

const rl = createInterface({ input: process.stdin, output: process.stdout });

function prompt(question) {
  return new Promise(res => rl.question(question, res));
}

function loadConfig() {
  const configPath = join(__dirname, 'config.json');
  if (!existsSync(configPath)) {
    throw new Error(`config.json not found at ${configPath}`);
  }
  return JSON.parse(readFileSync(configPath, 'utf-8'));
}

/** Resolve and normalize file path (Windows + Mac compatible). */
function resolveFilePath(inputPath) {
  const trimmed = inputPath.trim().replace(/^["']|["']$/g, '');
  return resolve(normalize(trimmed));
}

/**
 * Read rows from XLSX: column A = document type (CIF, NIF, DNI/NIF), column B = document number.
 * Returns array of { docType: 'CIF' | 'NIF', value: string }.
 * If only one column is present, value is taken from A and docType defaults to 'CIF'.
 */
function readRowsFromXlsx(filePath) {
  const buf = readFileSync(filePath);
  const workbook = XLSX.read(buf, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const a = data[i][0];
    const b = data[i][1];
    const typeRaw = a != null ? String(a).trim() : '';
    const valueB = b != null ? String(b).trim() : '';
    const valueA = a != null ? String(a).trim() : '';
    const hasTwoCols = valueB !== '';
    const value = hasTwoCols ? valueB : valueA;
    if (!value) continue;
    const headerLike = (v) => /^(tipo|type|id|n[uú]mero|documento|value|numero)$/i.test(String(v).trim());
    if (i === 0 && hasTwoCols && (headerLike(typeRaw) || headerLike(value))) continue;
    const docType = hasTwoCols ? ((typeRaw.toUpperCase() === 'CIF') ? 'CIF' : 'NIF') : 'CIF';
    rows.push({ docType, value });
  }
  return rows;
}

/** Write result matrix to XLSX: same directory as input, name = {basename}_resultado.xlsx */
function writeResultXlsx(inputFilePath, headers, rows) {
  const dir = dirname(inputFilePath);
  const base = basename(inputFilePath, extname(inputFilePath));
  const outPath = join(dir, `${base}_resultado.xlsx`);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Resultado');
  XLSX.writeFile(wb, outPath);
  return outPath;
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

const STEP_DELAY_MS = 1500;

/** Clicks the "Buscar Cliente" button inside frame W1DefaultW1 (image-only button). */
async function clickBuscarCliente(page) {
  const w1Frame = page.frameLocator('frame[name="W1DefaultW1"]');
  const btn = w1Frame.getByRole('button').filter({ hasText: /^$/ }).first();
  await btn.waitFor({ state: 'visible', timeout: 20000 });
  await btn.click();
  await sleep(STEP_DELAY_MS);
}

const AVISO_LEGAL_URL = 'https://w1.tiendas.ztna.telefonicaservices.com/w1/avisoLegal';
const W1_HOME_URL = 'https://w1.tiendas.ztna.telefonicaservices.com/w1/home';

/** Navigate to avisoLegal, then home, wait for frames, click Buscar Cliente, wait for search form (same as first-time flow). */
async function goToSearchForm(page, frameCabeceraName, frameContenidoName) {
  await page.goto(AVISO_LEGAL_URL, { waitUntil: 'load', timeout: 30000 });
  await sleep(STEP_DELAY_MS);
  await page.goto(W1_HOME_URL, { waitUntil: 'load', timeout: 60000 });
  await sleep(STEP_DELAY_MS * 2);

  let cabeceraFrame = null;
  for (let w = 0; w < 20; w++) {
    cabeceraFrame = page.frame({ name: frameCabeceraName });
    if (cabeceraFrame) break;
    await sleep(1000);
  }
  if (!cabeceraFrame) throw new Error('Frame W1CabPagW1 not found after 20s');

  await cabeceraFrame.waitForLoadState('load').catch(() => null);
  const contentFrame = page.frame({ name: frameContenidoName });
  if (contentFrame) await contentFrame.waitForLoadState('load').catch(() => null);
  await sleep(2000);

  await clickBuscarCliente(page);
  const content = page.frame({ name: frameContenidoName }) || page;
  await content.locator('h2.ctc-titles:has-text("Búsqueda de Cliente")').waitFor({ state: 'visible', timeout: 20000 });
  await sleep(STEP_DELAY_MS);
}

/** Run diagnostics and save screenshot + report when something is not found. */
async function runDiagnostics(page, contextLabel = 'page') {
  const outDir = join(__dirname, 'diagnostics');
  try {
    if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });
  } catch {
    // ignore
  }

  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const prefix = join(outDir, `diagnostic_${ts}`);

  console.log('\n--- DIAGNOSTIC ---');
  console.log('Page URL:', page.url());
  console.log('Page title:', await page.title().catch(() => '(failed)'));

  const frames = page.frames();
  console.log('Frames count:', frames.length);
  const frameNamesFromDom = await page.evaluate(() =>
    Array.from(document.querySelectorAll('frame[name], iframe[name]')).map(el => ({ name: el.getAttribute('name'), src: el.getAttribute('src') }))
  ).catch(() => []);
  console.log('Frame names in DOM (name, src):', JSON.stringify(frameNamesFromDom, null, 2));

  for (let i = 0; i < frames.length; i++) {
    const f = frames[i];
    const isMain = f === page.mainFrame();
    console.log(`  Frame ${i} ${isMain ? '(main)' : ''}: ${f.url().slice(0, 80)}${f.url().length > 80 ? '...' : ''}`);
  }

  const bodySnippet = await page.evaluate(() => {
    const body = document.body;
    if (!body) return '(no body)';
    const html = body.innerHTML;
    return html.length > 2000 ? html.slice(0, 2000) + '\n... [truncated]' : html;
  }).catch(e => `(error: ${e.message})`);

  const report = [
    `URL: ${page.url()}`,
    `Title: ${await page.title().catch(() => '')}`,
    `Frames: ${frames.length}`,
    `Frame names in DOM: ${JSON.stringify(frameNamesFromDom)}`,
    '',
    '--- Main page body (snippet) ---',
    bodySnippet,
  ].join('\n');

  const reportPath = `${prefix}_report.txt`;
  try {
    writeFileSync(reportPath, report, 'utf-8');
    console.log('Report saved:', reportPath);
  } catch (e) {
    console.log('Could not save report:', e.message);
  }

  const screenshotPath = `${prefix}_screenshot.png`;
  try {
    await page.screenshot({ path: screenshotPath, fullPage: true });
    console.log('Screenshot saved:', screenshotPath);
  } catch (e) {
    console.log('Could not save screenshot:', e.message);
  }

  const imgsMain = await page.locator('img').evaluateAll(els =>
    els.map(el => ({ id: el.id, src: (el.getAttribute('src') || '').slice(0, 60), className: el.className }))
  ).catch(() => []);
  console.log('Images in main page (id, src, class):', JSON.stringify(imgsMain, null, 2));

  for (let i = 0; i < frames.length; i++) {
    if (frames[i] === page.mainFrame()) continue;
    const imgsFrame = await frames[i].locator('img').evaluateAll(els =>
      els.map(el => ({ id: el.id, src: el.getAttribute('src') || '', className: el.className }))
    ).catch(() => []);
    console.log(`Images in frame ${i}:`, JSON.stringify(imgsFrame, null, 2));
  }

  const cabecera = page.frame({ name: 'W1CabPagW1' });
  if (cabecera) {
    const cabeceraSnippet = await cabecera.evaluate(() => {
      const b = document.body;
      return b ? b.innerHTML.slice(0, 5000) : '(no body)';
    }).catch(e => `(error: ${e.message})`);
    const cabeceraImgs = await cabecera.locator('img').evaluateAll(els =>
      els.map(el => ({ id: el.id, src: el.getAttribute('src') || '', className: el.className }))
    ).catch(() => []);
    const cabeceraBuscarRelated = await cabecera.evaluate(() => {
      const all = document.querySelectorAll('*');
      const out = [];
      const lower = (s) => (s || '').toLowerCase();
      all.forEach(el => {
        const id = el.id || '', cl = el.className || '', src = (el.getAttribute && el.getAttribute('src')) || '', href = (el.getAttribute && el.getAttribute('href')) || '', title = (el.getAttribute && el.getAttribute('title')) || '';
        const str = [id, cl, src, href, title].join(' ');
        if (lower(str).includes('buscar') || lower(str).includes('cliente')) {
          out.push({ tag: el.tagName, id, className: cl.slice(0, 80), src: src.slice(0, 80), href: href.slice(0, 80), title: title.slice(0, 80) });
        }
      });
      return out;
    }).catch(() => []);
    console.log('Cabecera frame images:', JSON.stringify(cabeceraImgs, null, 2));
    console.log('Cabecera elements with buscar/cliente:', JSON.stringify(cabeceraBuscarRelated, null, 2));
    try {
      writeFileSync(`${prefix}_cabecera_frame.txt`, `--- Cabecera frame images ---\n${JSON.stringify(cabeceraImgs, null, 2)}\n\n--- Elements with buscar/cliente ---\n${JSON.stringify(cabeceraBuscarRelated, null, 2)}\n\n--- Body snippet ---\n${cabeceraSnippet}`, 'utf-8');
      console.log('Cabecera frame dump saved:', `${prefix}_cabecera_frame.txt`);
    } catch {
      // ignore
    }
  }

  console.log('--- END DIAGNOSTIC ---\n');
}

// ── Tiendas flow (after main login) ────────────────────────────────────────

async function runTiendasFlow(page, config, rows) {
  const tiendas = config.tiendas || {};
  const homeUrl = tiendas.homeUrl || 'https://w1.tiendas.ztna.telefonicaservices.com/w1/inicio/home';
  const baseUrl = tiendas.baseUrl || 'https://w1.tiendas.ztna.telefonicaservices.com/w1/';
  const loginUser = tiendas.loginUser || '';
  const loginPassword = tiendas.loginPassword || '';

  console.log('\n--- Tiendas: navigating to home ---');
  await page.goto(homeUrl, { waitUntil: 'load', timeout: 30000 });
  await sleep(STEP_DELAY_MS);

  const loginBtn = page.locator('button:has-text("Iniciar sesión")');
  const pwdField = page.locator('input[type="password"]').first();
  const hasPwd = await pwdField.isVisible().catch(() => false);
  const hasLoginBtn = await loginBtn.isVisible().catch(() => false);
  if (hasPwd || hasLoginBtn) {
    console.log('Filling tiendas login modal...');
    const userInput = page.locator('input[name="username"], input[type="text"]').first();
    await userInput.waitFor({ state: 'visible', timeout: 5000 }).catch(() => null);
    await userInput.fill(loginUser);
    await pwdField.fill(loginPassword);
    await sleep(STEP_DELAY_MS);
    await loginBtn.click();
    await sleep(STEP_DELAY_MS);
    await page.waitForLoadState('load').catch(() => null);
    await sleep(STEP_DELAY_MS);
  }

  console.log('Navigating to base URL...');
  const baseGotoOptions = { waitUntil: 'load', timeout: 30000 };
  let baseLoaded = false;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      await page.goto(baseUrl, baseGotoOptions);
      baseLoaded = true;
      break;
    } catch (err) {
      const aborted = err.message?.includes('ERR_ABORTED') || err.message?.includes('net::ERR_ABORTED');
      if (aborted && attempt < 3) {
        console.log(`Base URL navigation aborted (attempt ${attempt}/3), waiting and retrying...`);
        await sleep(STEP_DELAY_MS * 2);
      } else {
        throw err;
      }
    }
  }
  await sleep(STEP_DELAY_MS);

  await page.waitForURL(/\/avisoLegal/i, { timeout: 25000 });
  await page.waitForLoadState('load').catch(() => null);
  await sleep(STEP_DELAY_MS);

  const frameCabeceraName = 'W1CabPagW1';
  const frameContenidoName = 'W1DefaultW1';
  const resultData = [];

  for (let i = 0; i < rows.length; i++) {
    const { docType, value } = rows[i];
    console.log(`Processing ${docType} ${i + 1}/${rows.length}: ${value}`);

    try {
      await goToSearchForm(page, frameCabeceraName, frameContenidoName);
      const content = page.frame({ name: frameContenidoName }) || page;

      const tipoDocDropdown = content.locator('button[data-id="_Tipodocumento_id"]');
      await tipoDocDropdown.click();
      await sleep(STEP_DELAY_MS);
      if (docType === 'CIF') {
        await content.locator('li a:has-text("CIF")').click();
      } else {
        await content.locator('li a:has-text("DNI/NIF")').click();
      }
      await sleep(STEP_DELAY_MS);

      const contextDropdown = content.locator('button[data-id="_Contexto_id"]');
      const docInput = content.locator('input#_Numerodocumento_id');

      const maxSubmitAttempts = 3;
      let descripciones = [];
      let abonos = [];

      for (let attempt = 1; attempt <= maxSubmitAttempts; attempt++) {
        await docInput.fill('');
        await docInput.fill(value);
        await sleep(STEP_DELAY_MS);

        await contextDropdown.click();
        await sleep(STEP_DELAY_MS);
        await content.locator('li[data-original-index="1"] a:has-text("Móvil")').click();
        await sleep(STEP_DELAY_MS);

        await content.locator('button#botonBuscar').click();
        await sleep(1500);

        const aceptarFrame = content.getByRole('button', { name: 'Aceptar' }).first();
        const aceptarPage = page.getByRole('button', { name: 'Aceptar' }).first();
        const modalAceptarVisible = await aceptarFrame.isVisible().catch(() => false) || await aceptarPage.isVisible().catch(() => false);
        if (modalAceptarVisible) {
          console.log(`Modal Aceptar shown (attempt ${attempt}/${maxSubmitAttempts}), clicking and retrying...`);
          try {
            await aceptarFrame.click({ timeout: 3000 });
          } catch {
            await aceptarPage.click({ timeout: 3000 });
          }
          await sleep(STEP_DELAY_MS);
          continue;
        }

        await page.frameLocator('frame[name="W1DefaultW1"]').getByRole('heading', { name: 'Indicadores', exact: true }).first().waitFor({ state: 'visible', timeout: 15000 });
        await sleep(STEP_DELAY_MS);

        const rawDesc = await content.locator('.panel.panel-info.alarmas .panel-body .datos_sociales table tbody tr td:nth-child(2)').evaluateAll(cells =>
          cells.map(c => (c.textContent || '').trim()).filter(Boolean)
        );
        const skipPhrases = [
          'no se encuentran datos para los criterios',
          'número comercial',
          'averías masivas',
          'código',
          'servicio',
          'prueba asociada',
          'fecha de alta',
          'fecha prevista',
          'fecha reparación',
          'estado',
        ];
        descripciones = rawDesc.filter(t => {
          const s = t.replace(/\s+/g, ' ').trim();
          if (!s) return false;
          if (s.length > 200 || (s.match(/\n/g) || []).length > 2) return false;
          if (/^×\s*/.test(s)) return false;
          const lower = s.toLowerCase();
          if (skipPhrases.some(p => lower.includes(p))) return false;
          return true;
        });

        await content.locator('button#botonAceptarAlarmas').click();
        await sleep(STEP_DELAY_MS);

        const masBtn = content.locator('button.btn.btn-xs.btn-info-cabcliente:has(img[src*="movil-contrato-pequeno"])');
        const isDisabled = await masBtn.getAttribute('disabled').then(d => d != null).catch(() => true);
        if (!isDisabled) {
          await masBtn.click();
          for (let w = 0; w < 40; w++) {
            const disabled = await content.evaluate(() => {
              const btn = document.querySelector('button.btn.btn-xs.btn-info-cabcliente img[src*="movil-contrato-pequeno"]');
              const b = btn?.closest('button');
              return b ? b.hasAttribute('disabled') === true : true;
            }).catch(() => true);
            if (disabled) break;
            await sleep(500);
          }
          await sleep(STEP_DELAY_MS);
        }

        abonos = await content.locator('table#tablaCabLineasCabecera tbody tr td.col-xs-3').evaluateAll(cells =>
          cells.map(c => (c.textContent || '').trim()).filter(Boolean)
        );
        break;
      }

      resultData.push({ cif: value, descripciones, abonos });
    } catch (err) {
      console.error(`${docType} ${value} failed:`, err.message);
      resultData.push({ cif: value, descripciones: [`Error: ${err.message}`], abonos: [] });
    }

    if (i < rows.length - 1) await sleep(STEP_DELAY_MS);
  }

  const numDescCols = Math.max(...resultData.map(d => d.descripciones.length), 1);
  const numAbonoCols = Math.max(...resultData.map(d => d.abonos.length), 1);
  const headers = ['CIF', ...Array.from({ length: numDescCols }, (_, i) => `Descripción_${i + 1}`), ...Array.from({ length: numAbonoCols }, (_, i) => `Abono_${i + 1}`)];

  const outputRows = resultData.map(({ cif, descripciones, abonos }) => {
    const descPart = [...descripciones];
    while (descPart.length < numDescCols) descPart.push('');
    const abonoPart = [...abonos];
    while (abonoPart.length < numAbonoCols) abonoPart.push('');
    return [cif, ...descPart, ...abonoPart];
  });

  return { headers, rows: outputRows };
}

// ── Main ───────────────────────────────────────────────────────────────────

async function main() {
  const filePath = await prompt('Enter the XLSX file path (column A = document type CIF/NIF, column B = number): ');
  const resolvedPath = resolveFilePath(filePath);

  if (!existsSync(resolvedPath)) {
    console.error(`Error: File not found at "${resolvedPath}"`);
    rl.close();
    process.exit(1);
  }

  if (!resolvedPath.toLowerCase().endsWith('.xlsx')) {
    console.error('Error: File must be .xlsx');
    rl.close();
    process.exit(1);
  }

  let rows;
  try {
    rows = readRowsFromXlsx(resolvedPath);
  } catch (e) {
    console.error('Error reading XLSX:', e.message);
    rl.close();
    process.exit(1);
  }

  if (!rows.length) {
    console.error('Error: No rows found (column A = document type CIF/NIF, column B = number)');
    rl.close();
    process.exit(1);
  }

  console.log(`File: ${resolvedPath}`);
  console.log(`Rows found: ${rows.length}`);

  const config = loadConfig();
  console.log(`\nLogging in as: ${config.username}`);
  console.log(`URL: ${config.url}\n`);

  const artifactsDir = join(__dirname, 'playwright-artifacts');
  const videosDir = join(artifactsDir, 'videos');
  const tracesDir = join(artifactsDir, 'traces');
  try {
    mkdirSync(videosDir, { recursive: true });
    mkdirSync(tracesDir, { recursive: true });
  } catch {
    // ignore
  }
  const traceTimestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const tracePath = join(tracesDir, `trace-${traceTimestamp}.zip`);

  const browser = await chromium.launch({ headless: config.headless || false });
  const tiendas = config.tiendas || {};
  const context = await browser.newContext({
    ignoreHTTPSErrors: true,
    viewport: { width: 1280, height: 720 },
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    httpCredentials: tiendas.loginUser && tiendas.loginPassword
      ? { username: tiendas.loginUser, password: tiendas.loginPassword }
      : undefined,
    recordVideo: { dir: videosDir },
  });
  await context.addInitScript(() => {
    Object.defineProperty(navigator, 'webdriver', { get: () => false });
  });
  const page = await context.newPage();

  await context.tracing.start({ screenshots: true, snapshots: true });
  console.log('Tracing started. Trace will be saved to:', tracePath);
  console.log('Video will be saved to:', videosDir);

  try {
    console.log('Navigating to login page...');
    await page.goto(config.url, { waitUntil: 'networkidle', timeout: 30000 });
    await sleep(STEP_DELAY_MS);

    console.log('Filling credentials...');
    await page.fill('#usuario', config.username);
    await page.fill('#password', config.password);
    await sleep(STEP_DELAY_MS);
    await page.locator('.content-check').click();
    console.log('Clicking Entrar...');
    await page.click('#entrar');
    await sleep(STEP_DELAY_MS);

    console.log('Waiting for OTP page...');
    await page.waitForSelector('#nffc', { timeout: 30000 });
    await sleep(STEP_DELAY_MS);

    let success = false;
    for (let attempt = 1; attempt <= 2; attempt++) {
      const otp = await prompt(`Enter OTP code (attempt ${attempt}/2): `);
      await page.fill('#nffc', '');
      await page.fill('#nffc', otp.trim());
      await page.click('#loginButton2');

      try {
        await page.waitForFunction(
          () => {
            const body = document.body?.innerText ?? '';
            const otpField = document.getElementById('nffc');
            return body.includes('Mis notificaciones') || (otpField && otpField.offsetParent !== null);
          },
          { timeout: 15000 }
        );

        const bodyText = await page.evaluate(() => document.body?.innerText ?? '');
        if (bodyText.includes('Mis notificaciones')) {
          success = true;
          break;
        }
        console.log(`\nOTP incorrect. ${attempt < 2 ? 'Please try again.' : ''}`);
        if (attempt === 2) throw new Error('OTP validation failed after 2 attempts.');
      } catch (waitErr) {
        if (waitErr.message?.includes('OTP validation failed')) throw waitErr;
        const bodyText = await page.evaluate(() => document.body?.innerText ?? '');
        if (bodyText.includes('Mis notificaciones')) {
          success = true;
          break;
        }
        if (attempt === 2) throw new Error('OTP validation failed after 2 attempts.');
      }
    }

    if (!success) {
      throw new Error('Login failed');
    }

    console.log('\nLogin successful. Running tiendas flow...');

    const { headers, rows: resultRows } = await runTiendasFlow(page, config, rows);

    const outPath = writeResultXlsx(resolvedPath, headers, resultRows);
    console.log(`\nResult written to: ${outPath}`);

  } catch (err) {
    console.error('\nError:', err.message);
  } finally {
    try {
      await context.tracing.stop({ path: tracePath });
      console.log('Trace saved:', tracePath);
      console.log('View with: npx playwright show-trace', tracePath);
    } catch (e) {
      console.warn('Could not save trace:', e?.message);
    }
    await browser.close();
    rl.close();
  }
}

main();
