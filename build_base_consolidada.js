const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const input = process.argv[2] || path.join(process.env.USERPROFILE || '', 'Downloads', 'YAKUPARK ADVENTURE S.R.L._Periodo _ 05-2026 (17).xlsx');
const extractDir = path.join(process.env.TEMP || __dirname, 'yakupark-base-extract');
const sharedStringsPath = path.join(extractDir, 'xl', 'sharedStrings.xml');
const sheetPath = path.join(extractDir, 'xl', 'worksheets', 'sheet1.xml');
const outputPath = path.join(__dirname, 'BASE_CONSOLIDADA.json');

function extractWorkbook() {
  if (!fs.existsSync(input)) throw new Error(`No se encontro el Excel: ${input}`);
  fs.rmSync(extractDir, { recursive: true, force: true });
  fs.mkdirSync(extractDir, { recursive: true });
  const zipPath = path.join(extractDir, 'workbook.zip');
  fs.copyFileSync(input, zipPath);
  execFileSync('powershell.exe', [
    '-NoProfile',
    '-Command',
    `Expand-Archive -LiteralPath ${JSON.stringify(zipPath)} -DestinationPath ${JSON.stringify(extractDir)} -Force`,
  ], { stdio: 'pipe' });
}

function decodeXml(value) {
  return String(value || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function columnIndex(ref) {
  return ref.split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

function loadSharedStrings() {
  const xml = fs.readFileSync(sharedStringsPath, 'utf8');
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(match => (
    decodeXml([...match[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(text => text[1]).join(''))
  ));
}

function readRows() {
  const sharedStrings = loadSharedStrings();
  const xml = fs.readFileSync(sheetPath, 'utf8');
  return [...xml.matchAll(/<row[^>]* r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)].map(rowMatch => {
    const row = [];
    for (const cellMatch of rowMatch[2].matchAll(/<c([^>]*)>([\s\S]*?)<\/c>/g)) {
      const attrs = cellMatch[1];
      const body = cellMatch[2];
      const ref = (attrs.match(/ r="([A-Z]+)\d+"/) || [])[1];
      if (!ref) continue;
      const type = (attrs.match(/ t="(\w+)"/) || [])[1];
      const valueMatch = body.match(/<v>([\s\S]*?)<\/v>/);
      const inlineMatch = body.match(/<t[^>]*>([\s\S]*?)<\/t>/);
      const raw = valueMatch ? valueMatch[1] : (inlineMatch ? inlineMatch[1] : '');
      row[columnIndex(ref)] = type === 's' ? sharedStrings[Number(raw)] || '' : decodeXml(raw);
    }
    return { number: Number(rowMatch[1]), values: row };
  });
}

function parseDate(value) {
  const match = String(value || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (Number.isNaN(date.getTime())) return null;
  return { day, month, year };
}

function pad(value) {
  return String(value).padStart(2, '0');
}

function dateParts(value) {
  const parsed = parseDate(value);
  if (!parsed) return null;
  return {
    exact: `${pad(parsed.day)}/${pad(parsed.month)}/${String(parsed.year).slice(-2)}`,
    iso: `${parsed.year}-${pad(parsed.month)}-${pad(parsed.day)}`,
    year: parsed.year,
    month: parsed.month,
    day: parsed.day,
  };
}

function toNumber(value) {
  const number = Number(String(value || '').replace(/,/g, ''));
  return Number.isFinite(number) ? number : 0;
}

function normalizeItem(value) {
  const text = String(value || '').trim().replace(/\s+/g, ' ').toUpperCase();
  if (text.includes('PULSERA PREMIUM')) return 'PULSERA FULL PASS';
  return text;
}

function buildBase() {
  extractWorkbook();
  if (!fs.existsSync(sheetPath)) {
    throw new Error(`No se encontro la hoja Reporte Soles en ${input}`);
  }

  const rows = readRows();
  const headers = rows.find(row => row.number === 6).values.map(value => String(value || '').trim());
  const dataRows = rows.filter(row => row.number > 6);

  const consolidated = dataRows.map(row => {
    const source = {};
    headers.forEach((header, index) => {
      if (header) source[header] = row.values[index] ?? '';
    });

    const parts = dateParts(source['Emisión']);
    if (!parts || parts.iso > '2026-05-17') return null;

    const normalizedItem = normalizeItem(source.Item);
    const quantity = toNumber(source['Cantidad Item']);
    const total = toNumber(source.Total);

    return {
      'Fecha Exacta': parts.exact,
      'Fecha ISO': parts.iso,
      'Año': parts.year,
      'Mes': parts.month,
      'Día': parts.day,
      'Item Normalizado': normalizedItem,
      'Cantidad': quantity,
      'Total': total,
      'Método de pago': source['Método de pago'] || '',
      'Horario': source.Horario || '',
      'Cliente': source.Cliente || '',
      'Clasificación Cliente': source['Clasificación Cliente'] || '',
      'Usuario': source.Usuario || '',
      'Nombre Cajero': source['Nombre Cajero'] || '',
      'Emisión': source['Emisión'] || '',
      'Hora Emisión': source['Hora Emisión'] || '',
      'Item': normalizedItem,
      'Cantidad Item': quantity,
      'Precio Unitario': toNumber(source['Precio Unitario']),
      'Comisionista': source.Comisionista || '',
      'Fecha reservada': source['Fecha reservada'] || '',
      'Condición': source['Condición'] || '',
      'Pago': source.Pago || '',
      'Nº': source['Nº'] || '',
      'Dni / Ruc': source['Dni / Ruc'] || '',
    };
  }).filter(Boolean);

  const fechas = [...new Set(consolidated.map(row => row['Fecha Exacta']))].sort((a, b) => {
    const [da, ma, ya] = a.split('/').map(Number);
    const [db, mb, yb] = b.split('/').map(Number);
    return new Date(2000 + ya, ma - 1, da) - new Date(2000 + yb, mb - 1, db);
  });

  fs.writeFileSync(outputPath, JSON.stringify({
    sheetName: 'BASE_CONSOLIDADA',
    sourceFile: path.basename(input),
    updatedThrough: '17/05/26',
    generatedAt: new Date().toISOString(),
    rowCount: consolidated.length,
    dates: fechas,
    rows: consolidated,
  }, null, 2));

  console.log('BASE_CONSOLIDADA generada:', outputPath);
  console.log('Filas:', consolidated.length);
  console.log('Fechas:', fechas.join(', '));
}

buildBase();
