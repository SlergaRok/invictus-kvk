/**
 * INVICTUS — API de jugadores y MGE.
 *
 * Instalación (una sola vez):
 *  1. En el Google Sheet "MGE & Players": Extensiones > Apps Script.
 *  2. Pega este fichero entero y guarda.
 *  3. Implementar > Nueva implementación > Tipo: Aplicación web
 *       Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
 *  4. Copia la URL que termina en /exec y ponla en DATA_API_URL de index.html.
 *  Cada vez que cambies este código: Implementar > Gestionar implementaciones > Editar > Nueva versión.
 *
 * Endpoints:
 *  ?view=players -> { view, data: [filas de la pestaña Players tal cual, cabecera -> valor] }
 *  ?view=mge     -> { view, costs, balances, events, unmatched }
 */

const TABS = { players: 'Players', costs: 'MGE_Costes', grid: 'MGE_Asignacion', balances: 'MGE_Saldos' };
const MEMBER_COLS = 9; // columnas B..J de MGE_Asignacion (un MGE por columna)

function doGet(e) {
  const view = String((e && e.parameter && e.parameter.view) || '').toLowerCase();
  let body;
  try {
    if (view === 'players') body = { view: 'players', data: readPlayers_() };
    else if (view === 'mge') body = Object.assign({ view: 'mge' }, readMge_());
    else body = { error: 'Usa ?view=players o ?view=mge' };
  } catch (err) {
    body = { error: String((err && err.message) || err) };
  }
  body.updated = new Date().toISOString();
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('No existe la pestaña "' + name + '"');
  return sh;
}

function clean_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

// Número o null ("-", vacío o texto no numérico -> null).
function num_(v) {
  if (typeof v === 'number') return v;
  const s = clean_(v);
  if (!s) return null;
  const n = Number(s.replace(/,/g, ''));
  return isFinite(n) ? n : null;
}

/* ---------------- Players ---------------- */

function readPlayers_() {
  const values = sheet_(TABS.players).getDataRange().getValues();
  const headers = values[0].map(clean_);
  return values.slice(1)
    .filter(row => clean_(row[0]) !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { if (h) obj[h] = (i === 0) ? clean_(row[i]) : row[i]; });
      return obj;
    });
}

/* ---------------- MGE ---------------- */

function readMge_() {
  // Costes por puesto
  const costs = sheet_(TABS.costs).getDataRange().getValues().slice(1)
    .map(r => ({ spot: num_(r[0]), heads: num_(r[1]), cost: num_(r[2]) }))
    .filter(c => c.spot);
  const costBySpot = {};
  costs.forEach(c => { costBySpot[c.spot] = c; });

  // Saldos (calculados con fórmulas en el Sheet)
  const balances = sheet_(TABS.balances).getDataRange().getValues().slice(1)
    .filter(r => clean_(r[0]) !== '' && clean_(r[1]) !== '')
    .map(r => ({ id: clean_(r[0]), name: clean_(r[1]), earned: num_(r[2]), spent: num_(r[3]), balance: num_(r[4]) }));
  const idByName = {};
  balances.forEach(b => { idByName[b.name] = b.id; });

  // Cuadrícula de asignaciones: bloques "MGE xxx", fila "TOPS" con cabeceras, filas 1..15
  const grid = sheet_(TABS.grid).getDataRange().getValues();
  const blocks = [];
  let block = null;
  grid.forEach(row => {
    const a = clean_(row[0]);
    if (/^mge\b/i.test(a)) {
      block = { name: a, headers: [], cells: [] };
      blocks.push(block);
      return;
    }
    if (!block) { block = { name: 'MGE', headers: [], cells: [] }; blocks.push(block); }
    if (/^tops?$/i.test(a)) {
      block.headers = row.slice(1, 1 + MEMBER_COLS).map(clean_);
      return;
    }
    const spot = num_(row[0]);
    if (!spot) return;
    for (let c = 1; c <= MEMBER_COLS; c++) {
      const name = clean_(row[c]);
      if (name) block.cells.push({ col: c, spot: spot, name: name });
    }
  });

  // Un evento por columna con al menos un nombre; el último de cada bloque es el actual
  const events = [];
  const unmatched = {};
  blocks.forEach(b => {
    const cols = Array.from(new Set(b.cells.map(x => x.col))).sort((x, y) => x - y);
    cols.forEach((col, i) => {
      const header = b.headers[col - 1] || '';
      const assignments = b.cells.filter(x => x.col === col)
        .sort((x, y) => x.spot - y.spot)
        .map(x => {
          const cost = costBySpot[x.spot] || {};
          const id = idByName[x.name] || null;
          if (!id) unmatched[x.name] = true;
          return { spot: x.spot, name: x.name, id: id, heads: cost.heads || null, cost: cost.cost || null };
        });
      events.push({
        block: b.name,
        number: i + 1,
        column: col,
        label: /^miembros?$/i.test(header) ? '' : header,
        current: i === cols.length - 1,
        assignments: assignments
      });
    });
  });

  return { costs: costs, balances: balances, events: events, unmatched: Object.keys(unmatched) };
}
