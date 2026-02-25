// =================== Configurações ===================
const CSV_DELIMITER = ';';                           // Troque para ',' se o seu CSV usar vírgula
const FILENAME_BASE = 'Quebra de Transporte';

// === Manter apenas estas colunas (ordem e nomes exatos) ===
const KEEP_COLS = ['DESCRICAO', 'UNID.ORIGEM', 'UNID.DESTINO', 'TOT.DESC', 'CFOP'];

// >>> ADICIONADO: Whitelist de DESCRICAO
const ALLOWED_DESCRICOES = new Set([
  'MILHO COMERCIAL T',
  'SOJA CML TRANSG',
  'TRIGO PÃO',
  'TRIGO PAO',         // variação sem acento
]);

// Nomes da planilha/arquivo Excel
const EXCEL_SHEET = 'Base_Limpa'; // cria planilha e Tabela "Tabela_dados" dentro dela

// =================== Elementos da UI =================
const $file   = document.getElementById('file');
const $start  = document.getElementById('start');
const $bar    = document.getElementById('bar');
const $status = document.getElementById('status');
const $dlCsv  = document.getElementById('dlCsv');
const $dlXlsx = document.getElementById('dlXlsx');

// ======== Remover o botão CSV (oculta e não gera CSV) ========
if ($dlCsv) $dlCsv.style.display = 'none';

// =================== Loader dinâmico do ExcelJS (sem HTML) ===================
let exceljsLoading = null;
async function ensureExcelJS() {
  if (window.ExcelJS) return 'already-loaded';
  if (exceljsLoading) return exceljsLoading;

  // CDN (pode trocar para unpkg se preferir)
  const CDN = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';

  exceljsLoading = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = CDN;
    s.async = true;
    s.onload = () => {
      const ok = !!(window.ExcelJS && window.ExcelJS.Workbook);
      if (!ok) { reject(new Error('ExcelJS carregado mas objeto inválido.')); return; }
      try {
        const testWb = new window.ExcelJS.Workbook();
        const testWs = testWb.addWorksheet('test');
        const supportsTable = typeof testWs.addTable === 'function';
        console.log('[ExcelJS] carregado. Suporte a Table:', supportsTable);
      } catch (e) {
        console.warn('[ExcelJS] carregado, mas não foi possível inspecionar addTable:', e);
      }
      resolve('loaded');
    };
    s.onerror = () => reject(new Error('Falha ao baixar ExcelJS do CDN.'));
    document.head.appendChild(s);
  });

  return exceljsLoading;
}

// =================== Helpers de UI ===================
function setStatus(html, cls = '') {
  $status.className = 'status ' + cls;
  $status.innerHTML = html;
}
function setProgress(value, max = 100) {
  const pct = Math.max(0, Math.min(100, Math.round((value / max) * 100)));
  $bar.style.width = pct + '%';
}

window.addEventListener('load', async () => {
  try {
    await ensureExcelJS();           // baixa ExcelJS
    $dlXlsx.style.display = '';      // mostra botão XLSX
  } catch (e) {
    console.warn('[ExcelJS] não carregado:', e.message);
    $dlXlsx.style.display = 'none';
  }
});

// =================== Parsing/Tratamento ==============

// Normaliza chave de coluna (para casar nomes equivalentes como "TOT.DESC", "Tot_Desc", "tot desc")
const normKey = (s) => String(s || '')
  .toUpperCase()
  .replace(/\s+/g, '')
  .replace(/[.\-_]/g, '');

function parseCSVLine(line, delimiter = CSV_DELIMITER) {
  // Parser simples com suporte a aspas duplas
  const out = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') { cur += '"'; i++; }
      else { inQuotes = !inQuotes; }
    } else if (ch === delimiter && !inQuotes) {
      out.push(cur); cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

function makeUnique(cols) {
  const seen = {};
  return cols.map(c => {
    c = (c || '').trim();
    if (!c) return '';
    if (seen[c] === undefined) { seen[c] = 0; return c; }
    seen[c] += 1; return `${c}_${seen[c]}`;
  });
}

function isHeaderRowFirst3(values, expected = ['DATA', 'NUMCMP', 'PRT']) {
  const v = [0,1,2].map(i => (values[i] ?? '').toString().trim().toUpperCase());
  return v[0] === expected[0] && v[1] === expected[1] && v[2] === expected[2];
}

const TEXT_LIKE = new Set(['PRT','UNID.ORIGEM','UNID.DESTINO','DESCRICAO','PLACA','CEP', 'CFOP']);
function tryParseBrNumber(s) {
  if (s === null || s === undefined) return s;
  const t = String(s).trim();
  if (!t) return t;
  if (/[A-Za-z]/.test(t)) return t;            // evita PLACA/DESCRICAO
  if (!/\d/.test(t)) return t;
  const n = Number(t.replace(/\./g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : s;
}

function normalizeNumericColumns(rows, columns) {
  const out = rows.map(r => ({ ...r }));
  for (const col of columns) {
    if (TEXT_LIKE.has(col)) continue;
    for (const obj of out) {
      if (col in obj) obj[col] = tryParseBrNumber(obj[col]);
    }
  }
  return out;
}

function consolidateRows(rows, preferFirst = ['DATA','NUMCMP','PRT']) {
  const set = new Set();
  rows.forEach(r => Object.keys(r).forEach(k => set.add(k)));
  const all = Array.from(set);
  const ordered = [...preferFirst.filter(p => set.has(p)), ...all.filter(c => !preferFirst.includes(c))];
  return { columns: ordered, rows };
}

function normalizeNullsToZero(rows, columns) {
  // Converte NaN/NaT/None/''/"nan"/"NULL" -> 0
  const NULL_STRS = new Set(['', 'nan', 'NaN', 'NAN', 'None', 'NULL', 'null']);
  return rows.map(r => {
    const o = {};
    for (const c of columns) {
      let v = r[c];
      if (v === undefined || v === null) v = 0;
      else {
        const s = String(v).trim();
        if (NULL_STRS.has(s)) v = 0;
      }
      o[c] = v;
    }
    return o;
  });
}

// ------------- Agregação por DESCRICAO + UNID.ORIGEM + UNID.DESTINO -------------
const GROUP_COLS = ['DESCRICAO', 'UNID.ORIGEM', 'UNID.DESTINO', 'CFOP'];
const SUM_COL    = 'TOT.DESC';

function aggregateByGroup(rows) {
  const map = new Map();
  for (const r of rows) {
    const key = GROUP_COLS.map(c => String(r[c] ?? '').trim().toUpperCase()).join('|');
    const cur = map.get(key);
    const addVal = Number(tryParseBrNumber(r[SUM_COL])) || 0;

    if (!cur) {
      const base = { ...r };
      // DATA agregada perde sentido: defina como 0 (ou mude aqui para 'primeira' ou 'vazia')
      if ('DATA' in base) base['DATA'] = 0;
      base[SUM_COL] = addVal; // numérico, preserva sinal
      map.set(key, base);
    } else {
      cur[SUM_COL] = (Number(cur[SUM_COL]) || 0) + addVal;
    }
  }
  // arredonda para 2 casas (se quiser)
  return Array.from(map.values()).map(o => ({ ...o, [SUM_COL]: Number((o[SUM_COL] || 0).toFixed(2)) }));
}

// =================== CSV helper removido (não geramos CSV) ===================

// Helper: letra da coluna (1->A etc.)
function excelCol(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ============== Excel: cria planilha + Tabela (+fallback) ==============
async function toExcelWithTable(columns, rows, tableName = 'Tabela_dados') {
  if (!window.ExcelJS) throw new Error('ExcelJS não disponível.');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(EXCEL_SHEET);

  // Escreve cabeçalho + dados (garante dados no WS mesmo se addTable falhar)
  ws.addRow(columns);
  rows.forEach(r => ws.addRow(columns.map(c => r[c])));

  const lastRow = rows.length + 1; // +1 cabeçalho
  const lastColLetter = excelCol(columns.length);
  const fullRange = `A1:${lastColLetter}${lastRow}`;

  // Tenta Tabela oficial
  try {
    const supportsTable = typeof ws.addTable === 'function';
    if (supportsTable) {
      try {
        if (ws.getTable && ws.getTable(tableName)) {
          ws.getTable(tableName).remove?.();
        }
      } catch (_) {}

      ws.addTable({
        name: tableName,
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        style: {
          theme: 'TableStyleMedium9',
          showRowStripes: true,
          showColumnStripes: false,
        },
        columns: columns.map(name => ({ name, filterButton: true })),
        rows: rows.map(r => columns.map(c => r[c])),
      });

      if (ws.getTable && !ws.getTable(tableName)) {
        console.warn('[ExcelJS] Tabela não registrada; aplicando autoFilter como fallback.');
        ws.autoFilter = fullRange;
      }
    } else {
      console.warn('[ExcelJS] addTable não disponível nesta build; aplicando autoFilter.');
      ws.autoFilter = fullRange;
    }
  } catch (e) {
    console.error('[ExcelJS] Falha ao criar Tabela:', e);
    ws.autoFilter = fullRange; // fallback
  }

  // Freeze header
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

  // Largura automática
  columns.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    const maxLen = Math.max(
      String(c).length,
      ...rows.slice(0, 200).map(r => String(r[c] ?? '').length)
    );
    col.width = Math.min(60, Math.max(10, Math.ceil(maxLen * 0.9)));
  });

  // Formatação
  const colIndex = (name) => columns.findIndex(c => c.toUpperCase() === name.toUpperCase()) + 1;
  const idxData = colIndex('DATA');
  if (idxData > 0) ws.getColumn(idxData).numFmt = 'dd/mm/yyyy';
  const idxTotDesc = colIndex('TOT.DESC');
  if (idxTotDesc > 0) ws.getColumn(idxTotDesc).numFmt = '#,##0.00';

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

// =================== Processo principal =================
async function processFile(file) {
  const text = await file.text();
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const total = lines.length;
  setProgress(0, total);

  const rows = [];
  let headersUnique = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) { setProgress(i + 1, total); continue; }

    // Detecta novo cabeçalho
    if (line.toUpperCase().startsWith('DATA' + CSV_DELIMITER)) {
      const vals = parseCSVLine(line);
      if (isHeaderRowFirst3(vals)) {
        const hdrs = vals.map(h => (h || '').trim());
        while (hdrs.length && !hdrs[hdrs.length - 1]) hdrs.pop();
        headersUnique = makeUnique(hdrs);
        setStatus(`Novo cabeçalho: <b>${headersUnique.length}</b> colunas.`, 'ok');
        setProgress(i + 1, total);
        continue;
      }
    }

    // Linha de dados do bloco atual
    if (headersUnique) {
      let vals = parseCSVLine(line);
      if (vals.length < headersUnique.length) vals = vals.concat(Array(headersUnique.length - vals.length).fill(''));
      else if (vals.length > headersUnique.length) vals = vals.slice(0, headersUnique.length);

      const obj = {};
      headersUnique.forEach((h, idx) => { if (h) obj[h] = vals[idx]; });
      const any = Object.values(obj).some(v => String(v).trim() !== '');
      if (any) rows.push(obj);
    }
    setProgress(i + 1, total);
  }

  // Consolida e trata (antes de projetar as 4 colunas)
  let { columns, rows: aligned } = consolidateRows(rows);
  aligned = normalizeNumericColumns(aligned, columns);
  aligned = normalizeNullsToZero(aligned, columns);

  return { columns, rows: aligned };
}

// =================== Eventos da UI =====================
let lastXLSXBlob = null;

function updateStartState() {
  const hasFile = $file && $file.files && $file.files.length > 0;
  $start.disabled = !hasFile;
  setStatus(hasFile ? `Arquivo selecionado: <b>${$file.files[0].name}</b>` : 'Aguardando arquivo…', hasFile ? 'ok' : '');
}

$file.addEventListener('change', updateStartState);
$file.addEventListener('input', updateStartState);
$file.addEventListener('click', () => setTimeout(updateStartState, 0));
document.addEventListener('DOMContentLoaded', updateStartState);

$start.addEventListener('click', async () => {
  const hasFile = $file && $file.files && $file.files.length > 0;
  if (!hasFile) { updateStartState(); return; }

  $start.disabled = true;
  $dlXlsx.disabled = true;

  setProgress(0, 100);
  setStatus('Processando… isso pode levar alguns segundos em arquivos grandes.');

  try {
    // garante ExcelJS antes de criar XLSX
    try {
      await ensureExcelJS();
      $dlXlsx.style.display = '';
    } catch (e) {
      console.warn('[ExcelJS] indisponível. Não será possível gerar Excel.', e.message);
    }

    // 1) Trata o CSV
    const { columns, rows } = await processFile($file.files[0]);

    // 2) Projeção: 4 colunas
    const present = new Map(columns.map(c => [normKey(c), c]));
    const outCols = [...KEEP_COLS];
    const projectedRows = rows.map(r => {
      const obj = {};
      for (const wanted of KEEP_COLS) {
        const actual = present.get(normKey(wanted));
        const val = actual ? r[actual] : 0;
        obj[wanted] = (val === undefined || val === null || String(val).trim() === '') ? 0 : val;
      }
      return obj;
    });

    // >>> ADICIONADO: 2.1) Filtro pela DESCRICAO (mantém apenas os permitidos)
    const filteredRows = projectedRows.filter(r => {
      const v = String(r['DESCRICAO'] ?? '').trim().toUpperCase();
      return ALLOWED_DESCRICOES.has(v);
    });

    // 3) AGREGAÇÃO (DESCRICAO + UNID.ORIGEM + UNID.DESTINO) somando TOT.DESC com sinal
    const outRows = aggregateByGroup(filteredRows);

    // 4) XLSX com Tabela (se ExcelJS disponível)
    if (window.ExcelJS) {
      lastXLSXBlob = await toExcelWithTable(outCols, outRows, 'Tabela_dados');
      $dlXlsx.disabled = false;
      $dlXlsx.style.display = '';
      setStatus(`Finalizado! Linhas (após filtrar/agregar): <b>${outRows.length.toLocaleString('pt-BR')}</b> | Colunas: <b>${outCols.length}</b>.`, 'ok');
    } else {
      lastXLSXBlob = null;
      $dlXlsx.style.display = 'none';
      setStatus(`Erro: ExcelJS não disponível para gerar o arquivo Excel.`, 'err');
    }

    setProgress(100, 100);
  } catch (err) {
    console.error(err);
    setStatus('Erro: ' + (err && err.message ? err.message : err), 'err');
  } finally {
    $start.disabled = false;
  }
});

// Download Excel
$dlXlsx.addEventListener('click', () => {
  if (!lastXLSXBlob) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(lastXLSXBlob);
  a.download = `${FILENAME_BASE}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
});
``


