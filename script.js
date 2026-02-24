// =================== Configurações ===================
const CSV_DELIMITER = ';';                           // Troque para ',' se o seu CSV usar vírgula
const FILENAME_BASE = 'Quebra de Transporte';

// === Manter apenas estas colunas (ordem e nomes exatos) ===
const KEEP_COLS = ['DATA', 'DESCRICAO', 'UNID.ORIGEM', 'UNID.DESTINO', 'TOT.DESC'];

// Nomes da planilha/arquivo Excel
const EXCEL_SHEET   = 'Base_Limpa';

// =================== Elementos da UI =================
const $file   = document.getElementById('file');
const $start  = document.getElementById('start');
const $bar    = document.getElementById('bar');
const $status = document.getElementById('status');
const $dlCsv  = document.getElementById('dlCsv');
const $dlXlsx = document.getElementById('dlXlsx');

// Esconde botão Excel se ExcelJS não estiver disponível
window.addEventListener('load', () => {
  if (!window.ExcelJS) $dlXlsx.style.display = 'none';
});

// =================== Helpers de UI ===================
function setStatus(html, cls = '') {
  $status.className = 'status ' + cls;
  $status.innerHTML = html;
}
function setProgress(value, max = 100) {
  const pct = Math.max(0, Math.min(100, Math.round((value / max) * 100)));
  $bar.style.width = pct + '%';
}

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

const TEXT_LIKE = new Set(['PRT','UNID.ORIGEM','UNID.DESTINO','DESCRICAO','PLACA','CEP']);
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

function toCSV(columns, rows, delimiter = CSV_DELIMITER) {
  const esc = (v) => {
    if (v === null || v === undefined) v = '';
    let s = String(v);
    if (s.includes('"') || s.includes(delimiter) || /\r|\n/.test(s)) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  let out = '';
  out += columns.map(esc).join(delimiter) + '\r\n';
  for (const r of rows) {
    out += columns.map(c => esc(r[c])).join(delimiter) + '\r\n';
  }
  return out;
}

async function toExcelSimple(columns, rows) {
  if (!window.ExcelJS) throw new Error('ExcelJS não disponível.');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(EXCEL_SHEET);

  // Cabeçalho + dados
  ws.addRow(columns);
  rows.forEach(r => ws.addRow(columns.map(c => r[c])));

  // Largura automática básica
  columns.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    const maxLen = Math.max(
      String(c).length,
      ...rows.slice(0, 200).map(r => String(r[c] ?? '').length)
    );
    col.width = Math.min(60, Math.max(10, Math.ceil(maxLen * 0.9)));
  });

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

  // Consolida e trata (antes de projetar as 5 colunas)
  let { columns, rows: aligned } = consolidateRows(rows);
  aligned = normalizeNumericColumns(aligned, columns);
  aligned = normalizeNullsToZero(aligned, columns);

  return { columns, rows: aligned };
}

// =================== Eventos da UI (hotfix) =====================
let lastCSVBlob = null;
let lastXLSXBlob = null;

// Habilita/desabilita o botão Iniciar e atualiza status
function updateStartState() {
  const hasFile = $file && $file.files && $file.files.length > 0;
  $start.disabled = !hasFile;
  setStatus(hasFile ? `Arquivo selecionado: <b>${$file.files[0].name}</b>` : 'Aguardando arquivo…', hasFile ? 'ok' : '');
}

// Registrar múltiplos eventos para garantir em qualquer navegador/host
$file.addEventListener('change', updateStartState);
$file.addEventListener('input', updateStartState);
$file.addEventListener('click', () => setTimeout(updateStartState, 0));
document.addEventListener('DOMContentLoaded', updateStartState);

// Iniciar processamento
$start.addEventListener('click', async () => {
  const hasFile = $file && $file.files && $file.files.length > 0;
  if (!hasFile) { updateStartState(); return; }

  $start.disabled = true; 
  $dlCsv.disabled = true; 
  $dlXlsx.disabled = true;

  setProgress(0, 100);
  setStatus('Processando… isso pode levar alguns segundos em arquivos grandes.');

  try {
    // 1) Trata o CSV “bagunçado” em blocos
    const { columns, rows } = await processFile($file.files[0]);

    // 2) === PROJEÇÃO: mantém só as 5 colunas desejadas, na ordem pedida ===
    const present = new Map(columns.map(c => [normKey(c), c]));  // mapa: chave normalizada -> nome real
    const outCols = [...KEEP_COLS];                               // nomes finais
    const outRows = rows.map(r => {
      const obj = {};
      for (const wanted of KEEP_COLS) {
        const actual = present.get(normKey(wanted));              // nome real no arquivo equivalente
        const val = actual ? r[actual] : 0;
        obj[wanted] = (val === undefined || val === null || String(val).trim() === '') ? 0 : val;
      }
      return obj;
    });

    // 3) CSV (apenas as 5 colunas)
    const csvText = toCSV(outCols, outRows, CSV_DELIMITER);
    lastCSVBlob = new Blob([csvText], { type: 'text/csv;charset=utf-8;' });

    // 4) Excel simples (sem Tabela, como você fará manualmente)
    if (window.ExcelJS) {
      lastXLSXBlob = await toExcelSimple(outCols, outRows);
      $dlXlsx.disabled = false;
      $dlXlsx.style.display = '';
      setStatus(`Finalizado! Linhas: <b>${outRows.length.toLocaleString('pt-BR')}</b> | Colunas: <b>${outCols.length}</b>.`, 'ok');
    } else {
      lastXLSXBlob = null;
      $dlXlsx.style.display = 'none';
      setStatus(`Finalizado! Linhas: <b>${outRows.length.toLocaleString('pt-BR')}</b> | Colunas: <b>${outCols.length}</b>.`, 'ok');
    }

    $dlCsv.disabled = false;
    setProgress(100, 100);
  } catch (err) {
    console.error(err);
    setStatus('Erro: ' + (err && err.message ? err.message : err), 'err');
  } finally {
    $start.disabled = false;
  }
});

// Downloads
$dlCsv.addEventListener('click', () => {
  if (!lastCSVBlob) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(lastCSVBlob);
  a.download = `${FILENAME_BASE}.csv`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
});

$dlXlsx.addEventListener('click', () => {
  if (!lastXLSXBlob) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(lastXLSXBlob);
  a.download = `${FILENAME_BASE}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
});
