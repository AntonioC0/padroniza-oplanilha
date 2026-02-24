// =================== Configurações ===================
const CSV_DELIMITER = ';';                           // Troque para ',' se o seu CSV usar vírgula
const FILENAME_BASE = 'Base_Limpa - quebra de transporte';
const EXCEL_SHEET   = 'Base_Limpa';
const TABLE_NAME    = 'Tabela_dados';

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
  // Converte NaN/NaT/None/''/"nan"/"NULL" -> 0 em todas as colunas
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

function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1) / 26 | 0; }
  return s;
}

// Gera nomes únicos só para o cabeçalho do Excel (CSV mantém original)
function uniqueForExcel(cols) {
  const seen = {};
  return cols.map(name => {
    let n = String(name || '').trim();
    if (!n) n = 'COLUNA';
    if (seen[n] === undefined) { seen[n] = 0; return n; }
    seen[n] += 1; return `${n}_${seen[n]}`;
  });
}

// ===== Criação do Excel com Tabela 'Tabela_dados' =====
async function toExcel(columns, rows) {
  if (!window.ExcelJS) throw new Error('ExcelJS não disponível.');

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(EXCEL_SHEET);

  // Cabeçalhos exclusivos para Excel (evita erro por duplicidade)
  const columnsExcel = uniqueForExcel(columns);

  // Dados da tabela: cada linha vira um array na ordem de 'columns'
  const tableRows = rows.map(r => columns.map(c => r[c]));

  if (typeof ws.addTable !== 'function') {
    throw new Error('worksheet.addTable() não existe nesta versão do ExcelJS. Atualize lib/exceljs.min.js (4.x).');
  }

  // Define o intervalo total da tabela
  const ref = `A1:${colLetter(columnsExcel.length)}${tableRows.length + 1}`;

  // Cria a Tabela (o Excel preencherá cabeçalho + linhas)
  ws.addTable({
    name: TABLE_NAME,                 // <<<<<< Tabela_dados
    ref,
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium9', showRowStripes: true },
    columns: columnsExcel.map(name => ({ name })),
    rows: tableRows
  });

  // (Opcional) confirma a tabela quando a API expõe getTable()
  try {
    if (typeof ws.getTable === 'function') {
      const tbl = ws.getTable(TABLE_NAME);
      if (tbl && typeof tbl.commit === 'function') tbl.commit();
    }
  } catch (e) {
    // silencioso: algumas builds não expõem getTable/commit
  }

  // Largura automática
  columnsExcel.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    const maxLen = Math.max(
      String(c).length,
      ...tableRows.slice(0, 200).map(r => String(r[i] ?? '').length)
    );
    col.width = Math.min(60, Math.max(10, Math.ceil(maxLen * 0.9)));
  });

  console.log(`Tabela '${TABLE_NAME}' criada com ${tableRows.length} linhas e ${columnsExcel.length} colunas em '${EXCEL_SHEET}'.`);

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

  // Consolida e trata
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
    const { columns, rows } = await processFile($file.files[0]);

    // CSV
    const csvText = toCSV(columns, rows, CSV_DELIMITER);
    lastCSVBlob = new Blob([csvText], { type: 'text/csv;charset=utf-8;' });

    // Excel (se disponível)
    if (window.ExcelJS) {
      lastXLSXBlob = await toExcel(columns, rows);
      $dlXlsx.disabled = false;
      $dlXlsx.style.display = '';
    } else {
      lastXLSXBlob = null;
      $dlXlsx.style.display = 'none';
    }

    $dlCsv.disabled = false;
    setStatus(`Finalizado! Linhas: <b>${rows.length.toLocaleString('pt-BR')}</b> | Colunas: <b>${columns.length}</b>.`, 'ok');
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