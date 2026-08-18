import { createZipStore } from './zip.js';

const HEADERS = ['ルールID', 'カテゴリ', '重要度', '場所', '該当文字列', '検出内容', '修正候補', '文脈'];

export function downloadCsv(findings, baseName) {
  const rows = [HEADERS, ...findings.map(rowValues)];
  const csv = '\ufeff' + rows.map(r => r.map(csvCell).join(',')).join('\r\n');
  downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), `${safeBase(baseName)}_qc.csv`);
}

export function downloadXlsx(findings, baseName) {
  const rows = [HEADERS, ...findings.map(rowValues)];
  const files = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="QC Results" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': sheetXml(rows)
  };
  const zip = createZipStore(files);
  downloadBlob(new Blob([zip], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${safeBase(baseName)}_qc.xlsx`);
}

function rowValues(f) { return [f.ruleId, f.category, severityLabel(f.severity), f.location, f.matched, f.message, f.suggestion, f.context]; }
function severityLabel(v) { return v === 'error' ? '要修正' : v === 'confirm' ? '確認' : '参考'; }
function csvCell(v) { const s = String(v ?? ''); return `"${s.replace(/"/g, '""')}"`; }
function safeBase(name) { return (name || 'document').replace(/\.(docx|xlsx)$/i, '').replace(/[\\/:*?"<>|]+/g, '_'); }
function escapeXml(s) { return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function colName(n) { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; }
function sheetXml(rows) {
  const body = rows.map((row, rIdx) => `<row r="${rIdx + 1}">${row.map((v, cIdx) => `<c r="${colName(cIdx + 1)}${rIdx + 1}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(v)}</t></is></c>`).join('')}</row>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
}
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.rel = 'noopener';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
