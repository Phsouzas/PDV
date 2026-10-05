// PDV - Apps Script ligado a uma planilha (Extensões > Apps Script).
// Implantar > Nova implantação > App da Web > Executar como: eu | Acesso: qualquer pessoa.
const ABAS = ['Produtos', 'Comandas', 'Vendas', 'Senhas'];
const ss_ = () => SpreadsheetApp.getActiveSpreadsheet();
const aba_ = n => ss_().getSheetByName(n) || ss_().insertSheet(n);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  const p = e.parameter || {};
  if (p.action === 'auth') { // liste os e-mails permitidos na aba "Usuarios", coluna A
    const em = String(p.email || '').trim().toLowerCase();
    const ok = aba_('Usuarios').getDataRange().getValues().flat().some(x => String(x).trim().toLowerCase() === em);
    return out_({ ok });
  }
  const d = {};
  ABAS.forEach(n => {
    const v = aba_(n).getDataRange().getValues();
    d[n] = v.length < 2 ? [] : v.slice(1).map(r => JSON.parse(r[v[0].length - 1]));
  });
  return out_(d);
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (ABAS.indexOf(b.sheet) < 0) return out_({ ok: false });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = aba_(b.sheet), rows = b.rows || [];
    s.clear();
    if (rows.length) {
      const keys = [...new Set(rows.flatMap(r => Object.keys(r).filter(k => typeof r[k] !== 'object')))];
      const vals = [keys.concat('_json')].concat(rows.map(r => keys.map(k => r[k] === undefined ? '' : r[k]).concat(JSON.stringify(r))));
      s.getRange(1, 1, vals.length, vals[0].length).setValues(vals);
    }
  } finally { lock.releaseLock(); }
  return out_({ ok: true });
}
