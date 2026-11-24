/**
 * Haushaltsplan – Abgleich für Christoffer und Ariane
 * In eine neue Google-Tabelle einfügen: Erweiterungen → Apps Script.
 * 1. Unten bei CODE euren eigenen Zugangscode eintragen (beliebiges Wort, z. B. "Luise2026").
 * 2. Speichern → Bereitstellen → Neue Bereitstellung → Typ "Web-App"
 *    Ausführen als: Ich · Zugriff: Jeder → Bereitstellen → Zugriff erlauben.
 * 3. Die Web-App-URL (endet auf /exec) in der App unter "Einstellungen" einfügen.
 */
const CODE = 'HIER-EUREN-CODE-EINTRAGEN';

function doGet(e) {
  const p = e.parameter || {};
  if (p.c !== CODE) return out({ error: 'code' });
  if (!/^\d{4}-W\d{2}$/.test(p.w || '')) return out({ error: 'woche' });
  return out(readWeek(p.w).data);
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out({ error: 'json' }); }
  if (b.c !== CODE) return out({ error: 'code' });
  if (!/^\d{4}-W\d{2}$/.test(b.w || '')) return out({ error: 'woche' });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const wk = readWeek(b.w);
    if (b.k) {
      if (!/^[a-z0-9-]+@[0-6]$/.test(b.k)) return out({ error: 'aufgabe' });
      wk.data.items[b.k] = b.v && typeof b.v === 'object' ? b.v : {};
    }
    if (b.wd && typeof b.wd === 'object') wk.data.wd = b.wd;
    const sh = sheet();
    const row = [b.w, JSON.stringify(wk.data), new Date()];
    if (wk.row) sh.getRange(wk.row, 1, 1, 3).setValues([row]); else sh.appendRow(row);
    return out(wk.data);
  } finally {
    lock.releaseLock();
  }
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Daten');
  if (!sh) {
    sh = ss.insertSheet('Daten');
    sh.appendRow(['Woche', 'Stand (JSON)', 'Geändert']);
  }
  return sh;
}

function readWeek(w) {
  const sh = sheet();
  const last = sh.getLastRow();
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === w) {
        let data;
        try { data = JSON.parse(vals[i][1]); } catch (err) { data = {}; }
        return { row: i + 2, data: { items: data.items || {}, wd: data.wd || {} } };
      }
    }
  }
  return { row: 0, data: { items: {}, wd: {} } };
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
