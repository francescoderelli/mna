const elPremi  = document.getElementById("filePremi");
const elSum    = document.getElementById("fileSum");
const elObi    = document.getElementById("fileObi");
const btnRun   = document.getElementById("btnRun");
const statusEl = document.getElementById("status");
const checksEl = document.getElementById("checks");

// evita race condition: solo l’ultima validazione può abilitare/disabilitare
let VALIDATION_SEQ = 0;

function setStatus(msg, kind="") {
  statusEl.className = "status " + (kind || "");
  statusEl.textContent = msg;
}
function setChecks(lines) {
  checksEl.innerHTML = "";
  for (const l of lines) {
    const div = document.createElement("div");
    div.textContent = l;
    checksEl.appendChild(div);
  }
}
function readyBtn(enabled) { btnRun.disabled = !enabled; }

function normStr(v) { return (v ?? "").toString().trim(); }
function toNum(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function assert(cond, msg) { if (!cond) throw new Error(msg); }

function isReadyFiles() {
  return !!(elPremi.files?.[0] && elSum.files?.[0] && elObi.files?.[0]);
}

function readXlsxToAOA(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Errore lettura file: " + file.name));
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      resolve({ wb, sheetName, aoa });
    };
    reader.readAsArrayBuffer(file);
  });
}

// ============================================================
// VALIDAZIONI (v1.1)
// ============================================================
function validatePremiAOA(aoa) {
  assert(Array.isArray(aoa) && aoa.length > 5, "Premi_Mensili_x_Sede: file vuoto o non leggibile.");

  let headerIdx = -1;
  for (let i = 0; i < aoa.length; i++) {
    if (normStr(aoa[i][0]) === "Rows") { headerIdx = i; break; }
  }
  assert(headerIdx >= 1, "Premi_Mensili_x_Sede: non trovo 'Rows' in colonna A.");
  assert(headerIdx - 1 >= 0, "Premi_Mensili_x_Sede: manca la riga mesi sopra 'Rows'.");

  let hasEdac = false, hasRA = false;
  for (let i = headerIdx + 1; i < aoa.length; i++) {
    const v = normStr(aoa[i][0]);
    if (v.toUpperCase().startsWith("EDAC_")) hasEdac = true;
    if (v === "responsabili_area") hasRA = true;
  }
  assert(hasEdac, "Premi_Mensili_x_Sede: non trovo nessuna riga che inizi con 'EDAC_'.");
  assert(hasRA, "Premi_Mensili_x_Sede: non trovo la riga 'responsabili_area'.");
  return true;
}

function validateObiettiviAOA(aoa) {
  assert(Array.isArray(aoa) && aoa.length > 5, "Sum_of_Obiettivoprev: file vuoto o non leggibile.");

  // deve arrivare almeno alla colonna Y (indice 24)
  let maxRowLen = 0;
  for (let i = 0; i < Math.min(80, aoa.length); i++) {
    maxRowLen = Math.max(maxRowLen, (aoa[i] || []).length);
  }
  assert(maxRowLen >= 25, "Sum_of_Obiettivoprev: layout non valido (mancano colonne fino a Y).");

  // deve contenere almeno una riga CORE (colonna G)
  let hasCore = false;
  for (let i = 0; i < aoa.length; i++) {
    const ramo = normStr((aoa[i] || [])[6]).toUpperCase(); // G
    if (ramo === "CORE") { hasCore = true; break; }
  }
  assert(hasCore, "Sum_of_Obiettivoprev: non trovo righe con Ramo=CORE (colonna G).");
  return true;
}

function validateSumImportoAOA(aoa) {
  assert(Array.isArray(aoa) && aoa.length > 5, "Sum_of_Importo: file vuoto o non leggibile.");

  // trova "Sezione" in col A
  let headerRow = -1;
  for (let i = 0; i < aoa.length; i++) {
    if (normStr((aoa[i] || [])[0]) === "Sezione") { headerRow = i; break; }
  }
  assert(headerRow >= 1, "Sum_of_Importo: non trovo 'Sezione' in colonna A.");

  const monthsRow = headerRow - 1;
  assert(monthsRow >= 0, "Sum_of_Importo: manca la riga mesi sopra 'Sezione'.");

  // deve esistere la riga MNA in col B
  let hasMna = false;
  for (let i = 0; i < aoa.length; i++) {
    const label = normStr((aoa[i] || [])[1]).toUpperCase();
    if (label.includes("MNA - MARGINE NETTO DI AREA")) { hasMna = true; break; }
  }
  assert(hasMna, "Sum_of_Importo: non trovo la riga 'MNA - Margine Netto di Area' (colonna B).");

  // ✅ deve esserci lo split mensile completo 1..12 sulle colonne "Importo*"
  const header = aoa[headerRow] || [];
  const months = aoa[monthsRow] || [];

  const found = new Set();
  for (let col = 2; col < header.length; col++) {
    const field = normStr(header[col]).toLowerCase();
    if (!field.startsWith("importo")) continue;

    const m = toNum(months[col]);
    if (m && Number.isInteger(m) && m >= 1 && m <= 12) found.add(m);
  }

  const missing = [];
  for (let m = 1; m <= 12; m++) if (!found.has(m)) missing.push(m);

  assert(
    missing.length === 0,
    "Sum_of_Importo: hai caricato il report GLOBALE (non mensilizzato). " +
    "Serve la versione con split mesi 1–12. Mesi mancanti: " + missing.join(", ")
  );

  return true;
}

// ============================================================
// LOGICA 1.0 (immutata)
// ============================================================
function kFromRatio(r) {
  if (r === null || r === undefined || Number.isNaN(r)) return null;
  if (r >= 0.9) return 1.0;
  if (r >= 0.8) return 0.8;
  if (r >= 0.7) return 0.5;
  return 0.0;
}
function clamp01(x) {
  if (x === null || x === undefined || Number.isNaN(x)) return null;
  if (x < 0) return 0;
  if (x > 1) return 1;
  return x;
}

// Premi_Mensili_x_Sede
function parsePremiMensili(aoa) {
  let headerIdx = -1;
  for (let i = 0; i < aoa.length; i++) {
    if (normStr(aoa[i][0]) === "Rows") { headerIdx = i; break; }
  }
  const monthRowIdx = headerIdx - 1;

  const headers = aoa[headerIdx].map(x => normStr(x));
  const monthRow = aoa[monthRowIdx];

  const monthStarts = [];
  for (let j = 0; j < monthRow.length; j++) {
    const m = toNum(monthRow[j]);
    if (m && Number.isInteger(m)) monthStarts.push([m, j]);
  }

  const idxRows = headers.indexOf("Rows");
  let currentSede = null;
  const out = [];

  for (let i = headerIdx + 1; i < aoa.length; i++) {
    const row = aoa[i];
    const key = normStr(row[idxRows]);

    if (key.toUpperCase().startsWith("EDAC_")) {
      currentSede = key;
      continue;
    }
    if (!currentSede) continue;

    if (key === "responsabili_area") {
      for (const [mese, startCol] of monthStarts) {
        const addon = toNum(row[startCol + 1]) ?? 0;
        const premioFinale = toNum(row[startCol + 2]) ?? 0;
        out.push({
          Sede: currentSede,
          Mese: mese,
          Addon: addon,
          VecchioCore: premioFinale - addon
        });
      }
    }
  }
  return out;
}

// Sum_of_Obiettivoprev (A,C,D,G,Q,U,Y)
function parseObiettivoprev(aoa) {
  const idxSede=0, idxAnno=2, idxMese=3, idxRamo=6, idxProd=16, idxFatt=20, idxInca=24;

  const map = new Map();
  for (let i = 0; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const sede = normStr(r[idxSede]);
    const ramo = normStr(r[idxRamo]).toUpperCase();
    if (!sede || ramo !== "CORE") continue;

    const anno = toNum(r[idxAnno]);
    const mese = toNum(r[idxMese]);
    if (!anno || !mese) continue;

    const prod = toNum(r[idxProd]) ?? 0;
    const fatt = toNum(r[idxFatt]) ?? 0;
    const inca = toNum(r[idxInca]) ?? 0;

    const key = `${sede}__${anno}__${mese}`;
    if (!map.has(key)) map.set(key, { Sede:sede, Anno:anno, Mese:mese, Prodotto:0, Fatturato:0, Incassato:0 });
    const obj = map.get(key);
    obj.Prodotto += prod;
    obj.Fatturato += fatt;
    obj.Incassato += inca;
  }
  return Array.from(map.values());
}

// Sum_of_Importo -> MNA by month
function parseMnaByMonth(aoa) {
  let headerRow = -1;
  for (let i = 0; i < aoa.length; i++) {
    if (normStr((aoa[i] || [])[0]) === "Sezione") { headerRow = i; break; }
  }
  const monthsRow = headerRow - 1;

  let mnaRow = -1;
  for (let i = 0; i < aoa.length; i++) {
    const label = normStr((aoa[i] || [])[1]).toUpperCase();
    if (label.includes("MNA - MARGINE NETTO DI AREA")) {
      if (!label.includes("TOTALE")) { mnaRow = i; break; }
      if (mnaRow < 0) mnaRow = i;
    }
  }

  const out = {};
  for (let col = 2; col < (aoa[headerRow] || []).length; col++) {
    const field = normStr((aoa[headerRow] || [])[col]).toLowerCase();
    if (!field.startsWith("importo")) continue;

    const mese = toNum((aoa[monthsRow] || [])[col]);
    if (!mese || !Number.isInteger(mese)) continue;

    out[mese] = toNum((aoa[mnaRow] || [])[col]);
  }
  return out;
}

// ============================================================
// OUTPUT Excel
// ============================================================
function hexRgb(hex) {
  const h = hex.replace("#","").trim();
  return { argb: "FF" + h.toUpperCase() };
}
function setFill(cell, hex) { cell.fill = { type:"pattern", pattern:"solid", fgColor: hexRgb(hex) }; }
function setMoney(cell)   { cell.numFmt = '€ #,##0.00'; }
function setPercent(cell) { cell.numFmt = '0.00%'; }
function setInt(cell)     { cell.numFmt = '0'; }
function setK(cell)       { cell.numFmt = '0.0'; }

function autoFitWorksheet(ws, minW=8, maxW=55, padding=2) {
  ws.columns.forEach(col => {
    let maxLen = 0;
    col.eachCell({ includeEmpty:true }, (cell) => {
      const v = cell.value;
      const s = (v === null || v === undefined) ? "" : String(v);
      maxLen = Math.max(maxLen, s.length);
    });
    col.width = Math.min(maxW, Math.max(minW, maxLen + padding));
  });
}

// ============================================================
// VALIDAZIONE LIVE (con anti-race)
// ============================================================
async function validateAllIfPossible() {
  const mySeq = ++VALIDATION_SEQ;

  setStatus("", "");
  setChecks([]);

  if (!isReadyFiles()) {
    readyBtn(false);
    return;
  }

  try {
    setStatus("Controllo file...", "");

    const [premi, sum, obi] = await Promise.all([
      readXlsxToAOA(elPremi.files[0]),
      readXlsxToAOA(elSum.files[0]),
      readXlsxToAOA(elObi.files[0]),
    ]);

    if (mySeq !== VALIDATION_SEQ) return;

    validatePremiAOA(premi.aoa);
    validateSumImportoAOA(sum.aoa);
    validateObiettiviAOA(obi.aoa);

    setChecks([
      "✅ Premi_Mensili_x_Sede OK",
      "✅ Sum_of_Importo OK (mensilizzato 1–12)",
      "✅ Sum_of_Obiettivoprev OK",
    ]);

    setStatus("✅ File OK. Puoi generare l’Excel.", "ok");
    readyBtn(true);
  } catch (e) {
    if (mySeq !== VALIDATION_SEQ) return;
    setChecks([]);
    setStatus("❌ " + (e?.message || String(e)), "err");
    readyBtn(false);
  }
}

elPremi.addEventListener("change", validateAllIfPossible);
elSum.addEventListener("change", validateAllIfPossible);
elObi.addEventListener("change", validateAllIfPossible);

// ============================================================
// RUN
// ============================================================
btnRun.addEventListener("click", async () => {
  try {
    readyBtn(false);
    setStatus("Elaboro...", "");

    const [premi, sum, obi] = await Promise.all([
      readXlsxToAOA(elPremi.files[0]),
      readXlsxToAOA(elSum.files[0]),
      readXlsxToAOA(elObi.files[0]),
    ]);

    // validazioni dure
    validatePremiAOA(premi.aoa);
    validateSumImportoAOA(sum.aoa);
    validateObiettiviAOA(obi.aoa);

    // parse (1.0)
    const premiRows = parsePremiMensili(premi.aoa);
    const obiRows   = parseObiettivoprev(obi.aoa);
    const mnaByMonth= parseMnaByMonth(sum.aoa);

    // coerenza sede (almeno una in comune)
    const sediPremi = new Set(premiRows.map(x => x.Sede));
    const sediObi   = new Set(obiRows.map(x => x.Sede));
    const intersect = [...sediPremi].some(s => sediObi.has(s));
    assert(intersect, "Mismatch: sedi in Premi non compaiono in Obiettivo (file non coerenti).");

    // indicizzo premi
    const premiMap = new Map();
    for (const r of premiRows) premiMap.set(`${r.Sede}__${r.Mese}`, r);

    const outRows = [];
    const sediSet = new Set();

    for (const r of obiRows) {
      const key = `${r.Sede}__${r.Mese}`;
      const p = premiMap.get(key) || { Addon: 0, VecchioCore: 0 };

      const incProd = (r.Prodotto && r.Prodotto !== 0) ? (r.Incassato / r.Prodotto) : null;
      const incProdCap = incProd === null ? null : clamp01(incProd);
      const K = incProdCap === null ? null : kFromRatio(incProdCap);

      const mna = (mnaByMonth[r.Mese] ?? null);

      outRows.push({
        Anno: r.Anno,
        Mese: r.Mese,
        Sede: r.Sede,
        Prodotto: r.Prodotto,
        Fatturato: r.Fatturato,
        Incassato: r.Incassato,
        MNA: mna,
        IncProd: incProd,
        K: K,
        Addon: p.Addon ?? 0,
        VecchioCore: p.VecchioCore ?? 0,
        BaseNew: 0,
        PremioNuovo: 0,
        DeltaCore: 0
      });
      sediSet.add(r.Sede);
    }

    const sedi = Array.from(sediSet);
    const singleSede = (sedi.length === 1) ? sedi[0] : "multi";

    // calcolo premio nuovo per sede (1.0)
    for (const sede of sedi) {
      const rowsSede = outRows.filter(x => x.Sede === sede).sort((a,b) => a.Mese - b.Mese);

      const mnaTotal = rowsSede.reduce((acc, x) => acc + (Number.isFinite(x.MNA) ? x.MNA : 0), 0);
      const target = 0.05 * mnaTotal;

      const sumPos = rowsSede.reduce((acc, x) => acc + ((Number.isFinite(x.MNA) && x.MNA > 0) ? x.MNA : 0), 0);

      for (const x of rowsSede) {
        x.BaseNew = 0;
        if (sumPos > 0 && Number.isFinite(x.MNA) && x.MNA > 0) {
          x.BaseNew = target * (x.MNA / sumPos);
        }
        const kVal = Number.isFinite(x.K) ? x.K : 0;
        x.PremioNuovo = (x.BaseNew || 0) * kVal;
        x.DeltaCore = x.PremioNuovo - (x.VecchioCore || 0);
      }
    }

    // build excel
    const wbOut = new ExcelJS.Workbook();
    const ws = wbOut.addWorksheet("Confronto");

    const headers = [
      "Anno","Mese","Sede",
      "Prodotto","Fatturato","Incassato",
      "MNA","Inc/Prod","K",
      "Addon","Vecchio core","Premio nuovo (core)","Δ core"
    ];
    ws.addRow(headers);
    ws.getRow(1).font = { bold: true };

    const GREEN = "#C6EFCE";
    const RED   = "#FFC7CE";

    const sorted = outRows.sort((a,b) => (a.Anno-b.Anno) || (a.Mese-b.Mese) || a.Sede.localeCompare(b.Sede));
    for (const x of sorted) {
      const row = ws.addRow([
        x.Anno, x.Mese, x.Sede,
        x.Prodotto, x.Fatturato, x.Incassato,
        x.MNA, x.IncProd, x.K,
        x.Addon, x.VecchioCore, x.PremioNuovo, x.DeltaCore
      ]);

      setInt(row.getCell(1));
      setInt(row.getCell(2));
      setMoney(row.getCell(4));
      setMoney(row.getCell(5));
      setMoney(row.getCell(6));
      setMoney(row.getCell(7));
      setPercent(row.getCell(8));
      setK(row.getCell(9));
      setMoney(row.getCell(10));
      setMoney(row.getCell(11));
      setMoney(row.getCell(12));
      setMoney(row.getCell(13));

      // semaforo + premio verde
      if (Number.isFinite(x.MNA)) setFill(row.getCell(7), x.MNA > 0 ? GREEN : RED);
      if (Number.isFinite(x.IncProd)) setFill(row.getCell(8), x.IncProd >= 0.7 ? GREEN : RED);
      if (Number.isFinite(x.K)) setFill(row.getCell(9), x.K > 0 ? GREEN : RED);
      if (Number.isFinite(x.PremioNuovo) && x.PremioNuovo > 0) setFill(row.getCell(12), GREEN);
    }

    autoFitWorksheet(ws);
    ws.getColumn(3).width = Math.max(ws.getColumn(3).width || 0, 18);

    const safe = singleSede.replace(/[\\/:*?"<>|]/g, "").replace(/\s+/g, "_");
    const outName = `confronto_${safe}_ra_mna.xlsx`;

    const buffer = await wbOut.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = outName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);

    setStatus(`✅ Creato: ${outName}`, "ok");
    readyBtn(true);
  } catch (e) {
    console.error(e);
    setStatus("❌ Errore: " + (e?.message || String(e)), "err");
    readyBtn(false);
  }
});
