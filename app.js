const elPremi = document.getElementById("filePremi");
const elSum   = document.getElementById("fileSum");
const elObi   = document.getElementById("fileObi");
const btnRun  = document.getElementById("btnRun");
const statusEl= document.getElementById("status");

function setStatus(msg, kind="") {
  statusEl.className = "status " + (kind || "");
  statusEl.textContent = msg;
}

function ready() {
  btnRun.disabled = !(elPremi.files?.[0] && elSum.files?.[0] && elObi.files?.[0]);
}
elPremi.addEventListener("change", ready);
elSum.addEventListener("change", ready);
elObi.addEventListener("change", ready);

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

function isString(v) { return typeof v === "string"; }
function normStr(v) { return (v ?? "").toString().trim(); }
function toNum(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

// ------------------------
// PARSE: Premi_Mensili_x_Sede
// ------------------------
function parsePremiMensili(aoa) {
  // trova riga header dove col0 == "Rows"
  let headerIdx = -1;
  for (let i = 0; i < aoa.length; i++) {
    if (normStr(aoa[i][0]) === "Rows") { headerIdx = i; break; }
  }
  if (headerIdx < 0) throw new Error("Nel file Premi_Mensili_x_Sede non trovo header 'Rows'.");

  const monthRowIdx = headerIdx - 1;
  if (monthRowIdx < 0) throw new Error("Nel file Premi: manca la riga mesi sopra 'Rows'.");

  const headers = aoa[headerIdx].map(x => normStr(x));
  const monthRow = aoa[monthRowIdx];

  // mappa mese -> start col
  const monthStarts = [];
  for (let j = 0; j < monthRow.length; j++) {
    const m = toNum(monthRow[j]);
    if (m && Number.isInteger(m)) monthStarts.push([m, j]);
  }
  if (!monthStarts.length) throw new Error("Nel file Premi: non trovo i mesi (riga sopra 'Rows').");

  const idxRows = headers.indexOf("Rows");
  if (idxRows < 0) throw new Error("Nel file Premi: colonna 'Rows' non trovata.");

  // scorro righe dati
  let currentSede = null;
  const out = []; // {Sede,Mese,Addon,VecchioCore}
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
        // ordine blocco: PremioCalcolato, Addon, PremioFinalePostSaldoNegativo, Gettone
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

  if (!out.length) throw new Error("Nel file Premi: non trovo la riga 'responsabili_area' sotto una EDAC_*."); 
  return out;
}

// ------------------------
// PARSE: Sum_of_Obiettivoprev
// columns by letter (0-based):
// A=0, C=2, D=3, G=6, Q=16, U=20, Y=24
// ------------------------
function parseObiettivoprev(aoa) {
  const idxSede=0, idxAnno=2, idxMese=3, idxRamo=6, idxProd=16, idxFatt=20, idxInca=24;

  // aggrego per Sede/Anno/Mese SOLO CORE
  const map = new Map(); // key -> {Sede,Anno,Mese,Prodotto,Fatturato,Incassato}
  for (let i = 0; i < aoa.length; i++) {
    const r = aoa[i];
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
    if (!map.has(key)) {
      map.set(key, { Sede: sede, Anno: anno, Mese: mese, Prodotto: 0, Fatturato: 0, Incassato: 0 });
    }
    const obj = map.get(key);
    obj.Prodotto += prod;
    obj.Fatturato += fatt;
    obj.Incassato += inca;
  }

  return Array.from(map.values());
}

// ------------------------
// PARSE: Sum_of_Importo -> MNA by month
// replica la logica: trova riga "Sezione" in col A, headerRow = quella, monthsRow = headerRow-1
// cerca in col B una riga che contiene "MNA - Margine Netto di Area"
// legge solo colonne "Importo €" (sulla headerRow) e usa il mese dalla monthsRow
// ------------------------
function parseMnaByMonth(aoa) {
  let headerRow = -1;
  for (let i = 0; i < aoa.length; i++) {
    if (normStr(aoa[i][0]) === "Sezione") { headerRow = i; break; }
  }
  if (headerRow < 0) throw new Error("Sum_of_Importo: non trovo la tabella (riga con 'Sezione' in col A).");

  const monthsRow = headerRow - 1;
  if (monthsRow < 0) throw new Error("Sum_of_Importo: manca la riga dei mesi sopra 'Sezione'.");

  // trova riga MNA in col B (index 1)
  let mnaRow = -1;
  for (let i = 0; i < aoa.length; i++) {
    const label = normStr(aoa[i][1]).toUpperCase();
    if (label.includes("MNA - MARGINE NETTO DI AREA")) {
      // evita "Totale" se esistono più righe
      if (!label.includes("TOTALE")) { mnaRow = i; break; }
      if (mnaRow < 0) mnaRow = i;
    }
  }
  if (mnaRow < 0) throw new Error("Sum_of_Importo: non trovo la riga 'MNA - Margine Netto di Area'.");

  const out = {}; // mese -> value
  for (let col = 2; col < aoa[headerRow].length; col++) {
    const field = normStr(aoa[headerRow][col]).toLowerCase();
    if (!field.startsWith("importo")) continue;

    const mese = toNum(aoa[monthsRow][col]);
    if (!mese || !Number.isInteger(mese)) continue;

    out[mese] = toNum(aoa[mnaRow][col]);
  }
  return out;
}

// ------------------------
// EXCEL build
// ------------------------
function hexRgb(hex) {
  const h = hex.replace("#","").trim();
  return { argb: "FF" + h.toUpperCase() };
}

function setFill(cell, hex) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: hexRgb(hex) };
}

function setMoney(cell) { cell.numFmt = '€ #,##0.00'; }
function setPercent(cell){ cell.numFmt = '0.00%'; }
function setInt(cell){ cell.numFmt = '0'; }
function setK(cell){ cell.numFmt = '0.0'; }

function autoFitWorksheet(ws, minW=8, maxW=55, padding=2) {
  ws.columns.forEach(col => {
    let maxLen = 0;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const v = cell.value;
      const s = (v === null || v === undefined) ? "" : (typeof v === "object" && v.richText ? "" : String(v));
      maxLen = Math.max(maxLen, s.length);
    });
    let w = Math.min(maxW, Math.max(minW, maxLen + padding));
    col.width = w;
  });
}

btnRun.addEventListener("click", async () => {
  try {
    btnRun.disabled = true;
    setStatus("Elaboro...", "");

    const [premi, sum, obi] = await Promise.all([
      readXlsxToAOA(elPremi.files[0]),
      readXlsxToAOA(elSum.files[0]),
      readXlsxToAOA(elObi.files[0]),
    ]);

    const premiRows = parsePremiMensili(premi.aoa);          // array {Sede,Mese,Addon,VecchioCore}
    const obiRows   = parseObiettivoprev(obi.aoa);           // array {Sede,Anno,Mese,Prodotto,Fatturato,Incassato}
    const mnaByMonth= parseMnaByMonth(sum.aoa);              // {mese: MNA}

    // unisco per (Sede,Mese) (anno dall'obiettivo)
    // indicizzo premi per key
    const premiMap = new Map();
    for (const r of premiRows) premiMap.set(`${r.Sede}__${r.Mese}`, r);

    // costruisco output per ogni riga CORE
    const outRows = [];
    const sediSet = new Set();

    for (const r of obiRows) {
      const key = `${r.Sede}__${r.Mese}`;
      const p = premiMap.get(key) || { Addon: 0, VecchioCore: 0 };

      const incProd = (r.Prodotto && r.Prodotto !== 0) ? (r.Incassato / r.Prodotto) : null;
      const incProdCap = incProd === null ? null : clamp01(incProd);
      const K = incProdCap === null ? null : kFromRatio(incProdCap);

      const mna = mnaByMonth[r.Mese] ?? null;

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

    // (per ora: 1 sede → come il tuo uso; se ce ne sono più di una, produce multi)
    const sedi = Array.from(sediSet);
    const singleSede = (sedi.length === 1) ? sedi[0] : "multi";

    // calcolo premio nuovo per sede (riparto su mesi MNA>0)
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

    // ------- crea xlsx -------
    const wbOut = new ExcelJS.Workbook();
    const ws = wbOut.addWorksheet("Confronto");

    const headers = [
      "Anno","Mese","Sede",
      "Prodotto","Fatturato","Incassato",
      "MNA","Inc/Prod","K",
      "Addon","Vecchio core","Premio nuovo (core)","Δ core"
    ];
    ws.addRow(headers);

    // stile header
    ws.getRow(1).font = { bold: true };

    // colori
    const GREEN = "#C6EFCE";
    const RED   = "#FFC7CE";

    // righe dati
    for (const x of outRows.sort((a,b) => (a.Anno-b.Anno) || (a.Mese-b.Mese) || a.Sede.localeCompare(b.Sede))) {
      const row = ws.addRow([
        x.Anno, x.Mese, x.Sede,
        x.Prodotto, x.Fatturato, x.Incassato,
        x.MNA, x.IncProd, x.K,
        x.Addon, x.VecchioCore, x.PremioNuovo, x.DeltaCore
      ]);

      // formati
      setInt(row.getCell(1)); // Anno
      setInt(row.getCell(2)); // Mese
      setMoney(row.getCell(4));
      setMoney(row.getCell(5));
      setMoney(row.getCell(6));
      setMoney(row.getCell(7));     // MNA
      setPercent(row.getCell(8));   // Inc/Prod
      setK(row.getCell(9));         // K
      setMoney(row.getCell(10));    // Addon
      setMoney(row.getCell(11));    // Vecchio core
      setMoney(row.getCell(12));    // Premio nuovo
      setMoney(row.getCell(13));    // Delta

      // semaforo richiesto:
      // MNA verde se >0 altrimenti rosso
      if (Number.isFinite(x.MNA)) {
        setFill(row.getCell(7), x.MNA > 0 ? GREEN : RED);
      }

      // Inc/Prod verde se >=0.70 altrimenti rosso
      if (Number.isFinite(x.IncProd)) {
        setFill(row.getCell(8), x.IncProd >= 0.7 ? GREEN : RED);
      }

      // K verde se >0 altrimenti rosso
      if (Number.isFinite(x.K)) {
        setFill(row.getCell(9), x.K > 0 ? GREEN : RED);
      }

      // Premio nuovo (core) verde se >0
      if (Number.isFinite(x.PremioNuovo)) {
        if (x.PremioNuovo > 0) setFill(row.getCell(12), GREEN);
      }
    }

    // autofit + patch colonna Sede (C)
    autoFitWorksheet(ws);
    ws.getColumn(3).width = Math.max(ws.getColumn(3).width || 0, 18);

    // nome file output
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
  } catch (e) {
    console.error(e);
    setStatus("❌ Errore: " + (e?.message || String(e)), "err");
  } finally {
    btnRun.disabled = false;
  }
});
