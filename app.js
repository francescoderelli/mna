<!doctype html>
<html lang="it">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Confronto premi RA (MNA + Inc/Prod)</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:24px;max-width:920px}
    .box{border:1px solid #ddd;border-radius:12px;padding:16px;margin:12px 0}
    label{display:block;margin:10px 0 6px;font-weight:700}
    input[type=file]{width:100%}
    button{padding:10px 14px;border:0;border-radius:10px;background:#111;color:#fff;font-weight:800;cursor:pointer}
    button:disabled{opacity:.45;cursor:not-allowed}
    .status{margin-top:12px;font-family:ui-monospace,Menlo,Consolas,monospace;white-space:pre-wrap}
    .ok{color:#0a7a2f}
    .err{color:#b00020}
    .hint{color:#555;font-size:14px;line-height:1.4}
    .checks{margin-top:10px;font-family:ui-monospace,Menlo,Consolas,monospace}
    .checks div{margin:4px 0}
  </style>
</head>
<body>

  <h1>Confronto premi RA</h1>
  <p class="hint">Tutto gira nel browser. I file non vengono caricati online.</p>

  <div class="box">
    <label>1) Premi_Mensili_x_Sede (xlsx)</label>
    <input id="filePremi" type="file" accept=".xlsx" />

    <label>2) Sum_of_Importo (xlsx)</label>
    <input id="fileSum" type="file" accept=".xlsx" />

    <label>3) Sum_of_Obiettivoprev (xlsx)</label>
    <input id="fileObi" type="file" accept=".xlsx" />

    <div class="checks" id="checks"></div>
  </div>

  <div class="box">
    <button id="btnRun" disabled>Genera Excel confronto</button>
    <div id="status" class="status"></div>
  </div>

  <!-- SheetJS (lettura xlsx) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- ExcelJS (scrittura xlsx) -->
  <script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>

  <script src="app.js"></script>
</body>
</html>
