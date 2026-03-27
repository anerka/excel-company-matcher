/* global XLSX */

function normalizeText(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function digitsOnly(s) {
  return String(s ?? "").replace(/\D/g, "");
}

/** Rad i A: valfritt prefix, första streck (-) och därefter organisationsnumret (jämförs mot varje cell i B). */
/** @param {unknown[][]} rows */
function extractOrgNumbersFromColumnA(rows) {
  /** @type {{ display: string; norm: string; digitKey: string }[]} */
  const out = [];
  for (const row of rows) {
    const cell = row[0];
    if (cell === undefined || cell === null || String(cell).trim() === "") continue;
    const text = String(cell).trim();
    const idx = text.indexOf("-");
    if (idx === -1) continue;
    const org = text.slice(idx + 1).trim();
    if (!org) continue;
    const norm = normalizeText(org);
    const digitKey = digitsOnly(org);
    out.push({ display: org, norm, digitKey });
  }
  return dedupeByOrg(out);
}

/** @param {{ display: string; norm: string; digitKey: string }[]} items */
function dedupeByOrg(items) {
  const seen = new Map();
  for (const n of items) {
    const key = n.digitKey.length >= 6 ? n.digitKey : n.norm;
    if (!key || seen.has(key)) continue;
    seen.set(key, n);
  }
  return [...seen.values()].map((n) => ({
    display: n.display,
    norm: n.norm,
    digitKey: n.digitKey,
  }));
}

/**
 * @param {unknown[]} row
 * @param {{ norm: string; display: string; digitKey: string }[]} orgList
 * @param {boolean} exactOnly
 */
function findMatchesInRow(row, orgList, exactOnly) {
  /** @type {{ display: string; colIndex: number }[]} */
  const hits = [];
  for (let c = 0; c < row.length; c++) {
    const raw = row[c];
    const cellNorm = normalizeText(raw);
    const cellDigits = digitsOnly(raw);
    if (!cellNorm && !cellDigits) continue;
    for (const { norm, display, digitKey } of orgList) {
      let ok = false;
      if (digitKey.length >= 6 && cellDigits.length >= digitKey.length) {
        ok = exactOnly
          ? cellDigits === digitKey
          : cellDigits === digitKey || cellDigits.includes(digitKey);
      }
      if (!ok && norm) {
        ok = exactOnly ? cellNorm === norm : cellNorm === norm || cellNorm.includes(norm);
      }
      if (ok) hits.push({ display, colIndex: c });
    }
  }
  const byOrg = new Map();
  for (const h of hits) {
    if (!byOrg.has(h.display)) byOrg.set(h.display, h.colIndex);
  }
  return [...byOrg.entries()].map(([display, colIndex]) => ({ display, colIndex }));
}

/** @param {File} file */
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        resolve(wb);
      } catch (e) {
        reject(e);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function firstSheetToAoA(wb) {
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
}

/** @param {unknown[][]} rows */
function padRowsToWidth(rows, width) {
  return rows.map((row) => {
    const r = [...row];
    while (r.length < width) r.push("");
    return r;
  });
}

function setStatus(el, text, isError) {
  el.hidden = false;
  el.textContent = text;
  el.classList.toggle("error", !!isError);
}

function main() {
  const fileA = document.getElementById("fileA");
  const fileB = document.getElementById("fileB");
  const runBtn = document.getElementById("runBtn");
  const status = document.getElementById("status");
  const downloadRow = document.getElementById("downloadRow");
  const downloadLink = document.getElementById("downloadLink");
  const exactOnly = document.getElementById("exactOnly");
  const skipFirstRowB = document.getElementById("skipFirstRowB");

  function updateRunEnabled() {
    runBtn.disabled = !(fileA.files?.length && fileB.files?.length);
  }

  fileA.addEventListener("change", updateRunEnabled);
  fileB.addEventListener("change", updateRunEnabled);

  let lastObjectUrl = null;

  runBtn.addEventListener("click", async () => {
    const fa = fileA.files?.[0];
    const fb = fileB.files?.[0];
    if (!fa || !fb) return;

    downloadRow.hidden = true;
    if (lastObjectUrl) {
      URL.revokeObjectURL(lastObjectUrl);
      lastObjectUrl = null;
    }

    setStatus(status, "Läser filer…", false);

    try {
      const [wbA, wbB] = await Promise.all([readWorkbook(fa), readWorkbook(fb)]);
      const rowsA = firstSheetToAoA(wbA);
      const rowsB = firstSheetToAoA(wbB);

      const extracted = extractOrgNumbersFromColumnA(rowsA);
      const orgList = [...extracted].sort(
        (a, b) =>
          Math.max(b.norm.length, b.digitKey.length) - Math.max(a.norm.length, a.digitKey.length)
      );
      if (orgList.length === 0) {
        setStatus(
          status,
          "Hittade inga organisationsnummer i fil A. Varje rad i kolumn A ska ha ett streck (-): text före första strecket ignoreras, text efter är organisationsnumret som jämförs mot B.",
          true
        );
        return;
      }

      const skipHeader = skipFirstRowB.checked;
      const headerRowB = skipHeader && rowsB.length > 0 ? rowsB[0] : null;
      const dataRowsB = skipHeader ? rowsB.slice(1) : rowsB;

      /** @type {unknown[][]} */
      const outRows = [];
      let matchCount = 0;

      for (let i = 0; i < dataRowsB.length; i++) {
        const row = dataRowsB[i];
        const matches = findMatchesInRow(row, orgList, exactOnly.checked);
        if (matches.length === 0) continue;
        matchCount += 1;
        const namesCell = matches.map((m) => m.display).join("; ");
        const colsCell = matches.map((m) => m.colIndex + 1).join("; ");
        const outRow = [namesCell, colsCell, ...row];
        outRows.push(outRow);
      }

      if (outRows.length === 0) {
        setStatus(
          status,
          `Inga träffar. Fil A hade ${orgList.length} unika organisationsnummer att söka efter. Prova \"innehåller\" (låt exakt vara avstängd) eller kontrollera format i B.`,
          false
        );
        return;
      }

      const extraHeader = [
        "Matchade org.nr från lista A",
        "Kolumn (1-baserat) där träff först hittades",
      ];
      /** @type {unknown[][]} */
      let finalAoA;
      if (skipHeader && headerRowB) {
        const w = Math.max(headerRowB.length, ...outRows.map((r) => r.length - 2));
        const paddedHeader = [...headerRowB];
        while (paddedHeader.length < w) paddedHeader.push("");
        finalAoA = [ [...extraHeader, ...paddedHeader], ...padRowsToWidth(outRows, w + 2) ];
      } else {
        const w = Math.max(0, ...outRows.map((r) => r.length));
        finalAoA = [ [...extraHeader, ...Array(Math.max(0, w - 2)).fill("")], ...padRowsToWidth(outRows, w) ];
      }

      const outWb = XLSX.utils.book_new();
      const outWs = XLSX.utils.aoa_to_sheet(finalAoA);
      XLSX.utils.book_append_sheet(outWb, outWs, "Träffar");

      const outBuf = XLSX.write(outWb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([outBuf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      lastObjectUrl = URL.createObjectURL(blob);
      downloadLink.href = lastObjectUrl;
      downloadLink.download = `matchning-resultat-${new Date().toISOString().slice(0, 10)}.xlsx`;

      setStatus(
        status,
          `Klart. ${orgList.length} unika organisationsnummer från A. ${matchCount} rader från B matchade (${outRows.length} rader i resultatet).`,
        false
      );
      downloadRow.hidden = false;
    } catch (e) {
      console.error(e);
      setStatus(
        status,
        `Ett fel uppstod: ${e instanceof Error ? e.message : String(e)}`,
        true
      );
    }
  });
}

main();