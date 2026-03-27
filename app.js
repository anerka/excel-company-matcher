/* global XLSX */

function normalizeText(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/** Normaliserade namn i A som är platshållare och inte ska sökas i B (t.ex. dummy-rader). */
const SKIPPED_LIST_A_NAMES = new Set(["aaa"]);

/** @param {string} norm lowercased & trimmad enligt normalizeText */
function shouldSkipListAName(norm) {
  return !norm || SKIPPED_LIST_A_NAMES.has(norm);
}

/** Kolumn A: prefix (kan innehålla egna streck), sedan ` - `, sedan företagsnamnet som matchas mot B. */
/** @param {unknown[][]} rows */
function extractCompanyNamesFromColumnA(rows) {
  /** @type {{ display: string; norm: string }[]} */
  const out = [];
  const sep = " - ";
  for (const row of rows) {
    const cell = row[0];
    if (cell === undefined || cell === null || String(cell).trim() === "") continue;
    const text = String(cell).trim();
    const idx = text.indexOf(sep);
    if (idx === -1) continue;
    const name = text.slice(idx + sep.length).trim();
    const norm = normalizeText(name);
    if (shouldSkipListAName(norm)) continue;
    out.push({ display: name, norm });
  }
  return dedupeByNorm(out);
}

/** @param {{ display: string; norm: string }[]} names */
function dedupeByNorm(names) {
  const seen = new Map();
  for (const n of names) {
    if (!n.norm || seen.has(n.norm)) continue;
    seen.set(n.norm, n.display);
  }
  return [...seen.entries()].map(([norm, display]) => ({ norm, display }));
}

/**
 * @param {unknown[]} row
 * @param {{ norm: string; display: string }[]} nameList
 * @param {boolean} exactOnly
 */
function findMatchesInRow(row, nameList, exactOnly) {
  /** @type {{ display: string; colIndex: number }[]} */
  const hits = [];
  for (let c = 0; c < row.length; c++) {
    const cellNorm = normalizeText(row[c]);
    if (!cellNorm) continue;
    for (const { norm, display } of nameList) {
      if (!norm) continue;
      const ok = exactOnly ? cellNorm === norm : cellNorm === norm || cellNorm.includes(norm);
      if (ok) hits.push({ display, colIndex: c });
    }
  }
  const byName = new Map();
  for (const h of hits) {
    if (!byName.has(h.display)) byName.set(h.display, h.colIndex);
  }
  return [...byName.entries()].map(([display, colIndex]) => ({ display, colIndex }));
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

      const extracted = extractCompanyNamesFromColumnA(rowsA);
      const nameList = [...extracted].sort((a, b) => b.norm.length - a.norm.length);
      if (nameList.length === 0) {
        setStatus(
          status,
          "Hittade inga företagsnamn i fil A. Använd mellanslag runt strecket: \"… - Företagsnamn\" (prefixet får ha streck, t.ex. 91000-003-00 - Namn AB).",
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
        const matches = findMatchesInRow(row, nameList, exactOnly.checked);
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
          `Inga träffar. Fil A hade ${nameList.length} unika namn att söka efter. Prova \"innehåller\" (låt exakt vara avstängd) eller kontrollera stavning.`,
          false
        );
        return;
      }

      const extraHeader = ["Matchade namn från lista A", "Kolumn (1-baserat) där träff först hittades"];
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
          `Klart. ${nameList.length} unika namn från A. ${matchCount} rader från B matchade (${outRows.length} rader i resultatet).`,
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