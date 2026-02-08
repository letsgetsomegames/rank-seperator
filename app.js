/* global XLSX */

const fileInput = document.getElementById("file");
const drop = document.getElementById("drop");
const newerSelect = document.getElementById("newer");
const olderSelect = document.getElementById("older");
const rerankBox = document.getElementById("rerank");
const runBtn = document.getElementById("run");
const note = document.getElementById("note");

let workbook = null;

function setNote(msg) {
  note.textContent = msg || "";
}

function normalizeCol(col) {
  return String(col || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/-/g, "");
}

function findCol(headers, targets) {
  const map = new Map();
  headers.forEach((h) => map.set(normalizeCol(h), h));
  for (const t of targets) {
    const key = normalizeCol(t);
    if (map.has(key)) return map.get(key);
  }
  return null;
}

function toNumber(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function findHeaderRow(rows) {
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i].map((v) => String(v || "").trim().toLowerCase());
    const hasTeam = row.some((v) => v.includes("team"));
    const hasNumber = row.some((v) => v.includes("number") || v === "team #" || v === "#");
    if (hasTeam && hasNumber) return i;
  }
  return 0;
}

function sheetToTable(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  const headerRow = findHeaderRow(rows);
  const headers = rows[headerRow] || [];
  const data = rows.slice(headerRow + 1).filter((r) => r.some((v) => v !== null && v !== undefined && v !== ""));
  return { headers, data };
}

function buildRecords(headers, data) {
  return data.map((row) => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });
}

function getTeamNumberCol(headers) {
  const found = headers.find((h) => {
    const s = String(h || "").toLowerCase();
    return s.includes("team") && (s.includes("number") || s.includes("#"));
  });
  return found || headers[1];
}

function createOutput() {
  const newer = newerSelect.value;
  const older = olderSelect.value;
  if (!newer || !older || !workbook) return;

  const newSheet = workbook.Sheets[newer];
  const oldSheet = workbook.Sheets[older];
  const newTable = sheetToTable(newSheet);
  const oldTable = sheetToTable(oldSheet);
  const newRecords = buildRecords(newTable.headers, newTable.data);
  const oldRecords = buildRecords(oldTable.headers, oldTable.data);

  const teamNewCol = getTeamNumberCol(newTable.headers);
  const teamOldCol = getTeamNumberCol(oldTable.headers);

  const keepCols = newTable.headers.slice(0, 3);
  const statCols = newTable.headers.slice(3);

  const oldByTeam = new Map();
  oldRecords.forEach((r) => oldByTeam.set(r[teamOldCol], r));

  const output = newRecords.map((r) => {
    const out = {};
    keepCols.forEach((c) => (out[c] = r[c]));
    statCols.forEach((c) => {
      const old = oldByTeam.get(r[teamNewCol]);
      const oldVal = old ? old[c] : 0;
      out[c] = toNumber(r[c]) - toNumber(oldVal);
    });
    return out;
  });

  const winsCol = findCol(statCols, ["wins", "w"]);
  const lossesCol = findCol(statCols, ["losses", "loss", "l"]);
  const tiesCol = findCol(statCols, ["ties", "tie", "t"]);
  const matchesCol = findCol(statCols, ["matches played", "matches", "played"]);
  const totalPointsCol = findCol(statCols, ["total points", "points total", "tp"]);
  const winPctCol = findCol(statCols, ["win percentage", "win %", "win%"]);
  const avgPointsCol = findCol(statCols, ["average points", "avg points", "avg pts", "average pts"]);
  const highScoreCol = findCol(statCols, ["high score", "highscore", "hs"]);
  const autoCol = findCol(statCols, ["autonomous points", "auton points", "auto points", "autonomous"]);
  const strengthCol = findCol(statCols, ["strength points", "strength", "sp"]);
  const rankCol = findCol(newTable.headers, ["rank", "ranking", "place"]);

  if (highScoreCol) {
    output.forEach((o, i) => {
      const old = oldByTeam.get(newRecords[i][teamNewCol]);
      const oldVal = old ? old[highScoreCol] : 0;
      o[highScoreCol] = Math.max(toNumber(newRecords[i][highScoreCol]), toNumber(oldVal));
    });
  }

  if (winsCol && lossesCol && tiesCol && matchesCol && winPctCol) {
    output.forEach((o) => {
      const wins = toNumber(o[winsCol]);
      const ties = toNumber(o[tiesCol]);
      const matches = toNumber(o[matchesCol]);
      const pct = matches > 0 ? ((wins + 0.5 * ties) / matches) * 100 : 0;
      o[winPctCol] = pct;
    });
  }

  if (totalPointsCol && matchesCol && avgPointsCol) {
    output.forEach((o) => {
      const total = toNumber(o[totalPointsCol]);
      const matches = toNumber(o[matchesCol]);
      o[avgPointsCol] = matches > 0 ? total / matches : 0;
    });
  }

  if (rerankBox.checked && winPctCol) {
    output.sort((a, b) => {
      const w = toNumber(b[winPctCol]) - toNumber(a[winPctCol]);
      if (w !== 0) return w;
      if (autoCol) {
        const aDiff = toNumber(b[autoCol]) - toNumber(a[autoCol]);
        if (aDiff !== 0) return aDiff;
      }
      if (strengthCol) {
        const sDiff = toNumber(b[strengthCol]) - toNumber(a[strengthCol]);
        if (sDiff !== 0) return sDiff;
      }
      return 0;
    });

    if (rankCol) {
      output.forEach((o, i) => (o[rankCol] = i + 1));
    } else {
      output.forEach((o, i) => (o.Rank = i + 1));
      keepCols.unshift("Rank");
    }
  }

  const outHeaders = keepCols.concat(statCols);
  const outRows = output.map((o) => outHeaders.map((h) => o[h]));

  const outSheet = XLSX.utils.aoa_to_sheet([outHeaders, ...outRows]);
  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, outSheet, "New Competition Only");

  XLSX.writeFile(outWb, "new_competition_only.xlsx");
  setNote("Downloaded new_competition_only.xlsx");
}

function fillSelects(names) {
  newerSelect.innerHTML = "";
  olderSelect.innerHTML = "";
  names.forEach((n) => {
    const opt1 = document.createElement("option");
    opt1.value = n;
    opt1.textContent = n;
    newerSelect.appendChild(opt1);

    const opt2 = document.createElement("option");
    opt2.value = n;
    opt2.textContent = n;
    olderSelect.appendChild(opt2);
  });
  newerSelect.disabled = false;
  olderSelect.disabled = false;
  runBtn.disabled = false;
  if (names.length > 1) {
    newerSelect.selectedIndex = 0;
    olderSelect.selectedIndex = 1;
  }
}

function loadFile(file) {
  if (!file || !file.name.toLowerCase().endsWith(".xlsx")) {
    setNote("Please select a .xlsx file.");
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: "array" });
    fillSelects(workbook.SheetNames);
    setNote(`Loaded ${file.name}`);
  };
  reader.readAsArrayBuffer(file);
}

fileInput.addEventListener("change", (e) => loadFile(e.target.files[0]));

drop.addEventListener("dragover", (e) => {
  e.preventDefault();
  drop.classList.add("dragover");
});
drop.addEventListener("dragleave", () => drop.classList.remove("dragover"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("dragover");
  const file = e.dataTransfer.files[0];
  loadFile(file);
});

runBtn.addEventListener("click", createOutput);
