const state = {
  files: {
    splitter: null,
    correction: null,
    merger: null,
  },
};

const selectors = {
  tabs: document.querySelectorAll(".tab-button"),
  panels: document.querySelectorAll(".tool-panel"),
  libraryStatus: document.querySelector("#libraryStatus"),
};

function setMessage(id, message, isError = false) {
  const element = document.querySelector(id);
  element.textContent = message;
  element.classList.toggle("is-error", isError);
}

function safeName(fileName, suffix) {
  const base = fileName.replace(/\.[^.]+$/, "").replace(/[^a-z0-9_-]+/gi, "_");
  return `${base}_${suffix}`;
}

function extensionOf(fileName) {
  return fileName.split(".").pop().toLowerCase();
}

function isExcel(file) {
  return ["xlsx", "xls"].includes(extensionOf(file.name));
}

function downloadBlob(blob, fileName) {
  if (window.saveAs) {
    window.saveAs(blob, fileName);
    return;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function sheetToRows(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" });
}

function rowsToCsv(rows) {
  return rows
    .map((row) =>
      row
        .map((value) => {
          const text = String(value ?? "");
          return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
        })
        .join(","),
    )
    .join("\n");
}

async function readRows(file) {
  const buffer = await file.arrayBuffer();
  if (isExcel(file)) {
    const workbook = XLSX.read(buffer, { type: "array" });
    return sheetToRows(workbook);
  }

  const text = new TextDecoder().decode(buffer);
  const workbook = XLSX.read(text, { type: "string" });
  return sheetToRows(workbook);
}

function workbookBlob(rows) {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Output");
  const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

async function buildSplitZip(file, size, keepHeader) {
  const rows = await readRows(file);
  if (!rows.length) throw new Error("This file has no rows to split.");

  const header = keepHeader ? rows[0] : null;
  const body = keepHeader ? rows.slice(1) : rows;
  if (!body.length) throw new Error("There are no data rows after the header.");

  const zip = new JSZip();
  const totalParts = Math.ceil(body.length / size);
  const outputAsExcel = isExcel(file);

  for (let index = 0; index < totalParts; index += 1) {
    const chunk = body.slice(index * size, (index + 1) * size);
    const partRows = header ? [header, ...chunk] : chunk;
    const partNumber = String(index + 1).padStart(2, "0");
    const name = `${safeName(file.name, `part_${partNumber}`)}.${outputAsExcel ? "xlsx" : "csv"}`;
    const content = outputAsExcel ? workbookBlob(partRows) : new Blob([rowsToCsv(partRows)], { type: "text/csv;charset=utf-8" });
    zip.file(name, content);
  }

  return {
    blob: await zip.generateAsync({ type: "blob" }),
    parts: totalParts,
    rows: body.length,
  };
}

async function splitFile(kind, rowsPerFile, keepHeader) {
  const file = state.files[kind];
  const resultId = kind === "correction" ? "#correctionResult" : "#splitResult";
  if (!file) {
    setMessage(resultId, "Choose a CSV or Excel file first.", true);
    return;
  }

  if (!window.XLSX || !window.JSZip) {
    setMessage(resultId, "File tools are still loading. Try again in a moment.", true);
    return;
  }

  try {
    setMessage(resultId, "Preparing your download...");
    const result = await buildSplitZip(file, rowsPerFile, keepHeader);
    downloadBlob(result.blob, `${safeName(file.name, `${rowsPerFile}_rows`)}.zip`);
    setMessage(resultId, `Done: ${result.rows} rows split into ${result.parts} files.`);
  } catch (error) {
    setMessage(resultId, error.message || "Could not split this file.", true);
  }
}

async function mergeColumns() {
  const file = state.files.merger;
  if (!file) {
    setMessage("#mergeResult", "Choose a CSV or Excel file first.", true);
    return;
  }

  if (!window.XLSX) {
    setMessage("#mergeResult", "File tools are still loading. Try again in a moment.", true);
    return;
  }

  try {
    setMessage("#mergeResult", "Merging columns...");
    const rows = await readRows(file);
    const joiner = document.querySelector("#mergeJoiner").value;
    const start = document.querySelector("#mergeSkipHeader").checked ? 1 : 0;

    for (let index = start; index < rows.length; index += 1) {
      const left = rows[index][0] ?? "";
      const right = rows[index][1] ?? "";
      rows[index][2] = [left, right].filter((value) => String(value).trim() !== "").join(joiner);
    }

    if (isExcel(file)) {
      downloadBlob(workbookBlob(rows), `${safeName(file.name, "merged")}.xlsx`);
    } else {
      downloadBlob(new Blob([rowsToCsv(rows)], { type: "text/csv;charset=utf-8" }), `${safeName(file.name, "merged")}.csv`);
    }

    setMessage("#mergeResult", `Done: column C updated for ${Math.max(rows.length - start, 0)} rows.`);
  } catch (error) {
    setMessage("#mergeResult", error.message || "Could not merge this file.", true);
  }
}

function convertLines() {
  const input = document.querySelector("#separatorInput").value;
  const separator = document.querySelector("#separatorValue").value || ",";
  const values = input
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  document.querySelector("#separatorOutput").value = values.join(separator);
  setMessage("#separatorResult", `Converted ${values.length} values.`);
}

async function copySeparatorOutput() {
  const output = document.querySelector("#separatorOutput").value;
  if (!output) {
    setMessage("#separatorResult", "Convert values before copying.", true);
    return;
  }

  await navigator.clipboard.writeText(output);
  setMessage("#separatorResult", "Copied to clipboard.");
}

function connectTabs() {
  selectors.tabs.forEach((button) => {
    button.addEventListener("click", () => {
      const tool = button.dataset.tool;
      selectors.tabs.forEach((item) => item.classList.toggle("is-active", item === button));
      selectors.panels.forEach((panel) => panel.classList.toggle("is-visible", panel.id === tool));
    });
  });
}

function connectDropZones() {
  document.querySelectorAll(".drop-zone").forEach((zone) => {
    const input = zone.querySelector("input");
    const kind = zone.dataset.drop;

    input.addEventListener("change", () => {
      state.files[kind] = input.files[0] || null;
      zone.querySelector("strong").textContent = state.files[kind]?.name || "Drop your CSV or Excel file";
    });

    ["dragenter", "dragover"].forEach((eventName) => {
      zone.addEventListener(eventName, (event) => {
        event.preventDefault();
        zone.classList.add("is-dragging");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      zone.addEventListener(eventName, (event) => {
        event.preventDefault();
        zone.classList.remove("is-dragging");
      });
    });

    zone.addEventListener("drop", (event) => {
      const file = event.dataTransfer.files[0];
      if (!file) return;
      state.files[kind] = file;
      zone.querySelector("strong").textContent = file.name;
    });
  });
}

function connectActions() {
  document.querySelector("#splitButton").addEventListener("click", () => {
    const rows = Number(document.querySelector("#splitRows").value);
    if (!Number.isFinite(rows) || rows < 1) {
      setMessage("#splitResult", "Enter a valid row count.", true);
      return;
    }
    splitFile("splitter", rows, document.querySelector("#splitHeader").checked);
  });

  document.querySelector("#correctionButton").addEventListener("click", () => {
    splitFile("correction", 999, true);
  });

  document.querySelector("#mergeButton").addEventListener("click", mergeColumns);
  document.querySelector("#convertButton").addEventListener("click", convertLines);
  document.querySelector("#separatorInput").addEventListener("input", convertLines);
  document.querySelector("#separatorValue").addEventListener("input", convertLines);
  document.querySelector("#copyButton").addEventListener("click", copySeparatorOutput);
}

function updateLibraryStatus() {
  if (window.XLSX && window.JSZip) {
    selectors.libraryStatus.innerHTML = '<i data-lucide="check-circle-2"></i> Ready';
    selectors.libraryStatus.classList.add("is-ready");
  } else {
    selectors.libraryStatus.innerHTML = '<i data-lucide="wifi-off"></i> File libraries unavailable';
  }

  if (window.lucide) lucide.createIcons();
}

connectTabs();
connectDropZones();
connectActions();
window.addEventListener("load", updateLibraryStatus);
setTimeout(updateLibraryStatus, 1500);
