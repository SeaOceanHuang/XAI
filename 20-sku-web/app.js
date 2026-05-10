const STORAGE_KEY = "sku-weight-check:v1";
const DEFAULT_THRESHOLD = 20;

const state = {
  estimates: new Map(),
  records: [],
  threshold: DEFAULT_THRESHOLD,
  filter: "all",
};

const el = {
  estimateFile: document.querySelector("#estimateFile"),
  actualFile: document.querySelector("#actualFile"),
  estimateStatus: document.querySelector("#estimateStatus"),
  estimateCount: document.querySelector("#estimateCount"),
  actualCount: document.querySelector("#actualCount"),
  exceptionCount: document.querySelector("#exceptionCount"),
  thresholdInput: document.querySelector("#thresholdInput"),
  weighForm: document.querySelector("#weighForm"),
  skuInput: document.querySelector("#skuInput"),
  actualWeightInput: document.querySelector("#actualWeightInput"),
  lastResult: document.querySelector("#lastResult"),
  recordBody: document.querySelector("#recordBody"),
  exceptionList: document.querySelector("#exceptionList"),
  emptyExceptions: document.querySelector("#emptyExceptions"),
  exportExceptionsBtn: document.querySelector("#exportExceptionsBtn"),
  exportAllBtn: document.querySelector("#exportAllBtn"),
  downloadTemplateBtn: document.querySelector("#downloadTemplateBtn"),
  clearBtn: document.querySelector("#clearBtn"),
  modal: document.querySelector("#alertModal"),
  modalMessage: document.querySelector("#modalMessage"),
  modalCloseBtn: document.querySelector("#modalCloseBtn"),
  modalExportBtn: document.querySelector("#modalExportBtn"),
  filterButtons: [...document.querySelectorAll("[data-filter]")],
};

const skuHeaders = [
  "sku",
  "sku编码",
  "sku编号",
  "平台sku",
  "店铺sku",
  "商品sku",
  "商品编码",
  "商品编号",
  "产品编码",
  "产品编号",
  "货号",
  "seller sku",
  "seller_sku",
  "msku",
];
const estimateHeaders = [
  "预估重量",
  "预估重量g",
  "预估重量(g)",
  "预计重量",
  "预计重量(g)",
  "估重",
  "商品重量",
  "商品重量(g)",
  "产品重量",
  "产品重量(g)",
  "包裹重量",
  "包裹重量(g)",
  "重量",
  "重量(g)",
  "estimated_weight",
  "estimate weight",
  "estimated weight",
  "weight",
];
const actualHeaders = [
  "实际重量",
  "实际重量g",
  "实际重量(g)",
  "实测重量",
  "实测重量(g)",
  "称重重量",
  "称重重量(g)",
  "商品重量",
  "商品重量(g)",
  "产品重量",
  "产品重量(g)",
  "重量",
  "重量(g)",
  "real_weight",
  "actual_weight",
  "actual weight",
  "weight",
];

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/^\uFEFF/, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[_-]+/g, "")
    .replace(/[（]/g, "(")
    .replace(/[）]/g, ")")
    .replace(/克/g, "g")
    .replace(/[：:]/g, "");
}

function normalizeSku(value) {
  return String(value ?? "").trim();
}

function parseWeight(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const cleaned = String(value ?? "")
    .trim()
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatWeight(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "-";
  return Number(value).toFixed(1).replace(/\.0$/, "");
}

function findColumn(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader);
  const normalizedHeaders = headers.map(normalizeHeader);
  let index = normalizedHeaders.findIndex((header) => normalizedCandidates.includes(header));
  if (index >= 0) return headers[index];
  index = normalizedHeaders.findIndex((header) =>
    normalizedCandidates.some((candidate) => header.includes(candidate))
  );
  if (index >= 0) return headers[index];
  index = normalizedHeaders.findIndex((header) =>
    candidates === skuHeaders
      ? header.includes("sku") || header.includes("编码") || header.includes("货号")
      : header.includes("重量") || header.includes("weight")
  );
  return index >= 0 ? headers[index] : null;
}

function looksLikeHeaderRow(row) {
  const headers = row.map((cell) => String(cell ?? "").trim()).filter(Boolean);
  if (!headers.length) return false;
  return Boolean(findColumn(headers, skuHeaders) && findColumn(headers, [...estimateHeaders, ...actualHeaders]));
}

function rowsToObjects(rows) {
  const headerRowIndex = rows.findIndex((row, index) => index < 20 && looksLikeHeaderRow(row));
  const fallbackIndex = rows.findIndex((row) => row.some((cell) => String(cell ?? "").trim()));
  const index = headerRowIndex >= 0 ? headerRowIndex : fallbackIndex;
  if (index < 0) return [];

  const headers = rows[index].map((cell) => String(cell ?? "").trim());
  return rows.slice(index + 1).map((values) =>
    headers.reduce((item, header, columnIndex) => {
      if (header) item[header] = values[columnIndex] ?? "";
      return item;
    }, {})
  );
}

async function readTable(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "xlsx") {
    return readXlsx(file);
  }

  if (ext === "xls") {
    throw new Error("离线版本暂不支持旧版 .xls，请另存为 .xlsx 或 .csv 后导入。");
  }

  const text = await file.text();
  return parseCsv(text);
}

async function readXlsx(file) {
  if (!("DecompressionStream" in window)) {
    throw new Error("当前浏览器不支持离线解析 XLSX，请使用新版 Chrome / Edge，或另存为 CSV 导入。");
  }

  const entries = await unzipEntries(await file.arrayBuffer());
  const workbook = parseXml(await getZipText(entries, "xl/workbook.xml"));
  const workbookRels = parseRels(await getZipText(entries, "xl/_rels/workbook.xml.rels"));
  const sharedStrings = entries.has("xl/sharedStrings.xml")
    ? parseSharedStrings(await getZipText(entries, "xl/sharedStrings.xml"))
    : [];
  const sheetPath = getFirstSheetPath(workbook, workbookRels);
  return parseSheet(await getZipText(entries, sheetPath), sharedStrings);
}

async function unzipEntries(buffer) {
  const view = new DataView(buffer);
  const bytes = new Uint8Array(buffer);
  const eocdOffset = findEndOfCentralDirectory(bytes);
  const entryCount = view.getUint16(eocdOffset + 10, true);
  const centralDirOffset = view.getUint32(eocdOffset + 16, true);
  const entries = new Map();
  let offset = centralDirOffset;

  for (let i = 0; i < entryCount; i += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error("XLSX 文件结构异常，无法读取目录。");
    }

    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const nameBytes = bytes.slice(offset + 46, offset + 46 + fileNameLength);
    const name = new TextDecoder().decode(nameBytes).replace(/\\/g, "/");

    entries.set(name, {
      method,
      compressedSize,
      localHeaderOffset,
      buffer,
    });

    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function findEndOfCentralDirectory(bytes) {
  const minOffset = Math.max(0, bytes.length - 65557);
  for (let i = bytes.length - 22; i >= minOffset; i -= 1) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) {
      return i;
    }
  }
  throw new Error("这不是有效的 XLSX 文件。");
}

async function getZipText(entries, path) {
  const entry = entries.get(path);
  if (!entry) throw new Error(`XLSX 缺少必要文件：${path}`);
  const view = new DataView(entry.buffer);
  const bytes = new Uint8Array(entry.buffer);
  const localOffset = entry.localHeaderOffset;

  if (view.getUint32(localOffset, true) !== 0x04034b50) {
    throw new Error("XLSX 文件内容异常，无法读取。");
  }

  const fileNameLength = view.getUint16(localOffset + 26, true);
  const extraLength = view.getUint16(localOffset + 28, true);
  const dataStart = localOffset + 30 + fileNameLength + extraLength;
  const compressed = bytes.slice(dataStart, dataStart + entry.compressedSize);

  let output;
  if (entry.method === 0) {
    output = compressed;
  } else if (entry.method === 8) {
    output = new Uint8Array(
      await new Response(new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).arrayBuffer()
    );
  } else {
    throw new Error("XLSX 使用了当前离线解析器不支持的压缩格式。");
  }

  return new TextDecoder("utf-8").decode(output);
}

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("XLSX XML 解析失败。");
  return doc;
}

function parseRels(text) {
  const doc = parseXml(text);
  const rels = new Map();
  doc.querySelectorAll("Relationship").forEach((node) => {
    rels.set(node.getAttribute("Id"), node.getAttribute("Target"));
  });
  return rels;
}

function getFirstSheetPath(workbook, rels) {
  const sheet = workbook.querySelector("sheet");
  if (!sheet) throw new Error("XLSX 没有可读取的工作表。");
  const relId = sheet.getAttribute("r:id") || sheet.getAttribute("id");
  const target = rels.get(relId);
  if (!target) throw new Error("XLSX 工作表关系缺失。");
  if (target.startsWith("/")) return target.slice(1);
  return `xl/${target}`.replace(/\/+/g, "/");
}

function parseSharedStrings(text) {
  const doc = parseXml(text);
  return [...doc.querySelectorAll("si")].map((item) =>
    [...item.querySelectorAll("t")]
      .map((node) => node.textContent ?? "")
      .join("")
  );
}

function parseSheet(text, sharedStrings) {
  const doc = parseXml(text);
  const rows = [...doc.querySelectorAll("sheetData row")].map((row) => {
    const cells = [];
    let fallbackIndex = 0;
    row.querySelectorAll("c").forEach((cell) => {
      const ref = cell.getAttribute("r");
      const index = ref ? columnNameToIndex(ref.replace(/\d+/g, "")) : fallbackIndex;
      cells[index] = getCellValue(cell, sharedStrings);
      fallbackIndex = index + 1;
    });
    return cells;
  });

  return rowsToObjects(rows);
}

function getCellValue(cell, sharedStrings) {
  const type = cell.getAttribute("t");
  if (type === "s") {
    const index = Number(cell.querySelector("v")?.textContent ?? "");
    return sharedStrings[index] ?? "";
  }
  if (type === "inlineStr") {
    return [...cell.querySelectorAll("t")].map((node) => node.textContent ?? "").join("");
  }
  if (type === "b") {
    return cell.querySelector("v")?.textContent === "1" ? "TRUE" : "FALSE";
  }
  return cell.querySelector("v")?.textContent ?? "";
}

function columnNameToIndex(name) {
  return [...name.toUpperCase()].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (char === '"' && quoted && next === '"') {
      field += '"';
      i += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(field);
      field = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(field);
      if (row.some((item) => String(item).trim())) rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }

  row.push(field);
  if (row.some((item) => String(item).trim())) rows.push(row);
  if (!rows.length) return [];

  return rowsToObjects(rows);
}

function importEstimates(rows) {
  if (!rows.length) throw new Error("表格中没有可读取的数据。");
  const headers = Object.keys(rows[0]);
  const skuColumn = findColumn(headers, skuHeaders);
  const weightColumn = findColumn(headers, estimateHeaders);
  if (!skuColumn || !weightColumn) {
    throw new Error("未找到 SKU 或预估重量列，请检查表头。");
  }

  let imported = 0;
  rows.forEach((row) => {
    const sku = normalizeSku(row[skuColumn]);
    const weight = parseWeight(row[weightColumn]);
    if (sku && weight !== null) {
      state.estimates.set(sku, weight);
      imported += 1;
    }
  });

  if (!imported) throw new Error("没有导入有效的 SKU 重量。");
  syncRecordsWithEstimates();
  saveState();
  render();
  return imported;
}

function importActuals(rows) {
  if (!rows.length) throw new Error("表格中没有可读取的数据。");
  const headers = Object.keys(rows[0]);
  const skuColumn = findColumn(headers, skuHeaders);
  const weightColumn = findColumn(headers, actualHeaders);
  if (!skuColumn || !weightColumn) {
    throw new Error("未找到 SKU 或实际重量列，请检查表头。");
  }

  let imported = 0;
  let firstException = null;
  rows.forEach((row) => {
    const record = createRecord(row[skuColumn], row[weightColumn]);
    if (record) {
      state.records.unshift(record);
      imported += 1;
      if (!firstException && record.isException) firstException = record;
    }
  });

  if (!imported) throw new Error("没有导入有效的称重记录。");
  saveState();
  render();
  if (firstException) showExceptionModal(firstException);
  return imported;
}

function createRecord(skuValue, actualValue) {
  const sku = normalizeSku(skuValue);
  const actual = parseWeight(actualValue);
  if (!sku || actual === null) return null;

  const estimate = state.estimates.get(sku);
  const diff = estimate === undefined ? null : actual - estimate;
  return {
    id: crypto.randomUUID(),
    sku,
    estimate: estimate ?? null,
    actual,
    diff,
    isException: diff !== null && Math.abs(diff) >= state.threshold,
    createdAt: new Date().toISOString(),
  };
}

function refreshRecordFlags() {
  state.records = state.records.map((record) => ({
    ...record,
    isException: record.diff !== null && Math.abs(record.diff) >= state.threshold,
  }));
}

function syncRecordsWithEstimates() {
  state.records = state.records.map((record) => {
    const estimate = state.estimates.get(record.sku);
    const diff = estimate === undefined ? null : record.actual - estimate;
    return {
      ...record,
      estimate: estimate ?? null,
      diff,
      isException: diff !== null && Math.abs(diff) >= state.threshold,
    };
  });
}

function getExceptions() {
  return state.records.filter((record) => record.isException);
}

function saveState() {
  const payload = {
    estimates: [...state.estimates.entries()],
    records: state.records,
    threshold: state.threshold,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;
  try {
    const payload = JSON.parse(raw);
    state.estimates = new Map(payload.estimates ?? []);
    state.records = payload.records ?? [];
    state.threshold = Number(payload.threshold) || DEFAULT_THRESHOLD;
    refreshRecordFlags();
  } catch {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function render() {
  refreshRecordFlags();
  const exceptions = getExceptions();
  el.estimateCount.textContent = state.estimates.size;
  el.actualCount.textContent = state.records.length;
  el.exceptionCount.textContent = exceptions.length;
  el.thresholdInput.value = state.threshold;

  el.estimateStatus.textContent = state.estimates.size ? `已导入 ${state.estimates.size}` : "待导入";
  el.estimateStatus.className = `status-pill ${state.estimates.size ? "good" : ""}`;

  renderExceptions(exceptions);
  renderRecords();
}

function renderExceptions(exceptions) {
  el.emptyExceptions.hidden = exceptions.length > 0;
  el.exceptionList.innerHTML = exceptions
    .slice(0, 30)
    .map(
      (record) => `
        <article class="exception-card">
          <div>
            <strong>${escapeHtml(record.sku)}</strong>
            <small>预估 ${formatWeight(record.estimate)}g / 实际 ${formatWeight(record.actual)}g</small>
          </div>
          <span>${formatWeight(Math.abs(record.diff))}g</span>
        </article>
      `
    )
    .join("");
}

function renderRecords() {
  const filtered = state.records.filter((record) => {
    if (state.filter === "exception") return record.isException;
    if (state.filter === "normal") return !record.isException;
    return true;
  });

  if (!filtered.length) {
    el.recordBody.innerHTML = `<tr><td colspan="6" class="empty-cell">暂无匹配记录</td></tr>`;
    return;
  }

  el.recordBody.innerHTML = filtered
    .map((record) => {
      const unknown = record.estimate === null;
      const statusText = unknown ? "未匹配" : record.isException ? "异常" : "正常";
      const statusClass = unknown ? "neutral" : record.isException ? "bad" : "good";
      const diffText = record.diff === null ? "-" : formatWeight(record.diff);
      return `
        <tr>
          <td class="sku-cell">${escapeHtml(record.sku)}</td>
          <td>${formatWeight(record.estimate)}</td>
          <td>${formatWeight(record.actual)}</td>
          <td class="${record.isException ? "diff-bad" : ""}">${diffText}</td>
          <td><span class="status-tag ${statusClass}">${statusText}</span></td>
          <td>${formatDate(record.createdAt)}</td>
        </tr>
      `;
    })
    .join("");
}

function setLastResult(message, type = "neutral") {
  el.lastResult.textContent = message;
  el.lastResult.className = `status-pill ${type}`;
}

function showExceptionModal(record) {
  el.modalMessage.textContent = `${record.sku} 预估 ${formatWeight(record.estimate)}g，实际 ${formatWeight(
    record.actual
  )}g，误差 ${formatWeight(Math.abs(record.diff))}g。`;
  el.modal.classList.remove("hidden");
}

function hideModal() {
  el.modal.classList.add("hidden");
}

function exportCsv(records, filename) {
  if (!records.length) {
    alert("没有可导出的记录。");
    return;
  }
  const headers = ["SKU", "预估重量(g)", "实际重量(g)", "误差(g)", "状态", "记录时间"];
  const rows = records.map((record) => [
    record.sku,
    formatWeight(record.estimate),
    formatWeight(record.actual),
    record.diff === null ? "" : formatWeight(record.diff),
    record.estimate === null ? "未匹配" : record.isException ? "异常" : "正常",
    formatDate(record.createdAt),
  ]);
  downloadText(toCsv([headers, ...rows]), filename, "text/csv;charset=utf-8");
}

function toCsv(rows) {
  return `\uFEFF${rows
    .map((row) =>
      row
        .map((cell) => {
          const value = String(cell ?? "");
          return /[",\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
        })
        .join(",")
    )
    .join("\n")}`;
}

function downloadText(text, filename, type) {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function downloadTemplate() {
  const rows = [
    ["SKU", "预估重量(g)"],
    ["SKU-001", "125"],
    ["SKU-002", "248.5"],
  ];
  downloadText(toCsv(rows), "sku-weight-template.csv", "text/csv;charset=utf-8");
}

function formatDate(value) {
  return new Intl.DateTimeFormat("zh-CN", {
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

async function handleFileImport(file, importer, successMessage) {
  if (!file) return;
  try {
    const count = importer(await readTable(file));
    setLastResult(`${successMessage} ${count} 条`, "good");
  } catch (error) {
    alert(error.message);
  }
}

el.estimateFile.addEventListener("change", (event) => {
  handleFileImport(event.target.files[0], importEstimates, "已导入预估");
  event.target.value = "";
});

el.actualFile.addEventListener("change", (event) => {
  handleFileImport(event.target.files[0], importActuals, "已导入称重");
  event.target.value = "";
});

el.weighForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const record = createRecord(el.skuInput.value, el.actualWeightInput.value);
  if (!record) return;
  state.records.unshift(record);
  saveState();
  render();

  if (record.estimate === null) {
    setLastResult("未匹配预估", "neutral");
  } else if (record.isException) {
    setLastResult("发现异常", "bad");
    showExceptionModal(record);
  } else {
    setLastResult("称重正常", "good");
  }

  el.skuInput.value = "";
  el.actualWeightInput.value = "";
  el.skuInput.focus();
});

el.thresholdInput.addEventListener("change", () => {
  state.threshold = Math.max(0, Number(el.thresholdInput.value) || DEFAULT_THRESHOLD);
  refreshRecordFlags();
  saveState();
  render();
});

el.filterButtons.forEach((button) => {
  button.addEventListener("click", () => {
    state.filter = button.dataset.filter;
    el.filterButtons.forEach((item) => item.classList.toggle("active", item === button));
    renderRecords();
  });
});

el.exportExceptionsBtn.addEventListener("click", () => {
  exportCsv(getExceptions(), `abnormal-sku-${new Date().toISOString().slice(0, 10)}.csv`);
});

el.modalExportBtn.addEventListener("click", () => {
  exportCsv(getExceptions(), `abnormal-sku-${new Date().toISOString().slice(0, 10)}.csv`);
});

el.exportAllBtn.addEventListener("click", () => {
  exportCsv(state.records, `sku-weight-records-${new Date().toISOString().slice(0, 10)}.csv`);
});

el.downloadTemplateBtn.addEventListener("click", downloadTemplate);
el.modalCloseBtn.addEventListener("click", hideModal);
el.modal.addEventListener("click", (event) => {
  if (event.target === el.modal) hideModal();
});

el.clearBtn.addEventListener("click", () => {
  if (!confirm("确认清空所有导入和称重记录？")) return;
  state.estimates.clear();
  state.records = [];
  state.threshold = DEFAULT_THRESHOLD;
  saveState();
  setLastResult("已清空", "neutral");
  render();
});

loadState();
render();
