const data = window.BOM_DATA;

const state = {
  inputs: new Map(),
  activeTab: "machine",
  filter: "",
  productionDate: todayIsoDate(),
  shiftsPerDay: 3,
  generatedAt: "",
  results: {
    machine: [],
    material: [],
    items: [],
    missing: [],
  },
};

const bomByItem = new Map();
for (const row of data.records) {
  if (!bomByItem.has(row.i)) bomByItem.set(row.i, []);
  bomByItem.get(row.i).push(row);
}

const itemNameByCode = new Map(data.items.map((item) => [item.code, item.component]));

const el = {
  sourceMeta: document.querySelector("#sourceMeta"),
  fileInput: document.querySelector("#fileInput"),
  itemInput: document.querySelector("#itemInput"),
  qtyInput: document.querySelector("#qtyInput"),
  itemList: document.querySelector("#itemList"),
  uploadStatus: document.querySelector("#uploadStatus"),
  productionDateInput: document.querySelector("#productionDateInput"),
  productionDateLabel: document.querySelector("#productionDateLabel"),
  shiftsPerDayInput: document.querySelector("#shiftsPerDayInput"),
  addLine: document.querySelector("#addLine"),
  pasteBox: document.querySelector("#pasteBox"),
  loadPaste: document.querySelector("#loadPaste"),
  clearAll: document.querySelector("#clearAll"),
  inputRows: document.querySelector("#inputRows"),
  totalQty: document.querySelector("#totalQty"),
  totalHours: document.querySelector("#totalHours"),
  totalShifts: document.querySelector("#totalShifts"),
  totalDays: document.querySelector("#totalDays"),
  totalMaterials: document.querySelector("#totalMaterials"),
  tabs: document.querySelectorAll(".tab"),
  exportPlan: document.querySelector("#exportPlan"),
  exportMachinePlan: document.querySelector("#exportMachinePlan"),
  exportBom: document.querySelector("#exportBom"),
  exportCsv: document.querySelector("#exportCsv"),
  exportStatus: document.querySelector("#exportStatus"),
  tableTitle: document.querySelector("#tableTitle"),
  filterBox: document.querySelector("#filterBox"),
  resultHead: document.querySelector("#resultHead"),
  resultBody: document.querySelector("#resultBody"),
};

const format = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 2 });
const formatPrecise = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 6 });
const formatQty = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 });
const parserEndpoint =
  window.location.protocol.startsWith("http") && window.location.port === "4175"
    ? "/parse-upload"
    : "http://127.0.0.1:4175/parse-upload";

function todayIsoDate() {
  const now = new Date();
  const offset = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - offset).toISOString().slice(0, 10);
}

function formatDateLabel(value) {
  if (!value) return "-";
  const [year, month, day] = value.split("-").map(Number);
  if (!year || !month || !day) return value;
  return new Intl.DateTimeFormat("en-IN", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(new Date(year, month - 1, day));
}

function markGenerated() {
  state.generatedAt = new Date().toLocaleString("en-IN");
}

function daysFromShifts(shifts) {
  return Math.ceil((Number(shifts) || 0) / Math.max(1, Number(state.shiftsPerDay) || 1));
}

function sourceCycleTime(value) {
  const cycleTime = Number(value) || 0;
  return Number((cycleTime * 60).toFixed(6));
}

function exportCycleTime(row) {
  return sourceCycleTime(row.ct);
}

function addDays(dateValue, days) {
  const [year, month, day] = dateValue.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  date.setDate(date.getDate() + Math.max(0, Number(days) || 0));
  const offset = date.getTimezoneOffset() * 60000;
  return new Date(date.getTime() - offset).toISOString().slice(0, 10);
}

function machineSortValue(machine) {
  const value = String(machine || "").toUpperCase();
  const order = [25, 45, 50, 80, 120, 160];
  const match = value.match(/(?:^|\D)(25|45|50|80|120|160)(?:\D|$)/);
  if (!match) return 1000;
  return order.indexOf(Number(match[1]));
}

function machineOrderLabel(machine) {
  return String(machine || "").toUpperCase().match(/(?:^|\D)(25|45|50|80|120|160)(?:\D|$)/)?.[1] || "";
}

function number(value) {
  const cleaned = String(value ?? "").replace(/,/g, "").trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function parseCsv(text) {
  const rows = [];
  let current = "";
  let row = [];
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (char === '"' && quoted && next === '"') {
      current += '"';
      i += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if ((char === "," || char === "\t") && !quoted) {
      row.push(current.trim());
      current = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(current.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }

  row.push(current.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
}

function setUploadStatus(message, type = "") {
  el.uploadStatus.textContent = message;
  el.uploadStatus.className = `upload-status ${type}`.trim();
}

function setUploadStatusHtml(html, type = "") {
  el.uploadStatus.innerHTML = html;
  el.uploadStatus.className = `upload-status ${type}`.trim();
}

function setExportStatus(message, type = "") {
  if (!el.exportStatus) return;
  el.exportStatus.textContent = message;
  el.exportStatus.className = `export-status ${type}`.trim();
}

function setExportStatusHtml(html, type = "") {
  if (!el.exportStatus) return;
  el.exportStatus.innerHTML = html;
  el.exportStatus.className = `export-status ${type}`.trim();
}

function loadRows(rows) {
  if (!rows.length) {
    setUploadStatus("No usable rows found in the uploaded file.", "error");
    return 0;
  }
  const header = rows[0].map((x) => x.toLowerCase());
  const hasHeader = header.some((x) => x.includes("item")) || header.some((x) => x.includes("qty"));
  const itemIndex = hasHeader ? Math.max(0, header.findIndex((x) => x.includes("item"))) : 0;
  const qtyIndex = hasHeader ? Math.max(1, header.findIndex((x) => x.includes("qty") || x.includes("quantity"))) : 1;
  const body = hasHeader ? rows.slice(1) : rows;
  let loaded = 0;

  for (const row of body) {
    const code = String(row[itemIndex] ?? "").trim();
    const qty = number(row[qtyIndex]);
    if (code && qty > 0) {
      state.inputs.set(code, (state.inputs.get(code) || 0) + qty);
      loaded += 1;
    }
  }
  if (loaded > 0) markGenerated();
  render();
  return loaded;
}

async function parseExcelFile(file) {
  const form = new FormData();
  form.append("file", file);
  const response = await fetch(parserEndpoint, {
    method: "POST",
    body: form,
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || "Unable to read the Excel file.");
  }
  return payload.rows || [];
}

async function handleUpload(file) {
  const name = file.name.toLowerCase();
  setUploadStatus(`Reading ${file.name}...`);
  try {
    let rows;
    if (name.endsWith(".csv") || name.endsWith(".txt") || name.endsWith(".tsv")) {
      rows = parseCsv(await file.text());
    } else if (name.endsWith(".xlsx")) {
      rows = await parseExcelFile(file);
    } else if (name.endsWith(".xls")) {
      throw new Error("Please save the Excel file as .xlsx. Legacy .xls is not supported by this local parser.");
    } else {
      throw new Error("Upload a .csv, .tsv, or .xlsx file.");
    }
    const loaded = loadRows(rows);
    if (loaded > 0) {
      setUploadStatus(`Loaded ${formatQty.format(loaded)} production rows from ${file.name}.`, "success");
    }
  } catch (error) {
    setUploadStatus(error.message || "Unable to read this file.", "error");
  }
}

function calculate() {
  const machine = new Map();
  const material = new Map();
  const items = [];
  const missing = [];
  let totalQty = 0;
  let totalHours = 0;

  for (const [code, qty] of state.inputs) {
    totalQty += qty;
    const rows = bomByItem.get(code);
    if (!rows) {
      missing.push({ code, productionDate: state.productionDate, qty, reason: "Item code not found in BOM Reference" });
      continue;
    }

    const first = rows[0];
    const hourQty = Number(first.hq) || 0;
    const hours = hourQty > 0 ? qty / hourQty : qty * (Number(first.ct) || 0);
    const shifts = Math.ceil(hours / 8);
    const days = daysFromShifts(shifts);
    totalHours += hours;

    const machineKey = `${first.ma}||${first.p}`;
    const machineRow = machine.get(machineKey) || {
      productionDate: state.productionDate,
      machine: first.ma || "Unassigned",
      process: first.p || "Unassigned",
      itemCount: 0,
      qty: 0,
      hours: 0,
      shifts: 0,
      days: 0,
    };
    machineRow.itemCount += 1;
    machineRow.qty += qty;
    machineRow.hours += hours;
    machineRow.shifts += shifts;
    machineRow.days = daysFromShifts(machineRow.shifts);
    machine.set(machineKey, machineRow);

    for (const bom of rows) {
      const materialKey = `${bom.mc}||${bom.m}`;
      const materialRow = material.get(materialKey) || {
        productionDate: state.productionDate,
        materialCode: bom.mc || "No code",
        material: bom.m || "No material name",
        factor: 0,
        qty: 0,
        itemCount: 0,
      };
      materialRow.factor += Number(bom.w) || 0;
      materialRow.qty += qty * (Number(bom.w) || 0);
      materialRow.itemCount += 1;
      material.set(materialKey, materialRow);
    }

    items.push({
      productionDate: state.productionDate,
      code,
      component: first.c || itemNameByCode.get(code) || "",
      qty,
      process: first.p,
      machine: first.ma,
      hourQty,
      hours,
      shifts,
      days,
      materialLines: rows.length,
    });
  }

  state.results.machine = [...machine.values()].sort((a, b) => b.hours - a.hours);
  state.results.material = [...material.values()].sort((a, b) => b.qty - a.qty);
  state.results.items = items.sort((a, b) => b.hours - a.hours);
  state.results.missing = missing;

  el.totalQty.textContent = formatQty.format(totalQty);
  el.productionDateLabel.textContent = formatDateLabel(state.productionDate);
  el.totalHours.textContent = format.format(totalHours);
  el.totalShifts.textContent = formatQty.format(state.results.machine.reduce((sum, row) => sum + row.shifts, 0));
  el.totalDays.textContent = formatQty.format(state.results.machine.reduce((max, row) => Math.max(max, row.days), 0));
  el.totalMaterials.textContent = formatQty.format(state.results.material.length);
}

function renderInputs() {
  if (!state.inputs.size) {
    el.inputRows.innerHTML = `<tr><td colspan="3" class="muted">No production quantities loaded</td></tr>`;
    return;
  }
  el.inputRows.innerHTML = [...state.inputs]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([code, qty]) => {
      return `<tr>
        <td>${escapeHtml(code)}<div class="muted">${escapeHtml(itemNameByCode.get(code) || "")}</div></td>
        <td class="num">${formatQty.format(qty)}</td>
        <td><button class="remove-line" data-code="${escapeHtml(code)}" type="button" title="Remove">×</button></td>
      </tr>`;
    })
    .join("");
}

function currentRows() {
  const rows = state.results[state.activeTab] || [];
  if (!state.filter) return rows;
  const needle = state.filter.toLowerCase();
  return rows.filter((row) => Object.values(row).join(" ").toLowerCase().includes(needle));
}

function renderTable() {
  const views = {
    machine: {
      title: "Machine Consumption",
      columns: [
        ["productionDate", "Production Date"],
        ["machine", "Machine"],
        ["process", "Process"],
        ["itemCount", "Items", "num"],
        ["qty", "Production Qty", "num0"],
        ["hours", "Required Hrs", "num"],
        ["shifts", "Shifts", "num0"],
        ["days", "Days", "num0"],
      ],
    },
    material: {
      title: "Material Consumption",
      columns: [
        ["productionDate", "Production Date"],
        ["materialCode", "Material Code"],
        ["material", "Material"],
        ["itemCount", "BOM Lines", "num0"],
        ["qty", "Consumption", "nump"],
      ],
    },
    items: {
      title: "Item Calculation",
      columns: [
        ["code", "Item Code"],
        ["productionDate", "Production Date"],
        ["component", "Component"],
        ["qty", "Production Qty", "num0"],
        ["machine", "Machine"],
        ["process", "Process"],
        ["hourQty", "Hour Qty", "num"],
        ["hours", "Required Hrs", "num"],
        ["shifts", "Shifts", "num0"],
        ["days", "Days", "num0"],
        ["materialLines", "Materials", "num0"],
      ],
    },
    missing: {
      title: "Missing Items",
      columns: [
        ["code", "Item Code"],
        ["productionDate", "Production Date"],
        ["qty", "Production Qty", "num0"],
        ["reason", "Reason"],
      ],
    },
  };
  const view = views[state.activeTab];
  const rows = currentRows();
  el.tableTitle.textContent = view.title;
  el.resultHead.innerHTML = `<tr>${view.columns
    .map(([, label, type]) => `<th class="${type ? "num" : ""}">${label}</th>`)
    .join("")}</tr>`;
  el.resultBody.innerHTML = rows.length
    ? rows
        .map((row) => {
          const klass = state.activeTab === "missing" ? " class=\"warning-row\"" : "";
          return `<tr${klass}>${view.columns
            .map(([key, , type]) => {
              const raw = row[key];
              const value =
                type === "num0"
                  ? formatQty.format(raw || 0)
                  : type === "nump"
                    ? formatPrecise.format(raw || 0)
                    : type === "num"
                      ? format.format(raw || 0)
                      : escapeHtml(raw);
              return `<td class="${type ? "num" : ""}">${value}</td>`;
            })
            .join("")}</tr>`;
        })
        .join("")
    : `<tr><td colspan="${view.columns.length}" class="muted">No rows to show</td></tr>`;
}

function render() {
  calculate();
  renderInputs();
  renderTable();
}

async function exportCurrentCsv() {
  const rows = currentRows();
  if (!rows.length) return;
  await downloadCsv(`${state.activeTab}_consumption.csv`, rows);
}

function rowsToCsv(rows) {
  const keys = Object.keys(rows[0]);
  return [
    keys.join(","),
    ...rows.map((row) =>
      keys
        .map((key) => {
          const value = row[key] ?? "";
          return `"${String(value).replaceAll('"', '""')}"`;
        })
        .join(","),
    ),
  ].join("\n");
}

async function saveCsvOnServer(filename, content) {
  const response = await fetch("http://127.0.0.1:4175/save-export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filename, content }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || "Unable to save export.");
  return payload;
}

async function savePlanXlsxOnServer(filename, rows) {
  const response = await fetch("http://127.0.0.1:4175/save-plan-xlsx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filename, rows }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || "Unable to save Excel plan.");
  return payload;
}

async function exportPlanXlsxOnServer() {
  const inputs = [...state.inputs].map(([code, qty]) => ({ code, qty }));
  const response = await fetch("http://127.0.0.1:4175/export-plan-xlsx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      inputs,
      productionDate: state.productionDate,
      shiftsPerDay: state.shiftsPerDay,
      generatedAt: state.generatedAt,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || "Unable to export Excel plan.");
  return payload;
}

async function exportMachinePlanXlsxOnServer() {
  const inputs = [...state.inputs].map(([code, qty]) => ({ code, qty }));
  const response = await fetch("http://127.0.0.1:4175/export-machine-plan-xlsx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      inputs,
      productionDate: state.productionDate,
      shiftsPerDay: state.shiftsPerDay,
      generatedAt: state.generatedAt,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || "Unable to export machine plan.");
  return payload;
}

function exportReadyHtml(label, saved) {
  const url = new URL(saved.url, "http://127.0.0.1:4175/").href;
  const pathText = saved.path ? `<span class="saved-path">Saved to Downloads: ${escapeHtml(saved.path)}</span>` : "";
  return `${label}: <a href="${escapeHtml(url)}" download="${escapeHtml(saved.filename)}" target="_blank" rel="noopener">Open ${escapeHtml(saved.filename)}</a>${pathText}`;
}

async function downloadCsv(filename, rows) {
  if (!rows.length) {
    setUploadStatus("No production rows available to export.", "error");
    setExportStatus("No production rows available to export.", "error");
    return;
  }
  setExportStatus(`Creating ${filename}...`);
  const csv = rowsToCsv(rows);
  try {
    const saved = await saveCsvOnServer(filename, csv);
    const html = exportReadyHtml("Export ready", saved);
    setUploadStatusHtml(html, "success");
    setExportStatusHtml(html, "success");
    return;
  } catch {
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setExportStatus(`Export created: ${filename}`, "success");
  }
}

async function downloadPlanXlsx(filename, rows) {
  if (!rows.length) {
    setUploadStatus("No production rows available to export.", "error");
    setExportStatus("No production rows available to export.", "error");
    return;
  }
  setExportStatus(`Creating ${filename}...`);
  try {
    const saved = await savePlanXlsxOnServer(filename, rows);
    const html = exportReadyHtml("Excel plan ready", saved);
    setUploadStatusHtml(html, "success");
    setExportStatusHtml(html, "success");
  } catch (error) {
    setUploadStatus(error.message || "Unable to create Excel plan.", "error");
    setExportStatus(error.message || "Unable to create Excel plan.", "error");
  }
}

function buildPlanExportRows() {
  const rows = [];
  for (const [code, productionQty] of state.inputs) {
    const bomRows = bomByItem.get(code);
    if (!bomRows) {
      rows.push({
        "Item Code": code,
        "Production Date": state.productionDate,
        "Generated At": state.generatedAt,
        "Shifts per Day": state.shiftsPerDay,
        Component: itemNameByCode.get(code) || "",
        "Production Qty": productionQty,
        Machine: "",
        Process: "",
        "Required Hrs": "",
        Shifts: "",
        "Days to Complete": "",
        "Material Code": "",
        Material: "",
        "Material Factor": "",
        "Material Consumption": "",
        Status: "Item code not found in BOM Reference",
      });
      continue;
    }

    const first = bomRows[0];
    const hourQty = Number(first.hq) || 0;
    const requiredHours = hourQty > 0 ? productionQty / hourQty : productionQty * (Number(first.ct) || 0);
    const shifts = Math.ceil(requiredHours / 8);
    const days = daysFromShifts(shifts);

    for (const bom of bomRows) {
      const factor = Number(bom.w) || 0;
      rows.push({
        "Item Code": code,
        "Production Date": state.productionDate,
        "Generated At": state.generatedAt,
        "Shifts per Day": state.shiftsPerDay,
        Component: first.c || itemNameByCode.get(code) || "",
        "Production Qty": productionQty,
        Machine: first.ma || "",
        Process: first.p || "",
        "Required Hrs": Number(requiredHours.toFixed(4)),
        Shifts: shifts,
        "Days to Complete": days,
        "Material Code": bom.mc || "",
        Material: bom.m || "",
        "Material Factor": factor,
        "Material Consumption": Number((productionQty * factor).toFixed(6)),
        Status: "OK",
      });
    }
  }
  return rows;
}

async function exportPlanCsv() {
  if (!state.inputs.size) {
    setUploadStatus("No production rows available to export.", "error");
    setExportStatus("No production rows available to export.", "error");
    return;
  }
  setUploadStatus("Creating Excel plan on server...");
  setExportStatus("Creating Excel plan on server...");
  try {
    const saved = await exportPlanXlsxOnServer();
    const html = exportReadyHtml("Excel plan ready", saved);
    setUploadStatusHtml(html, "success");
    setExportStatusHtml(html, "success");
  } catch (error) {
    setUploadStatus(error.message || "Unable to export Excel plan.", "error");
    setExportStatus(error.message || "Unable to export Excel plan.", "error");
  }
}

function buildExcelPlanRows() {
  const rows = [];
  const machineFinishDay = new Map();
  const itemRows = state.results.items
    .slice()
    .sort((a, b) => {
      const machineOrder = machineSortValue(a.machine) - machineSortValue(b.machine);
      if (machineOrder !== 0) return machineOrder;
      return String(a.machine || "").localeCompare(String(b.machine || "")) || b.hours - a.hours;
    });

  for (const item of itemRows) {
    const machine = item.machine || "Unassigned";
    const offset = machineFinishDay.get(machine) || 0;
    const startDate = addDays(state.productionDate, offset);
    machineFinishDay.set(machine, offset + item.days);
    const bomRows = bomByItem.get(item.code) || [];

    for (const bom of bomRows) {
      const factor = Number(bom.w) || 0;
      rows.push({
        Date: startDate,
        Machine: machine,
        "Item Code": item.code,
        Component: item.component,
        "Plan Qty": item.qty,
        "BOM Cycle Time": exportCycleTime(bom),
        Material: bom.m || "",
        "Material Required": Number((item.qty * factor).toFixed(6)),
        "Production Hrs": Number(item.hours.toFixed(4)),
      });
    }
  }

  for (const missing of state.results.missing) {
    rows.push({
      Date: state.productionDate,
      Machine: "",
      "Item Code": missing.code,
      Component: itemNameByCode.get(missing.code) || missing.reason,
      "Plan Qty": missing.qty,
      "BOM Cycle Time": "",
      Material: "",
      "Material Required": "",
      "Production Hrs": "",
    });
  }

  return rows;
}

function buildMachineWisePlanRows() {
  const rows = [];
  const machineFinishDay = new Map();
  const itemRows = state.results.items
    .slice()
    .sort((a, b) => {
      const machineOrder = machineSortValue(a.machine) - machineSortValue(b.machine);
      if (machineOrder !== 0) return machineOrder;
      return String(a.machine || "").localeCompare(String(b.machine || "")) || b.hours - a.hours;
    });

  for (const item of itemRows) {
    const machine = item.machine || "Unassigned";
    const offset = machineFinishDay.get(machine) || 0;
    const startDate = addDays(state.productionDate, offset);
    const endDate = addDays(startDate, Math.max(0, item.days - 1));
    machineFinishDay.set(machine, offset + item.days);
    rows.push({
      "Machine Order": machineSortValue(machine) >= 1000 ? "Other" : machineOrderLabel(machine),
      Machine: machine,
      Process: item.process || "",
      "Item Code": item.code,
      Component: item.component,
      "Production Qty": item.qty,
      "Start Date": startDate,
      "End Date": endDate,
      "Required Hrs": Number(item.hours.toFixed(4)),
      Shifts: item.shifts,
      "Shifts per Day": state.shiftsPerDay,
      "Days to Complete": item.days,
      "Generated At": state.generatedAt,
      Status: "OK",
    });
  }

  for (const missing of state.results.missing) {
    rows.push({
      "Machine Order": "",
      Machine: "",
      Process: "",
      "Item Code": missing.code,
      Component: itemNameByCode.get(missing.code) || "",
      "Production Qty": missing.qty,
      "Start Date": state.productionDate,
      "End Date": "",
      "Required Hrs": "",
      Shifts: "",
      "Shifts per Day": state.shiftsPerDay,
      "Days to Complete": "",
      "Generated At": state.generatedAt,
      Status: missing.reason,
    });
  }

  return rows;
}

async function exportMachineWisePlanCsv() {
  if (!state.inputs.size) {
    setUploadStatus("No production rows available to export.", "error");
    setExportStatus("No production rows available to export.", "error");
    return;
  }
  setUploadStatus("Creating machine plan on server...");
  setExportStatus("Creating machine plan on server...");
  try {
    const saved = await exportMachinePlanXlsxOnServer();
    const html = exportReadyHtml("Machine plan ready", saved);
    setUploadStatusHtml(html, "success");
    setExportStatusHtml(html, "success");
  } catch (error) {
    setUploadStatus(error.message || "Unable to export machine plan.", "error");
    setExportStatus(error.message || "Unable to export machine plan.", "error");
  }
}

function buildBomExportRows() {
  const selectedCodes = state.inputs.size ? [...state.inputs.keys()] : [...bomByItem.keys()];
  const rows = [];
  for (const code of selectedCodes) {
    const productionQty = state.inputs.get(code) || "";
    const bomRows = bomByItem.get(code);
    if (!bomRows) {
      rows.push({
        "Item Code": code,
        "Production Date": state.productionDate,
        "Generated At": state.generatedAt,
        "Shifts per Day": state.shiftsPerDay,
        Component: itemNameByCode.get(code) || "",
        "Production Qty": productionQty,
        "Material Code": "",
        Material: "",
        "Part Wt / Factor": "",
        Process: "",
        Machine: "",
        "BOM Cycle Time": "",
        "Hour Qty": "",
        Shifts: "",
        "Days to Complete": "",
        "Material Consumption": "",
        Status: "Item code not found in BOM Reference",
      });
      continue;
    }

    for (const bom of bomRows) {
      const factor = Number(bom.w) || 0;
      rows.push({
        "Item Code": bom.i,
        "Production Date": state.productionDate,
        "Generated At": state.generatedAt,
        "Shifts per Day": state.shiftsPerDay,
        Component: bom.c,
        "Production Qty": productionQty,
        "Material Code": bom.mc,
        Material: bom.m,
        "Part Wt / Factor": factor,
        Process: bom.p,
        Machine: bom.ma,
        "BOM Cycle Time": exportCycleTime(bom),
        "Hour Qty": bom.hq,
        Shifts: productionQty === "" ? "" : Math.ceil(((Number(bom.hq) || 0) > 0 ? productionQty / Number(bom.hq) : productionQty * (Number(bom.ct) || 0)) / 8),
        "Days to Complete":
          productionQty === ""
            ? ""
            : daysFromShifts(Math.ceil(((Number(bom.hq) || 0) > 0 ? productionQty / Number(bom.hq) : productionQty * (Number(bom.ct) || 0)) / 8)),
        "Material Consumption": productionQty === "" ? "" : Number((productionQty * factor).toFixed(6)),
        Status: "OK",
      });
    }
  }
  return rows;
}

async function exportBomCsv() {
  const rows = buildBomExportRows();
  const filename = state.inputs.size ? "bom_details_for_selected_items.csv" : "full_bom_reference_details.csv";
  await downloadCsv(filename, rows);
}

function hydrateItemList() {
  el.itemList.innerHTML = data.items
    .map((item) => `<option value="${escapeHtml(item.code)}">${escapeHtml(item.component)}</option>`)
    .join("");
}

el.sourceMeta.textContent = `${formatQty.format(data.itemCount)} items · ${formatQty.format(data.recordCount)} BOM rows`;
el.productionDateInput.value = state.productionDate;
el.shiftsPerDayInput.value = state.shiftsPerDay;
hydrateItemList();

el.fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  await handleUpload(file);
  event.target.value = "";
});

document.querySelector(".dropzone").addEventListener("dragover", (event) => {
  event.preventDefault();
});

document.querySelector(".dropzone").addEventListener("drop", async (event) => {
  event.preventDefault();
  const file = event.dataTransfer.files[0];
  if (!file) return;
  await handleUpload(file);
});

el.addLine.addEventListener("click", () => {
  const code = el.itemInput.value.trim();
  const qty = number(el.qtyInput.value);
  if (!code || qty <= 0) return;
  state.inputs.set(code, (state.inputs.get(code) || 0) + qty);
  markGenerated();
  el.itemInput.value = "";
  el.qtyInput.value = "";
  render();
});

el.loadPaste.addEventListener("click", () => {
  const loaded = loadRows(parseCsv(el.pasteBox.value));
  if (loaded > 0) setUploadStatus(`Loaded ${formatQty.format(loaded)} pasted production rows.`, "success");
});

el.clearAll.addEventListener("click", () => {
  state.inputs.clear();
  el.pasteBox.value = "";
  state.generatedAt = "";
  setUploadStatus("");
  setExportStatus("");
  render();
});

el.productionDateInput.addEventListener("change", () => {
  state.productionDate = el.productionDateInput.value || todayIsoDate();
  render();
});

el.shiftsPerDayInput.addEventListener("input", () => {
  const value = Math.max(1, Math.min(3, Math.round(number(el.shiftsPerDayInput.value) || 1)));
  state.shiftsPerDay = value;
  el.shiftsPerDayInput.value = value;
  render();
});

el.inputRows.addEventListener("click", (event) => {
  const button = event.target.closest(".remove-line");
  if (!button) return;
  state.inputs.delete(button.dataset.code);
  render();
});

for (const tab of el.tabs) {
  tab.addEventListener("click", () => {
    state.activeTab = tab.dataset.tab;
    for (const item of el.tabs) item.classList.toggle("active", item === tab);
    renderTable();
  });
}

el.filterBox.addEventListener("input", () => {
  state.filter = el.filterBox.value.trim();
  renderTable();
});

el.exportCsv.addEventListener("click", exportCurrentCsv);
el.exportPlan.addEventListener("click", exportPlanCsv);
el.exportMachinePlan.addEventListener("click", exportMachineWisePlanCsv);
el.exportBom.addEventListener("click", exportBomCsv);

render();
