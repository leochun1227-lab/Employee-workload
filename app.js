const state = {
  data: null,
  temp: { problemFlags: {}, manualTickets: {} },
  firebaseWarning: "",
  selectedEmployee: "",
  tab: "removed",
  selectedDate: yesterdayKey(),
  employeeFilter: "",
  tableFilter: "",
};

const FIREBASE_BASE = "https://snowy-hr-report-default-rtdb.asia-southeast1.firebasedatabase.app";
const DATA_URL = `${FIREBASE_BASE}/ctmTicketStatusMonitorV44/tempEmployeeTicketDashboard/data.json`;
const TEMP_URL = `${FIREBASE_BASE}/ctmTicketStatusMonitorV44/tempEmployeeTicketTable.json`;

const columns = {
  tickets: ["Ticket ID", "Status", "Customer", "Claim Type", "Dealer", "Aging Days", "Amount Including Tax", "Created On", "Last Update"],
  removed: ["Ticket ID", "Removed Date", "From Status", "To Status", "Customer", "Claim Type", "Dealer", "Aging Days", "Amount"],
  approved: ["Ticket ID", "Decision Date", "Status", "Customer", "Claim Type", "Dealer", "Aging Days", "Amount Including Tax"],
  unapproved: ["Ticket ID", "Decision Date", "Status", "Customer", "Claim Type", "Dealer", "Aging Days", "Amount Including Tax"],
  manual: ["Move", "Ticket ID", "Status", "Employee", "Note", "Created On"],
};

const allowedEmployees = new Set([
  "marissa colosimo",
  "ford hapuku",
  "mark bertoncini",
  "leanne pulford",
  "kylie clayton",
  "rosemary johnstone",
  "michael scordia",
  "robert stella",
  "chloe bolger",
]);

const el = {
  fileInput: document.querySelector("#fileInput"),
  addTicketButton: document.querySelector("#addTicketButton"),
  exportChangesButton: document.querySelector("#exportChangesButton"),
  refreshButton: document.querySelector("#refreshButton"),
  dateFilter: document.querySelector("#dateFilter"),
  employeeSearch: document.querySelector("#employeeSearch"),
  tableSearch: document.querySelector("#tableSearch"),
  employeeList: document.querySelector("#employeeList"),
  employeeCount: document.querySelector("#employeeCount"),
  employeeName: document.querySelector("#employeeName"),
  removedCount: document.querySelector("#removedCount"),
  approvedCount: document.querySelector("#approvedCount"),
  unapprovedCount: document.querySelector("#unapprovedCount"),
  datasetStatus: document.querySelector("#datasetStatus"),
  sourceStatus: document.querySelector("#sourceStatus"),
  tableHead: document.querySelector("#tableHead"),
  tableBody: document.querySelector("#tableBody"),
  emptyState: document.querySelector("#emptyState"),
  markHint: document.querySelector("#markHint"),
  tabs: document.querySelector("#tabs"),
  modalBackdrop: document.querySelector("#modalBackdrop"),
  modalClose: document.querySelector("#modalClose"),
  manualTicketForm: document.querySelector("#manualTicketForm"),
  manualEmployee: document.querySelector("#manualEmployee"),
};

function clean(value) {
  return String(value ?? "").trim();
}

function yesterdayKey() {
  const date = new Date();
  date.setDate(date.getDate() - 1);
  return localDateKey(date);
}

function localDateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function localDateKeyFromValue(value) {
  const text = clean(value);
  if (!text) return "";
  const date = new Date(text);
  if (!Number.isNaN(date.getTime())) return localDateKey(date);
  return normalizeDateKey(text);
}

function employeeKey(value) {
  return clean(value).toLowerCase();
}

function isAssignedEmployee(value) {
  const name = employeeKey(value);
  return name && name !== "-" && name !== "unassigned" && name !== "not assigned";
}

function formatMoney(value) {
  const number = Number(clean(value).replace(/,/g, ""));
  if (!Number.isFinite(number)) return clean(value);
  return number.toLocaleString("en-AU", { style: "currency", currency: "AUD" });
}

function recordKey(row) {
  return [row.type || state.tab, clean(row.ticketId), clean(row.employee)].join("|").toLowerCase();
}

function isProblem(row) {
  return Boolean(state.temp?.problemFlags?.[recordKey(row)]);
}

function manualRows() {
  return Object.entries(state.temp?.manualTickets || {}).map(([key, row]) => ({
    type: "tickets",
    source: "manual",
    manualKey: key,
    ticketId: clean(row.ticketId),
    employee: clean(row.employee),
    status: clean(row.status) || "Manual added",
    customer: clean(row.note),
    note: clean(row.note),
    claimType: "Manual",
    dealer: "",
    agingDays: "",
    amount: "",
    createdOn: manualReviewDate(row),
    lastUpdate: localDateKeyFromValue(row.updatedAt),
    createdAt: clean(row.createdAt),
    updatedAt: clean(row.updatedAt),
    reviewDate: manualReviewDate(row),
  })).filter((row) => isAssignedEmployee(row.employee));
}

function manualReviewDate(row) {
  return localDateKeyFromValue(row?.reviewDate || row?.auditDate || row?.createdAt);
}

function parseWorkbookXml(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const error = doc.querySelector("parsererror");
  if (error) throw new Error("Excel XML could not be parsed.");
  return [...doc.getElementsByTagName("Worksheet")].map((sheet) => {
    const name = sheet.getAttribute("ss:Name") || sheet.getAttribute("Name") || "Sheet";
    const rows = [...sheet.getElementsByTagName("Row")].map((row) => {
      const values = [];
      let cursor = 0;
      [...row.getElementsByTagName("Cell")].forEach((cell) => {
        const index = Number(cell.getAttribute("ss:Index") || cell.getAttribute("Index"));
        if (index) cursor = index - 1;
        const data = cell.getElementsByTagName("Data")[0];
        values[cursor] = clean(data?.textContent);
        cursor += 1;
      });
      return values;
    });
    return { name, rows };
  });
}

function rowsToObjects(sheet) {
  if (!sheet || sheet.rows.length < 2) return [];
  const headers = sheet.rows[0].map(clean);
  return sheet.rows.slice(1).map((row) => {
    const item = {};
    headers.forEach((header, index) => {
      if (!header) return;
      if (item[header] === undefined) item[header] = clean(row[index]);
    });
    return item;
  });
}

function sheetByName(sheets, name) {
  return sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase());
}

async function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

function normalizeTicketRecord(row, type) {
  const employee = clean(row.Employee || row["Employee Group"]);
  return {
    type,
    employee,
    ticketId: clean(row["Ticket ID"] || row["Raw TicketID"]),
    status: clean(row.Status || row["Status Text"] || row["To Status"]),
    customer: clean(row.Customer || row["Raw TicketName"]),
    claimType: clean(row["Claim Type"] || row["Raw TicketTypeText"]),
    repair: clean(row.Repair),
    dealer: clean(row.Dealer || row["Raw DealerName"]),
    agingDays: clean(row["Aging Days"]),
    amount: clean(row["Amount Including Tax"] || row.Amount || row["Raw AmountIncludingTax"]),
    createdOn: clean(row["Created On"] || row["Raw CreatedOn"]),
    lastUpdate: clean(row["Last Update"] || row["Raw LastUpdateDateTime"]),
    decisionDate: clean(row["Decision Date"]),
    removedDate: clean(row["Removed Date"]),
    fromStatus: clean(row["From Status"]),
    toStatus: clean(row["To Status"]),
  };
}

function mergeImportedData(imports) {
  const records = { tickets: [], removed: [], approved: [], unapproved: [] };
  const summaryMap = new Map();
  const sourceNames = [];

  imports.forEach(({ fileName, sheets }) => {
    sourceNames.push(fileName);
    rowsToObjects(sheetByName(sheets, "By Employee")).forEach((row) => {
      const employee = clean(row.Employee);
      if (!isAssignedEmployee(employee)) return;
      const key = employeeKey(employee);
      if (!summaryMap.has(key)) {
        summaryMap.set(key, { name: employee, tickets: 0, removed: 0, approved: 0, unapproved: 0 });
      }
      const summary = summaryMap.get(key);
      summary.tickets = Math.max(summary.tickets, Number(clean(row["Current Critical Tickets"]).replace(/,/g, "")) || 0);
      summary.removed = Math.max(summary.removed, Number(clean(row.Removed).replace(/,/g, "")) || 0);
      summary.approved = Math.max(summary.approved, Number(clean(row.Approved).replace(/,/g, "")) || 0);
      summary.unapproved = Math.max(summary.unapproved, Number(clean(row.Unapproved).replace(/,/g, "")) || 0);
    });

    const current = rowsToObjects(sheetByName(sheets, "Current Critical Detail"));
    const removed = rowsToObjects(sheetByName(sheets, "Removed Ticket Detail"));
    const approved = rowsToObjects(sheetByName(sheets, "Approved Ticket Detail"));
    const unapproved = rowsToObjects(sheetByName(sheets, "Unapproved Ticket Detail"));

    records.tickets.push(...current.map((row) => normalizeTicketRecord(row, "tickets")).filter((row) => isAssignedEmployee(row.employee)));
    records.removed.push(...removed.map((row) => normalizeTicketRecord(row, "removed")).filter((row) => isAssignedEmployee(row.employee)));
    records.approved.push(...approved.map((row) => normalizeTicketRecord(row, "approved")).filter((row) => isAssignedEmployee(row.employee)));
    records.unapproved.push(...unapproved.map((row) => normalizeTicketRecord(row, "unapproved")).filter((row) => isAssignedEmployee(row.employee)));
  });

  const employeeMap = new Map();
  const detailCountMap = new Map();
  summaryMap.forEach((value, key) => employeeMap.set(key, { ...value }));
  Object.entries(records).forEach(([bucket, rows]) => {
    rows.forEach((row) => {
      const key = employeeKey(row.employee);
      if (!detailCountMap.has(key)) {
        detailCountMap.set(key, { name: row.employee, tickets: 0, removed: 0, approved: 0, unapproved: 0 });
      }
      detailCountMap.get(key)[bucket] += 1;
    });
  });

  detailCountMap.forEach((detail, key) => {
    if (!employeeMap.has(key)) {
      employeeMap.set(key, detail);
      return;
    }
    const summary = employeeMap.get(key);
    ["tickets", "removed", "approved", "unapproved"].forEach((bucket) => {
      summary[bucket] = Math.max(Number(summary[bucket] || 0), Number(detail[bucket] || 0));
    });
  });

  return {
    meta: {
      sourceNames,
      importedBy: "Web import",
    },
    employees: [...employeeMap.values()].sort((a, b) => a.name.localeCompare(b.name)),
    records,
  };
}

async function loadData() {
  const [dataResponse, tempResponse] = await Promise.all([
    fetch(DATA_URL, { cache: "no-store" }),
    fetch(TEMP_URL, { cache: "no-store" }),
  ]);
  state.data = normalizeDashboardData(await dataResponse.json());
  const rawTemp = await tempResponse.json();
  state.temp = normalizeTempData(rawTemp);
  if (needsTempMigration(rawTemp)) {
    await persistTempData(state.temp);
  }
  state.firebaseWarning = "";
  const employees = visibleEmployees();
  if (!employees.some((employee) => employee.name === state.selectedEmployee)) {
    state.selectedEmployee = employees[0]?.name || "";
  }
  render();
}

function normalizeDashboardData(data) {
  return {
    meta: {
      updatedAt: null,
      sourceNames: [],
      importedBy: "Web import",
      ...(data?.meta || {}),
    },
    employees: Array.isArray(data?.employees) ? data.employees : [],
    records: {
      tickets: Array.isArray(data?.records?.tickets) ? data.records.tickets : [],
      removed: Array.isArray(data?.records?.removed) ? data.records.removed : [],
      approved: Array.isArray(data?.records?.approved) ? data.records.approved : [],
      unapproved: Array.isArray(data?.records?.unapproved) ? data.records.unapproved : [],
    },
  };
}

function normalizeTempData(data) {
  const manualTickets = {};
  Object.entries(data?.manualTickets || {}).forEach(([key, row]) => {
    manualTickets[key] = {
      ...row,
      reviewDate: row?.reviewDate || row?.auditDate || localDateKeyFromValue(row?.createdAt),
    };
  });

  const problemFlags = {};
  Object.entries(data?.problemFlags || {}).forEach(([key, row]) => {
    problemFlags[key] = {
      ...row,
      reviewDate: row?.reviewDate || localDateKeyFromValue(row?.markedAt),
    };
  });

  return {
    meta: { updatedAt: null, source: "employee-ticket-dashboard", ...(data?.meta || {}) },
    problemFlags,
    manualTickets,
  };
}

function needsTempMigration(rawTemp) {
  const rawManual = rawTemp?.manualTickets || {};
  const rawProblem = rawTemp?.problemFlags || {};
  return Object.values(rawManual).some((row) => !clean(row?.reviewDate || row?.auditDate)) ||
    Object.values(rawProblem).some((row) => !clean(row?.reviewDate));
}

async function importFiles(files) {
  const excelFiles = [...files];
  if (!excelFiles.length) return;
  el.datasetStatus.textContent = "Importing Excel files...";
  const imports = [];
  for (const file of excelFiles) {
    const text = await readFileAsText(file);
    imports.push({ fileName: file.name, sheets: parseWorkbookXml(text) });
  }
  const data = mergeImportedData(imports);
  data.meta = {
    ...(data.meta || {}),
    updatedAt: new Date().toISOString(),
    source: "public-site",
  };
  const response = await fetch(DATA_URL, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  if (!response.ok) throw new Error("Import failed.");
  await loadData();
}

function selectedRows(bucket) {
  const baseRows = state.data?.records?.[bucket] || [];
  const rows = bucket === "manual" ? manualRows() : baseRows;
  return rows.filter((row) => {
    const employeeOk = state.selectedEmployee && row.employee === state.selectedEmployee;
    return employeeOk && rowDateKey(row, bucket) === state.selectedDate;
  });
}

function rowDateKey(row, bucket) {
  const raw = bucket === "removed"
    ? row.removedDate
    : bucket === "approved" || bucket === "unapproved"
      ? row.decisionDate
      : row.createdOn;
  return normalizeDateKey(raw);
}

function normalizeDateKey(value) {
  const text = clean(value);
  if (!text) return "";
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const slash = text.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$/);
  if (slash) return `${slash[3]}-${slash[2].padStart(2, "0")}-${slash[1].padStart(2, "0")}`;
  return text.slice(0, 10);
}

function filteredRows() {
  const term = state.tableFilter.toLowerCase();
  const rows = selectedRows(state.tab);
  const filtered = term ? rows.filter((row) => Object.values(row).join(" ").toLowerCase().includes(term)) : rows;
  return filtered.sort((a, b) => Number(isProblem(b)) - Number(isProblem(a)));
}

function renderEmployees() {
  const employees = visibleEmployees();
  const term = state.employeeFilter.toLowerCase();
  const visible = employees.filter((employee) => employee.name.toLowerCase().includes(term));
  el.employeeCount.textContent = employees.length;
  el.employeeList.innerHTML = visible.map((employee) => {
    const active = employee.name === state.selectedEmployee ? " active" : "";
    return `<button class="employeeItem${active}" data-employee="${escapeHtml(employee.name)}"><span>${escapeHtml(employee.name)}</span></button>`;
  }).join("");
  el.manualEmployee.innerHTML = employees.map((employee) => `<option value="${escapeHtml(employee.name)}">${escapeHtml(employee.name)}</option>`).join("");
}

function visibleEmployees() {
  return (state.data?.employees || []).filter((employee) => allowedEmployees.has(employeeKey(employee.name)));
}

function renderSummary() {
  const employee = state.data?.employees?.find((item) => item.name === state.selectedEmployee);
  const summary = {
    name: employee?.name || "Select employee",
    removed: selectedRows("removed").length,
    approved: selectedRows("approved").length,
    unapproved: selectedRows("unapproved").length,
  };
  el.employeeName.textContent = summary.name;
  el.removedCount.textContent = summary.removed;
  el.approvedCount.textContent = summary.approved;
  el.unapprovedCount.textContent = summary.unapproved;
}

function renderTable() {
  const rows = filteredRows();
  const visibleColumns = activeColumns(rows);
  el.tableHead.innerHTML = `<tr>${visibleColumns.map((name) => `<th>${escapeHtml(name)}</th>`).join("")}</tr>`;
  el.tableBody.innerHTML = rows.map((row) => {
    const classes = [isProblem(row) ? "problemRow" : "", row.source === "manual" ? "manualRow" : ""].join(" ").trim();
    return `<tr class="${classes}">${visibleColumns.map((name) => `<td class="${name.includes("Amount") || name === "Amount" ? "amount" : ""}">${formatCell(row, name)}</td>`).join("")}</tr>`;
  }).join("");
  if (el.markHint) el.markHint.style.display = rows.length ? "block" : "none";
  el.emptyState.style.display = rows.length ? "none" : "block";
  if (!rows.length) {
    el.emptyState.textContent = state.data?.employees?.length ? "You're all good" : "Import Excel files to begin.";
    el.emptyState.classList.toggle("cheerful", Boolean(state.data?.employees?.length));
  }
}

function activeColumns(rows) {
  const base = ["Mark", ...columns[state.tab]];
  if (!rows.length) return base;
  return base.filter((name) => name === "Mark" || name === "Move" || rows.some((row) => clean(cellValue(row, name)) && clean(cellValue(row, name)) !== "-"));
}

function formatCell(row, name) {
  if (name === "Move") {
    return row.source === "manual"
      ? `<button class="moveButton" data-move-manual="${escapeHtml(row.manualKey)}" title="Move this manual ticket out">Move</button>`
      : "";
  }
  if (name === "Mark") {
    const active = isProblem(row) ? " active" : "";
    const label = isProblem(row) ? "Undo" : "!";
    const title = isProblem(row) ? "Restore this ticket" : "Mark ticket problem";
    return `<button class="markButton${active}" data-mark="${escapeHtml(recordKey(row))}" title="${title}">${label}</button>`;
  }
  const raw = cellValue(row, name);
  const value = name.includes("Amount") || name === "Amount" ? formatMoney(raw) : clean(raw) || "-";
  if (name === "Status" || name === "From Status" || name === "To Status") return `<span class="pill">${escapeHtml(value)}</span>`;
  return escapeHtml(value);
}

function cellValue(row, name) {
  const map = {
    "Ticket ID": row.ticketId,
    Status: row.status,
    Customer: row.customer,
    "Claim Type": row.claimType,
    Dealer: row.dealer,
    "Aging Days": row.agingDays,
    "Amount Including Tax": row.amount,
    Amount: row.amount,
    "Created On": row.createdOn,
    "Last Update": row.lastUpdate,
    "Decision Date": row.decisionDate,
    "Removed Date": row.removedDate,
    "From Status": row.fromStatus,
    "To Status": row.toStatus,
    Employee: row.employee,
    Note: row.note || row.customer,
  };
  return map[name];
}

function renderStatus() {
  const meta = state.data?.meta || {};
  el.datasetStatus.textContent = meta.updatedAt
    ? `Showing ${state.selectedDate}. Last updated ${new Date(meta.updatedAt).toLocaleString()}`
    : `Showing ${state.selectedDate}. No data imported yet.`;
  el.sourceStatus.textContent = state.firebaseWarning || (meta.sourceNames?.length ? meta.sourceNames.join(" + ") : "");
}

function render() {
  renderStatus();
  renderEmployees();
  renderSummary();
  renderTable();
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + Number(row[key] || 0), 0);
}

function escapeHtml(value) {
  return clean(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[char]);
}

function exportChanges() {
  const rows = [];
  Object.entries(state.temp?.problemFlags || {}).forEach(([key, value]) => {
    const parts = key.split("|");
    rows.push({
      changeType: "Wrong ticket marked",
      ticketType: parts[0] || "",
      ticketId: parts[1] || "",
      employee: parts[2] || "",
      reviewDate: localDateKeyFromValue(value?.reviewDate || value?.markedAt),
      status: "",
      note: "",
      changedAt: value?.markedAt || "",
      createdAt: "",
      updatedAt: value?.markedAt || "",
      firebaseKey: key,
    });
  });
  Object.entries(state.temp?.manualTickets || {}).forEach(([key, value]) => {
    rows.push({
      changeType: "Manual added ticket",
      ticketType: "manual",
      ticketId: value?.ticketId || "",
      employee: value?.employee || "",
      reviewDate: manualReviewDate(value),
      status: value?.status || "",
      note: value?.note || "",
      changedAt: value?.updatedAt || value?.createdAt || "",
      createdAt: value?.createdAt || "",
      updatedAt: value?.updatedAt || "",
      firebaseKey: key,
    });
  });
  const headers = ["Change Type", "Ticket Type", "Ticket ID", "Employee", "Review Date", "Status", "Note", "Changed At", "Created At", "Updated At", "Firebase Key"];
  const bodyRows = rows.length ? rows.map((row) => [
    row.changeType,
    row.ticketType,
    row.ticketId,
    row.employee,
    row.reviewDate,
    row.status,
    row.note,
    formatExportDate(row.changedAt),
    formatExportDate(row.createdAt),
    formatExportDate(row.updatedAt),
    row.firebaseKey,
  ]) : [["No changes", "", "", "", "", "", "", "", "", "", ""]];
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><table border="1"><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${bodyRows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table></body></html>`;
  const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `ticket-change-export-${localDateKey(new Date())}.xls`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function formatExportDate(value) {
  const date = new Date(value || "");
  if (Number.isNaN(date.getTime())) return clean(value);
  return date.toLocaleString("en-AU", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

async function saveTemp() {
  el.datasetStatus.textContent = "Saving changes...";
  const payload = buildTempPayload(state.temp);
  await persistTempData(payload);
  state.temp = normalizeTempData(payload);
  state.firebaseWarning = "";
  render();
}

function buildTempPayload(temp) {
  return {
    ...temp,
    meta: { ...(temp.meta || {}), updatedAt: new Date().toISOString(), source: "public-site" },
  };
}

async function persistTempData(temp) {
  const response = await fetch(TEMP_URL, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(buildTempPayload(temp)),
  });
  if (!response.ok) throw new Error("Firebase save failed.");
}

function openManualModal() {
  if (state.selectedEmployee) el.manualEmployee.value = state.selectedEmployee;
  el.modalBackdrop.classList.add("open");
}

function closeManualModal() {
  el.modalBackdrop.classList.remove("open");
  el.manualTicketForm.reset();
}

el.fileInput.addEventListener("change", async (event) => {
  try {
    await importFiles(event.target.files);
  } catch (error) {
    alert(error.message);
    el.datasetStatus.textContent = "Import failed.";
  } finally {
    event.target.value = "";
  }
});

el.refreshButton.addEventListener("click", loadData);
el.exportChangesButton.addEventListener("click", exportChanges);
el.dateFilter.value = state.selectedDate;
el.dateFilter.addEventListener("change", (event) => {
  state.selectedDate = event.target.value || yesterdayKey();
  el.dateFilter.value = state.selectedDate;
  render();
});
el.addTicketButton.addEventListener("click", openManualModal);
el.modalClose.addEventListener("click", closeManualModal);
el.modalBackdrop.addEventListener("click", (event) => {
  if (event.target === el.modalBackdrop) closeManualModal();
});
el.manualTicketForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const form = new FormData(event.currentTarget);
  const ticketId = clean(form.get("ticketId"));
  const employee = clean(form.get("employee"));
  if (!ticketId || !employee) return;
  const key = ["manual", ticketId, employee].join("|").toLowerCase();
  const existing = state.temp.manualTickets[key] || {};
  const now = new Date().toISOString();
  state.temp.manualTickets[key] = {
    ...existing,
    ticketId,
    employee,
    status: clean(form.get("status")) || "Manual added",
    note: clean(form.get("note")),
    createdAt: existing.createdAt || now,
    updatedAt: now,
    reviewDate: state.selectedDate || yesterdayKey(),
  };
  closeManualModal();
  state.tab = "manual";
  [...el.tabs.querySelectorAll("button")].forEach((tab) => tab.classList.toggle("active", tab.dataset.tab === "manual"));
  try {
    await saveTemp();
  } catch (error) {
    alert(error.message);
    el.datasetStatus.textContent = "Save failed.";
  }
});
el.employeeSearch.addEventListener("input", (event) => {
  state.employeeFilter = event.target.value;
  renderEmployees();
});
el.tableSearch.addEventListener("input", (event) => {
  state.tableFilter = event.target.value;
  renderTable();
});
el.tableBody.addEventListener("click", async (event) => {
  const moveButton = event.target.closest("[data-move-manual]");
  if (moveButton) {
    const key = moveButton.dataset.moveManual;
    const row = state.temp.manualTickets[key];
    if (row) {
      const problemKey = ["tickets", clean(row.ticketId), clean(row.employee)].join("|").toLowerCase();
      delete state.temp.problemFlags[problemKey];
      delete state.temp.manualTickets[key];
      try {
        await saveTemp();
      } catch (error) {
        alert(error.message);
        el.datasetStatus.textContent = "Save failed.";
      }
    }
    return;
  }
  const button = event.target.closest("[data-mark]");
  if (!button) return;
  const key = button.dataset.mark;
  if (state.temp.problemFlags[key]) {
    delete state.temp.problemFlags[key];
  } else {
    state.temp.problemFlags[key] = { key, markedAt: new Date().toISOString(), markedBy: "Web user", reviewDate: state.selectedDate || yesterdayKey() };
  }
  try {
    await saveTemp();
  } catch (error) {
    alert(error.message);
    el.datasetStatus.textContent = "Save failed.";
  }
});
el.employeeList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-employee]");
  if (!button) return;
  state.selectedEmployee = button.dataset.employee;
  render();
});
el.tabs.addEventListener("click", (event) => {
  const button = event.target.closest("[data-tab]");
  if (!button) return;
  state.tab = button.dataset.tab;
  [...el.tabs.querySelectorAll("button")].forEach((tab) => tab.classList.toggle("active", tab === button));
  renderTable();
});

loadData().catch((error) => {
  el.datasetStatus.textContent = error.message;
});
