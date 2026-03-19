// Configuration for Oplans and shared form fields
const OPLANS = [
  { id: "pagtugis", name: "Pagtugis" },
  { id: "paglalansag-omega", name: "Paglalansag Omega" },
  { id: "bolilyo", name: "Bolilyo" },
  { id: "lira", name: "Lira" },
  { id: "megashopper", name: "Megashopper" },
  { id: "kalikasan", name: "Kalikasan" },
  { id: "big-bertha", name: "Big Bertha" },
];

// Field configuration (mirrors the typical operational accomplishment Excel columns)
// Note: Flagship Project / Oplan and Number of Operations are handled automatically in code,
// not entered manually in the form.
const FIELDS = [
  {
    id: "date",
    label: "Date of Operation",
    type: "date",
    required: true,
  },
  {
    id: "time",
    label: "Time Sent",
    type: "time",
    required: false,
  },
  {
    id: "approvedPreOpsClearance",
    label: "Approved Pre-Ops Clearance",
    type: "text",
    required: false,
  },
  {
    id: "pfuCfuDfu",
    label: "Unit",
    type: "select",
    options: ["CAVITE", "LAGUNA", "BATANGAS", "RIZAL", "QUEZON"],
    required: false,
  },
  {
    id: "modeOfOperation",
    label: "Mode of Operation",
    type: "select",
    options: ["Search Warrant", "Buy Bust", "Entrapment"],
    required: false,
  },
  {
    id: "taggedListedAs",
    label: "Tagged/Listed As",
    type: "text",
    required: false,
  },
  {
    id: "placeRegion",
    label: "Place of Operation/Incident - Region",
    type: "text",
    required: false,
  },
  {
    id: "placeProvince",
    label: "Place of Operation/Incident - Province",
    type: "text",
    required: false,
  },
  {
    id: "placeMunicipalityCity",
    label: "Place of Operation/Incident - Municipality/City",
    type: "text",
    required: false,
  },
  {
    id: "numArrests",
    label: "Number of Persons Arrested",
    type: "number",
    min: 0,
    step: 1,
    required: true,
  },
  {
    id: "secondaryField",
    label: "Secondary (optional)",
    type: "text",
    required: false,
  },
  {
    id: "numFirearmsConfiscated",
    label: "Number of Firearms Confiscated",
    type: "text",
    inputMode: "numeric",
    required: false,
  },
  {
    id: "smallArms",
    label: "Small Arms",
    type: "text",
    inputMode: "numeric",
    required: false,
  },
  {
    id: "bigArms",
    label: "Big Arms",
    type: "text",
    inputMode: "numeric",
    required: false,
  },
  {
    id: "confiscatedItems",
    label: "Confiscated / Seized Items",
    type: "textarea",
    required: false,
  },
  {
    id: "amountOfConfiscated",
    label: "Amount of Confiscated",
    type: "text",
    inputMode: "decimal",
    placeholder: "e.g. 3,000",
    required: false,
  },
  {
    id: "reportFile",
    label: "Report Upload (PDF or DOCX)",
    type: "file",
    required: false,
  },
  {
    id: "remarks",
    label: "Remarks / Other Accomplishments",
    type: "textarea",
    required: false,
  },
];

const STORAGE_KEY = "oplan_daily_accomplishments_v1";
const CONFISCATION_COUNT_FIELDS = [
  "numFirearmsConfiscated",
  "smallArms",
  "bigArms",
];
const CONFISCATION_AMOUNT_FIELD = "amountOfConfiscated";

const state = {
  currentView: "dashboard",
  currentOplanId: null,
  editingId: null,
  isSaving: false,
  drawerOpen: false,
  searchQuery: "",
  recordsByOplan: {}, // { [oplanId]: Record[] }
  dashboardDateFilter: { month: "", day: "", year: "" },
  oplanDateFilters: {}, // { [oplanId]: { month, day, year } }
  reportFilters: {
    fromDate: "",
    toDate: "",
    month: "",
    year: "",
    oplanId: "",
    unit: "",
    modeOfOperation: "",
    status: "",
  },
};

// Attachments are stored in IndexedDB to support optional PDF/DOCX uploads.
const ATTACHMENTS_DB = "oplan_attachments_v1";
const ATTACHMENTS_STORE = "attachments";

function openAttachmentsDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(ATTACHMENTS_DB, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(ATTACHMENTS_STORE)) {
        db.createObjectStore(ATTACHMENTS_STORE, { keyPath: "id" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveAttachment(file) {
  const db = await openAttachmentsDb();
  const id = `att-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  const record = {
    id,
    name: file.name,
    type: file.type || "",
    size: file.size || 0,
    lastModified: file.lastModified || 0,
    blob: file,
  };
  await new Promise((resolve, reject) => {
    const tx = db.transaction(ATTACHMENTS_STORE, "readwrite");
    tx.objectStore(ATTACHMENTS_STORE).put(record);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
  return record;
}

async function getAttachment(id) {
  if (!id) return null;
  const db = await openAttachmentsDb();
  const result = await new Promise((resolve, reject) => {
    const tx = db.transaction(ATTACHMENTS_STORE, "readonly");
    const req = tx.objectStore(ATTACHMENTS_STORE).get(id);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
  db.close();
  return result;
}

function loadFromStorage() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function saveToStorage() {
  try {
    window.localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify(state.recordsByOplan)
    );
  } catch {
    // Ignore storage errors (e.g., quota exceeded)
  }
}

function ensureAllOplanArrays() {
  for (const o of OPLANS) {
    if (!Array.isArray(state.recordsByOplan[o.id])) {
      state.recordsByOplan[o.id] = [];
    }
    if (!state.oplanDateFilters[o.id]) {
      state.oplanDateFilters[o.id] = { month: "", day: "", year: "" };
    }
  }
}

function initState() {
  state.recordsByOplan = loadFromStorage();
  ensureAllOplanArrays();
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function formatNumber(value) {
  const n = Number(value);
  if (Number.isNaN(n)) return "";
  return n.toLocaleString();
}

function sanitizeNumericInput(value, allowDecimal = false) {
  const raw = String(value || "");
  let cleaned = raw.replace(/,/g, "").replace(/[^\d.]/g, "");

  if (!allowDecimal) {
    cleaned = cleaned.replace(/\./g, "");
  } else {
    const firstDot = cleaned.indexOf(".");
    if (firstDot !== -1) {
      cleaned =
        cleaned.slice(0, firstDot + 1) + cleaned.slice(firstDot + 1).replace(/\./g, "");
    }
  }

  return cleaned;
}

function formatAmountInputValue(value) {
  const cleaned = sanitizeNumericInput(value, true);
  if (!cleaned) return "";
  const parts = cleaned.split(".");
  const whole = parts[0] ? Number(parts[0]).toLocaleString() : "0";
  if (parts.length === 1) return whole;
  return `${whole}.${parts[1]}`;
}

function formatPeso(value) {
  const cleaned = sanitizeNumericInput(value, true);
  if (!cleaned) return "";
  const n = Number(cleaned);
  if (Number.isNaN(n)) return "";
  const hasDecimal = cleaned.includes(".");
  return `₱${n.toLocaleString(undefined, {
    minimumFractionDigits: hasDecimal ? 2 : 0,
    maximumFractionDigits: 2,
  })}`;
}

function getAllRecords() {
  const all = [];
  for (const o of OPLANS) {
    const arr = state.recordsByOplan[o.id] || [];
    for (const rec of arr) {
      all.push({ ...rec, oplanId: o.id });
    }
  }
  return all;
}

function getNextOperationNumber() {
  const all = getAllRecords();
  const maxOp = all.reduce(
    (max, r) => Math.max(max, Number(r.numOperations) || 0),
    0
  );
  return maxOp + 1;
}

function extractDateParts(dateStr) {
  const raw = String(dateStr || "").trim();
  if (!raw) return null;

  const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    return {
      year: Number(isoMatch[1]),
      month: Number(isoMatch[2]),
      day: Number(isoMatch[3]),
    };
  }

  const slashMatch = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (slashMatch) {
    return {
      year: Number(slashMatch[3]),
      month: Number(slashMatch[1]),
      day: Number(slashMatch[2]),
    };
  }

  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return null;
  return {
    year: d.getFullYear(),
    month: d.getMonth() + 1,
    day: d.getDate(),
  };
}

function passesDateFilter(dateStr, filter) {
  const parts = extractDateParts(dateStr);
  if (!parts) return false;

  if (filter.year && parts.year !== Number(filter.year)) return false;
  if (filter.month && parts.month !== Number(filter.month)) return false;
  if (filter.day && parts.day !== Number(filter.day)) return false;
  return true;
}

function updateTodayLabel() {
  const el = document.getElementById("today-label");
  if (!el) return;
  const now = new Date();
  const datePart = now.toLocaleDateString("en-GB", {
    weekday: "short",
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
  const timePart = now.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  el.textContent = `${datePart}, ${timePart}`;
}

function setContainerMode(modeClass) {
  const container = document.getElementById("view-container");
  if (!container) return null;
  container.classList.remove("dashboard-view", "oplan-view", "reports-view");
  if (modeClass) container.classList.add(modeClass);
  return container;
}

function enableResizableTable(tableId, storageKey) {
  const table = document.getElementById(tableId);
  if (!table) return;

  const headerRow = table.querySelector("thead tr");
  if (!headerRow) return;

  const headers = Array.from(headerRow.children);
  if (!headers.length) return;

  let colgroup = table.querySelector("colgroup.resizable-cols");
  if (!colgroup) {
    colgroup = document.createElement("colgroup");
    colgroup.className = "resizable-cols";
    table.insertBefore(colgroup, table.firstChild);
  }

  if (colgroup.children.length !== headers.length) {
    colgroup.innerHTML = "";
    headers.forEach(() => colgroup.appendChild(document.createElement("col")));
  }

  const defaultWidths = headers.map((th) =>
    Math.max(90, Math.round(th.getBoundingClientRect().width || 120))
  );

  let saved = null;
  try {
    const raw = window.localStorage.getItem(storageKey);
    saved = raw ? JSON.parse(raw) : null;
  } catch {
    saved = null;
  }

  const widths = headers.map((_, idx) => {
    const fromSaved = saved && Number(saved[idx]);
    return Number.isFinite(fromSaved) && fromSaved >= 90
      ? fromSaved
      : defaultWidths[idx];
  });

  const cols = Array.from(colgroup.children);
  cols.forEach((col, idx) => {
    col.style.width = `${widths[idx]}px`;
  });

  table.style.width = `${widths.reduce((sum, w) => sum + w, 0)}px`;

  headers.forEach((th, idx) => {
    th.classList.add("resizable-header");
    const existing = th.querySelector(".col-resizer");
    if (existing) existing.remove();

    const handle = document.createElement("span");
    handle.className = "col-resizer";
    handle.setAttribute("aria-hidden", "true");

    handle.addEventListener("mousedown", (e) => {
      e.preventDefault();
      const startX = e.clientX;
      const startWidth = widths[idx];

      function onMove(ev) {
        const delta = ev.clientX - startX;
        const next = Math.max(90, startWidth + delta);
        widths[idx] = next;
        cols[idx].style.width = `${next}px`;
        table.style.width = `${widths.reduce((sum, w) => sum + w, 0)}px`;
      }

      function onUp() {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        try {
          window.localStorage.setItem(storageKey, JSON.stringify(widths));
        } catch {
          // ignore storage errors
        }
      }

      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });

    th.appendChild(handle);
  });
}

// Rendering helpers
function setViewTitle(title, subtitle) {
  const titleEl = document.getElementById("view-title");
  const subtitleEl = document.getElementById("view-subtitle");
  if (titleEl) titleEl.textContent = title;
  if (subtitleEl) subtitleEl.textContent = subtitle || "";
}

function setActiveNav(button) {
  document
    .querySelectorAll(".nav-item")
    .forEach((btn) => btn.classList.remove("active"));
  if (button) button.classList.add("active");
}

// Dashboard
function renderDashboard() {
  state.currentView = "dashboard";
  state.currentOplanId = null;
  state.editingId = null;

  setViewTitle(
    "Dashboard",
    "Summary of operational accomplishments across all Oplans."
  );

  const container = setContainerMode("dashboard-view");

  const filter = state.dashboardDateFilter || {
    month: "",
    day: "",
    year: "",
  };
  const all = getAllRecords().filter((r) => {
    if (!filter.month && !filter.day && !filter.year) return true;
    return passesDateFilter(r.date, filter);
  });

  const totalOperations = all.length;
  const totalArrests = all.reduce(
    (sum, r) => sum + (Number(r.numArrests) || 0),
    0
  );

  const perOplan = {};
  for (const o of OPLANS) {
    perOplan[o.id] = { operations: 0, arrests: 0, count: 0 };
  }
  for (const r of all) {
    const agg = perOplan[r.oplanId];
    if (!agg) continue;
    agg.operations += 1;
    agg.arrests += Number(r.numArrests) || 0;
    agg.count += 1;
  }

  const perOplanRows = OPLANS.map((o) => {
    const agg = perOplan[o.id];
    return `
      <tr>
        <td class="text-left">${o.name}</td>
        <td>${formatNumber(agg.operations)}</td>
        <td>${formatNumber(agg.arrests)}</td>
        <td>${formatNumber(agg.count)}</td>
      </tr>
    `;
  }).join("");

  const detailedRows = all
    .map((r) => {
      const oplan = OPLANS.find((o) => o.id === r.oplanId);
      return `
        <tr data-oplan-id="${r.oplanId}" data-record-id="${r.id}">
          <td class="text-left">${oplan ? oplan.name : r.oplanId}</td>
          <td class="text-left">${formatDate(r.date)}</td>
          <td class="text-left">${r.approvedPreOpsClearance || ""}</td>
          <td>${formatNumber(r.numOperations || 1)}</td>
          <td>${formatNumber(r.numArrests)}</td>
          <td class="arrested-list-cell text-left">${formatArrestedPersonsNumberedHtml(
            r.arrestedPersons || []
          )}</td>
          <td class="text-left">${r.pfuCfuDfu || ""}</td>
          <td class="text-left">${r.modeOfOperation || ""}</td>
          <td class="text-left">${[r.placeRegion, r.placeProvince, r.placeMunicipalityCity]
            .filter(Boolean)
            .join(" / ")}</td>
          <td>${formatNumber(r.numFirearmsConfiscated)}</td>
          <td>${formatNumber(r.smallArms)}</td>
          <td>${formatNumber(r.bigArms)}</td>
          <td>${formatPeso(r.amountOfConfiscated)}</td>
          <td class="text-left">${r.reportAttachmentName || ""}</td>
          <td>
            <button type="button" class="btn secondary sm" data-dashboard-action="edit">Edit</button>
            <button type="button" class="btn danger sm" data-dashboard-action="delete">Delete</button>
          </td>
        </tr>
      `;
    })
    .join("");

  container.innerHTML = `
    <div class="card">
      <div class="form-header">
        <div>
          <div class="report-section-title">Filter by Date</div>
          <p class="small">Filter dashboard statistics by Month / Day / Year.</p>
        </div>
        <div class="report-filters">
          <div class="filter-group">
            <label for="dashboard-month">Month</label>
            <input type="number" id="dashboard-month" min="1" max="12" placeholder="MM">
          </div>
          <div class="filter-group">
            <label for="dashboard-day">Day</label>
            <input type="number" id="dashboard-day" min="1" max="31" placeholder="DD">
          </div>
          <div class="filter-group">
            <label for="dashboard-year">Year</label>
            <input type="number" id="dashboard-year" placeholder="YYYY">
          </div>
        </div>
      </div>
    </div>

    <div class="cards-row">
      <div class="card">
        <div class="card-title">Total Operations (All Oplans)</div>
        <div class="card-value">${formatNumber(totalOperations)}</div>
        <div class="card-subvalue">
          <span class="badge total">Total Arrests: ${formatNumber(
            totalArrests
          )}</span>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Total Records</div>
        <div class="card-value">${formatNumber(all.length)}</div>
        <div class="card-subvalue">Each record represents one operation entry.</div>
      </div>
    </div>

    <div class="card">
      <div class="form-header">
        <div>
          <div class="report-section-title">Operations by Oplan</div>
          <p class="small">Shows total number of operations and arrested persons per Oplan.</p>
        </div>
      </div>
      <div class="card-table">
        <table>
          <thead>
            <tr>
              <th class="text-left">Oplan</th>
              <th>Total Operations</th>
              <th>Total Arrests</th>
              <th>Records</th>
            </tr>
          </thead>
          <tbody>
            ${
              all.length === 0
                ? `<tr><td colspan="4" class="table-empty">No records yet. Use the Oplan forms to add daily accomplishments.</td></tr>`
                : perOplanRows
            }
          </tbody>
        </table>
      </div>
    </div>

    <div class="card dashboard-records-card">
      <div class="report-section-title">All Oplan Record Details</div>
      <div class="card-table dashboard-records-table-wrapper">
        <table id="dashboard-records-table">
          <thead>
            <tr>
              <th class="text-left">Oplan</th>
              <th class="text-left">Date</th>
              <th class="text-left">Approved Pre-Ops Clearance</th>
              <th>No. of Operation</th>
              <th>No. Arrested</th>
              <th class="text-left">Names of Arrested Persons</th>
              <th class="text-left">Unit</th>
              <th class="text-left">Mode of Operation</th>
              <th class="text-left">Place of Operation/Incident</th>
              <th>Firearms Confiscated</th>
              <th>Small Arms</th>
              <th>Big Arms</th>
              <th>Amount of Confiscated</th>
              <th class="text-left">Uploaded Report File</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${
              all.length === 0
                ? `<tr><td colspan="15" class="table-empty">No records yet across all Oplans.</td></tr>`
                : detailedRows
            }
          </tbody>
        </table>
      </div>
    </div>
  `;

  const monthInput = document.getElementById("dashboard-month");
  const dayInput = document.getElementById("dashboard-day");
  const yearInput = document.getElementById("dashboard-year");
  if (monthInput) monthInput.value = filter.month || "";
  if (dayInput) dayInput.value = filter.day || "";
  if (yearInput) yearInput.value = filter.year || "";

  const applyDashboardFilter = () => {
    state.dashboardDateFilter = {
      month: monthInput?.value.trim() || "",
      day: dayInput?.value.trim() || "",
      year: yearInput?.value.trim() || "",
    };
    renderDashboard();
  };

  [monthInput, dayInput, yearInput].forEach((el) => {
    if (!el) return;
    el.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      applyDashboardFilter();
    });
  });

  enableResizableTable("dashboard-records-table", "dashboard_records_table_cols");

  container.querySelectorAll("button[data-dashboard-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const action = btn.getAttribute("data-dashboard-action");
      const tr = btn.closest("tr");
      const oplanId = tr ? tr.getAttribute("data-oplan-id") : "";
      const recordId = tr ? tr.getAttribute("data-record-id") : "";
      if (!oplanId || !recordId) return;

      if (action === "edit") {
        renderOplanForm(oplanId);
        loadRecordIntoForm(recordId);
      } else if (action === "delete") {
        deleteDashboardRecord(oplanId, recordId);
      }
    });
  });
}

function deleteDashboardRecord(oplanId, recordId) {
  const confirmed = window.confirm(
    "Delete this record? This action cannot be undone."
  );
  if (!confirmed) return;

  const records = state.recordsByOplan[oplanId] || [];
  const idx = records.findIndex((r) => r.id === recordId);
  if (idx < 0) return;
  records.splice(idx, 1);
  state.recordsByOplan[oplanId] = records;
  saveToStorage();
  renderDashboard();
}

// Oplan form rendering
function renderFormField(field, labelOverride = "") {
  const requiredMark = field.required ? '<span class="required">*</span>' : "";
  const base = `
    <div class="form-field" data-field-id="${field.id}">
      <label for="${field.id}">
        ${labelOverride || field.label}
        ${requiredMark}
      </label>
  `;

  let inputEl = "";
  if (field.type === "textarea") {
    inputEl = `<textarea id="${field.id}" name="${field.id}"></textarea>`;
  } else if (field.type === "select") {
    const optionsHtml = (field.options || [])
      .map((opt) => `<option value="${opt}">${opt}</option>`)
      .join("");
    inputEl = `<select id="${field.id}" name="${field.id}">
      <option value="">Select</option>
      ${optionsHtml}
    </select>`;
  } else if (field.type === "file") {
    inputEl = `<input id="${field.id}" name="${field.id}" type="file" accept=".pdf,.docx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document">`;
  } else {
    const extraAttrs = [
      field.min != null ? `min="${field.min}"` : "",
      field.step != null ? `step="${field.step}"` : "",
      field.inputMode ? `inputmode="${field.inputMode}"` : "",
      field.placeholder ? `placeholder="${field.placeholder}"` : "",
    ]
      .filter(Boolean)
      .join(" ");
    inputEl = `<input id="${field.id}" name="${field.id}" type="${field.type}" ${extraAttrs}>`;
  }

  const hint =
    field.id === "numArrests"
      ? '<div class="field-hint">Enter the number of individuals arrested during the operation(s). The number of name fields below will adjust automatically.</div>'
      : "";

  const extra =
    field.id === "numArrests"
      ? '<div id="arrested-persons-container" class="arrested-persons-container"></div>'
      : "";

  return `${base}${inputEl}${hint}${extra}</div>`;
}

function openRecordDrawer() {
  const overlay = document.getElementById("record-drawer-overlay");
  if (!overlay) return;
  overlay.classList.add("open");
  document.body.classList.add("drawer-open");
  state.drawerOpen = true;
}

function closeRecordDrawer() {
  const overlay = document.getElementById("record-drawer-overlay");
  if (!overlay) return;
  overlay.classList.remove("open");
  document.body.classList.remove("drawer-open");
  state.drawerOpen = false;
}

function exportCurrentOplanRecords() {
  if (!state.currentOplanId) return;

  const filter = state.oplanDateFilters[state.currentOplanId] || {
    month: "",
    day: "",
    year: "",
  };
  const q = (state.searchQuery || "").toLowerCase();

  const records = (state.recordsByOplan[state.currentOplanId] || []).filter((r) => {
    if (filter.month || filter.day || filter.year) {
      if (!passesDateFilter(r.date, filter)) return false;
    }

    if (!q) return true;
    const haystack = [
      r.confiscatedItems,
      r.remarks,
      r.taggedListedAs,
      r.placeRegion,
      r.placeProvince,
      r.placeMunicipalityCity,
      r.modeOfOperation,
      r.pfuCfuDfu,
      ...(r.arrestedPersons || []).map(formatArrestedPersonName),
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase();
    return haystack.includes(q);
  });

  if (!records.length) {
    alert("No records to export.");
    return;
  }

  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert(
      "PDF library not loaded. Please make sure you are connected to the internet and try again."
    );
    return;
  }

  const oplan = OPLANS.find((o) => o.id === state.currentOplanId);
  const { rows: summaryRows } = buildSummaryByOplan(records);
  const summaryRow = summaryRows.find((r) => r.oplanId === state.currentOplanId);
  const totalOperations = records.length;
  const totalArrests = records.reduce(
    (sum, r) => sum + (Number(r.numArrests) || 0),
    0
  );

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const title = `${oplan ? oplan.name : state.currentOplanId} Daily Accomplishment Report`;
  doc.setFontSize(14);
  doc.text(title, 36, 28);
  doc.setFontSize(10);
  doc.text(
    `Filters: Month: ${filter.month || "All"} | Day: ${filter.day || "All"} | Year: ${
      filter.year || "All"
    } | Search: ${state.searchQuery || "None"}`,
    36,
    46
  );
  doc.text(
    `Total Operations: ${formatNumber(
      totalOperations
    )} | Total Arrests: ${formatNumber(totalArrests)}`,
    36,
    62
  );

  doc.autoTable({
    startY: 74,
    margin: { left: 36, right: 36 },
    head: [
      [
        "CIDG Flagship Projects",
        "No. of Operation Conducted",
        "Search Warrant",
        "Buy Bust",
        "Entrapment",
        "Total No. of Arrested",
      ],
    ],
    body: [
      [
        `OPLAN ${(oplan?.name || state.currentOplanId || "").toUpperCase()}`,
        String(summaryRow?.operations || 0),
        String(summaryRow?.searchWarrant || 0),
        String(summaryRow?.buyBust || 0),
        String(summaryRow?.entrapment || 0),
        String(summaryRow?.totalArrested || 0),
      ],
    ],
    styles: { fontSize: 9, halign: "center", valign: "middle" },
    columnStyles: { 0: { halign: "left" } },
    headStyles: {
      fillColor: [31, 79, 143],
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
  });

  const detailBody = records.map((r) => [
    formatDate(r.date),
    r.approvedPreOpsClearance || "",
    r.pfuCfuDfu || "",
    r.modeOfOperation || "",
    [r.placeRegion, r.placeProvince, r.placeMunicipalityCity]
      .filter(Boolean)
      .join(" / "),
    String(Number(r.numArrests) || 0),
    formatPeso(r.amountOfConfiscated),
    r.remarks || "",
  ]);

  doc.autoTable({
    startY: (doc.lastAutoTable?.finalY || 74) + 14,
    margin: { left: 36, right: 36, bottom: 28 },
    head: [
      [
        "Date",
        "Approved Pre-Ops Clearance",
        "Unit",
        "Mode of Operation",
        "Place",
        "No. Arrested",
        "Amount of Confiscated",
        "Remarks",
      ],
    ],
    body: detailBody,
    styles: { fontSize: 8.5, valign: "top", overflow: "linebreak" },
    headStyles: {
      fillColor: [31, 79, 143],
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    columnStyles: {
      0: { halign: "left" },
      1: { halign: "left" },
      2: { halign: "left" },
      3: { halign: "left" },
      4: { halign: "left" },
      5: { halign: "center" },
      6: { halign: "right" },
      7: { halign: "left" },
    },
  });

  doc.save(`${state.currentOplanId}-report.pdf`);
}

function renderOplanForm(oplanId) {
  const oplan = OPLANS.find((o) => o.id === oplanId);
  if (!oplan) return;

  state.currentView = "oplan";
  state.currentOplanId = oplanId;
  state.editingId = null;
  state.searchQuery = "";

  setViewTitle(
    `${oplan.name} Form`,
    "Record daily accomplishments for this Oplan. Each record represents one operation entry."
  );

  const container = setContainerMode("oplan-view");

  const fieldById = Object.fromEntries(FIELDS.map((f) => [f.id, f]));
  const spotlightFieldIds = [
    "modeOfOperation",
    "pfuCfuDfu",
    "time",
    "date",
    "numArrests",
  ];
  const spotlightOverrides = {
    modeOfOperation: "Action Taken",
    pfuCfuDfu: "Unit",
    time: "Time Sent",
    date: "Date Sent",
  };
  const spotlightFieldsHtml = spotlightFieldIds
    .map((fieldId) => {
      const field = fieldById[fieldId];
      if (!field) return "";
      return renderFormField(field, spotlightOverrides[fieldId] || "");
    })
    .join("");

  const additionalFieldsHtml = FIELDS.filter(
    (field) => !spotlightFieldIds.includes(field.id)
  )
    .map((field) => renderFormField(field))
    .join("");

  container.innerHTML = `
    <div class="records-shell">
      <div class="form-header">
        <div>
          <h2>${oplan.name} Daily Accomplishment</h2>
          <p class="muted">
            Manage records using quick filters and open the side panel to add new entries.
          </p>
        </div>
        <div class="form-actions">
          <button type="button" id="open-record-drawer-btn" class="btn primary">+ New Record</button>
          <button type="button" id="refresh-records-btn" class="btn secondary">Refresh</button>
          <button type="button" id="export-oplan-btn" class="btn secondary">Export report</button>
        </div>
      </div>

      <div class="cards-row oplan-summary-row">
        <div class="card oplan-summary-card">
          <div class="card-title">Total Operations (this Oplan)</div>
          <div class="card-value" id="oplan-total-operations">0</div>
        </div>
        <div class="card oplan-summary-card">
          <div class="card-title">Total Arrested (this Oplan)</div>
          <div class="card-value" id="oplan-total-arrests">0</div>
        </div>
      </div>

      <div class="form-toolbar">
        <div class="report-filters">
          <div class="filter-group">
            <label for="oplan-month">Month</label>
            <input type="number" id="oplan-month" min="1" max="12" placeholder="MM">
          </div>
          <div class="filter-group">
            <label for="oplan-day">Day</label>
            <input type="number" id="oplan-day" min="1" max="31" placeholder="DD">
          </div>
          <div class="filter-group">
            <label for="oplan-year">Year</label>
            <input type="number" id="oplan-year" placeholder="YYYY">
          </div>
        </div>
        <div class="form-toolbar-right">
          <div class="search-input">
            <input
              id="search-input"
              type="search"
              placeholder="Search by place, tagged, remarks…"
            >
          </div>
          <div class="small">
            <span class="badge" id="oplan-record-count">Records: 0</span>
          </div>
        </div>
      </div>

      <div class="table-wrapper">
        <table id="records-table">
          <thead>
            <tr>
              <th class="text-left">Date</th>
              <th class="text-left">Approved Pre-Ops Clearance</th>
              <th>No. of Operations</th>
              <th>No. Arrested</th>
              <th class="text-left">Names of Arrested Persons</th>
              <th class="text-left">Unit</th>
              <th class="text-left">Mode of Operation</th>
              <th class="text-left">Place (Region/Province/City)</th>
              <th class="text-left">Tagged/Listed As</th>
              <th>Firearms Confiscated</th>
              <th>Small Arms</th>
              <th>Big Arms</th>
              <th class="text-left">Confiscated Items</th>
              <th>Amount of Confiscated</th>
              <th class="text-left">Report Upload</th>
              <th class="text-left">Remarks</th>
              <th style="width: 120px;">Actions</th>
            </tr>
          </thead>
          <tbody id="records-tbody"></tbody>
        </table>
      </div>
    </div>

    <div id="record-drawer-overlay" class="record-drawer-overlay" aria-hidden="true">
      <aside id="record-drawer" class="record-drawer" role="dialog" aria-modal="true" aria-label="${oplan.name} Report">
        <div class="record-drawer-header">
          <h3>${oplan.name} Report</h3>
          <button type="button" id="close-record-drawer-btn" class="drawer-close-btn" aria-label="Close form">×</button>
        </div>

        <div class="record-drawer-body">
          <form id="record-form" class="record-form drawer-form" novalidate>
            <section class="drawer-group">
              <h4>Disposition</h4>
              ${spotlightFieldsHtml}
            </section>
            <section class="drawer-group">
              <h4>Additional Details</h4>
              ${additionalFieldsHtml}
            </section>
          </form>
        </div>

        <div class="record-drawer-actions">
          <button type="button" id="save-record-btn" class="btn primary">Save</button>
          <button type="button" id="clear-record-btn" class="btn secondary">Clear</button>
        </div>
      </aside>
    </div>
  `;

  document
    .getElementById("open-record-drawer-btn")
    .addEventListener("click", () => {
      state.editingId = null;
      resetForm();
      openRecordDrawer();
    });

  document
    .getElementById("close-record-drawer-btn")
    .addEventListener("click", closeRecordDrawer);

  document
    .getElementById("record-drawer-overlay")
    .addEventListener("click", (e) => {
      if (e.target && e.target.id === "record-drawer-overlay") {
        closeRecordDrawer();
      }
    });

  document
    .getElementById("clear-record-btn")
    .addEventListener("click", resetForm);

  document
    .getElementById("refresh-records-btn")
    .addEventListener("click", renderRecordsTable);

  document
    .getElementById("export-oplan-btn")
    .addEventListener("click", exportCurrentOplanRecords);

  document
    .getElementById("save-record-btn")
    .addEventListener("click", handleSaveRecord);

  setupFormUppercaseBehavior();
  setupConfiscationInputBehavior();

  const recordForm = document.getElementById("record-form");
  if (recordForm) {
    recordForm.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      const target = e.target;
      if (target && target.tagName && target.tagName.toLowerCase() === "textarea") return;
      e.preventDefault();
      handleSaveRecord();
    });
  }

  const numArrestsInput = document.getElementById("numArrests");
  if (numArrestsInput) {
    numArrestsInput.addEventListener("input", handleNumArrestsChange);
  }

  const searchInput = document.getElementById("search-input");
  if (searchInput) {
    searchInput.addEventListener("input", (e) => {
      state.searchQuery = e.target.value || "";
      renderRecordsTable();
    });
    searchInput.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      state.searchQuery = searchInput.value || "";
      renderRecordsTable();
    });
  }

  const applyOplanFilter = () => {
    state.oplanDateFilters[oplanId] = {
      month: document.getElementById("oplan-month")?.value.trim() || "",
      day: document.getElementById("oplan-day")?.value.trim() || "",
      year: document.getElementById("oplan-year")?.value.trim() || "",
    };
    renderRecordsTable();
  };

  ["oplan-month", "oplan-day", "oplan-year"].forEach((id) => {
    const input = document.getElementById(id);
    if (!input) return;
    input.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      applyOplanFilter();
    });
  });

  const oplanFilter = state.oplanDateFilters[oplanId] || {
    month: "",
    day: "",
    year: "",
  };
  const monthInput = document.getElementById("oplan-month");
  const dayInput = document.getElementById("oplan-day");
  const yearInput = document.getElementById("oplan-year");
  if (monthInput) monthInput.value = oplanFilter.month || "";
  if (dayInput) dayInput.value = oplanFilter.day || "";
  if (yearInput) yearInput.value = oplanFilter.year || "";

  renderRecordsTable();
}

function getFormElement() {
  return document.getElementById("record-form");
}

function normalizeToUpper(value) {
  return String(value ?? "").toUpperCase();
}

function setupFormUppercaseBehavior() {
  const form = getFormElement();
  if (!form) return;

  form.addEventListener("input", (e) => {
    const target = e.target;
    if (!target || !target.tagName) return;

    const tag = target.tagName.toLowerCase();
    if (tag !== "input" && tag !== "textarea") return;

    if (tag === "input") {
      const type = (target.getAttribute("type") || "text").toLowerCase();
      if (type !== "text") return;
    }

    const start = target.selectionStart;
    const end = target.selectionEnd;
    const upper = normalizeToUpper(target.value);
    if (upper === target.value) return;

    target.value = upper;
    if (
      start != null &&
      end != null &&
      typeof target.setSelectionRange === "function"
    ) {
      target.setSelectionRange(start, end);
    }
  });
}

function setupConfiscationInputBehavior() {
  const form = getFormElement();
  if (!form) return;

  for (const fieldId of CONFISCATION_COUNT_FIELDS) {
    const input = form.elements[fieldId];
    if (!input) continue;
    input.addEventListener("input", () => {
      input.value = sanitizeNumericInput(input.value, false);
    });
  }

  const amountInput = form.elements[CONFISCATION_AMOUNT_FIELD];
  if (amountInput) {
    amountInput.addEventListener("input", () => {
      amountInput.value = formatAmountInputValue(amountInput.value);
    });
  }
}

function resetForm() {
  const form = getFormElement();
  if (!form) return;
  form.reset();
  state.editingId = null;
  form
    .querySelectorAll(".form-field")
    .forEach((f) => f.classList.remove("invalid"));

  updateArrestedPersonsInputs(0, []);
}

function readFormValues() {
  const form = getFormElement();
  if (!form) return null;
  const result = {};
  for (const field of FIELDS) {
    const el = form.elements[field.id];
    if (!el) continue;
    if (field.type === "file") {
      result[field.id] = "";
      continue;
    }
    const raw = el.value != null ? el.value.trim() : "";
    if (CONFISCATION_COUNT_FIELDS.includes(field.id)) {
      result[field.id] = sanitizeNumericInput(raw, false);
      continue;
    }
    if (field.id === CONFISCATION_AMOUNT_FIELD) {
      result[field.id] = sanitizeNumericInput(raw, true);
      continue;
    }
    result[field.id] =
      field.type === "text" || field.type === "textarea"
        ? normalizeToUpper(raw)
        : raw;
  }

  const arrestedContainer = document.getElementById(
    "arrested-persons-container"
  );
  const arrestedPersons = [];
  if (arrestedContainer) {
    arrestedContainer.querySelectorAll(".arrested-person-field").forEach((row) => {
      const lastName = normalizeToUpper(
        (row.querySelector('input[data-arrested-part="lastName"]')?.value || "").trim()
      );
      const firstName = normalizeToUpper(
        (row.querySelector('input[data-arrested-part="firstName"]')?.value || "").trim()
      );
      const middleName = normalizeToUpper(
        (row.querySelector('input[data-arrested-part="middleName"]')?.value || "").trim()
      );
      const suffix = normalizeToUpper(
        (row.querySelector('input[data-arrested-part="suffix"]')?.value || "").trim()
      );
      arrestedPersons.push({ lastName, firstName, middleName, suffix });
    });
  }
  result.arrestedPersons = arrestedPersons;

  return result;
}

function formatArrestedPersonName(p) {
  if (!p) return "";
  // Backward compatibility: if older data stored a string, return it.
  if (typeof p === "string") return normalizeToUpper(p);
  const last = normalizeToUpper((p.lastName || "").trim());
  const first = normalizeToUpper((p.firstName || "").trim());
  const middle = normalizeToUpper((p.middleName || "").trim());
  const suffix = normalizeToUpper((p.suffix || "").trim());
  const middleInitial = middle ? ` ${middle}` : "";
  const suffixPart = suffix ? ` ${suffix}` : "";
  if (!last && !first && !middle && !suffix) return "";
  if (last && first) return `${last}, ${first}${middleInitial}${suffixPart}`.trim();
  return `${[last, first, middle, suffix].filter(Boolean).join(" ")}`.trim();
}

function formatArrestedPersonsNumberedText(persons) {
  const names = (persons || []).map(formatArrestedPersonName).filter(Boolean);
  if (!names.length) return "";
  return names.map((name, idx) => `${idx + 1}) ${name}`).join("\n");
}

function formatArrestedPersonsNumberedHtml(persons) {
  const names = (persons || []).map(formatArrestedPersonName).filter(Boolean);
  if (!names.length) return "";
  return names
    .map((name, idx) => `<div>${idx + 1}) ${name}</div>`)
    .join("");
}

function isAllowedReportFile(file) {
  if (!file) return false;
  const name = (file.name || "").toLowerCase();
  const okExt = name.endsWith(".pdf") || name.endsWith(".docx");
  const okType =
    file.type === "application/pdf" ||
    file.type ===
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    file.type === "";
  return okExt && okType;
}

function validateForm(values) {
  const form = getFormElement();
  if (!form) return false;
  let valid = true;
  for (const field of FIELDS) {
    const wrapper = form.querySelector(`[data-field-id="${field.id}"]`);
    if (!wrapper) continue;
    wrapper.classList.remove("invalid");
    if (field.required) {
      const value = (values[field.id] ?? "").trim();
      if (!value) {
        valid = false;
        wrapper.classList.add("invalid");
      }
    }
  }
  return valid;
}

async function handleSaveRecord() {
  if (!state.currentOplanId || state.isSaving) return;
  const values = readFormValues();
  if (!values) return;

  const isValid = validateForm(values);
  if (!isValid) {
    alert("Please fill in all required fields.");
    return;
  }

  state.isSaving = true;
  const oplanId = state.currentOplanId;
  const records = state.recordsByOplan[oplanId] || [];

  try {
    // Optional report upload
    const reportFileInput = document.getElementById("reportFile");
    let reportAttachment = null;
    if (reportFileInput && reportFileInput.files && reportFileInput.files[0]) {
      const file = reportFileInput.files[0];
      if (!isAllowedReportFile(file)) {
        alert("Report upload must be a PDF or DOCX file only.");
        return;
      }
      reportAttachment = await saveAttachment(file);
    }

    if (state.editingId) {
      const idx = records.findIndex((r) => r.id === state.editingId);
      if (idx >= 0) {
        records[idx] = {
          ...records[idx],
          ...values,
          flagshipProject: oplanId,
          numOperations: records[idx].numOperations || 1,
          reportAttachmentId: reportAttachment
            ? reportAttachment.id
            : records[idx].reportAttachmentId || "",
          reportAttachmentName: reportAttachment
            ? reportAttachment.name
            : records[idx].reportAttachmentName || "",
          reportAttachmentType: reportAttachment
            ? reportAttachment.type
            : records[idx].reportAttachmentType || "",
        };
      }
    } else {
      const newRecord = {
        id: `${oplanId}-${Date.now()}-${Math.random()
          .toString(16)
          .slice(2, 8)}`,
        ...values,
        flagshipProject: oplanId,
        numOperations: getNextOperationNumber(),
        reportAttachmentId: reportAttachment ? reportAttachment.id : "",
        reportAttachmentName: reportAttachment ? reportAttachment.name : "",
        reportAttachmentType: reportAttachment ? reportAttachment.type : "",
      };
      records.push(newRecord);
    }

    state.recordsByOplan[oplanId] = records;
    saveToStorage();
    resetForm();
    closeRecordDrawer();
    renderRecordsTable();
    // Update dashboard counters in case the user navigates back
  } finally {
    state.isSaving = false;
  }
}

function renderRecordsTable() {
  if (!state.currentOplanId) return;
  const tbody = document.getElementById("records-tbody");
  if (!tbody) return;

  const filter = state.oplanDateFilters[state.currentOplanId] || {
    month: "",
    day: "",
    year: "",
  };

  const records = (state.recordsByOplan[state.currentOplanId] || []).filter(
    (r) => {
      if (!filter.month && !filter.day && !filter.year) return true;
      return passesDateFilter(r.date, filter);
    }
  );
  const q = (state.searchQuery || "").toLowerCase();

  const filtered = records.filter((r) => {
    if (!q) return true;
    const haystack = [
      r.confiscatedItems,
      r.remarks,
      r.taggedListedAs,
      r.placeRegion,
      r.placeProvince,
      r.placeMunicipalityCity,
      r.modeOfOperation,
      r.pfuCfuDfu,
      ...(r.arrestedPersons || []).map(formatArrestedPersonName),
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase();
    return haystack.includes(q);
  });

  const countBadge = document.getElementById("oplan-record-count");
  if (countBadge) {
    countBadge.textContent = `Records: ${filtered.length}`;
  }

  const totalOpsEl = document.getElementById("oplan-total-operations");
  const totalArrestsEl = document.getElementById("oplan-total-arrests");
  const totalOperations = filtered.length;
  const totalArrests = filtered.reduce(
    (sum, r) => sum + (Number(r.numArrests) || 0),
    0
  );
  if (totalOpsEl) totalOpsEl.textContent = formatNumber(totalOperations);
  if (totalArrestsEl) totalArrestsEl.textContent = formatNumber(totalArrests);

  enableResizableTable(
    "records-table",
    `records_table_cols_${state.currentOplanId || "default"}`
  );

  if (!filtered.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="17" class="table-empty">
          No records yet for this Oplan. Use the form above to add a new record.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = filtered
    .map(
      (r) => `
      <tr data-record-id="${r.id}">
        <td class="text-left">${formatDate(r.date)}</td>
        <td class="text-left">${r.approvedPreOpsClearance || ""}</td>
        <td>${formatNumber(r.numOperations || 1)}</td>
        <td>${formatNumber(r.numArrests)}</td>
        <td class="arrested-list-cell">${formatArrestedPersonsNumberedHtml(
          r.arrestedPersons || []
        )}</td>
        <td class="text-left">${r.pfuCfuDfu || ""}</td>
        <td class="text-left">${r.modeOfOperation || ""}</td>
        <td class="text-left">${[r.placeRegion, r.placeProvince, r.placeMunicipalityCity]
          .filter(Boolean)
          .join(" / ")}</td>
        <td class="text-left">${r.taggedListedAs || ""}</td>
        <td>${formatNumber(r.numFirearmsConfiscated)}</td>
        <td>${formatNumber(r.smallArms)}</td>
        <td>${formatNumber(r.bigArms)}</td>
        <td class="text-left">${r.confiscatedItems || ""}</td>
        <td>${formatPeso(r.amountOfConfiscated)}</td>
        <td class="text-left">${
          r.reportAttachmentName
            ? `<button type="button" class="btn secondary sm" data-action="view-attachment">${r.reportAttachmentName}</button>`
            : ""
        }</td>
        <td class="text-left">${r.remarks || ""}</td>
        <td>
          <button type="button" class="btn secondary sm" data-action="edit">Edit</button>
          <button type="button" class="btn danger sm" data-action="delete">Delete</button>
        </td>
      </tr>
    `
    )
    .join("");

  tbody.querySelectorAll("button[data-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const action = btn.getAttribute("data-action");
      const tr = btn.closest("tr");
      const id = tr ? tr.getAttribute("data-record-id") : null;
      if (!id) return;
      if (action === "edit") {
        loadRecordIntoForm(id);
      } else if (action === "delete") {
        deleteRecord(id);
      } else if (action === "view-attachment") {
        viewAttachmentForRecord(id);
      }
    });
  });
}

async function viewAttachmentForRecord(recordId) {
  if (!state.currentOplanId) return;
  const records = state.recordsByOplan[state.currentOplanId] || [];
  const rec = records.find((r) => r.id === recordId);
  if (!rec || !rec.reportAttachmentId) return;
  const att = await getAttachment(rec.reportAttachmentId);
  if (!att || !att.blob) {
    alert("Attachment not found (it may have been cleared).");
    return;
  }
  const url = URL.createObjectURL(att.blob);
  const lower = (att.name || "").toLowerCase();
  if (lower.endsWith(".pdf")) {
    window.open(url, "_blank");
    setTimeout(() => URL.revokeObjectURL(url), 5 * 60 * 1000);
  } else {
    const a = document.createElement("a");
    a.href = url;
    a.download = att.name || "report";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}

function loadRecordIntoForm(id) {
  if (!state.currentOplanId) return;
  const records = state.recordsByOplan[state.currentOplanId] || [];
  const rec = records.find((r) => r.id === id);
  if (!rec) return;
  const form = getFormElement();
  if (!form) return;

  for (const field of FIELDS) {
    const el = form.elements[field.id];
    if (!el) continue;
    if (field.type === "file") continue;
    const raw = rec[field.id] != null ? rec[field.id] : "";
    if (CONFISCATION_COUNT_FIELDS.includes(field.id)) {
      el.value = sanitizeNumericInput(raw, false);
      continue;
    }
    if (field.id === CONFISCATION_AMOUNT_FIELD) {
      el.value = formatAmountInputValue(raw);
      continue;
    }
    el.value =
      field.type === "text" || field.type === "textarea"
        ? normalizeToUpper(raw)
        : raw;
  }

  const numArrestsInput = form.elements["numArrests"];
  if (numArrestsInput) {
    const count =
      (Array.isArray(rec.arrestedPersons) && rec.arrestedPersons.length) ||
      Number(numArrestsInput.value) ||
      0;
    updateArrestedPersonsInputs(count, rec.arrestedPersons || []);
  }
  state.editingId = id;
  openRecordDrawer();
}

function handleNumArrestsChange(e) {
  const value = e.target.value;
  const count = Math.max(0, Math.min(50, parseInt(value || "0", 10) || 0));
  const existing = [];
  const container = document.getElementById("arrested-persons-container");
  if (container) {
    container
      .querySelectorAll("input[data-arrested-index]")
      .forEach((input) => existing.push(input.value || ""));
  }
  updateArrestedPersonsInputs(count, existing);
}

function updateArrestedPersonsInputs(count, existingValues) {
  const container = document.getElementById("arrested-persons-container");
  if (!container) return;
  container.innerHTML = "";
  if (!count || count <= 0) return;
  const values = Array.isArray(existingValues) ? existingValues : [];
  for (let i = 0; i < count; i++) {
    const wrapper = document.createElement("div");
    wrapper.className = "arrested-person-field";

    const title = document.createElement("div");
    title.className = "arrested-person-title";
    title.textContent = `Arrested Person ${i + 1}`;

    const row = document.createElement("div");
    row.className = "arrested-person-grid";

    const existing = values[i] || {};
    const normalized =
      typeof existing === "string"
        ? { lastName: existing, firstName: "", middleName: "", suffix: "" }
        : existing;

    row.appendChild(makeArrestedNameInput("Last Name", "lastName", normalized.lastName || ""));
    row.appendChild(makeArrestedNameInput("First Name", "firstName", normalized.firstName || ""));
    row.appendChild(makeArrestedNameInput("Middle Name", "middleName", normalized.middleName || ""));
    row.appendChild(makeArrestedNameInput("Suffix", "suffix", normalized.suffix || ""));

    wrapper.appendChild(title);
    wrapper.appendChild(row);
    container.appendChild(wrapper);
  }
}

function makeArrestedNameInput(labelText, part, value) {
  const wrap = document.createElement("div");
  wrap.className = "arrested-name-part";
  const label = document.createElement("label");
  label.textContent = labelText;
  const input = document.createElement("input");
  input.type = "text";
  input.value = normalizeToUpper(value || "");
  input.setAttribute("data-arrested-part", part);
  wrap.appendChild(label);
  wrap.appendChild(input);
  return wrap;
}

function deleteRecord(id) {
  if (!state.currentOplanId) return;
  const confirmed = window.confirm(
    "Delete this record? This action cannot be undone."
  );
  if (!confirmed) return;
  const oplanId = state.currentOplanId;
  const records = state.recordsByOplan[oplanId] || [];
  const idx = records.findIndex((r) => r.id === id);
  if (idx >= 0) {
    records.splice(idx, 1);
    state.recordsByOplan[oplanId] = records;
    saveToStorage();
    if (state.editingId === id) {
      resetForm();
    }
    renderRecordsTable();
  }
}

// Reports
function filterRecordsByDateRange(fromDateStr, toDateStr) {
  const all = getAllRecords();
  if (!fromDateStr && !toDateStr) return all;

  const from = fromDateStr ? new Date(fromDateStr) : null;
  const to = toDateStr ? new Date(toDateStr) : null;
  if (to) {
    to.setHours(23, 59, 59, 999);
  }

  return all.filter((r) => {
    if (!r.date) return false;
    const d = new Date(r.date);
    if (Number.isNaN(d.getTime())) return false;
    if (from && d < from) return false;
    if (to && d > to) return false;
    return true;
  });
}

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function readReportFiltersFromDom() {
  return {
    fromDate: (document.getElementById("from-date")?.value || "").trim(),
    toDate: (document.getElementById("to-date")?.value || "").trim(),
    month: (document.getElementById("report-month")?.value || "").trim(),
    year: (document.getElementById("report-year")?.value || "").trim(),
    oplanId: (document.getElementById("report-oplan")?.value || "").trim(),
    unit: (document.getElementById("report-unit")?.value || "").trim(),
    modeOfOperation: (
      document.getElementById("report-mode")?.value || ""
    ).trim(),
    status: (document.getElementById("report-status")?.value || "").trim(),
  };
}

function getFilteredReportRecords(filters) {
  let rows = filterRecordsByDateRange(filters.fromDate, filters.toDate);

  rows = rows.filter((r) => {
    if (filters.month || filters.year) {
      if (
        !passesDateFilter(r.date, {
          month: filters.month,
          year: filters.year,
        })
      ) {
        return false;
      }
    }

    if (filters.oplanId && String(r.oplanId) !== String(filters.oplanId)) return false;
    if (filters.unit && normalizeText(r.pfuCfuDfu) !== normalizeText(filters.unit)) {
      return false;
    }
    if (
      filters.modeOfOperation &&
      normalizeText(r.modeOfOperation) !== normalizeText(filters.modeOfOperation)
    ) {
      return false;
    }
    if (filters.status && normalizeText(r.taggedListedAs) !== normalizeText(filters.status)) {
      return false;
    }

    return true;
  });

  return rows;
}

function buildSummaryByOplan(records) {
  const perOplan = {};
  for (const o of OPLANS) {
    perOplan[o.id] = {
      oplanId: o.id,
      oplanName: o.name,
      operations: 0,
      searchWarrant: 0,
      buyBust: 0,
      entrapment: 0,
      totalArrested: 0,
    };
  }

  for (const r of records) {
    const agg = perOplan[r.oplanId];
    if (!agg) continue;

    agg.operations += 1;
    agg.totalArrested += Number(r.numArrests) || 0;

    const mode = normalizeText(r.modeOfOperation);
    if (mode === "search warrant") agg.searchWarrant += 1;
    if (mode === "buy bust") agg.buyBust += 1;
    if (mode === "entrapment") agg.entrapment += 1;
  }

  const rows = OPLANS.map((o) => perOplan[o.id]);
  const total = rows.reduce(
    (acc, row) => ({
      operations: acc.operations + row.operations,
      searchWarrant: acc.searchWarrant + row.searchWarrant,
      buyBust: acc.buyBust + row.buyBust,
      entrapment: acc.entrapment + row.entrapment,
      totalArrested: acc.totalArrested + row.totalArrested,
    }),
    {
      operations: 0,
      searchWarrant: 0,
      buyBust: 0,
      entrapment: 0,
      totalArrested: 0,
    }
  );

  return { rows, total };
}

function getAppliedReportFilterItems(filters) {
  const items = [];

  if (filters.fromDate || filters.toDate) {
    items.push(`Date Range: ${filters.fromDate || "-"} to ${filters.toDate || "-"}`);
  }
  if (filters.month) items.push(`Month: ${filters.month}`);
  if (filters.year) items.push(`Year: ${filters.year}`);
  if (filters.oplanId) {
    const oplan = OPLANS.find((o) => o.id === filters.oplanId);
    items.push(`Oplan: ${oplan ? oplan.name : filters.oplanId}`);
  }
  if (filters.unit) items.push(`Unit: ${filters.unit}`);
  if (filters.modeOfOperation) items.push(`Mode of Operation: ${filters.modeOfOperation}`);
  if (filters.status) items.push(`Status: ${filters.status}`);

  return items;
}

function renderReportsView() {
  state.currentView = "reports";
  state.currentOplanId = null;
  state.editingId = null;

  setViewTitle(
    "Reports",
    "Generate filter-aware summaries and detailed records across all Oplans."
  );

  const container = setContainerMode("reports-view");
  const currentFilters = state.reportFilters || {};
  const unitField = FIELDS.find((f) => f.id === "pfuCfuDfu");
  const modeField = FIELDS.find((f) => f.id === "modeOfOperation");
  const statusOptions = Array.from(
    new Set(
      getAllRecords()
        .map((r) => String(r.taggedListedAs || "").trim())
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b));

  container.innerHTML = `
    <div class="card report-layout">
      <div class="form-header">
        <div>
          <div class="report-section-title">Report Filters</div>
          <p class="small">
            The preview and downloadable report both use the same filtered dataset shown below.
          </p>
        </div>
        <div class="form-actions">
          <button type="button" id="generate-report-btn" class="btn primary">View Report</button>
          <button type="button" id="download-report-btn" class="btn secondary">Download Report (PDF)</button>
        </div>
      </div>
      <div class="report-filters">
        <div class="filter-group">
          <label for="from-date">From date</label>
          <input type="date" id="from-date">
        </div>
        <div class="filter-group">
          <label for="to-date">To date</label>
          <input type="date" id="to-date">
        </div>
        <div class="filter-group">
          <label for="report-month">Month</label>
          <input type="number" id="report-month" min="1" max="12" placeholder="MM">
        </div>
        <div class="filter-group">
          <label for="report-year">Year</label>
          <input type="number" id="report-year" placeholder="YYYY">
        </div>
        <div class="filter-group">
          <label for="report-oplan">Oplan</label>
          <select id="report-oplan">
            <option value="">All</option>
            ${OPLANS.map((o) => `<option value="${o.id}">${o.name}</option>`).join("")}
          </select>
        </div>
        <div class="filter-group">
          <label for="report-unit">Unit</label>
          <select id="report-unit">
            <option value="">All</option>
            ${(unitField?.options || []).map((opt) => `<option value="${opt}">${opt}</option>`).join("")}
          </select>
        </div>
        <div class="filter-group">
          <label for="report-mode">Mode of Operation</label>
          <select id="report-mode">
            <option value="">All</option>
            ${(modeField?.options || []).map((opt) => `<option value="${opt}">${opt}</option>`).join("")}
          </select>
        </div>
        <div class="filter-group">
          <label for="report-status">Status</label>
          <select id="report-status">
            <option value="">All</option>
            ${statusOptions.map((opt) => `<option value="${opt}">${opt}</option>`).join("")}
          </select>
        </div>
      </div>
    </div>

    <div id="report-results" class="report-layout">
      <!-- Summary and tables will be rendered here -->
    </div>
  `;

  document
    .getElementById("generate-report-btn")
    .addEventListener("click", generateReport);
  document
    .getElementById("download-report-btn")
    .addEventListener("click", downloadReportPdf);

  const reportFilterIds = [
    "from-date",
    "to-date",
    "report-month",
    "report-year",
    "report-oplan",
    "report-unit",
    "report-mode",
    "report-status",
  ];
  reportFilterIds.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      generateReport();
    });
  });

  const byId = (id) => document.getElementById(id);
  byId("from-date").value = currentFilters.fromDate || "";
  byId("to-date").value = currentFilters.toDate || "";
  byId("report-month").value = currentFilters.month || "";
  byId("report-year").value = currentFilters.year || "";
  byId("report-oplan").value = currentFilters.oplanId || "";
  byId("report-unit").value = currentFilters.unit || "";
  byId("report-mode").value = currentFilters.modeOfOperation || "";
  byId("report-status").value = currentFilters.status || "";
}

function generateReport() {
  const filters = readReportFiltersFromDom();
  state.reportFilters = { ...filters };

  const resultsEl = document.getElementById("report-results");
  const filtered = getFilteredReportRecords(filters);
  const { rows: summaryRows, total: summaryTotal } = buildSummaryByOplan(filtered);

  const totalOperations = filtered.length;
  const totalArrests = filtered.reduce(
    (sum, r) => sum + (Number(r.numArrests) || 0),
    0
  );

  const summaryTableRows = summaryRows
    .map(
      (row) => `
      <tr>
        <td class="text-left">OPLAN ${row.oplanName.toUpperCase()}</td>
        <td>${formatNumber(row.operations)}</td>
        <td>${formatNumber(row.searchWarrant)}</td>
        <td>${formatNumber(row.buyBust)}</td>
        <td>${formatNumber(row.entrapment)}</td>
        <td>${formatNumber(row.totalArrested)}</td>
      </tr>
    `
    )
    .join("");

  const appliedFilters = getAppliedReportFilterItems(filters);
  const recordScopeLabel =
    filters.fromDate || filters.toDate
      ? `From <strong>${filters.fromDate || "–"}</strong> to <strong>${
          filters.toDate || "–"
        }</strong>`
      : "All available records";
  const filtersLabel = appliedFilters.length
    ? appliedFilters.join(" | ")
    : "None (All records)";

  resultsEl.innerHTML = `
    <div class="card report-output-card">
      <div class="report-output-title">Daily Accomplishment Report (Oplan)</div>
      <div class="report-output-meta small">
        <div>${recordScopeLabel}</div>
        <div>Filters: ${filtersLabel}</div>
        <div>Total Operations: ${formatNumber(totalOperations)} | Total Arrests: ${formatNumber(
    totalArrests
  )}</div>
      </div>

      <div class="card-table report-summary-table-wrap">
        <table>
          <thead>
            <tr>
              <th class="text-left">CIDG Flagship Projects</th>
              <th>No. of Operation Conducted</th>
              <th>Search Warrant</th>
              <th>Buy Bust</th>
              <th>Entrapment</th>
              <th>Total No. of Arrested</th>
            </tr>
          </thead>
          <tbody>
            ${summaryTableRows}
            <tr>
              <th class="text-left">TOTAL</th>
              <th>${formatNumber(summaryTotal.operations)}</th>
              <th>${formatNumber(summaryTotal.searchWarrant)}</th>
              <th>${formatNumber(summaryTotal.buyBust)}</th>
              <th>${formatNumber(summaryTotal.entrapment)}</th>
              <th>${formatNumber(summaryTotal.totalArrested)}</th>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function downloadReportPdf() {
  const filters = readReportFiltersFromDom();
  state.reportFilters = { ...filters };

  const filtered = getFilteredReportRecords(filters);
  const { rows: summaryRows, total: summaryTotal } = buildSummaryByOplan(filtered);

  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert(
      "PDF library not loaded. Please make sure you are connected to the internet and try again."
    );
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a3" });

  const title = "Daily Accomplishment Report (Oplan)";
  doc.setFontSize(14);
  doc.text(title, 36, 28);

  doc.setFontSize(10);
  const rangeLabel =
    filters.fromDate || filters.toDate
      ? `From ${filters.fromDate || "–"} to ${filters.toDate || "–"}`
      : "All available records";
  doc.text(rangeLabel, 36, 46);

  const appliedFilters = getAppliedReportFilterItems(filters);
  doc.text(
    appliedFilters.length
      ? `Filters: ${appliedFilters.join(" | ")}`
      : "Filters: None (All records)",
    36,
    62
  );

  const totalOperations = filtered.length;
  const totalArrests = filtered.reduce(
    (sum, r) => sum + (Number(r.numArrests) || 0),
    0
  );

  doc.text(
    `Total Operations: ${formatNumber(
      totalOperations
    )} | Total Arrests: ${formatNumber(totalArrests)}`,
    36,
    78
  );

  const summaryBody = summaryRows.map((row) => [
    `OPLAN ${row.oplanName.toUpperCase()}`,
    String(row.operations),
    String(row.searchWarrant),
    String(row.buyBust),
    String(row.entrapment),
    String(row.totalArrested),
  ]);
  summaryBody.push([
    "TOTAL",
    String(summaryTotal.operations),
    String(summaryTotal.searchWarrant),
    String(summaryTotal.buyBust),
    String(summaryTotal.entrapment),
    String(summaryTotal.totalArrested),
  ]);

  doc.autoTable({
    startY: 92,
    margin: { left: 36, right: 36 },
    head: [
      [
        "CIDG Flagship Projects",
        "No. of Operation Conducted",
        "Search Warrant",
        "Buy Bust",
        "Entrapment",
        "Total No. of Arrested",
      ],
    ],
    body: summaryBody,
    styles: { fontSize: 9, halign: "center", valign: "middle" },
    columnStyles: { 0: { halign: "left" } },
    headStyles: {
      fillColor: [31, 79, 143],
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
  });

  doc.save("oplan_daily_accomplishment_report.pdf");
}

// Navigation
function setupNavigation() {
  document.querySelectorAll(".nav-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      setActiveNav(btn);
      const view = btn.getAttribute("data-view");
      if (view === "dashboard") {
        renderDashboard();
      } else if (view === "reports") {
        renderReportsView();
      } else if (view === "oplan") {
        const oplanId = btn.getAttribute("data-oplan-id");
        renderOplanForm(oplanId);
      }
    });
  });
}

function initApp() {
  initState();
  setupNavigation();
  updateTodayLabel();
  renderDashboard();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initApp);
} else {
  initApp();
}

