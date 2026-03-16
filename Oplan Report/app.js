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
    label: "Time / Time Range",
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
    label: "PFUs / CFUs / DFUs",
    type: "select",
    options: ["CAVITE", "LAGUNA", "BATANGAS", "RIZAL", "QUEZON"],
    required: false,
  },
  {
    id: "modeOfOperation",
    label: "Mode of Operation",
    type: "select",
    options: ["Search Warrant", "Buy Bust"],
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
    type: "number",
    min: 0,
    step: 1,
    required: false,
  },
  {
    id: "smallArms",
    label: "Small Arms",
    type: "number",
    min: 0,
    step: 1,
    required: false,
  },
  {
    id: "bigArms",
    label: "Big Arms",
    type: "number",
    min: 0,
    step: 1,
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
    type: "number",
    min: 0,
    step: "0.01",
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

const state = {
  currentView: "dashboard",
  currentOplanId: null,
  editingId: null,
  searchQuery: "",
  recordsByOplan: {}, // { [oplanId]: Record[] }
  dashboardDateFilter: { month: "", day: "", year: "" },
  oplanDateFilters: {}, // { [oplanId]: { month, day, year } }
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

function passesDateFilter(dateStr, filter) {
  if (!dateStr) return false;
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return false;
  const month = d.getMonth() + 1;
  const day = d.getDate();
  const year = d.getFullYear();

  if (filter.year && String(year) !== String(filter.year)) return false;
  if (filter.month && String(month) !== String(filter.month)) return false;
  if (filter.day && String(day) !== String(filter.day)) return false;
  return true;
}

function updateTodayLabel() {
  const el = document.getElementById("today-label");
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleString(undefined, {
    weekday: "short",
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
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

  const container = document.getElementById("view-container");

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
        <td>${o.name}</td>
        <td>${formatNumber(agg.operations)}</td>
        <td>${formatNumber(agg.arrests)}</td>
        <td>${formatNumber(agg.count)}</td>
      </tr>
    `;
  }).join("");

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
          <div class="filter-group">
            <label>&nbsp;</label>
            <button type="button" id="dashboard-apply-filter" class="btn secondary sm">Enter</button>
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
              <th>Oplan</th>
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

    <div class="card">
      <div class="form-header">
        <div>
          <div class="report-section-title">Quick Access</div>
          <p class="small">Open the dedicated data entry form for each Oplan.</p>
        </div>
      </div>
      <div class="cards-row">
        ${OPLANS.map(
          (o) => `
          <div class="card">
            <div class="card-title">${o.name}</div>
            <button class="btn secondary sm" data-quick-oplan-id="${o.id}">
              Open ${o.name} Form
            </button>
          </div>
        `
        ).join("")}
      </div>
    </div>
  `;

  const monthInput = document.getElementById("dashboard-month");
  const dayInput = document.getElementById("dashboard-day");
  const yearInput = document.getElementById("dashboard-year");
  if (monthInput) monthInput.value = filter.month || "";
  if (dayInput) dayInput.value = filter.day || "";
  if (yearInput) yearInput.value = filter.year || "";

  const applyBtn = document.getElementById("dashboard-apply-filter");
  if (applyBtn) {
    applyBtn.addEventListener("click", () => {
      state.dashboardDateFilter = {
        month: monthInput.value.trim(),
        day: dayInput.value.trim(),
        year: yearInput.value.trim(),
      };
      renderDashboard();
    });
  }

  container
    .querySelectorAll("[data-quick-oplan-id]")
    .forEach((btn) =>
      btn.addEventListener("click", () =>
        renderOplanForm(btn.getAttribute("data-quick-oplan-id"))
      )
    );
}

// Oplan form rendering
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

  const container = document.getElementById("view-container");

  const fieldsHtml = FIELDS.map((field) => {
    const requiredMark = field.required ? '<span class="required">*</span>' : "";
    const base = `
      <div class="form-field" data-field-id="${field.id}">
        <label for="${field.id}">
          ${field.label}
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
  }).join("");

  container.innerHTML = `
    <div class="form-header">
      <div>
        <h2>${oplan.name} Daily Accomplishment</h2>
        <p class="muted">
          Fill in the operational fields below. Each saved record corresponds to one operation entry under this Oplan.
        </p>
      </div>
      <div class="form-actions">
        <button type="button" id="save-record-btn" class="btn primary">Save Record</button>
      </div>
    </div>

    <form id="record-form" class="record-form" novalidate>
      ${fieldsHtml}
    </form>

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
        <div class="filter-group">
          <label>&nbsp;</label>
          <button type="button" id="oplan-apply-filter" class="btn secondary sm">Enter</button>
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
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>No. of Operations</th>
            <th>No. Arrested</th>
            <th>Names of Arrested Persons</th>
            <th>PFU/CFU/DFU</th>
            <th>Mode of Operation</th>
            <th>Place (Region/Province/City)</th>
            <th>Tagged/Listed As</th>
            <th>Firearms Confiscated</th>
            <th>Small Arms</th>
            <th>Big Arms</th>
            <th>Confiscated Items</th>
            <th>Amount of Confiscated</th>
            <th>Report Upload</th>
            <th>Remarks</th>
            <th style="width: 120px;">Actions</th>
          </tr>
        </thead>
        <tbody id="records-tbody"></tbody>
      </table>
    </div>
  `;

  document
    .getElementById("save-record-btn")
    .addEventListener("click", handleSaveRecord);
  const numArrestsInput = document.getElementById("numArrests");
  if (numArrestsInput) {
    numArrestsInput.addEventListener("input", handleNumArrestsChange);
  }

  document
    .getElementById("search-input")
    .addEventListener("input", (e) => {
      state.searchQuery = e.target.value || "";
      renderRecordsTable();
    });

  const applyBtn = document.getElementById("oplan-apply-filter");
  if (applyBtn) {
    applyBtn.addEventListener("click", () => {
      state.oplanDateFilters[oplanId] = {
        month: document.getElementById("oplan-month").value.trim(),
        day: document.getElementById("oplan-day").value.trim(),
        year: document.getElementById("oplan-year").value.trim(),
      };
      renderRecordsTable();
    });
  }

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
    result[field.id] = el.value != null ? el.value.trim() : "";
  }

  const arrestedContainer = document.getElementById(
    "arrested-persons-container"
  );
  const arrestedPersons = [];
  if (arrestedContainer) {
    arrestedContainer.querySelectorAll(".arrested-person-field").forEach((row) => {
      const lastName = (row.querySelector('input[data-arrested-part="lastName"]')?.value || "").trim();
      const firstName = (row.querySelector('input[data-arrested-part="firstName"]')?.value || "").trim();
      const middleName = (row.querySelector('input[data-arrested-part="middleName"]')?.value || "").trim();
      const suffix = (row.querySelector('input[data-arrested-part="suffix"]')?.value || "").trim();
      arrestedPersons.push({ lastName, firstName, middleName, suffix });
    });
  }
  result.arrestedPersons = arrestedPersons;

  return result;
}

function formatArrestedPersonName(p) {
  if (!p) return "";
  // Backward compatibility: if older data stored a string, return it.
  if (typeof p === "string") return p;
  const last = (p.lastName || "").trim();
  const first = (p.firstName || "").trim();
  const middle = (p.middleName || "").trim();
  const suffix = (p.suffix || "").trim();
  const middleInitial = middle ? ` ${middle}` : "";
  const suffixPart = suffix ? ` ${suffix}` : "";
  if (!last && !first && !middle && !suffix) return "";
  if (last && first) return `${last}, ${first}${middleInitial}${suffixPart}`.trim();
  return `${[last, first, middle, suffix].filter(Boolean).join(" ")}`.trim();
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
  if (!state.currentOplanId) return;
  const values = readFormValues();
  if (!values) return;

  const isValid = validateForm(values);
  if (!isValid) {
    alert("Please fill in all required fields.");
    return;
  }

  const oplanId = state.currentOplanId;
  const records = state.recordsByOplan[oplanId] || [];

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
        numOperations: 1,
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
      numOperations: 1,
      reportAttachmentId: reportAttachment ? reportAttachment.id : "",
      reportAttachmentName: reportAttachment ? reportAttachment.name : "",
      reportAttachmentType: reportAttachment ? reportAttachment.type : "",
    };
    records.push(newRecord);
  }

  state.recordsByOplan[oplanId] = records;
  saveToStorage();
  resetForm();
  renderRecordsTable();
  // Update dashboard counters in case the user navigates back
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

  if (!filtered.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="16" class="table-empty">
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
        <td>${formatDate(r.date)}</td>
        <td>1</td>
        <td>${formatNumber(r.numArrests)}</td>
        <td>${
          (r.arrestedPersons || [])
            .map(formatArrestedPersonName)
            .filter(Boolean)
            .join("; ") || ""
        }</td>
        <td>${r.pfuCfuDfu || ""}</td>
        <td>${r.modeOfOperation || ""}</td>
        <td>${[r.placeRegion, r.placeProvince, r.placeMunicipalityCity]
          .filter(Boolean)
          .join(" / ")}</td>
        <td>${r.taggedListedAs || ""}</td>
        <td>${formatNumber(r.numFirearmsConfiscated)}</td>
        <td>${formatNumber(r.smallArms)}</td>
        <td>${formatNumber(r.bigArms)}</td>
        <td>${r.confiscatedItems || ""}</td>
        <td>${formatNumber(r.amountOfConfiscated)}</td>
        <td>${
          r.reportAttachmentName
            ? `<button type="button" class="btn secondary sm" data-action="view-attachment">${r.reportAttachmentName}</button>`
            : ""
        }</td>
        <td>${r.remarks || ""}</td>
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
    el.value = rec[field.id] != null ? rec[field.id] : "";
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
  input.value = value || "";
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
function renderReportsView() {
  state.currentView = "reports";
  state.currentOplanId = null;
  state.editingId = null;

  setViewTitle(
    "Reports",
    "Generate summaries of operations for a selected date range across all Oplans."
  );

  const container = document.getElementById("view-container");

  container.innerHTML = `
    <div class="card report-layout">
      <div class="form-header">
        <div>
          <div class="report-section-title">Date and Time Filter</div>
          <p class="small">
            Select a date range, then click <strong>View Report</strong> to preview all records within that period.
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
}

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

function generateReport() {
  const fromDateStr = document.getElementById("from-date").value;
  const toDateStr = document.getElementById("to-date").value;
  const resultsEl = document.getElementById("report-results");

  const filtered = filterRecordsByDateRange(fromDateStr, toDateStr);

  const totalOperations = filtered.length;
  const totalArrests = filtered.reduce(
    (sum, r) => sum + (Number(r.numArrests) || 0),
    0
  );

  const perOplan = {};
  for (const o of OPLANS) {
    perOplan[o.id] = { operations: 0, arrests: 0, count: 0 };
  }
  for (const r of filtered) {
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
        <td>${o.name}</td>
        <td>${formatNumber(agg.operations)}</td>
        <td>${formatNumber(agg.arrests)}</td>
        <td>${formatNumber(agg.count)}</td>
      </tr>
    `;
  }).join("");

  const detailedRows = filtered
    .map((r) => {
      const oplan = OPLANS.find((o) => o.id === r.oplanId);
      return `
        <tr>
          <td>${oplan ? oplan.name : r.oplanId}</td>
          <td>${formatDate(r.date)}</td>
          <td>${r.time || ""}</td>
          <td>1</td>
          <td>${formatNumber(r.numArrests)}</td>
          <td>${
            (r.arrestedPersons || [])
              .map(formatArrestedPersonName)
              .filter(Boolean)
              .join("; ") || ""
          }</td>
          <td>${r.pfuCfuDfu || ""}</td>
          <td>${r.modeOfOperation || ""}</td>
          <td>${[r.placeRegion, r.placeProvince, r.placeMunicipalityCity]
            .filter(Boolean)
            .join(" / ")}</td>
          <td>${r.taggedListedAs || ""}</td>
          <td>${formatNumber(r.numFirearmsConfiscated)}</td>
          <td>${formatNumber(r.smallArms)}</td>
          <td>${formatNumber(r.bigArms)}</td>
          <td>${r.confiscatedItems || ""}</td>
          <td>${formatNumber(r.amountOfConfiscated)}</td>
          <td>${r.reportAttachmentName || ""}</td>
          <td>${r.remarks || ""}</td>
        </tr>
      `;
    })
    .join("");

  const rangeLabel =
    fromDateStr || toDateStr
      ? `From <strong>${fromDateStr || "–"}</strong> to <strong>${
          toDateStr || "–"
        }</strong>`
      : "All available records";

  resultsEl.innerHTML = `
    <div class="report-summary-row">
      <div class="card">
        <div class="card-title">Total Operations</div>
        <div class="card-value">${formatNumber(totalOperations)}</div>
        <div class="card-subvalue">All records in selected date range.</div>
      </div>
      <div class="card">
        <div class="card-title">Total Arrested Individuals</div>
        <div class="card-value">${formatNumber(totalArrests)}</div>
        <div class="card-subvalue">All Oplans combined.</div>
      </div>
      <div class="card">
        <div class="card-title">Total Records in Range</div>
        <div class="card-value">${formatNumber(filtered.length)}</div>
        <div class="card-subvalue">${rangeLabel}</div>
      </div>
    </div>

    <div class="card">
      <div class="report-section-title">Summary per Oplan</div>
      <div class="card-table">
        <table>
          <thead>
            <tr>
              <th>Oplan</th>
              <th>Total Operations</th>
              <th>Total Arrests</th>
              <th>Records</th>
            </tr>
          </thead>
          <tbody>
            ${
              filtered.length === 0
                ? `<tr><td colspan="4" class="table-empty">No records found in the selected date range.</td></tr>`
                : perOplanRows
            }
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div class="report-section-title">Detailed Records</div>
      <div class="card-table">
        <table id="report-detailed-table">
          <thead>
            <tr>
              <th>Oplan</th>
              <th>Date</th>
              <th>Time</th>
              <th>No. of Operations</th>
              <th>No. Arrested</th>
              <th>Names of Arrested Persons</th>
              <th>PFU/CFU/DFU</th>
              <th>Mode of Operation</th>
              <th>Place (Region/Province/City)</th>
              <th>Tagged/Listed As</th>
              <th>Firearms Confiscated</th>
              <th>Small Arms</th>
              <th>Big Arms</th>
              <th>Confiscated Items</th>
              <th>Amount of Confiscated</th>
              <th>Report Upload</th>
              <th>Remarks</th>
            </tr>
          </thead>
          <tbody>
            ${
              filtered.length === 0
                ? `<tr><td colspan="18" class="table-empty">No records to display.</td></tr>`
                : detailedRows
            }
          </tbody>
        </table>
      </div>
      <p class="small" style="margin-top:8px;">
        Note: Uploaded report files are listed by filename here. You can download attachments from their originating Oplan record.
      </p>
    </div>
  `;
}

function downloadReportPdf() {
  const fromDateStr = document.getElementById("from-date").value;
  const toDateStr = document.getElementById("to-date").value;
  const filtered = filterRecordsByDateRange(fromDateStr, toDateStr);

  if (!filtered.length) {
    alert("No records in the selected date range.");
    return;
  }

  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert(
      "PDF library not loaded. Please make sure you are connected to the internet and try again."
    );
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape" });

  const title = "Daily Accomplishment Report (Oplan)";
  doc.setFontSize(14);
  doc.text(title, 14, 14);

  doc.setFontSize(10);
  const rangeLabel =
    fromDateStr || toDateStr
      ? `From ${fromDateStr || "–"} to ${toDateStr || "–"}`
      : "All available records";
  doc.text(rangeLabel, 14, 20);

  const totalOperations = filtered.reduce(
    (sum) => sum + 1,
    0
  );
  const totalArrests = filtered.reduce(
    (sum, r) => sum + (Number(r.numArrests) || 0),
    0
  );

  doc.text(
    `Total Operations: ${formatNumber(
      totalOperations
    )} | Total Arrests: ${formatNumber(totalArrests)}`,
    14,
    26
  );

  const body = filtered.map((r) => {
    const oplan = OPLANS.find((o) => o.id === r.oplanId);
    return [
      oplan ? oplan.name : r.oplanId,
      formatDate(r.date),
      r.time || "",
      "1",
      formatNumber(r.numArrests),
      (r.arrestedPersons || [])
        .map(formatArrestedPersonName)
        .filter(Boolean)
        .join("; ") || "",
      r.pfuCfuDfu || "",
      r.modeOfOperation || "",
      [r.placeRegion, r.placeProvince, r.placeMunicipalityCity]
        .filter(Boolean)
        .join(" / "),
      r.taggedListedAs || "",
      String(Number(r.numFirearmsConfiscated) || 0),
      String(Number(r.smallArms) || 0),
      String(Number(r.bigArms) || 0),
      r.confiscatedItems || "",
      String(Number(r.amountOfConfiscated) || 0),
      r.reportAttachmentName || "",
      r.remarks || "",
    ];
  });

  doc.autoTable({
    startY: 32,
    head: [
      [
        "Oplan",
        "Date",
        "Time",
        "No. of Operations",
        "No. Arrested",
        "Names of Arrested Persons",
        "PFU/CFU/DFU",
        "Mode of Operation",
        "Place (Region/Province/City)",
        "Tagged/Listed As",
        "Firearms Confiscated",
        "Small Arms",
        "Big Arms",
        "Confiscated Items",
        "Amount of Confiscated",
        "Report Upload",
        "Remarks",
      ],
    ],
    body,
    styles: {
      fontSize: 8,
      cellPadding: 2,
    },
    headStyles: {
      fillColor: [31, 79, 143],
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

