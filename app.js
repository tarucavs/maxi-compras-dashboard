const FALLBACK_FILE_NAME = "SEGUIMIENTOS COMPRAS FABRICACIONES MILITARES (10).xlsx";
const GOOGLE_SHEETS_EDIT_URL = "https://docs.google.com/spreadsheets/d/1kFaKvilRf1z97XC15fAMYgS9l0cPq4bC-UF7VRVR9_U/edit?gid=0#gid=0";
const GOOGLE_REFRESH_MS = 30 * 1000;

const STORAGE_KEYS = {
  statuses: "compras_dashboard_status_overrides_v2",
  kpis: "compras_dashboard_selected_kpis_v2",
  columns: "compras_dashboard_selected_columns_v1",
  filters: "compras_dashboard_filters_v1",
};

const STATUS_BASE_OPTIONS = ["Adjudicado", "En licitacion", "Con OC", "Pendiente", "Sin estado"];

const KPI_DEFS = [
  { key: "procesos", label: "Procesos", calc: ({ total }) => formatInt.format(total) },
  { key: "monto_total", label: "Monto total", calc: ({ totalEstimado }) => formatMoney.format(totalEstimado) },
  { key: "promedio", label: "Promedio por proceso", calc: ({ ticket }) => formatMoney.format(ticket) },
  { key: "adjudicados", label: "% adjudicados", calc: ({ pctAdjudicado }) => pctAdjudicado.toFixed(1) + "%" },
  { key: "con_oc", label: "% con OC", calc: ({ pctOc }) => pctOc.toFixed(1) + "%" },
  { key: "lead", label: "Dias promedio", calc: ({ leadAvg }) => leadAvg.toFixed(1) + " dias" },
];

const TABLE_COLUMNS = [
  { key: "solped", label: "Solped", value: (r) => r.solped },
  { key: "objeto", label: "Objeto", value: (r) => r.objeto },
  { key: "proceso", label: "Proceso", value: (r) => r.proceso },
  { key: "expediente", label: "Expediente", value: (r) => r.expediente },
  { key: "comprador", label: "Comprador", value: (r) => r.comprador },
  { key: "fabrica", label: "Fabrica", value: (r) => r.fabrica },
  { key: "estado", label: "Estado (editable)", editableStatus: true },
  { key: "estimado", label: "Estimado sin IVA", value: (r) => formatMoney.format(r.estimado || 0) },
  { key: "oc", label: "OC", value: (r) => r.oc || "-" },
  { key: "observaciones", label: "Observaciones", value: (r) => r.observaciones || "-" },
];

const DEFAULT_COLUMNS = ["solped", "objeto", "proceso", "comprador", "fabrica", "estado", "estimado", "oc"];

const state = {
  rawRows: [],
  filteredRows: [],
  selectedKpis: new Set(loadStoredSelection(STORAGE_KEYS.kpis, KPI_DEFS.map((k) => k.key))),
  selectedColumns: new Set(loadStoredSelection(STORAGE_KEYS.columns, DEFAULT_COLUMNS)),
  statusOverrides: loadStoredStatuses(),
  selectedFilters: loadStoredFilters(),
  charts: {
    status: null,
    amount: null,
    timeline: null,
  },
};

const dom = {
  fileInput: document.getElementById("excelFile"),
  refreshOnlineBtn: document.getElementById("refreshOnlineBtn"),
  loadDefaultBtn: document.getElementById("loadDefaultBtn"),
  dataSourceInfo: document.getElementById("dataSourceInfo"),
  loadHint: document.getElementById("loadHint"),
  dashboard: document.getElementById("dashboard"),
  emptyState: document.getElementById("emptyState"),
  kpiGrid: document.getElementById("kpiGrid"),
  kpiOptions: document.getElementById("kpiOptions"),
  columnOptions: document.getElementById("columnOptions"),
  timelineRange: document.getElementById("timelineRange"),
  chartTopN: document.getElementById("chartTopN"),
  searchInput: document.getElementById("searchInput"),
  buyerFilterOptions: document.getElementById("buyerFilterOptions"),
  factoryFilterOptions: document.getElementById("factoryFilterOptions"),
  statusFilterOptions: document.getElementById("statusFilterOptions"),
  buyerFilterCount: document.getElementById("buyerFilterCount"),
  factoryFilterCount: document.getElementById("factoryFilterCount"),
  statusFilterCount: document.getElementById("statusFilterCount"),
  clearFiltersBtn: document.getElementById("clearFiltersBtn"),
  rowsInfo: document.getElementById("rowsInfo"),
  tableHead: document.getElementById("tableHead"),
  tableBody: document.getElementById("tableBody"),
};

const formatInt = new Intl.NumberFormat("es-AR");
const formatMoney = new Intl.NumberFormat("es-AR", {
  style: "currency",
  currency: "ARS",
  maximumFractionDigits: 0,
});

init();

function init() {
  dom.fileInput.addEventListener("change", onFileChange);
  dom.refreshOnlineBtn.addEventListener("click", () => loadPreferredSource({ forceOnline: true }));
  dom.loadDefaultBtn.addEventListener("click", loadDefaultFromFolder);
  dom.clearFiltersBtn.addEventListener("click", clearAllFilters);

  [dom.searchInput].forEach((el) => {
    el.addEventListener("input", applyFilters);
  });

  [dom.timelineRange, dom.chartTopN].forEach((el) => {
    el.addEventListener("change", () => renderCharts(state.filteredRows));
  });

  dom.tableBody.addEventListener("change", handleStatusChange);

  renderKpiSelector();
  renderColumnSelector();

  loadPreferredSource();
  setInterval(() => {
    loadFromGoogleSheet({ silent: true });
  }, GOOGLE_REFRESH_MS);
}

async function loadPreferredSource(options = {}) {
  const { forceOnline = false } = options;

  if (forceOnline) {
    const ok = await loadFromGoogleSheet();
    if (!ok) {
      dom.loadHint.textContent =
        "Google no disponible. Revisa permisos de la hoja o usa archivo local/manual.";
    }
    return;
  }

  const onlineOk = await loadFromGoogleSheet({ silent: true });
  if (!onlineOk) {
    await loadDefaultFromFolder();
  }
}

async function loadFromGoogleSheet(options = {}) {
  const { silent = false } = options;
  const sheet = parseGoogleSheetRef(GOOGLE_SHEETS_EDIT_URL);

  if (!sheet) {
    if (!silent) {
      dom.loadHint.textContent = "Enlace de Google Sheets invalido.";
    }
    setSourceInfo({ connected: false, sourceName: "Enlace invalido", atDate: new Date() });
    return false;
  }

  if (!silent) {
    dom.loadHint.textContent = "Actualizando desde Google Sheets...";
  }

  try {
    const jsonpRows = await loadFromGoogleGvizJsonp(sheet);
    if (Array.isArray(jsonpRows) && jsonpRows.length) {
      ingestRows(jsonpRows);
      dom.loadHint.textContent = "Datos actualizados en vivo desde Google Sheets.";
      setSourceInfo({ connected: true, sourceName: "Google Sheets (tiempo real)", atDate: new Date() });
      return true;
    }
  } catch (_error) {
  }

  const urls = [
    {
      name: "Google CSV (gviz)",
      url: "https://docs.google.com/spreadsheets/d/" + sheet.id + "/gviz/tq?tqx=out:csv&gid=" + sheet.gid,
      type: "csv",
    },
    {
      name: "Google CSV (export)",
      url: "https://docs.google.com/spreadsheets/d/" + sheet.id + "/export?format=csv&gid=" + sheet.gid,
      type: "csv",
    },
    {
      name: "Google XLSX",
      url: "https://docs.google.com/spreadsheets/d/" + sheet.id + "/export?format=xlsx&gid=" + sheet.gid,
      type: "xlsx",
    },
  ];

  for (const candidate of urls) {
    try {
      const response = await fetch(candidate.url + "&ts=" + Date.now());
      if (!response.ok) {
        continue;
      }

      if (candidate.type === "csv") {
        const csvText = await response.text();
        parseCsv(csvText);
      } else {
        const arrayBuffer = await response.arrayBuffer();
        parseWorkbook(arrayBuffer);
      }

      dom.loadHint.textContent = "Datos actualizados desde Google Sheets.";
      setSourceInfo({ connected: true, sourceName: candidate.name, atDate: new Date() });
      return true;
    } catch (_error) {
    }
  }

  setSourceInfo({ connected: false, sourceName: "Respaldo local", atDate: new Date() });
  if (!silent) {
    dom.loadHint.textContent =
      "Google no disponible. En Google Sheets habilita: Compartir > Cualquier persona con el enlace (Lector).";
  }
  return false;
}

function loadFromGoogleGvizJsonp(sheet) {
  return new Promise((resolve, reject) => {
    const callbackName = "__gs_jsonp_cb_" + Date.now() + "_" + Math.floor(Math.random() * 10000);
    const timeoutMs = 12000;
    let timeoutId = null;

    const cleanup = (scriptNode) => {
      if (timeoutId) {
        clearTimeout(timeoutId);
      }
      try {
        delete window[callbackName];
      } catch (_error) {
        window[callbackName] = undefined;
      }
      if (scriptNode && scriptNode.parentNode) {
        scriptNode.parentNode.removeChild(scriptNode);
      }
    };

    const script = document.createElement("script");
    const queryUrl =
      "https://docs.google.com/spreadsheets/d/" +
      sheet.id +
      "/gviz/tq?gid=" +
      sheet.gid +
      "&headers=1&tqx=out:json;responseHandler:" +
      callbackName +
      "&ts=" +
      Date.now();

    window[callbackName] = (response) => {
      cleanup(script);
      try {
        const rows = gvizResponseToRows(response);
        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };

    script.onerror = () => {
      cleanup(script);
      reject(new Error("Fallo JSONP"));
    };

    timeoutId = setTimeout(() => {
      cleanup(script);
      reject(new Error("Timeout JSONP"));
    }, timeoutMs);

    script.src = queryUrl;
    document.head.appendChild(script);
  });
}

function gvizResponseToRows(response) {
  const table = response && response.table;
  if (!table || !Array.isArray(table.cols) || !Array.isArray(table.rows)) {
    return [];
  }

  const headers = table.cols.map((col, idx) => {
    const label = cleanText(col && col.label);
    const id = cleanText(col && col.id);
    return label || id || "Unnamed: " + idx;
  });

  return table.rows.map((row) => {
    const out = {};
    headers.forEach((header, idx) => {
      const cell = row && row.c ? row.c[idx] : null;
      const value = cell ? (cell.f ?? cell.v) : null;
      out[header] = value;
    });
    return out;
  });
}

async function loadDefaultFromFolder() {
  dom.loadHint.textContent = "Intentando cargar archivo de la carpeta...";
  try {
    const response = await fetch(encodeURIComponent(FALLBACK_FILE_NAME));
    if (!response.ok) {
      throw new Error("local not found");
    }
    const arrayBuffer = await response.arrayBuffer();
    parseWorkbook(arrayBuffer);
    dom.loadHint.textContent = "Archivo local cargado correctamente.";
    setSourceInfo({ connected: false, sourceName: "Archivo local", atDate: new Date() });
  } catch (_error) {
    dom.loadHint.textContent =
      "No se pudo cargar archivo local. Si abriste con file://, usa servidor local (python -m http.server).";
  }
}

function onFileChange(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    parseWorkbook(e.target.result);
    dom.loadHint.textContent = "Archivo manual cargado: " + file.name;
    setSourceInfo({ connected: false, sourceName: "Archivo manual", atDate: new Date() });
  };
  reader.readAsArrayBuffer(file);
}

function parseGoogleSheetRef(url) {
  try {
    const parsed = new URL(url);
    const parts = parsed.pathname.split("/").filter(Boolean);
    const dPos = parts.indexOf("d");
    const id = dPos >= 0 ? parts[dPos + 1] : null;
    const gid = parsed.searchParams.get("gid") || "0";
    if (!id) return null;
    return { id, gid: encodeURIComponent(gid) };
  } catch (_error) {
    return null;
  }
}

function setSourceInfo({ connected, sourceName, atDate }) {
  const hh = String(atDate.getHours()).padStart(2, "0");
  const mm = String(atDate.getMinutes()).padStart(2, "0");
  dom.dataSourceInfo.classList.remove("status-ok", "status-fallback");
  dom.dataSourceInfo.classList.add(connected ? "status-ok" : "status-fallback");
  const mode = connected ? "Conectado a Google Sheets" : "Modo respaldo local";
  dom.dataSourceInfo.textContent = mode + " | Fuente: " + sourceName + " | " + hh + ":" + mm;
}

function parseCsv(csvText) {
  const workbook = XLSX.read(csvText, { type: "string" });
  const firstSheet = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheet];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: null, raw: false });
  ingestRows(rows);
}

function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const firstSheet = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheet];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: null, raw: false });
  ingestRows(rows);
}

function ingestRows(rows) {
  state.rawRows = rows
    .map((row) => normalizeRow(row))
    .filter((row) => row.solped || row.objeto || row.proceso);

  state.rawRows.forEach((row) => {
    const override = state.statusOverrides[row.id];
    if (override) {
      row.estado = override;
    }
  });

  populateFilters(state.rawRows);
  applyFilters();

  dom.emptyState.classList.add("hidden");
  dom.dashboard.classList.remove("hidden");
}

function normalizeRow(row) {
  const estimated = parseNumber(row["ESTIMADO SIN IVA"]);
  const statusRaw = cleanText(row["Unnamed: 11"]) || cleanText(row["OBSERVACIONES"]);
  const solped = cleanText(row.SOLPED);
  const proceso = cleanText(row["Nº PROCESO"]);
  const expediente = cleanText(row["Nº EXPEDIENTE"]);

  return {
    id: buildRowId(solped, proceso, expediente),
    solped,
    objeto: cleanText(row.OBJETO),
    proceso,
    expediente,
    estimado: estimated,
    observaciones: cleanText(row.OBSERVACIONES),
    apertura: parseDate(row.APERTURA),
    desdeFecha: parseDate(row["DESDE (FECHA)"]),
    comprador: cleanText(row.COMPRADOR) || "Sin comprador",
    fabrica: cleanText(row["FÁBRICA"]) || "Sin fabrica",
    estado: normalizeStatus(statusRaw),
    oc: cleanText(row.OC),
  };
}

function applyFilters() {
  const search = cleanText(dom.searchInput.value).toLowerCase();

  const buyerSet = new Set(state.selectedFilters.buyers || []);
  const factorySet = new Set(state.selectedFilters.factories || []);
  const statusSet = new Set(state.selectedFilters.statuses || []);

  state.filteredRows = state.rawRows.filter((r) => {
    const inSearch =
      !search || [r.objeto, r.proceso, r.expediente, r.solped].join(" ").toLowerCase().includes(search);
    const inBuyer = !buyerSet.size || buyerSet.has(r.comprador);
    const inFactory = !factorySet.size || factorySet.has(r.fabrica);
    const inStatus = !statusSet.size || statusSet.has(r.estado);
    return inSearch && inBuyer && inFactory && inStatus;
  });

  renderKpis(state.filteredRows);
  renderTable(state.filteredRows);
  renderCharts(state.filteredRows);
}

function populateFilters(rows) {
  renderFilterGroup({
    container: dom.buyerFilterOptions,
    countBadge: dom.buyerFilterCount,
    type: "buyers",
    values: unique(rows.map((r) => r.comprador)),
    emptyText: "Todos",
  });

  renderFilterGroup({
    container: dom.factoryFilterOptions,
    countBadge: dom.factoryFilterCount,
    type: "factories",
    values: unique(rows.map((r) => r.fabrica)),
    emptyText: "Todas",
  });

  renderFilterGroup({
    container: dom.statusFilterOptions,
    countBadge: dom.statusFilterCount,
    type: "statuses",
    values: unique(rows.map((r) => r.estado)),
    emptyText: "Todos",
  });
}

function renderFilterGroup({ container, countBadge, type, values, emptyText }) {
  const selectedRaw = new Set(state.selectedFilters[type] || []);
  const selected = new Set([...selectedRaw].filter((value) => values.includes(value)));

  if (selected.size !== selectedRaw.size) {
    state.selectedFilters[type] = [...selected];
    storeFilters(state.selectedFilters);
  }

  container.innerHTML =
    '<div class="filter-tools">' +
    '<input type="text" class="filter-search" placeholder="Buscar..." data-filter-search="' +
    escapeHtml(type) +
    '" />' +
    '<div class="filter-tools-actions">' +
    '<button type="button" class="tiny-btn" data-filter-action="all" data-filter-type="' +
    escapeHtml(type) +
    '">Todos</button>' +
    '<button type="button" class="tiny-btn" data-filter-action="none" data-filter-type="' +
    escapeHtml(type) +
    '">Ninguno</button>' +
    "</div>" +
    "</div>" +
    '<div class="filter-options-list">' +
    values
      .map((value) => {
        const checked = selected.has(value) ? " checked" : "";
        return (
          '<label class="filter-item" data-filter-label="' +
          escapeHtml(value.toLowerCase()) +
          '">' +
          '<input type="checkbox" data-filter-type="' +
          escapeHtml(type) +
          '" data-filter-value="' +
          escapeHtml(value) +
          '"' +
          checked +
          " />" +
          "<span>" +
          escapeHtml(value) +
          "</span></label>"
        );
      })
      .join("") +
    "</div>";

  const listRoot = container.querySelector(".filter-options-list");

  container.querySelectorAll("input[type='checkbox']").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const filterType = checkbox.dataset.filterType;
      const filterValue = checkbox.dataset.filterValue;

      const target = new Set(state.selectedFilters[filterType] || []);
      if (checkbox.checked) target.add(filterValue);
      else target.delete(filterValue);

      state.selectedFilters[filterType] = [...target];
      storeFilters(state.selectedFilters);
      updateFilterCounts();
      applyFilters();
    });
  });

  const searchInput = container.querySelector(".filter-search");
  if (searchInput) {
    searchInput.addEventListener("input", () => {
      const query = cleanText(searchInput.value).toLowerCase();
      listRoot.querySelectorAll(".filter-item").forEach((item) => {
        const normalized = item.dataset.filterLabel || "";
        item.style.display = !query || normalized.includes(query) ? "flex" : "none";
      });
    });
  }

  container.querySelectorAll("button[data-filter-action]").forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.filterAction;
      if (action === "all") {
        state.selectedFilters[type] = [...values];
      } else {
        state.selectedFilters[type] = [];
      }
      storeFilters(state.selectedFilters);
      populateFilters(state.rawRows);
      applyFilters();
    });
  });

  const count = selected.size;
  countBadge.textContent = count ? String(count) : emptyText;
}

function updateFilterCounts() {
  dom.buyerFilterCount.textContent =
    state.selectedFilters.buyers && state.selectedFilters.buyers.length
      ? String(state.selectedFilters.buyers.length)
      : "Todos";
  dom.factoryFilterCount.textContent =
    state.selectedFilters.factories && state.selectedFilters.factories.length
      ? String(state.selectedFilters.factories.length)
      : "Todas";
  dom.statusFilterCount.textContent =
    state.selectedFilters.statuses && state.selectedFilters.statuses.length
      ? String(state.selectedFilters.statuses.length)
      : "Todos";
}

function clearAllFilters() {
  state.selectedFilters = { buyers: [], factories: [], statuses: [] };
  storeFilters(state.selectedFilters);
  populateFilters(state.rawRows);
  applyFilters();
}

function renderKpis(rows) {
  const total = rows.length;
  const totalEstimado = rows.reduce((acc, r) => acc + (r.estimado || 0), 0);
  const conMonto = rows.filter((r) => Number.isFinite(r.estimado)).length;
  const ticket = conMonto ? totalEstimado / conMonto : 0;
  const adjudicados = rows.filter((r) => r.estado === "Adjudicado").length;
  const pctAdjudicado = total ? (adjudicados / total) * 100 : 0;
  const conOc = rows.filter((r) => cleanText(r.oc) !== "").length;
  const pctOc = total ? (conOc / total) * 100 : 0;

  const leadTimes = rows
    .map((r) => {
      if (!r.apertura || !r.desdeFecha) return null;
      const diffDays = (r.apertura.getTime() - r.desdeFecha.getTime()) / (1000 * 60 * 60 * 24);
      return Number.isFinite(diffDays) && diffDays >= 0 ? diffDays : null;
    })
    .filter((v) => v !== null);

  const leadAvg = leadTimes.length ? leadTimes.reduce((a, b) => a + b, 0) / leadTimes.length : 0;

  const metrics = { total, totalEstimado, ticket, pctAdjudicado, pctOc, leadAvg };
  const defs = KPI_DEFS.filter((kpi) => state.selectedKpis.has(kpi.key));

  if (!defs.length) {
    dom.kpiGrid.innerHTML =
      '<article class="kpi-card"><p>Sin KPI seleccionados</p><strong>Marca al menos uno en el panel lateral.</strong></article>';
    return;
  }

  dom.kpiGrid.innerHTML = defs
    .map(
      (kpi) =>
        '<article class="kpi-card"><p>' +
        escapeHtml(kpi.label) +
        "</p><strong>" +
        escapeHtml(kpi.calc(metrics)) +
        "</strong></article>"
    )
    .join("");
}

function renderKpiSelector() {
  if (!state.selectedKpis.size) {
    KPI_DEFS.forEach((kpi) => state.selectedKpis.add(kpi.key));
  }

  dom.kpiOptions.innerHTML = KPI_DEFS.map((kpi) => {
    const checked = state.selectedKpis.has(kpi.key) ? " checked" : "";
    return (
      '<label class="kpi-option">' +
      '<input type="checkbox" data-kpi-key="' +
      escapeHtml(kpi.key) +
      '"' +
      checked +
      " />" +
      "<span>" +
      escapeHtml(kpi.label) +
      "</span></label>"
    );
  }).join("");

  dom.kpiOptions.querySelectorAll("input[data-kpi-key]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const key = checkbox.dataset.kpiKey;
      if (checkbox.checked) state.selectedKpis.add(key);
      else state.selectedKpis.delete(key);

      storeSelection(STORAGE_KEYS.kpis, [...state.selectedKpis]);
      renderKpis(state.filteredRows);
    });
  });
}

function renderColumnSelector() {
  if (!state.selectedColumns.size) {
    DEFAULT_COLUMNS.forEach((key) => state.selectedColumns.add(key));
  }

  dom.columnOptions.innerHTML = TABLE_COLUMNS.map((col) => {
    const checked = state.selectedColumns.has(col.key) ? " checked" : "";
    return (
      '<label class="kpi-option">' +
      '<input type="checkbox" data-column-key="' +
      escapeHtml(col.key) +
      '"' +
      checked +
      " />" +
      "<span>" +
      escapeHtml(col.label) +
      "</span></label>"
    );
  }).join("");

  dom.columnOptions.querySelectorAll("input[data-column-key]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const key = checkbox.dataset.columnKey;
      if (checkbox.checked) state.selectedColumns.add(key);
      else state.selectedColumns.delete(key);

      if (!state.selectedColumns.size) {
        state.selectedColumns.add("solped");
        const must = dom.columnOptions.querySelector('input[data-column-key="solped"]');
        if (must) must.checked = true;
      }

      storeSelection(STORAGE_KEYS.columns, [...state.selectedColumns]);
      renderTable(state.filteredRows);
    });
  });
}

function getSelectedColumns() {
  const selected = TABLE_COLUMNS.filter((col) => state.selectedColumns.has(col.key));
  return selected.length ? selected : TABLE_COLUMNS.filter((c) => c.key === "solped");
}

function renderTable(rows) {
  const selectedColumns = getSelectedColumns();
  const limit = 250;
  const viewRows = rows.slice(0, limit);

  dom.rowsInfo.textContent =
    "Mostrando " + formatInt.format(viewRows.length) + " de " + formatInt.format(rows.length) + " registros";

  dom.tableHead.innerHTML =
    "<tr>" + selectedColumns.map((col) => "<th>" + escapeHtml(col.label) + "</th>").join("") + "</tr>";

  dom.tableBody.innerHTML = viewRows
    .map((row) => {
      const cells = selectedColumns.map((col) => {
        if (col.editableStatus) {
          return '<td class="status-cell">' + statusSelectHtml(row) + "</td>";
        }
        const value = col.value ? col.value(row) : "";
        return "<td>" + escapeHtml(value) + "</td>";
      });
      return "<tr>" + cells.join("") + "</tr>";
    })
    .join("");
}

function statusSelectHtml(row) {
  const options = getStatusOptions();
  const opts = options
    .map((status) => {
      const selected = status === row.estado ? " selected" : "";
      return '<option value="' + escapeHtml(status) + '"' + selected + ">" + escapeHtml(status) + "</option>";
    })
    .join("");

  return '<select data-row-id="' + escapeHtml(row.id) + '">' + opts + "</select>";
}

function handleStatusChange(event) {
  const target = event.target;
  if (!target.matches("select[data-row-id]")) return;

  const rowId = target.dataset.rowId;
  const newStatus = cleanText(target.value) || "Sin estado";
  const row = state.rawRows.find((item) => item.id === rowId);
  if (!row) return;

  row.estado = newStatus;
  state.statusOverrides[rowId] = newStatus;
  localStorage.setItem(STORAGE_KEYS.statuses, JSON.stringify(state.statusOverrides));

  populateFilters(state.rawRows);
  applyFilters();
}

function renderCharts(rows) {
  const topN = Number(dom.chartTopN.value || "8");

  const byStatusCountRaw = countBy(rows, (r) => r.estado || "Sin estado");
  const byStatusCount = toTopNWithOthers(byStatusCountRaw, topN);

  const amountByStatusRaw = rows.reduce((acc, row) => {
    const key = row.estado || "Sin estado";
    acc[key] = (acc[key] || 0) + (row.estimado || 0);
    return acc;
  }, {});
  const amountByStatus = toTopNWithOthers(amountByStatusRaw, topN);

  const byMonthAll = countBy(rows, (r) => {
    const d = r.desdeFecha || r.apertura;
    if (!d) return null;
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
  });

  const monthLabelsAll = Object.keys(byMonthAll).sort();
  const rangeValue = dom.timelineRange.value;
  const monthLabels =
    rangeValue === "all" ? monthLabelsAll : monthLabelsAll.slice(-Number(rangeValue || "12"));

  state.charts.status = drawChart(state.charts.status, "statusChart", "doughnut", {
    labels: Object.keys(byStatusCount),
    datasets: [
      {
        data: Object.values(byStatusCount),
        backgroundColor: ["#0f766e", "#f77f00", "#7c3aed", "#dc2626", "#2563eb", "#4b5563"],
      },
    ],
  });

  state.charts.amount = drawChart(state.charts.amount, "amountChart", "bar", {
    labels: Object.keys(amountByStatus),
    datasets: [
      {
        label: "Monto estimado",
        data: Object.values(amountByStatus),
        backgroundColor: "#0f766e",
      },
    ],
  });

  state.charts.timeline = drawChart(state.charts.timeline, "timelineChart", "line", {
    labels: monthLabels,
    datasets: [
      {
        label: "Procesos por mes",
        data: monthLabels.map((key) => byMonthAll[key]),
        fill: true,
        borderColor: "#f77f00",
        backgroundColor: "rgba(247,127,0,0.18)",
        tension: 0.35,
      },
    ],
  });
}

function toTopNWithOthers(dict, topN) {
  const sorted = Object.entries(dict).sort((a, b) => b[1] - a[1]);
  const top = sorted.slice(0, topN);
  const rest = sorted.slice(topN);

  if (!rest.length) {
    return Object.fromEntries(top);
  }

  const othersValue = rest.reduce((acc, entry) => acc + entry[1], 0);
  return Object.fromEntries([...top, ["Otros", othersValue]]);
}

function drawChart(oldChart, canvasId, type, data) {
  if (oldChart) {
    oldChart.destroy();
  }

  return new Chart(document.getElementById(canvasId), {
    type,
    data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          labels: {
            font: { family: "Space Grotesk" },
          },
        },
      },
      scales:
        type === "doughnut"
          ? undefined
          : {
              x: {
                ticks: { color: "#4a5c6f" },
                grid: { color: "rgba(0,0,0,0.05)" },
              },
              y: {
                ticks: { color: "#4a5c6f" },
                grid: { color: "rgba(0,0,0,0.05)" },
              },
            },
    },
  });
}

function countBy(rows, accessor) {
  return rows.reduce((acc, row) => {
    const key = accessor(row);
    if (!key) return acc;
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function getStatusOptions() {
  const dynamic = unique(state.rawRows.map((r) => r.estado));
  return unique([...STATUS_BASE_OPTIONS, ...dynamic]);
}

function loadStoredStatuses() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.statuses);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (_error) {
    return {};
  }
}

function loadStoredSelection(key, defaults) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return defaults;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) && parsed.length ? parsed : defaults;
  } catch (_error) {
    return defaults;
  }
}

function storeSelection(key, values) {
  localStorage.setItem(key, JSON.stringify(values));
}

function loadStoredFilters() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.filters);
    if (!raw) return { buyers: [], factories: [], statuses: [] };
    const parsed = JSON.parse(raw);
    return {
      buyers: Array.isArray(parsed.buyers) ? parsed.buyers : [],
      factories: Array.isArray(parsed.factories) ? parsed.factories : [],
      statuses: Array.isArray(parsed.statuses) ? parsed.statuses : [],
    };
  } catch (_error) {
    return { buyers: [], factories: [], statuses: [] };
  }
}

function storeFilters(filters) {
  localStorage.setItem(STORAGE_KEYS.filters, JSON.stringify(filters));
}

function unique(list) {
  return [...new Set(list)].filter(Boolean).sort((a, b) => a.localeCompare(b));
}

function buildRowId(solped, proceso, expediente) {
  const base = [solped, proceso, expediente].map((v) => cleanText(v).toLowerCase()).join("|");
  return base || "row-" + Math.random().toString(36).slice(2, 9);
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;

  const normalized = String(value)
    .trim()
    .replace(/\./g, "")
    .replace(/,/g, ".")
    .replace(/[^0-9.-]/g, "");

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;

  const text = String(value).trim();
  if (!text) return null;

  const nativeDate = new Date(text);
  if (!isNaN(nativeDate.getTime())) return nativeDate;

  const parts = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (parts) {
    const dd = Number(parts[1]);
    const mm = Number(parts[2]) - 1;
    let yyyy = Number(parts[3]);
    if (yyyy < 100) yyyy += 2000;
    const parsed = new Date(yyyy, mm, dd);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  return null;
}

function normalizeStatus(statusRaw) {
  const text = cleanText(statusRaw).toLowerCase();
  if (!text) return "Sin estado";
  if (text.includes("adjudicado")) return "Adjudicado";
  if (text.includes("licit") || text.includes("proceso")) return "En licitacion";
  if (text.includes("pedido") || text.includes("oc") || text.includes("orden")) return "Con OC";
  if (text.includes("pendiente")) return "Pendiente";
  return sentenceCase(text);
}

function sentenceCase(value) {
  if (!value) return "";
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function cleanText(value) {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
