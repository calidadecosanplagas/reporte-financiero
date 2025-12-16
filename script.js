/**
 * Reporte financiero - /data/reporte.xlsx
 * Incluye:
 * - KPIs anuales desde Detalle Clientes (Total/Abono/Diferencia)
 * - KPIs anuales Visitas Únicas (Venta/Abono/Diferencia)
 * - % Cobrado + contadores con/sin deuda
 * - Tabla clientes con filtros: búsqueda + estado + min/max deuda + paginación + orden
 * - Click en cliente => modal con meses + gráfico mensual del cliente + export CSV cliente
 * - Top 20 deuda + export
 * - Export CSV tabla actual (filtrada)
 * - Gráficos mejorados (ticks/tooltip CLP, responsive real)
 */

const EXCEL_PATH = "./data/reporte.xlsx";

let state = {
  unicos: [],
  clientes: [],
  page: 1,
  pageSize: 25,
  query: "",
  sort: "nombre",

  // filtros adicionales
  minDeuda: null,
  maxDeuda: null,
  estadoDeuda: "all", // all | con_deuda | sin_deuda

  // charts
  charts: {
    c1: null,
    c2: null,
    c3: null,
    c4: null,
    modal: null
  },

  // selección modal
  selectedCliente: null,
};

const el = (id) => document.getElementById(id);

function norm(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function toNumberCLP(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return isFinite(value) ? value : 0;

  let s = String(value).trim();
  if (!s || s === "$") return 0;

  s = s.replace(/\$/g, "").replace(/\s+/g, "");
  s = s.replace(/\./g, "").replace(/,/g, ".");
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function formatCLP(n) {
  const sign = n < 0 ? "-" : "";
  const abs = Math.abs(Math.round(n));
  return `${sign}$${abs.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
}

function chartMoneyTicks(value) {
  const n = Number(value) || 0;
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return (n / 1_000_000).toFixed(1).replace(".", ",") + "M";
  if (abs >= 1_000) return (n / 1_000).toFixed(0) + "k";
  return String(Math.round(n));
}

function setText(id, text) {
  const node = el(id);
  if (node) node.textContent = text;
}

function clearTable(tid) {
  const tb = el(tid)?.querySelector("tbody");
  if (tb) tb.innerHTML = "";
}

function addRow(tid, cells, onClick) {
  const tb = el(tid)?.querySelector("tbody");
  if (!tb) return;
  const tr = document.createElement("tr");
  if (onClick) {
    tr.style.cursor = "pointer";
    tr.addEventListener("click", onClick);
  }
  cells.forEach((c) => {
    const td = document.createElement("td");
    if (c?.className) td.className = c.className;
    td.textContent = c?.text ?? "";
    tr.appendChild(td);
  });
  tb.appendChild(tr);
}

function findHeaderRow(rows, neededHeaders, maxScan = 12) {
  const limit = Math.min(maxScan, rows.length);
  for (let i = 0; i < limit; i++) {
    const row = rows[i].map(x => norm(x));
    const ok = neededHeaders.every(h => row.includes(norm(h)));
    if (ok) return i;
  }
  return -1;
}

function detectSheets(workbook) {
  let detalleSheetName = null;
  let unicosSheetName = null;

  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows || rows.length === 0) continue;

    const detalleHeaders = ["Nombre Cliente", "Enero", "Total", "Abono", "Diferencia"];
    const unicosHeadersA = ["Mes", "Venta", "Abono", "Diferencia"];
    const unicosHeadersB = ["Nombre", "Venta", "Abono", "Diferencia"];

    if (findHeaderRow(rows, detalleHeaders, 12) !== -1) detalleSheetName = name;

    if (
      findHeaderRow(rows, unicosHeadersA, 12) !== -1 ||
      findHeaderRow(rows, unicosHeadersB, 12) !== -1
    ) unicosSheetName = name;
  }

  return { detalleSheetName, unicosSheetName };
}

function parseUnicos(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows || rows.length === 0) return [];

  const headerRowIndex =
    findHeaderRow(rows, ["Mes", "Venta", "Abono"], 12) !== -1
      ? findHeaderRow(rows, ["Mes", "Venta", "Abono"], 12)
      : findHeaderRow(rows, ["Nombre", "Venta", "Abono"], 12);

  if (headerRowIndex === -1) return [];

  const header = rows[headerRowIndex].map(x => String(x).trim());
  const headerNorm = header.map(h => norm(h));

  const idxMes = headerNorm.indexOf("MES") !== -1 ? headerNorm.indexOf("MES") : headerNorm.indexOf("NOMBRE");
  const idxVenta = headerNorm.indexOf("VENTA");
  const idxAbono = headerNorm.indexOf("ABONO");
  const idxDif = headerNorm.indexOf("DIFERENCIA");

  const out = [];
  for (let i = headerRowIndex + 1; i < rows.length; i++) {
    const r = rows[i];
    const mes = String(r[idxMes] ?? "").trim();
    if (!mes) continue;

    const venta = toNumberCLP(r[idxVenta]);
    const abono = toNumberCLP(r[idxAbono]);
    const diferencia = (idxDif >= 0) ? toNumberCLP(r[idxDif]) : (abono - venta);

    out.push({ mes, venta, abono, diferencia });
  }
  return out;
}

function parseDetalleClientes(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows || rows.length === 0) return [];

  const headerRowIndex = findHeaderRow(rows, ["Nombre Cliente", "Total", "Abono", "Diferencia"], 12);
  if (headerRowIndex === -1) return [];

  const header = rows[headerRowIndex].map(x => String(x).trim());
  const headerNorm = header.map(h => norm(h));

  const idxNombre = headerNorm.indexOf("NOMBRE CLIENTE");
  const idxTotal = headerNorm.indexOf("TOTAL");
  const idxAbono = headerNorm.indexOf("ABONO");
  const idxDif = headerNorm.indexOf("DIFERENCIA");

  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  const idxMeses = Object.fromEntries(meses.map(m => [m, headerNorm.indexOf(norm(m))]));

  const out = [];
  for (let i = headerRowIndex + 1; i < rows.length; i++) {
    const r = rows[i];
    const nombre = String(r[idxNombre] ?? "").trim();
    if (!nombre) continue;

    const total = toNumberCLP(r[idxTotal]);
    const abono = toNumberCLP(r[idxAbono]);
    const diferencia = toNumberCLP(r[idxDif]);

    const mesesObj = {};
    for (const m of meses) {
      const idx = idxMeses[m];
      mesesObj[m] = idx >= 0 ? toNumberCLP(r[idx]) : 0;
    }

    out.push({ nombre, total, abono, diferencia, meses: mesesObj });
  }

  return out;
}

function computeKPIs(unicos, clientes) {
  const ventaTotal = clientes.reduce((a, x) => a + (x.total || 0), 0);
  const abonoTotal = clientes.reduce((a, x) => a + (x.abono || 0), 0);
  const deudaTotal = clientes.reduce((a, x) => a + (x.diferencia || 0), 0);

  const totalClientes = clientes.length;

  const conDeuda = clientes.filter(c => (c.diferencia || 0) < 0).length;
  const sinDeuda = clientes.filter(c => (c.diferencia || 0) >= 0).length;
  const pct = ventaTotal > 0 ? Math.round((abonoTotal / ventaTotal) * 1000) / 10 : 0;

  setText("kpiVentaTotal", formatCLP(ventaTotal));
  setText("kpiAbonoTotal", formatCLP(abonoTotal));
  setText("kpiDeudaTotal", formatCLP(deudaTotal));
  setText("kpiClientes", String(totalClientes));
  setText("kpiPctCobrado", `${pct}%`);
  setText("kpiConDeuda", String(conDeuda));
  setText("kpiSinDeuda", String(sinDeuda));

  const unicosVenta = unicos.reduce((a, x) => a + (x.venta || 0), 0);
  const unicosAbono = unicos.reduce((a, x) => a + (x.abono || 0), 0);
  const unicosDif = unicos.reduce((a, x) => a + (x.diferencia || 0), 0);

  setText("kpiUnicosVentaTotal", formatCLP(unicosVenta));
  setText("kpiUnicosAbonoTotal", formatCLP(unicosAbono));
  setText("kpiUnicosDeudaTotal", formatCLP(unicosDif));
  setText("kpiUnicosMeses", String(unicos.length));

  // coherencia simple (si diferencia ≈ abono - venta)
  const check = Math.round((unicosAbono - unicosVenta) - unicosDif);
  setText("kpiCoherencia", check === 0 ? "OK" : `Revisar (${formatCLP(check)})`);

  // delta ventas anual (clientes vs unicos)
  const delta = ventaTotal - unicosVenta;
  const sign = delta === 0 ? "" : (delta > 0 ? "+" : "-");
  setText("kpiDeltaVentas", `${sign}${formatCLP(delta)}`);
}

function renderUnicosTable(unicos) {
  clearTable("tablaUnicos");
  for (const x of unicos) {
    addRow("tablaUnicos", [
      { text: x.mes },
      { text: formatCLP(x.venta), className: "num" },
      { text: formatCLP(x.abono), className: "num" },
      { text: formatCLP(x.diferencia), className: `num ${x.diferencia < 0 ? "neg" : "pos"}` },
    ]);
  }
}

/** aplica búsqueda + filtros extra + sort */
function applyFiltersAndSort(clientes) {
  let out = [...clientes];

  // buscar
  const q = state.query.trim().toLowerCase();
  if (q) out = out.filter(c => c.nombre.toLowerCase().includes(q));

  // estado deuda
  if (state.estadoDeuda === "con_deuda") out = out.filter(c => (c.diferencia || 0) < 0);
  if (state.estadoDeuda === "sin_deuda") out = out.filter(c => (c.diferencia || 0) >= 0);

  // min / max deuda (se interpreta por magnitud negativa o por valor directo)
  // aquí usamos "deuda" como ABS de diferencias negativas para filtrar más natural:
  // deuda_cliente = max(0, -diferencia)
  const minD = Number.isFinite(state.minDeuda) ? state.minDeuda : null;
  const maxD = Number.isFinite(state.maxDeuda) ? state.maxDeuda : null;

  if (minD !== null) out = out.filter(c => Math.max(0, -(c.diferencia || 0)) >= minD);
  if (maxD !== null) out = out.filter(c => Math.max(0, -(c.diferencia || 0)) <= maxD);

  // ordenar
  switch (state.sort) {
    case "total_desc": out.sort((a,b) => (b.total||0) - (a.total||0)); break;
    case "abono_desc": out.sort((a,b) => (b.abono||0) - (a.abono||0)); break;
    case "diferencia_asc": out.sort((a,b) => (a.diferencia||0) - (b.diferencia||0)); break;
    default: out.sort((a,b) => a.nombre.localeCompare(b.nombre, "es"));
  }

  return out;
}

function renderClientesTable(clientes) {
  clearTable("tablaClientes");

  const filtered = applyFiltersAndSort(clientes);

  // KPI mostrando
  setText("kpiMostrando", `${filtered.length} / ${clientes.length}`);

  const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
  state.page = Math.min(Math.max(state.page, 1), totalPages);

  const start = (state.page - 1) * state.pageSize;
  const pageRows = filtered.slice(start, start + state.pageSize);

  for (const c of pageRows) {
    addRow("tablaClientes", [
      { text: c.nombre },
      { text: formatCLP(c.total), className: "num" },
      { text: formatCLP(c.abono), className: "num" },
      { text: formatCLP(c.diferencia), className: `num ${c.diferencia < 0 ? "neg" : "pos"}` },
    ], () => openClienteModal(c));
  }

  setText("pageInfo", `Página ${state.page} de ${totalPages} · Mostrando ${pageRows.length} de ${filtered.length}`);
}

function renderTopDeudaTable(clientes) {
  clearTable("tablaTopDeuda");
  const top20 = [...clientes].sort((a,b) => (a.diferencia||0) - (b.diferencia||0)).slice(0,20);
  top20.forEach((c, idx) => {
    addRow("tablaTopDeuda", [
      { text: String(idx + 1) },
      { text: c.nombre },
      { text: formatCLP(c.total), className:"num" },
      { text: formatCLP(c.abono), className:"num" },
      { text: formatCLP(c.diferencia), className:`num ${c.diferencia < 0 ? "neg" : "pos"}` },
    ], () => openClienteModal(c));
  });
}

function destroyCharts() {
  for (const k of ["c1","c2","c3","c4","modal"]) {
    if (state.charts[k]) {
      state.charts[k].destroy();
      state.charts[k] = null;
    }
  }
}

function baseChartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: "bottom" },
      tooltip: {
        callbacks: {
          label: (ctx) => `${ctx.dataset.label}: ${formatCLP(ctx.parsed.y)}`
        }
      }
    },
    scales: {
      y: { ticks: { callback: (v) => chartMoneyTicks(v) } }
    }
  };
}

function renderCharts(unicos, clientes) {
  // NO destruir modal chart aquí (si está abierto)
  for (const k of ["c1","c2","c3","c4"]) {
    if (state.charts[k]) { state.charts[k].destroy(); state.charts[k] = null; }
  }

  // 1) Line Venta vs Abono (Unicos)
  const labelsU = unicos.map(x => x.mes);
  const ventasU = unicos.map(x => x.venta);
  const abonosU = unicos.map(x => x.abono);
  const difsU = unicos.map(x => x.diferencia);

  const ctx1 = el("chartUnicosVentaAbono");
  if (ctx1) {
    state.charts.c1 = new Chart(ctx1, {
      type: "line",
      data: {
        labels: labelsU,
        datasets: [
          { label: "Venta", data: ventasU, tension: 0.28, pointRadius: 2 },
          { label: "Abono", data: abonosU, tension: 0.28, pointRadius: 2 },
        ]
      },
      options: {
        ...baseChartOptions(),
        interaction: { mode: "index", intersect: false }
      }
    });
  }

  // 2) Bar Diferencia (Unicos)
  const ctx2 = el("chartUnicosDiferencia");
  if (ctx2) {
    state.charts.c2 = new Chart(ctx2, {
      type: "bar",
      data: { labels: labelsU, datasets: [{ label: "Diferencia", data: difsU }] },
      options: baseChartOptions()
    });
  }

  // 3) Top 10 deuda (Clientes)
  const topDeuda10 = [...clientes]
    .sort((a,b) => (a.diferencia||0) - (b.diferencia||0))
    .slice(0,10);

  const labelsTop = topDeuda10.map(x => x.nombre);
  const valoresTop = topDeuda10.map(x => x.diferencia);

  const ctx3 = el("chartTopDeuda");
  if (ctx3) {
    state.charts.c3 = new Chart(ctx3, {
      type: "bar",
      data: { labels: labelsTop, datasets: [{ label: "Deuda", data: valoresTop }] },
      options: {
        ...baseChartOptions(),
        scales: {
          x: {
            ticks: {
              autoSkip: false,
              maxRotation: 35,
              callback: function(v) {
                const label = this.getLabelForValue(v);
                return label.length > 18 ? label.slice(0, 18) + "…" : label;
              }
            }
          },
          y: { ticks: { callback: (v) => chartMoneyTicks(v) } }
        }
      }
    });
  }

  // 4) Top 10 ventas (Clientes): Venta vs Abono (bar agrupado)
  const topVentas10 = [...clientes]
    .sort((a,b) => (b.total||0) - (a.total||0))
    .slice(0,10);

  const ctx4 = el("chartTopVentas");
  if (ctx4) {
    state.charts.c4 = new Chart(ctx4, {
      type: "bar",
      data: {
        labels: topVentas10.map(x => x.nombre),
        datasets: [
          { label: "Venta anual", data: topVentas10.map(x => x.total || 0) },
          { label: "Abono anual", data: topVentas10.map(x => x.abono || 0) },
        ]
      },
      options: {
        ...baseChartOptions(),
        scales: {
          x: {
            ticks: {
              autoSkip: false,
              maxRotation: 35,
              callback: function(v) {
                const label = this.getLabelForValue(v);
                return label.length > 16 ? label.slice(0, 16) + "…" : label;
              }
            }
          },
          y: { ticks: { callback: (v) => chartMoneyTicks(v) } }
        }
      }
    });
  }
}

async function loadExcel() {
  const res = await fetch(EXCEL_PATH, { cache: "no-store" });
  if (!res.ok) {
    throw new Error(`No se pudo cargar ${EXCEL_PATH}. Revisa que exista en el repo y esté commiteado (ruta exacta /data/reporte.xlsx).`);
  }
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });

  const { detalleSheetName, unicosSheetName } = detectSheets(wb);

  if (!detalleSheetName) throw new Error("No encontré la hoja Detalle Clientes (encabezado: Nombre Cliente, meses, Total, Abono, Diferencia).");
  if (!unicosSheetName) throw new Error("No encontré la hoja Visitas Únicas (encabezado: Mes/Nombre + Venta + Abono + Diferencia).");

  const clientes = parseDetalleClientes(wb.Sheets[detalleSheetName]);
  const unicos = parseUnicos(wb.Sheets[unicosSheetName]);

  return { clientes, unicos };
}

/* ==========================
   Export CSV helpers
========================== */
function toCSV(rows) {
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  return rows.map(r => r.map(esc).join(",")).join("\n");
}

function downloadText(filename, text) {
  const blob = new Blob([text], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function exportTablaActualCSV() {
  const filtered = applyFiltersAndSort(state.clientes);
  const rows = [
    ["Cliente","Total","Abono","Diferencia"],
    ...filtered.map(c => [c.nombre, c.total, c.abono, c.diferencia])
  ];
  downloadText("clientes_filtrados.csv", toCSV(rows));
}

function exportTop20CSV() {
  const top20 = [...state.clientes].sort((a,b) => (a.diferencia||0) - (b.diferencia||0)).slice(0,20);
  const rows = [
    ["#","Cliente","Total","Abono","Deuda"],
    ...top20.map((c,i) => [i+1, c.nombre, c.total, c.abono, c.diferencia])
  ];
  downloadText("top20_deuda.csv", toCSV(rows));
}

/* ==========================
   Modal cliente
========================== */
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

function openClienteModal(cliente) {
  state.selectedCliente = cliente;

  el("modal")?.classList.remove("hidden");
  setText("modalTitle", cliente.nombre);
  setText("modalSub", `Total: ${formatCLP(cliente.total)} · Abono: ${formatCLP(cliente.abono)} · Diferencia: ${formatCLP(cliente.diferencia)}`);

  setText("mTotal", formatCLP(cliente.total));
  setText("mAbono", formatCLP(cliente.abono));
  setText("mDif", formatCLP(cliente.diferencia));

  // tabla meses
  clearTable("tablaClienteMeses");
  for (const m of MESES) {
    addRow("tablaClienteMeses", [
      { text: m },
      { text: formatCLP(cliente.meses?.[m] || 0), className: "num" }
    ]);
  }

  // gráfico mensual cliente
  if (state.charts.modal) { state.charts.modal.destroy(); state.charts.modal = null; }
  const ctx = el("chartClienteMensual");
  if (ctx) {
    const dataMeses = MESES.map(m => cliente.meses?.[m] || 0);
    state.charts.modal = new Chart(ctx, {
      type: "bar",
      data: {
        labels: MESES,
        datasets: [{ label: "Monto mensual", data: dataMeses }]
      },
      options: {
        ...baseChartOptions(),
        scales: {
          x: { ticks: { maxRotation: 0, autoSkip: false } },
          y: { ticks: { callback: (v) => chartMoneyTicks(v) } }
        }
      }
    });
  }
}

function closeModal() {
  el("modal")?.classList.add("hidden");
}

function exportClienteCSV() {
  const c = state.selectedCliente;
  if (!c) return;

  const rows = [
    ["Cliente", c.nombre],
    ["Total", c.total],
    ["Abono", c.abono],
    ["Diferencia", c.diferencia],
    [],
    ["Mes","Monto"],
    ...MESES.map(m => [m, c.meses?.[m] || 0])
  ];
  downloadText(`cliente_${c.nombre.replace(/[^\w]+/g,"_").toLowerCase()}.csv`, toCSV(rows));
}

/* ==========================
   Init + UI
========================== */
async function init() {
  try {
    // placeholders
    ["kpiClientes","kpiVentaTotal","kpiAbonoTotal","kpiDeudaTotal","kpiPctCobrado","kpiSinDeuda","kpiConDeuda","kpiMostrando",
     "kpiUnicosVentaTotal","kpiUnicosAbonoTotal","kpiUnicosDeudaTotal","kpiUnicosMeses","kpiCoherencia","kpiDeltaVentas"
    ].forEach(id => setText(id, "…"));

    const { clientes, unicos } = await loadExcel();

    state.clientes = clientes;
    state.unicos = unicos;
    state.page = 1;

    computeKPIs(unicos, clientes);
    renderUnicosTable(unicos);
    renderClientesTable(clientes);
    renderTopDeudaTable(clientes);
    renderCharts(unicos, clientes);

  } catch (err) {
    console.error(err);
    alert(err.message || String(err));
  }
}

function wireUI() {
  el("btnReload")?.addEventListener("click", init);
  el("btnPrint")?.addEventListener("click", () => window.print());
  el("btnExportCSV")?.addEventListener("click", exportTablaActualCSV);
  el("btnExportTopDeuda")?.addEventListener("click", exportTop20CSV);

  el("searchCliente")?.addEventListener("input", (e) => {
    state.query = e.target.value || "";
    state.page = 1;
    renderClientesTable(state.clientes);
  });

  el("sortBy")?.addEventListener("change", (e) => {
    state.sort = e.target.value;
    state.page = 1;
    renderClientesTable(state.clientes);
  });

  el("prevPage")?.addEventListener("click", () => {
    state.page -= 1;
    renderClientesTable(state.clientes);
  });

  el("nextPage")?.addEventListener("click", () => {
    state.page += 1;
    renderClientesTable(state.clientes);
  });

  // filtros extra
  el("minDeuda")?.addEventListener("input", (e) => {
    const v = toNumberCLP(e.target.value);
    state.minDeuda = (e.target.value.trim() === "") ? null : v;
    state.page = 1;
    renderClientesTable(state.clientes);
  });

  el("maxDeuda")?.addEventListener("input", (e) => {
    const v = toNumberCLP(e.target.value);
    state.maxDeuda = (e.target.value.trim() === "") ? null : v;
    state.page = 1;
    renderClientesTable(state.clientes);
  });

  el("estadoDeuda")?.addEventListener("change", (e) => {
    state.estadoDeuda = e.target.value;
    state.page = 1;
    renderClientesTable(state.clientes);
  });

  el("pageSize")?.addEventListener("change", (e) => {
    state.pageSize = parseInt(e.target.value, 10) || 25;
    state.page = 1;
    renderClientesTable(state.clientes);
  });

  el("btnClearFilters")?.addEventListener("click", () => {
    state.query = "";
    state.sort = "nombre";
    state.minDeuda = null;
    state.maxDeuda = null;
    state.estadoDeuda = "all";
    state.pageSize = 25;
    state.page = 1;

    if (el("searchCliente")) el("searchCliente").value = "";
    if (el("sortBy")) el("sortBy").value = "nombre";
    if (el("minDeuda")) el("minDeuda").value = "";
    if (el("maxDeuda")) el("maxDeuda").value = "";
    if (el("estadoDeuda")) el("estadoDeuda").value = "all";
    if (el("pageSize")) el("pageSize").value = "25";

    renderClientesTable(state.clientes);
  });

  // modal
  el("modalClose")?.addEventListener("click", closeModal);
  el("modalCloseBtn")?.addEventListener("click", closeModal);
  el("btnExportClienteCSV")?.addEventListener("click", exportClienteCSV);

  // ESC para cerrar modal
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !el("modal")?.classList.contains("hidden")) closeModal();
  });
}

wireUI();
init();
