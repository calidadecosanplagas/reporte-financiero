/**
 * Reporte financiero - lee /data/reporte.xlsx directamente (sin CSV)
 * - Detecta hojas por encabezados
 * - Renderiza KPIs + tablas + paginación + búsqueda + gráficos
 * - Frecuencia mensual desde Detalle Clientes:
 *    clientes con ingreso > 0, total mes, promedio por cliente activo
 */

const EXCEL_URL = "data/reporte.xlsx";

if (typeof XLSX === "undefined") {
  alert("SheetJS (XLSX) no cargó. Revisa que xlsx.full.min.js esté disponible (CDN bloqueado o falta archivo local).");
  throw new Error("XLSX is not defined (SheetJS no cargó).");
}


let state = {
  unicos: [],       // [{mes, venta, abono, diferencia}]
  clientes: [],     // [{nombre,total,abono,diferencia, meses:{Enero:..}}]
  actividadMensual: [], // [{mes, activos, totalMes, promedioActivo}]
  page: 1,
  pageSize: 25,
  query: "",
  sort: "nombre",
  charts: {
    c1: null,
    c2: null,
    c3: null,
    c5: null,
  }
};

const el = (id) => document.getElementById(id);

/* ===================== HELPERS NUM ===================== */

function toNumberCLP(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return isFinite(value) ? value : 0;

  let s = String(value).trim();
  if (!s || s === "$") return 0;

  // quita $ y espacios
  s = s.replace(/\$/g, "").replace(/\s+/g, "");

  // normaliza miles/decimales
  // 1.234.567 -> 1234567
  // 1,23 (no deberías tener decimales, pero por si acaso)
  s = s.replace(/\./g, "").replace(/,/g, ".");

  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function formatCLP(n) {
  const sign = n < 0 ? "-" : "";
  const abs = Math.abs(Math.round(n));
  return `${sign}$${abs.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
}

function setText(id, text) {
  const node = el(id);
  if (node) node.textContent = text;
}

function clearTable(tid) {
  const tb = el(tid)?.querySelector("tbody");
  if (tb) tb.innerHTML = "";
}

function addRow(tid, cells) {
  const tb = el(tid)?.querySelector("tbody");
  if (!tb) return;
  const tr = document.createElement("tr");
  cells.forEach((c) => {
    const td = document.createElement("td");
    if (c?.className) td.className = c.className;
    td.textContent = c?.text ?? "";
    tr.appendChild(td);
  });
  tb.appendChild(tr);
}

function safeDiv(a, b) {
  if (!b) return 0;
  return a / b;
}

/* ===================== DETECT + PARSE ===================== */

function detectSheets(workbook) {
  let detalleSheetName = null;
  let unicosSheetName = null;

  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows || rows.length === 0) continue;

    const headCandidates = rows.slice(0, 10).map(r => r.map(x => String(x).trim()));

    // Detalle clientes
    const hasNombreCliente = headCandidates.some(r => r.includes("Nombre Cliente"));
    const hasTotal = headCandidates.some(r => r.includes("Total"));
    const hasEnero = headCandidates.some(r => r.includes("Enero"));
    const hasDiferencia = headCandidates.some(r => r.includes("Diferencia"));

    if (hasNombreCliente && hasTotal && hasEnero && hasDiferencia) {
      detalleSheetName = name;
      continue;
    }

    // Unicos
    const hasVenta = headCandidates.some(r => r.includes("Venta"));
    const hasAbono = headCandidates.some(r => r.includes("Abono"));
    const hasNombreOrMes = headCandidates.some(r => r.includes("Nombre") || r.includes("Mes"));

    if (hasNombreOrMes && hasVenta && hasAbono && hasDiferencia) {
      unicosSheetName = name;
      continue;
    }
  }

  return { detalleSheetName, unicosSheetName };
}

function parseUnicos(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows || rows.length === 0) return [];

  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const r = rows[i].map(x => String(x).trim());
    if ((r.includes("Nombre") || r.includes("Mes")) && r.includes("Venta") && r.includes("Abono")) {
      headerRowIndex = i;
      break;
    }
  }
  if (headerRowIndex === -1) return [];

  const header = rows[headerRowIndex].map(x => String(x).trim());
  const idxMes = header.indexOf("Mes") !== -1 ? header.indexOf("Mes") : header.indexOf("Nombre");
  const idxVenta = header.indexOf("Venta");
  const idxAbono = header.indexOf("Abono");
  const idxDif = header.indexOf("Diferencia");

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

  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const r = rows[i].map(x => String(x).trim());
    if (r.includes("Nombre Cliente") && r.includes("Total") && r.includes("Abono") && r.includes("Diferencia")) {
      headerRowIndex = i;
      break;
    }
  }
  if (headerRowIndex === -1) return [];

  const header = rows[headerRowIndex].map(x => String(x).trim());

  const idxNombre = header.indexOf("Nombre Cliente");
  const idxTotal = header.indexOf("Total");
  const idxAbono = header.indexOf("Abono");
  const idxDif = header.indexOf("Diferencia");

  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  const idxMeses = Object.fromEntries(meses.map(m => [m, header.indexOf(m)]));

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

/* ===================== FRECUENCIA MENSUAL (CLIENTES) ===================== */

function computeActividadMensualDesdeClientes(clientes) {
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

  return meses.map(m => {
    let activos = 0;
    let totalMes = 0;

    for (const c of clientes) {
      const v = (c.meses?.[m] ?? 0);
      if (v > 0) activos += 1;
      totalMes += v;
    }

    const promedioActivo = activos > 0 ? (totalMes / activos) : 0;

    return { mes: m, activos, totalMes, promedioActivo };
  });
}

function renderActividadMensualTable(rows) {
  clearTable("tablaActividadMensual");
  for (const r of rows) {
    addRow("tablaActividadMensual", [
      { text: r.mes },
      { text: String(r.activos), className: "num" },
      { text: formatCLP(r.totalMes), className: "num" },
      { text: formatCLP(r.promedioActivo), className: "num" },
    ]);
  }
}

/* ===================== KPIs ===================== */

function computeKPIs(unicos, clientes) {
  // ===================== CLIENTES (DETALLE CLIENTES) =====================
  const ventaTotalClientes = clientes.reduce((a, c) => a + (c.total || 0), 0);
  const abonoTotalClientes = clientes.reduce((a, c) => a + (c.abono || 0), 0);
  const deudaTotalClientes = clientes.reduce((a, c) => a + (c.diferencia || 0), 0);

  const totalClientes = clientes.length;
  const sinDeuda = clientes.filter(c => (c.diferencia || 0) >= 0).length;
  const conDeuda = clientes.filter(c => (c.diferencia || 0) < 0).length;

  const pctCobrado = ventaTotalClientes > 0 ? (abonoTotalClientes / ventaTotalClientes) * 100 : 0;

  setText("kpiVentaTotalClientes", formatCLP(ventaTotalClientes));
  setText("kpiAbonoTotalClientes", formatCLP(abonoTotalClientes));
  setText("kpiDeudaTotalClientes", formatCLP(deudaTotalClientes));
  setText("kpiClientes", String(totalClientes));
  setText("kpiSinDeuda", String(sinDeuda));
  setText("kpiConDeuda", String(conDeuda));
  setText("kpiPctCobrado", `${pctCobrado.toFixed(1)}%`);

  // “Mostrando” depende del filtro actual (se actualiza también en wireUI)
  const filtered = applyFiltersAndSort(clientes);
  setText("kpiMostrando", `${filtered.length} / ${clientes.length}`);

  // promedio mensual anual (clientes)
  setText("kpiPromMesClientes", formatCLP(ventaTotalClientes / 12));

  // ===================== ACTIVIDAD MENSUAL + KPIs EXTRA (CLIENTES) =====================
  const actividad = computeActividadMensualDesdeClientes(clientes);
  state.actividadMensual = actividad;

  // KPI: Promedio mensual por cliente activo (promedio de promedios mensuales)
  const promedios = actividad.map(x => x.promedioActivo).filter(x => x > 0);
  const promMesPorCliente = promedios.length ? (promedios.reduce((a,b)=>a+b,0) / promedios.length) : 0;
  setText("kpiPromMesPorCliente", formatCLP(promMesPorCliente));

  // KPI: mes con más activos
  const mesMasActivos = [...actividad].sort((a,b) => b.activos - a.activos)[0];
  setText("kpiMesMasActivos", mesMasActivos ? `${mesMasActivos.mes} (${mesMasActivos.activos})` : "—");

  // KPI: mes con mayor ingreso
  const mesMayorIngreso = [...actividad].sort((a,b) => b.totalMes - a.totalMes)[0];
  setText("kpiMesMayorIngreso", mesMayorIngreso ? `${mesMayorIngreso.mes} (${formatCLP(mesMayorIngreso.totalMes)})` : "—");

  // ===================== ÚNICOS (VISITAS ÚNICAS) =====================
  const ventaTotalUnicos = unicos.reduce((a, x) => a + (x.venta || 0), 0);
  const abonoTotalUnicos = unicos.reduce((a, x) => a + (x.abono || 0), 0);
  const mesesLeidos = unicos.length;

  setText("kpiVentaTotalUnicos", formatCLP(ventaTotalUnicos));
  setText("kpiAbonoTotalUnicos", formatCLP(abonoTotalUnicos));
  setText("kpiMesesUnicos", String(mesesLeidos));

  // coherencia (simple pero útil): suma difs vs (abono - venta)
  const difExcel = unicos.reduce((a, x) => a + (x.diferencia || 0), 0);
  const difCalc = abonoTotalUnicos - ventaTotalUnicos;
  const delta = Math.abs(difExcel - difCalc);
  const coherencia = mesesLeidos === 0
    ? "Sin datos"
    : (delta <= 2 ? "OK" : `Revisar (Δ ${formatCLP(delta)})`);
  setText("kpiCoherenciaUnicos", coherencia);

  // promedio mensual únicos (venta / meses leídos)
  const promMesUnicos = mesesLeidos > 0 ? (ventaTotalUnicos / mesesLeidos) : 0;
  setText("kpiPromMesUnicos", formatCLP(promMesUnicos));

  // Comparación venta anual (clientes - unicos)
  const diffVenta = ventaTotalClientes - ventaTotalUnicos;
  const sign = diffVenta >= 0 ? "+" : "";
  setText("kpiComparacionVenta", `${sign}${formatCLP(diffVenta)}`);

  // Render de tabla actividad mensual (ya calculada arriba)
  renderActividadMensualTable(actividad);
}

/* ===================== TABLES ===================== */

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

function applyFiltersAndSort(clientes) {
  let out = [...clientes];

  const q = state.query.trim().toLowerCase();
  if (q) out = out.filter(c => c.nombre.toLowerCase().includes(q));

  switch (state.sort) {
    case "total_desc":
      out.sort((a,b) => (b.total||0) - (a.total||0)); break;
    case "abono_desc":
      out.sort((a,b) => (b.abono||0) - (a.abono||0)); break;
    case "diferencia_asc":
      out.sort((a,b) => (a.diferencia||0) - (b.diferencia||0)); break; // más negativo primero
    default:
      out.sort((a,b) => a.nombre.localeCompare(b.nombre, "es"));
  }

  return out;
}

function renderClientesTable(clientes) {
  clearTable("tablaClientes");

  const filtered = applyFiltersAndSort(clientes);

  const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
  if (state.page > totalPages) state.page = totalPages;
  if (state.page < 1) state.page = 1;

  const start = (state.page - 1) * state.pageSize;
  const pageRows = filtered.slice(start, start + state.pageSize);

  for (const c of pageRows) {
    addRow("tablaClientes", [
      { text: c.nombre },
      { text: formatCLP(c.total), className: "num" },
      { text: formatCLP(c.abono), className: "num" },
      { text: formatCLP(c.diferencia), className: `num ${c.diferencia < 0 ? "neg" : "pos"}` },
    ]);
  }

  setText("pageInfo", `Página ${state.page} de ${totalPages} · Mostrando ${pageRows.length} de ${filtered.length}`);

  // mantiene KPI “Mostrando” coherente con búsqueda/filtros
  setText("kpiMostrando", `${filtered.length} / ${clientes.length}`);
}

/* ===================== CHART HELPERS ===================== */

function destroyCharts() {
  for (const k of ["c1","c2","c3","c5"]) {
    if (state.charts[k]) {
      state.charts[k].destroy();
      state.charts[k] = null;
    }
  }
}
function pct(n) {
  if (!isFinite(n)) return "0%";
  return `${(n * 100).toFixed(1)}%`;
}

function clamp01(x) {
  return Math.max(0, Math.min(1, x));
}

function scoreLabel(s) {
  if (s >= 0.75) return "Alto";
  if (s >= 0.45) return "Medio";
  return "Bajo";
}


function chartMoneyTicks(v) {
  const n = Number(v);
  if (!isFinite(n)) return v;

  // abreviación para no llenar el eje con números gigantes
  const abs = Math.abs(n);
  if (abs >= 1_000_000_000) return `${Math.round(n / 1_000_000_000)}B`;
  if (abs >= 1_000_000) return `${Math.round(n / 1_000_000)}M`;
  if (abs >= 1_000) return `${Math.round(n / 1_000)}K`;
  return String(Math.round(n));
}

const CHART_GRID_COLOR = "rgba(255,255,255,.08)";
const CHART_TICK_COLOR = "rgba(255,255,255,.65)";

/* ===================== CHARTS ===================== */

function renderCharts(unicos, clientes) {
  destroyCharts();

  // ========== 1) Unicos - Venta vs Abono (Line) ==========
  const labelsUnicos = unicos.map(x => x.mes);
  const ventas = unicos.map(x => x.venta);
  const abonos = unicos.map(x => x.abono);

  const ctx1 = el("chartUnicosVentaAbono");
  if (ctx1) {
    state.charts.c1 = new Chart(ctx1, {
      type: "line",
      data: {
        labels: labelsUnicos,
        datasets: [
          { label: "Venta", data: ventas, tension: 0.35, pointRadius: 3, borderWidth: 2 },
          { label: "Abono", data: abonos, tension: 0.35, pointRadius: 3, borderWidth: 2 },
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom", labels: { color: CHART_TICK_COLOR } },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatCLP(ctx.raw)}`
            }
          }
        },
        scales: {
          x: { ticks: { color: CHART_TICK_COLOR }, grid: { color: "rgba(255,255,255,.05)" } },
          y: { ticks: { color: CHART_TICK_COLOR, callback: (v) => chartMoneyTicks(v) }, grid: { color: CHART_GRID_COLOR } }
        }
      }
    });
  }

  // ========== 2) Unicos - Diferencia (Bar) ==========
  const difs = unicos.map(x => x.diferencia);

  const ctx2 = el("chartUnicosDiferencia");
  if (ctx2) {
    state.charts.c2 = new Chart(ctx2, {
      type: "bar",
      data: {
        labels: labelsUnicos,
        datasets: [{ label: "Diferencia", data: difs }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom", labels: { color: CHART_TICK_COLOR } },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatCLP(ctx.raw)}`
            }
          }
        },
        scales: {
          x: { ticks: { color: CHART_TICK_COLOR }, grid: { color: "rgba(255,255,255,.05)" } },
          y: { ticks: { color: CHART_TICK_COLOR, callback: (v) => chartMoneyTicks(v) }, grid: { color: CHART_GRID_COLOR } }
        }
      }
    });
  }

  // ========== 3) Top 10 Deuda (más negativo primero) ==========
  const topDeuda = [...clientes]
    .filter(c => typeof c.diferencia === "number")
    .sort((a,b) => (a.diferencia||0) - (b.diferencia||0))
    .slice(0, 10);

  const labelsTop = topDeuda.map(x => x.nombre);
  const valoresTop = topDeuda.map(x => x.diferencia);

  const ctx3 = el("chartTopDeuda");
  if (ctx3) {
    state.charts.c3 = new Chart(ctx3, {
      type: "bar",
      data: {
        labels: labelsTop,
        datasets: [{ label: "Deuda (Diferencia)", data: valoresTop }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom", labels: { color: CHART_TICK_COLOR } },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatCLP(ctx.raw)}`
            }
          }
        },
        scales: {
          x: {
            ticks: { color: CHART_TICK_COLOR, autoSkip: false, maxRotation: 0, minRotation: 0 },
            grid: { display: false }
          },
          y: { ticks: { color: CHART_TICK_COLOR, callback: (v) => chartMoneyTicks(v) }, grid: { color: CHART_GRID_COLOR } }
        }
      }
    });
  }

  // ========== 4) Actividad mensual - Ingreso total por mes ==========
  const act = state.actividadMensual?.length ? state.actividadMensual : computeActividadMensualDesdeClientes(clientes);

  const ctx5 = el("chartActividadMensual");
  if (ctx5) {
    state.charts.c5 = new Chart(ctx5, {
      type: "bar",
      data: {
        labels: act.map(x => x.mes),
        datasets: [
          { label: "Ingreso total del mes", data: act.map(x => x.totalMes) }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom", labels: { color: CHART_TICK_COLOR } },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatCLP(ctx.raw)}`
            }
          }
        },
        scales: {
          x: { ticks: { color: CHART_TICK_COLOR }, grid: { color: "rgba(255,255,255,.05)" } },
          y: { ticks: { color: CHART_TICK_COLOR, callback: (v) => chartMoneyTicks(v) }, grid: { color: CHART_GRID_COLOR } }
        }
      }
    });
  }
}

/* ===================== LOAD EXCEL ===================== */

async function loadExcel() {
  const res = await fetch(EXCEL_URL, { cache: "no-store" });
  if (!res.ok) throw new Error(`No se pudo cargar ${EXCEL_URL}. ¿Está en /data y se llama reporte.xlsx?`);

  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });

  const { detalleSheetName, unicosSheetName } = detectSheets(wb);

  if (!detalleSheetName) {
    throw new Error("No encontré la hoja de Detalle Clientes. Encabezado requerido: 'Nombre Cliente', meses, 'Total', 'Abono', 'Diferencia'.");
  }
  if (!unicosSheetName) {
    throw new Error("No encontré la hoja de Visitas Únicas. Encabezado requerido: 'Mes' o 'Nombre' + 'Venta' + 'Abono' + 'Diferencia'.");
  }

  const wsDetalle = wb.Sheets[detalleSheetName];
  const wsUnicos = wb.Sheets[unicosSheetName];

  const clientes = parseDetalleClientes(wsDetalle);
  const unicos = parseUnicos(wsUnicos);

  return { clientes, unicos };
}

/* ===================== EXPORT CSV ===================== */

function exportClientesCSV(clientes) {
  const header = ["Nombre Cliente","Total","Abono","Diferencia"];
  const lines = [header.join(",")];

  for (const c of clientes) {
    const row = [
      `"${String(c.nombre).replace(/"/g,'""')}"`,
      c.total ?? 0,
      c.abono ?? 0,
      c.diferencia ?? 0
    ];
    lines.push(row.join(","));
  }

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "detalle_clientes.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ===================== INIT + UI ===================== */

async function init() {
  try {
    // placeholders (evita “undefined” en pantalla)
    [
      "kpiVentaTotalClientes","kpiAbonoTotalClientes","kpiDeudaTotalClientes","kpiClientes",
      "kpiSinDeuda","kpiConDeuda","kpiPctCobrado","kpiMostrando",
      "kpiPromMesClientes","kpiPromMesPorCliente","kpiMesMasActivos","kpiMesMayorIngreso",
      "kpiVentaTotalUnicos","kpiAbonoTotalUnicos","kpiMesesUnicos","kpiCoherenciaUnicos",
      "kpiPromMesUnicos","kpiComparacionVenta"
    ].forEach(id => setText(id, "…"));

    const { clientes, unicos } = await loadExcel();

    state.clientes = clientes;
    state.unicos = unicos;
    state.page = 1;

    computeKPIs(state.unicos, state.clientes);
    renderUnicosTable(state.unicos);
    renderClientesTable(state.clientes);
    renderCharts(state.unicos, state.clientes);

  } catch (err) {
    console.error(err);
    alert(err.message || String(err));
  }
}
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function buildResumenHtml() {
  const clientesFiltrados = applyFiltersAndSort(state.clientes);

  const ventaTotalClientes = state.clientes.reduce((a,c)=>a+(c.total||0),0);
  const abonoTotalClientes = state.clientes.reduce((a,c)=>a+(c.abono||0),0);
  const deudaTotalClientes = state.clientes.reduce((a,c)=>a+(c.diferencia||0),0);

  const totalClientes = state.clientes.length;
  const pctCobrado = ventaTotalClientes > 0 ? (abonoTotalClientes / ventaTotalClientes) : 0;

  const actividad = state.actividadMensual?.length
    ? state.actividadMensual
    : computeActividadMensualDesdeClientes(state.clientes);

  const mesesConActividad = actividad.filter(x => x.totalMes > 0);
  const promMesClientes = ventaTotalClientes / 12;

  const mesMasActivos = [...actividad].sort((a,b)=>b.activos-a.activos)[0];
  const mesMayorIngreso = [...actividad].sort((a,b)=>b.totalMes-a.totalMes)[0];
  const mesMenorIngreso = [...mesesConActividad].sort((a,b)=>a.totalMes-b.totalMes)[0];

  // Volatilidad simple: coeficiente de variación (std/mean) sobre meses con ingreso > 0
  const mean = mesesConActividad.length
    ? (mesesConActividad.reduce((a,x)=>a+x.totalMes,0) / mesesConActividad.length)
    : 0;

  const variance = mesesConActividad.length
    ? (mesesConActividad.reduce((a,x)=>a + Math.pow(x.totalMes - mean, 2), 0) / mesesConActividad.length)
    : 0;

  const std = Math.sqrt(variance);
  const cv = mean > 0 ? (std / mean) : 0; // 0 = estable, >1 muy variable

  // Top deuda (más negativo primero)
  const topDeuda = [...state.clientes]
    .filter(c => typeof c.diferencia === "number")
    .sort((a,b) => (a.diferencia||0) - (b.diferencia||0))
    .slice(0, 10);

  const deudaTop10Abs = topDeuda.reduce((a,x)=>a + Math.abs(Math.min(0, x.diferencia||0)), 0);
  const deudaTotalAbs = state.clientes.reduce((a,c)=>a + Math.abs(Math.min(0, c.diferencia||0)), 0);
  const concentracionDeuda = deudaTotalAbs > 0 ? (deudaTop10Abs / deudaTotalAbs) : 0;

  // Top ventas
  const topVentas = [...state.clientes]
    .sort((a,b)=> (b.total||0) - (a.total||0))
    .slice(0, 10);

  const ventaTop10 = topVentas.reduce((a,x)=>a + (x.total||0),0);
  const concentracionVenta = ventaTotalClientes > 0 ? (ventaTop10 / ventaTotalClientes) : 0;

  // Activos promedio y ratio de actividad
  const promedioActivos = actividad.reduce((a,x)=>a + (x.activos||0),0) / 12;
  const ratioActivos = totalClientes > 0 ? (promedioActivos / totalClientes) : 0;

  // === ÚNICOS ===
  const ventaTotalUnicos = state.unicos.reduce((a,x)=>a+(x.venta||0),0);
  const abonoTotalUnicos = state.unicos.reduce((a,x)=>a+(x.abono||0),0);
  const mesesLeidos = state.unicos.length;
  const promMesUnicos = mesesLeidos ? (ventaTotalUnicos / mesesLeidos) : 0;

  // Comparación
  const diffVenta = ventaTotalClientes - ventaTotalUnicos;

  // === SCORES (para “semáforo”) ===
  // Cobranza: alto si pctCobrado >= 0.9, medio >= 0.75
  const scoreCobranza = clamp01((pctCobrado - 0.6) / 0.4);
  // Riesgo deuda concentrada: peor si concentracionDeuda alta
  const scoreConcentracionDeuda = 1 - clamp01((concentracionDeuda - 0.35) / 0.5);
  // Estabilidad: peor si cv alta
  const scoreEstabilidad = 1 - clamp01((cv - 0.35) / 0.9);
  // Dependencia ventas: peor si concentracionVenta alta
  const scoreDependencia = 1 - clamp01((concentracionVenta - 0.35) / 0.5);
  // Actividad base: mejor si ratioActivos alto
  const scoreBaseActiva = clamp01((ratioActivos - 0.15) / 0.45);

  const ahora = new Date();
  const fecha = ahora.toLocaleString("es-CL", { dateStyle:"medium", timeStyle:"short" });

  const filtroTexto = state.query?.trim()
    ? `Búsqueda: “${escapeHtml(state.query.trim())}”`
    : "Búsqueda: (sin filtro)";

  // === Texto de análisis basado en métricas (fundamentado) ===
  const bullets = [];

  // Cobranza
  bullets.push(`
    <li><b>Cobranza</b>: se ha cobrado <b>${(pctCobrado*100).toFixed(1)}%</b> del total anual.
    Esto se fundamenta en <b>Abono/Total</b> (${formatCLP(abonoTotalClientes)} / ${formatCLP(ventaTotalClientes)}).
    ${pctCobrado >= 0.9 ? "Nivel sano para operar sin presión de caja." : pctCobrado >= 0.75 ? "Nivel aceptable, pero conviene reforzar seguimiento de pagos." : "Nivel bajo: alto riesgo de caja, conviene plan de cobranza."}
    </li>
  `);

  // Deuda concentración
  bullets.push(`
    <li><b>Concentración de deuda</b>: el Top 10 explica <b>${(concentracionDeuda*100).toFixed(1)}%</b> de la deuda total (en valor absoluto).
    Fundamento: suma deuda Top10 / suma deuda total (solo deudas &lt; 0).
    ${concentracionDeuda >= 0.65 ? "Riesgo alto: con pocos clientes se te mueve toda la caja. Prioriza estos 10." : concentracionDeuda >= 0.45 ? "Riesgo medio: hay foco claro para cobrar." : "Riesgo bajo: deuda más distribuida, el problema es más “masivo”."}
    </li>
  `);

  // Dependencia por ventas
  bullets.push(`
    <li><b>Dependencia de ventas</b>: el Top 10 clientes aporta <b>${(concentracionVenta*100).toFixed(1)}%</b> del total anual.
    Fundamento: suma Total Top10 / Total anual.
    ${concentracionVenta >= 0.65 ? "Dependencia alta: perder 1–2 clientes impacta fuerte. Conveniente diversificar." : concentracionVenta >= 0.45 ? "Dependencia media: monitorea a los principales." : "Dependencia baja: buena distribución de ingresos."}
    </li>
  `);

  // Estacionalidad / estabilidad
  if (mesMayorIngreso && mesMenorIngreso) {
    bullets.push(`
      <li><b>Estacionalidad / estabilidad</b>: el mes con mayor ingreso fue <b>${escapeHtml(mesMayorIngreso.mes)}</b> (${formatCLP(mesMayorIngreso.totalMes)}) y el menor (con ingresos) fue <b>${escapeHtml(mesMenorIngreso.mes)}</b> (${formatCLP(mesMenorIngreso.totalMes)}).
      Coeficiente de variación (CV) aproximado: <b>${cv.toFixed(2)}</b> (0=estable, &gt;1=muy variable).
      ${cv <= 0.35 ? "Comportamiento bastante estable." : cv <= 0.8 ? "Variación moderada: planifica caja por ciclos." : "Variación alta: conviene controlar flujo y contratos para estabilizar."}
      </li>
    `);
  }

  // Base activa
  bullets.push(`
    <li><b>Base activa mensual</b>: en promedio hay <b>${promedioActivos.toFixed(1)}</b> clientes con ingreso &gt; 0 por mes (de ${totalClientes} totales).
    Ratio de actividad: <b>${(ratioActivos*100).toFixed(1)}%</b>.
    ${ratioActivos >= 0.45 ? "Buena base constante." : ratioActivos >= 0.25 ? "Base media: hay meses con poca actividad." : "Base baja: alta dependencia de pocos clientes/meses."}
    </li>
  `);

  // Únicos
  bullets.push(`
    <li><b>Clientes únicos</b>: hay ${mesesLeidos ? `<b>${mesesLeidos}</b> meses leídos` : "0 meses leídos"} en Visitas Únicas.
    Promedio mensual (únicos): <b>${formatCLP(promMesUnicos)}</b>.
    Diferencia vs anual (Clientes - Únicos): <b>${(diffVenta>=0?"+":"") + formatCLP(diffVenta)}</b>.
    Esto sirve para detectar si los “únicos” están incompletos o si están midiendo otra cosa.
    </li>
  `);

  const recomendaciones = [];
  if (pctCobrado < 0.9) recomendaciones.push("Implementar seguimiento semanal de cobranza (Top deuda primero) y recordatorios por fecha.");
  if (concentracionDeuda >= 0.55) recomendaciones.push("Plan específico para Top 10 deuda: acuerdos de pago, cortes de servicio, o reajuste de condiciones.");
  if (concentracionVenta >= 0.55) recomendaciones.push("Reducir dependencia: campaña para sumar nuevos contratos mensuales y paquetes estandarizados.");
  if (cv >= 0.8) recomendaciones.push("Estabilizar flujo: promover contratos mensuales y calendarizar servicios para suavizar la estacionalidad.");
  if (ratioActivos < 0.25) recomendaciones.push("Revisar cartera: clientes “inactivos” por meses; reactivación con ofertas, visitas preventivas o renovación de contrato.");

  if (!recomendaciones.length) {
    recomendaciones.push("Mantener control: revisar Top deuda y comportamiento mensual para anticipar cambios de caja.");
  }

  return `
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Resumen y Análisis - Reporte Financiero</title>
  <style>
    *{box-sizing:border-box}
    body{font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial; margin:24px; color:#0b1220}
    h1{margin:0 0 8px 0; font-size:20px}
    .meta{color:#555; font-size:12px; margin-bottom:16px}
    .grid{display:grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap:12px; margin:12px 0 18px}
    .card{border:1px solid #ddd; border-radius:12px; padding:12px}
    .label{color:#666; font-size:12px}
    .value{font-weight:800; font-size:18px; margin-top:6px}
    .hint{color:#666; font-size:12px; margin-top:4px}
    .pill{display:inline-block; padding:4px 8px; border-radius:999px; font-size:12px; border:1px solid #ddd; margin-right:6px}
    table{width:100%; border-collapse:collapse; margin-top:10px}
    th,td{border-bottom:1px solid #eee; padding:8px 6px; font-size:12px; text-align:left}
    th{color:#555}
    td.num, th.num{text-align:right; font-variant-numeric: tabular-nums}
    ul{margin:8px 0 0 18px}
    li{margin:8px 0; line-height:1.35}
    .section{margin-top:18px}
    @media print{ body{margin:12mm} }
  </style>
</head>
<body>
  <h1>Resumen y Análisis Ejecutivo</h1>
  <div class="meta">
    Generado: ${escapeHtml(fecha)} · ${filtroTexto} · Orden: ${escapeHtml(state.sort)}
  </div>

  <div class="section">
    <span class="pill"><b>Cobranza:</b> ${scoreLabel(scoreCobranza)} (${(pctCobrado*100).toFixed(1)}%)</span>
    <span class="pill"><b>Deuda concentrada:</b> ${scoreLabel(scoreConcentracionDeuda)} (${(concentracionDeuda*100).toFixed(1)}%)</span>
    <span class="pill"><b>Estabilidad:</b> ${scoreLabel(scoreEstabilidad)} (CV ${cv.toFixed(2)})</span>
    <span class="pill"><b>Dependencia:</b> ${scoreLabel(scoreDependencia)} (${(concentracionVenta*100).toFixed(1)}%)</span>
    <span class="pill"><b>Base activa:</b> ${scoreLabel(scoreBaseActiva)} (${(ratioActivos*100).toFixed(1)}%)</span>
  </div>

  <div class="section">
    <h2 style="margin:0 0 8px 0; font-size:14px;">KPIs principales</h2>
    <div class="grid">
      <div class="card">
        <div class="label">Venta total anual (Clientes con frecuencia)</div>
        <div class="value">${formatCLP(ventaTotalClientes)}</div>
        <div class="hint">Promedio mensual: ${formatCLP(promMesClientes)}</div>
      </div>
      <div class="card">
        <div class="label">Abono total anual</div>
        <div class="value">${formatCLP(abonoTotalClientes)}</div>
        <div class="hint">% cobrado: ${(pctCobrado*100).toFixed(1)}%</div>
      </div>
      <div class="card">
        <div class="label">Deuda total anual</div>
        <div class="value">${formatCLP(deudaTotalClientes)}</div>
        <div class="hint">Concentración deuda Top10: ${(concentracionDeuda*100).toFixed(1)}%</div>
      </div>
      <div class="card">
        <div class="label">Clientes (base)</div>
        <div class="value">${totalClientes}</div>
        <div class="hint">Promedio activos/mes: ${promedioActivos.toFixed(1)} (${(ratioActivos*100).toFixed(1)}%)</div>
      </div>
    </div>
  </div>

  <div class="section">
    <h2 style="margin:0 0 8px 0; font-size:14px;">Análisis ejecutivo (con fundamentos)</h2>
    <ul>
      ${bullets.join("")}
    </ul>
  </div>

  <div class="section">
    <h2 style="margin:0 0 8px 0; font-size:14px;">Recomendaciones accionables</h2>
    <ul>
      ${recomendaciones.map(r => `<li>${escapeHtml(r)}</li>`).join("")}
    </ul>
  </div>

  <div class="section">
    <h2 style="margin:0 0 8px 0; font-size:14px;">Top 10 clientes con mayor deuda</h2>
    <table>
      <thead>
        <tr><th>Cliente</th><th class="num">Diferencia</th></tr>
      </thead>
      <tbody>
        ${topDeuda.map(x => `
          <tr>
            <td>${escapeHtml(x.nombre)}</td>
            <td class="num">${formatCLP(x.diferencia || 0)}</td>
          </tr>`).join("")}
      </tbody>
    </table>
  </div>

  <div class="section">
    <h2 style="margin:0 0 8px 0; font-size:14px;">Actividad mensual (clientes con ingreso &gt; 0)</h2>
    <div class="meta">
      Mes con más activos: ${mesMasActivos ? `${escapeHtml(mesMasActivos.mes)} (${mesMasActivos.activos})` : "—"}
      · Mes con mayor ingreso: ${mesMayorIngreso ? `${escapeHtml(mesMayorIngreso.mes)} (${formatCLP(mesMayorIngreso.totalMes)})` : "—"}
    </div>
    <table>
      <thead>
        <tr>
          <th>Mes</th>
          <th class="num">Activos</th>
          <th class="num">Ingreso total</th>
          <th class="num">Promedio/activo</th>
        </tr>
      </thead>
      <tbody>
        ${actividad.map(r => `
          <tr>
            <td>${escapeHtml(r.mes)}</td>
            <td class="num">${r.activos}</td>
            <td class="num">${formatCLP(r.totalMes)}</td>
            <td class="num">${formatCLP(r.promedioActivo)}</td>
          </tr>`).join("")}
      </tbody>
    </table>
  </div>

  <div class="section">
    <h2 style="margin:0 0 8px 0; font-size:14px;">Clientes únicos (Visitas Únicas)</h2>
    <div class="grid">
      <div class="card">
        <div class="label">Venta total</div>
        <div class="value">${formatCLP(ventaTotalUnicos)}</div>
        <div class="hint">Meses leídos: ${mesesLeidos}</div>
      </div>
      <div class="card">
        <div class="label">Promedio mensual (únicos)</div>
        <div class="value">${formatCLP(promMesUnicos)}</div>
        <div class="hint">Comparación vs anual: ${(diffVenta>=0?"+":"") + formatCLP(diffVenta)}</div>
      </div>
    </div>
  </div>

</body>
</html>
  `;
}

function abrirResumenPdf() {
  const html = buildResumenHtml();

  // 1) Crear un iframe oculto
  const iframe = document.createElement("iframe");
  iframe.style.position = "fixed";
  iframe.style.right = "0";
  iframe.style.bottom = "0";
  iframe.style.width = "0";
  iframe.style.height = "0";
  iframe.style.border = "0";
  iframe.style.opacity = "0";
  iframe.setAttribute("aria-hidden", "true");
  document.body.appendChild(iframe);

  const doc = iframe.contentDocument || iframe.contentWindow.document;

  // 2) Escribir el HTML del resumen dentro del iframe
  doc.open();
  doc.write(html);
  doc.close();

  // 3) Esperar a que cargue y mandar a imprimir
  iframe.onload = () => {
    const w = iframe.contentWindow;

    // Algunos navegadores necesitan un pequeño delay
    setTimeout(() => {
      try {
        w.focus();
        w.print();
      } finally {
        // 4) Limpiar iframe cuando termine (o fallback)
        const cleanup = () => {
          iframe.remove();
          window.removeEventListener("afterprint", cleanup);
        };

        // afterprint del window principal suele funcionar mejor
        window.addEventListener("afterprint", cleanup);

        // fallback por si afterprint no dispara
        setTimeout(cleanup, 1500);
      }
    }, 50);
  };
}



function wireUI() {
  el("btnReload")?.addEventListener("click", init);
  el("btnPrint")?.addEventListener("click", () => window.print());

  el("btnExportCsv")?.addEventListener("click", () => exportClientesCSV(state.clientes));

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
  
  el("btnResumen")?.addEventListener("click", abrirResumenPdf);

}

wireUI();
init();
