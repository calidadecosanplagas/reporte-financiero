/**
 * Reporte financiero - lee /data/reporte.xlsx directamente (sin CSV)
 * - Detecta hojas por encabezados
 * - Renderiza KPIs + tablas + paginación + búsqueda + gráficos
 * - NUEVO: frecuencia mensual desde Detalle Clientes (clientes con ingreso > 0 y total mes)
 */

const EXCEL_URL = "./data/reporte.xlsx";

let state = {
  unicos: [],       // [{mes, venta, abono, diferencia}]
  clientes: [],     // [{nombre,total,abono,diferencia, meses:{...}}]
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

/* ===================== DETECT + PARSE ===================== */

function detectSheets(workbook) {
  let detalleSheetName = null;
  let unicosSheetName = null;

  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows || rows.length === 0) continue;

    const headCandidates = rows.slice(0, 8).map(r => r.map(x => String(x).trim()));

    const hasNombreCliente = headCandidates.some(r => r.includes("Nombre Cliente"));
    const hasTotal = headCandidates.some(r => r.includes("Total"));
    const hasEnero = headCandidates.some(r => r.includes("Enero"));
    const hasDiferencia = headCandidates.some(r => r.includes("Diferencia"));

    if (hasNombreCliente && hasTotal && hasEnero && hasDiferencia) {
      detalleSheetName = name;
      continue;
    }

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
  for (let i = 0; i < Math.min(12, rows.length); i++) {
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
  for (let i = 0; i < Math.min(12, rows.length); i++) {
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

/* ===================== KPIs ===================== */

function computeKPIs(unicos, clientes) {
  // Totales ANUALES desde Detalle Clientes (lo real del año)
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
  setText("kpiMostrando", `${Math.min(totalClientes, totalClientes)} / ${totalClientes}`);

  // NUEVO: promedio mensual del año (clientes con frecuencia)
  setText("kpiPromMesClientes", formatCLP(ventaTotalClientes / 12));

  // Totales desde Visitas Únicas (mensual, puede no ser 12 meses)
  const ventaTotalUnicos = unicos.reduce((a, x) => a + x.venta, 0);
  const abonoTotalUnicos = unicos.reduce((a, x) => a + x.abono, 0);
  const mesesLeidos = unicos.length;

  setText("kpiVentaTotalUnicos", formatCLP(ventaTotalUnicos));
  setText("kpiAbonoTotalUnicos", formatCLP(abonoTotalUnicos));
  setText("kpiMesesUnicos", String(mesesLeidos));

  // coherencia “simple”
  const coherencia = (mesesLeidos > 0) ? "OK" : "Sin datos";
  setText("kpiCoherenciaUnicos", coherencia);

  // Comparación venta anual (clientes - unicos)
  const diffVenta = ventaTotalClientes - ventaTotalUnicos;
  const sign = diffVenta >= 0 ? "+" : "";
  setText("kpiComparacionVenta", `${sign}${formatCLP(diffVenta)}`);
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
      out.sort((a,b) => (a.diferencia||0) - (b.diferencia||0)); break;
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

    return { mes: m, activos, totalMes };
  });
}

function renderActividadMensualTable(rows) {
  clearTable("tablaActividadMensual");
  for (const r of rows) {
    addRow("tablaActividadMensual", [
      { text: r.mes },
      { text: String(r.activos), className: "num" },
      { text: formatCLP(r.totalMes), className: "num" },
    ]);
  }
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

function chartMoneyTicks(v) {
  const n = Number(v);
  if (!isFinite(n)) return v;
  return formatCLP(n).replace("$", ""); // deja números con puntos, sin $
}

/* ===================== CHARTS ===================== */

function renderCharts(unicos, clientes) {
  destroyCharts();

  // Unicos
  const labelsUnicos = unicos.map(x => x.mes);
  const ventas = unicos.map(x => x.venta);
  const abonos = unicos.map(x => x.abono);
  const difs = unicos.map(x => x.diferencia);

  const ctx1 = el("chartUnicosVentaAbono");
  if (ctx1) {
    state.charts.c1 = new Chart(ctx1, {
      type: "line",
      data: {
        labels: labelsUnicos,
        datasets: [
          { label: "Venta", data: ventas, tension: 0.25, pointRadius: 3 },
          { label: "Abono", data: abonos, tension: 0.25, pointRadius: 3 },
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: "bottom" } },
        scales: { y: { ticks: { callback: (v) => chartMoneyTicks(v) } } }
      }
    });
  }

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
        plugins: { legend: { position: "bottom" } },
        scales: { y: { ticks: { callback: (v) => chartMoneyTicks(v) } } }
      }
    });
  }

  // Top deuda (más negativo primero)
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
        plugins: { legend: { position: "bottom" } },
        scales: {
          x: { ticks: { autoSkip: false, maxRotation: 0 } },
          y: { ticks: { callback: (v) => chartMoneyTicks(v) } }
        }
      }
    });
  }

  // Actividad mensual desde clientes (Ingreso total por mes)
  const act = computeActividadMensualDesdeClientes(clientes);
  renderActividadMensualTable(act);

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
        plugins: { legend: { position: "bottom" } },
        scales: { y: { ticks: { callback: (v) => chartMoneyTicks(v) } } }
      }
    });
  }
}

/* ===================== LOAD EXCEL ===================== */

async function loadExcel() {
  const res = await fetch(EXCEL_URL, { cache: "no-store" });
  if (!res.ok) {
    throw new Error(`No se pudo cargar ${EXCEL_URL}. ¿Está en /data y se llama reporte.xlsx?`);
  }
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
    setText("kpiVentaTotalClientes", "…");
    setText("kpiAbonoTotalClientes", "…");
    setText("kpiDeudaTotalClientes", "…");
    setText("kpiClientes", "…");
    setText("kpiSinDeuda", "…");
    setText("kpiConDeuda", "…");
    setText("kpiPctCobrado", "…");
    setText("kpiMostrando", "…");
    setText("kpiPromMesClientes", "…");

    setText("kpiVentaTotalUnicos", "…");
    setText("kpiAbonoTotalUnicos", "…");
    setText("kpiMesesUnicos", "…");
    setText("kpiCoherenciaUnicos", "…");
    setText("kpiComparacionVenta", "…");

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

function wireUI() {
  el("btnReload")?.addEventListener("click", init);
  el("btnPrint")?.addEventListener("click", () => window.print());

  el("btnExportCsv")?.addEventListener("click", () => {
    exportClientesCSV(state.clientes);
  });

  el("searchCliente")?.addEventListener("input", (e) => {
    state.query = e.target.value || "";
    state.page = 1;
    renderClientesTable(state.clientes);

    // actualiza “Mostrando”
    const filtered = applyFiltersAndSort(state.clientes);
    setText("kpiMostrando", `${filtered.length} / ${state.clientes.length}`);
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
}

wireUI();
init();
