/**
 * Reporte financiero - lee /data/reporte.xlsx directamente (sin CSV)
 * - Detecta hojas por encabezados
 * - Renderiza KPIs + tablas + paginación + búsqueda + gráficos
 * - KPIs anuales reales desde Detalle Clientes
 * - KPIs anuales de Visitas Únicas (extra)
 */

const EXCEL_PATH = "./data/reporte.xlsx"; // OJO: GitHub Pages es case-sensitive

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
  }
};

const el = (id) => document.getElementById(id);

/** Normaliza texto para comparar encabezados (trim + espacios + upper) */
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

  // Limpia $ y espacios
  s = s.replace(/\$/g, "").replace(/\s+/g, "");

  // Si viene con puntos de miles: 1.234.567
  // y/o coma decimal: 1.234,56 (no debería, pero por si acaso)
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
  // Busca:
  // - Hoja "Detalle clientes" por encabezado "Nombre Cliente" + meses + "Total" + "Abono" + "Diferencia"
  // - Hoja "Visitas únicas" por encabezado "Mes" o "Nombre" + "Venta" + "Abono" + "Diferencia"
  let detalleSheetName = null;
  let unicosSheetName = null;

  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows || rows.length === 0) continue;

    const detalleHeaders = ["Nombre Cliente", "Enero", "Total", "Abono", "Diferencia"];
    const unicosHeadersA = ["Mes", "Venta", "Abono", "Diferencia"];
    const unicosHeadersB = ["Nombre", "Venta", "Abono", "Diferencia"];

    const detalleRow = findHeaderRow(rows, detalleHeaders, 12);
    if (detalleRow !== -1) {
      detalleSheetName = name;
      continue;
    }

    const unicosRowA = findHeaderRow(rows, unicosHeadersA, 12);
    const unicosRowB = findHeaderRow(rows, unicosHeadersB, 12);
    if (unicosRowA !== -1 || unicosRowB !== -1) {
      unicosSheetName = name;
      continue;
    }
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

/** KPIs:
 * - Anual real desde Detalle Clientes
 * - Anual visitas únicas (extra) desde la hoja de Unicos
 */
function computeKPIs(unicos, clientes) {
  // KPI anual REAL desde Detalle Clientes
  const ventaTotal = clientes.reduce((a, x) => a + (x.total || 0), 0);
  const abonoTotal = clientes.reduce((a, x) => a + (x.abono || 0), 0);
  const deudaTotal = clientes.reduce((a, x) => a + (x.diferencia || 0), 0);
  const totalClientes = clientes.length;

  setText("kpiVentaTotal", formatCLP(ventaTotal));
  setText("kpiAbonoTotal", formatCLP(abonoTotal));
  setText("kpiDeudaTotal", formatCLP(deudaTotal));
  setText("kpiClientes", String(totalClientes));

  // KPIs de Visitas Únicas (si existen IDs en el HTML)
  const unicosVenta = unicos.reduce((a, x) => a + (x.venta || 0), 0);
  const unicosAbono = unicos.reduce((a, x) => a + (x.abono || 0), 0);
  const unicosDif = unicos.reduce((a, x) => a + (x.diferencia || 0), 0);

  setText("kpiUnicosVentaTotal", formatCLP(unicosVenta));
  setText("kpiUnicosAbonoTotal", formatCLP(unicosAbono));
  setText("kpiUnicosDeudaTotal", formatCLP(unicosDif));
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

function applyFiltersAndSort(clientes) {
  let out = [...clientes];

  const q = state.query.trim().toLowerCase();
  if (q) out = out.filter(c => c.nombre.toLowerCase().includes(q));

  switch (state.sort) {
    case "total_desc":
      out.sort((a,b) => b.total - a.total); break;
    case "abono_desc":
      out.sort((a,b) => b.abono - a.abono); break;
    case "diferencia_asc":
      out.sort((a,b) => a.diferencia - b.diferencia); break;
    default:
      out.sort((a,b) => a.nombre.localeCompare(b.nombre, "es"));
  }

  return out;
}

function renderClientesTable(clientes) {
  clearTable("tablaClientes");

  const filtered = applyFiltersAndSort(clientes);

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
    ]);
  }

  setText("pageInfo", `Página ${state.page} de ${totalPages} · Mostrando ${pageRows.length} de ${filtered.length}`);
}

function destroyCharts() {
  for (const k of ["c1","c2","c3"]) {
    if (state.charts[k]) {
      state.charts[k].destroy();
      state.charts[k] = null;
    }
  }
}

function chartMoneyTicks(value) {
  // Formato CLP compacto para ejes
  const n = Number(value) || 0;
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return (n / 1_000_000).toFixed(1).replace(".", ",") + "M";
  if (abs >= 1_000) return (n / 1_000).toFixed(0) + "k";
  return String(Math.round(n));
}

function renderCharts(unicos, clientes) {
  destroyCharts();

  const labelsUnicos = unicos.map(x => x.mes);
  const ventas = unicos.map(x => x.venta);
  const abonos = unicos.map(x => x.abono);
  const difs = unicos.map(x => x.diferencia);

  // ---- Chart 1: Line Venta vs Abono (Unicos)
  const ctx1 = el("chartUnicosVentaAbono");
  if (ctx1) {
    state.charts.c1 = new Chart(ctx1, {
      type: "line",
      data: {
        labels: labelsUnicos,
        datasets: [
          { label: "Venta", data: ventas, tension: 0.25, pointRadius: 2 },
          { label: "Abono", data: abonos, tension: 0.25, pointRadius: 2 },
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: { position: "bottom" },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatCLP(ctx.parsed.y)}`
            }
          }
        },
        scales: {
          y: {
            ticks: {
              callback: (v) => chartMoneyTicks(v)
            }
          }
        }
      }
    });
  }

  // ---- Chart 2: Bar Diferencia (Unicos)
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
          legend: { position: "bottom" },
          tooltip: {
            callbacks: {
              label: (ctx) => `Diferencia: ${formatCLP(ctx.parsed.y)}`
            }
          }
        },
        scales: {
          y: {
            ticks: { callback: (v) => chartMoneyTicks(v) }
          }
        }
      }
    });
  }

  // ---- Chart 3: Top 10 Deuda (clientes con diferencia más negativa)
  const topDeuda = [...clientes]
    .filter(c => typeof c.diferencia === "number")
    .sort((a,b) => a.diferencia - b.diferencia)
    .slice(0, 10);

  const labelsTop = topDeuda.map(x => x.nombre);
  const valoresTop = topDeuda.map(x => x.diferencia);

  const ctx3 = el("chartTopDeuda");
  if (ctx3) {
    state.charts.c3 = new Chart(ctx3, {
      type: "bar",
      data: {
        labels: labelsTop,
        datasets: [{ label: "Diferencia (Deuda)", data: valoresTop }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom" },
          tooltip: {
            callbacks: {
              label: (ctx) => `Deuda: ${formatCLP(ctx.parsed.y)}`
            }
          }
        },
        scales: {
          x: {
            ticks: {
              autoSkip: false,
              maxRotation: 35,
              minRotation: 0,
              callback: function(value) {
                // corta nombres largos en el eje X
                const label = this.getLabelForValue(value);
                return label.length > 18 ? label.slice(0, 18) + "…" : label;
              }
            }
          },
          y: {
            ticks: { callback: (v) => chartMoneyTicks(v) }
          }
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

  if (!detalleSheetName) {
    throw new Error("No encontré la hoja de Detalle Clientes. Debe tener encabezado: 'Nombre Cliente', meses, 'Total', 'Abono', 'Diferencia'.");
  }
  if (!unicosSheetName) {
    throw new Error("No encontré la hoja de Visitas Únicas. Debe tener encabezado: 'Mes' o 'Nombre' + 'Venta' + 'Abono' + 'Diferencia'.");
  }

  const wsDetalle = wb.Sheets[detalleSheetName];
  const wsUnicos = wb.Sheets[unicosSheetName];

  const clientes = parseDetalleClientes(wsDetalle);
  const unicos = parseUnicos(wsUnicos);

  return { clientes, unicos };
}

async function init() {
  try {
    setText("kpiClientes", "…");
    setText("kpiVentaTotal", "…");
    setText("kpiAbonoTotal", "…");
    setText("kpiDeudaTotal", "…");

    // extras (si existen)
    setText("kpiUnicosVentaTotal", "…");
    setText("kpiUnicosAbonoTotal", "…");
    setText("kpiUnicosDeudaTotal", "…");

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
}

wireUI();
init();
