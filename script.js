/**
 * Reporte financiero - lee /data/reporte.xlsx directamente (sin CSV)
 * - Detecta hojas por encabezados
 * - Renderiza KPIs + tablas + paginación + búsqueda + gráficos
 */

/**
 * IMPORTANTÍSIMO (GitHub Pages):
 * Si tu repo se llama "reporte-financiero", la página vive en:
 * https://calidadecosanplagas.github.io/reporte-financiero/
 *
 * Entonces los assets deben resolverse relativo a esa base.
 * Esto arma la ruta correcta aunque cambies de dominio/ruta.
 */
const BASE_PATH = window.location.pathname.endsWith("/")
  ? window.location.pathname
  : window.location.pathname + "/";

const EXCEL_URL = `${BASE_PATH}data/reporte.xlsx`; // <- carpeta data en el repo
// Si tu Excel está con otro nombre, cámbialo aquí.

let state = {
  unicos: [],       // [{mes, venta, abono, diferencia}]
  clientes: [],     // [{nombre,total,abono,diferencia, meses:{...}}]
  page: 1,
  pageSize: 25,
  query: "",
  sort: "nombre",
  charts: { c1: null, c2: null, c3: null }
};

const el = (id) => document.getElementById(id);

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

function detectSheets(workbook) {
  let detalleSheetName = null;
  let unicosSheetName = null;

  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows || rows.length === 0) continue;

    const headCandidates = rows.slice(0, 5).map(r => r.map(x => String(x).trim()));

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
  for (let i = 0; i < Math.min(10, rows.length); i++) {
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
  for (let i = 0; i < Math.min(10, rows.length); i++) {
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

function computeKPIs(unicos, clientes) {
  // === TOTALES REALES DESDE DETALLE CLIENTES ===
  const ventaTotal = clientes.reduce((a, c) => a + (c.total || 0), 0);
  const abonoTotal = clientes.reduce((a, c) => a + (c.abono || 0), 0);
  const deudaTotal = clientes.reduce((a, c) => a + (c.diferencia || 0), 0);

  // Conteo real de clientes
  const totalClientes = clientes.length;

  setText("kpiVentaTotal", formatCLP(ventaTotal));
  setText("kpiAbonoTotal", formatCLP(abonoTotal));
  setText("kpiDeudaTotal", formatCLP(deudaTotal));
  setText("kpiClientes", String(totalClientes));
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

function renderCharts(unicos, clientes) {
  destroyCharts();

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
          { label: "Venta", data: ventas },
          { label: "Abono", data: abonos },
        ]
      },
      options: { responsive: true, plugins: { legend: { position: "bottom" } } }
    });
  }

  const ctx2 = el("chartUnicosDiferencia");
  if (ctx2) {
    state.charts.c2 = new Chart(ctx2, {
      type: "bar",
      data: { labels: labelsUnicos, datasets: [{ label: "Diferencia", data: difs }] },
      options: { responsive: true, plugins: { legend: { position: "bottom" } } }
    });
  }

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
      data: { labels: labelsTop, datasets: [{ label: "Diferencia (Deuda)", data: valoresTop }] },
      options: {
        responsive: true,
        plugins: { legend: { position: "bottom" } },
        scales: { x: { ticks: { autoSkip: false, maxRotation: 0 } } }
      }
    });
  }
}

async function loadExcel() {
  const res = await fetch(EXCEL_URL, { cache: "no-store" });
  if (!res.ok) {
    throw new Error(`No se pudo cargar ${EXCEL_URL}. Verifica que exista en el repo: /data/reporte.xlsx`);
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

  const clientes = parseDetalleClientes(wb.Sheets[detalleSheetName]);
  const unicos = parseUnicos(wb.Sheets[unicosSheetName]);

  return { clientes, unicos };
}

async function init() {
  try {
    setText("kpiClientes", "…");
    setText("kpiVentaTotal", "…");
    setText("kpiAbonoTotal", "…");
    setText("kpiDeudaTotal", "…");

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
