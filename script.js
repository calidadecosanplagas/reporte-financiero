/*************************************************
 * UTILIDADES
 *************************************************/
const CLP = (n) => "$" + Math.trunc(n).toLocaleString("es-CL");

function toNumber(v) {
  if (v === undefined || v === null) return 0;
  const s = String(v).trim();
  if (s === "" || s === "$") return 0;
  const clean = s.replace(/\./g, "").replace(/,/g, "");
  const n = Number(clean);
  return Number.isFinite(n) ? n : 0;
}

/*************************************************
 * RESUMEN MENSUAL – CLIENTES VISITA ÚNICA
 *************************************************/
const resumenMensual = [
  { mes: "Enero", venta: 3842560, abono: 2807600 },
  { mes: "Febrero", venta: 2658900, abono: 1279200 },
  { mes: "Marzo", venta: 2696560, abono: 2536568 },
  { mes: "Abril", venta: 1908750, abono: 1645950 },
  { mes: "Mayo", venta: 1541200, abono: 1248400 },
  { mes: "Junio - Julio", venta: 1586940, abono: 1456940 },
  { mes: "Agosto - Septiembre - Octubre", venta: 3240300, abono: 2813200 },
  { mes: "Noviembre", venta: 2680650, abono: 1308500 },
  { mes: "Diciembre", venta: 537100, abono: 90000 }
];

/*************************************************
 * DETALLE CLIENTES (PEGADO TAL CUAL DESDE EXCEL)
 *************************************************/
const DETALLE_CLIENTES_TSV = `
Nombre Cliente	Enero	Febrero	Marzo	Abril	Mayo	Junio	Julio	Agosto	Septiembre	Octubre	Noviembre	Diciembre	Total	Abono	Diferencia
Afp Plan Vital	59423	0	0	60439	0	0	60735	0	0	61140	0	0	241737	241742	5
EL COMINO	681288	681471	687585	958907	694017	695092	696385	694053	700119	700119	910141	0	8099177	7189036	-910141
AGRÍCOLA PRAVIA	91361	91838	92520	92984	93259	93430	93282	93596	93975	94206	94352		1024803	1024803	0
AGRÍCOLA PRIMOS Z	95966	96317	97071	97504	97839	98017	98162	97936	98675	98687	99069	0	1075243	0	-1075243
SOCIEDAD AGRICOLA DON ENRIQUE	0	0	0	55669	0	0	0	55963	0	0	0	0	111632	0	-111632
COMBUSTIBLES PEUMO	0	0	$	55605	0	0	0	56174	0	0	0	0	111779	0	-111779
IVAN OLEA GARCIA	0	54966	$		0	55995	0	0	0	56429	0	0	167390	0	-167390
AGRICOLA SANTA ISABEL SPA	0	374850	$	226100	0	226100	0	226100	0	226100	0	0	1279250	1053150	-226100
AGROCRECES	205734	205831	207823	208589	209657	210037	210347	209680	211201	211445	212127	212292	2514763	1667698	-847065
AL TIRO PIZZA	59381	59555	60069	60239	60559	60661	60774	60609	61066	61084	84654	61305	749956	542913	-207043
APÍCOLA SAN VICENTE	73093	73556	74016	74375	74602	74739	74655	74726	75181	75316	75482		819741	295040	-524701
`.trim();

/*************************************************
 * PARSER TSV → OBJETOS
 *************************************************/
function parseDetalle(tsv) {
  const lines = tsv.split(/\r?\n/).filter(l => l.trim() !== "");
  const header = lines[0].split("\t");

  const idxNombre = header.indexOf("Nombre Cliente");
  const idxTotal = header.indexOf("Total");
  const idxAbono = header.indexOf("Abono");
  const idxDif = header.indexOf("Diferencia");

  return lines.slice(1).map(line => {
    const cols = line.split("\t");
    const total = toNumber(cols[idxTotal]);
    const abono = toNumber(cols[idxAbono]);
    const diferencia = toNumber(cols[idxDif]);

    return {
      nombre: (cols[idxNombre] || "").trim(),
      total,
      abono,
      diferencia
    };
  }).filter(r => r.nombre);
}

const clientes = parseDetalle(DETALLE_CLIENTES_TSV);

/*************************************************
 * RENDER RESUMEN MENSUAL
 *************************************************/
const tablaMensual = document.getElementById("tablaMensual");
resumenMensual.forEach(r => {
  const tr = document.createElement("tr");
  const diff = r.venta - r.abono;
  tr.innerHTML = `
    <td>${r.mes}</td>
    <td class="num">${CLP(r.venta)}</td>
    <td class="num">${CLP(r.abono)}</td>
    <td class="num deuda">${CLP(diff)}</td>
  `;
  tablaMensual.appendChild(tr);
});

/*************************************************
 * RENDER CLIENTES + KPIs
 *************************************************/
let totalVentas = 0;
let totalAbonos = 0;

const tablaClientes = document.getElementById("tablaClientes");

clientes.forEach(c => {
  totalVentas += c.total;
  totalAbonos += c.abono;

  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td>${c.nombre}</td>
    <td class="num">${CLP(c.total)}</td>
    <td class="num">${CLP(c.abono)}</td>
    <td class="num deuda">${CLP(c.diferencia)}</td>
  `;
  tablaClientes.appendChild(tr);
});

/*************************************************
 * KPIs
 *************************************************/
document.getElementById("kpiVentas").textContent = CLP(totalVentas);
document.getElementById("kpiAbonos").textContent = CLP(totalAbonos);
document.getElementById("kpiDeuda").textContent = CLP(totalVentas - totalAbonos);
document.getElementById("kpiClientes").textContent = clientes.length;
