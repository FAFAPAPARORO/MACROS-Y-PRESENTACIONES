// Genera el Análisis Financiero de JURIDICA ABOGADOS Y CONSULTORES S.A.S.
// con el formato gráfico ACONTIS (mismo estilo del informe Santa Vicenta).
//
// Uso:  npm run juridica      (o)   node presentaciones/juridica/generar_juridica.js
//
// Para actualizar un mes:
//   1. Cifras: cambiar los valores EN PESOS del bloque DATOS (estado de resultados
//      mes a mes, balance de los dos meses comparados y detalle de gastos).
//      Tablas, tarjetas, gráficos, márgenes, indicadores y el puente de utilidad
//      se calculan solos a partir de esas cifras.
//   2. Textos: ajustar las lecturas del bloque TEXTOS (historia del mes, factores,
//      conclusiones...), que son redacción de análisis.
// Los gráficos son nativos de PowerPoint y las tablas son tablas reales, así que
// también se pueden editar directamente en el .pptx.

const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");
const fs = require("fs");

const ROOT = path.resolve(__dirname, "..", "..");
const ASSETS = path.join(ROOT, "assets", "acontis");

// ───────────────────────────── DATOS (pesos colombianos) ─────────────────────────────
const DATOS = {
  archivo: "Analisis_Financiero_Juridica_Agosto_2026.pptx",
  empresa: "JURIDICA ABOGADOS Y CONSULTORES S.A.S.",
  empresaNombre: "Juridica Abogados y Consultores S.A.S.",
  nit: "900.944.440-3",
  preparadoPor: "Preparado por Asesorías Contables del Caribe S.A.S. (ACONTIS)",
  anio: 2026,
  corteLargo: "Corte a 31 de agosto de 2026",

  // Meses del año con cierre. El último es el mes actual y el penúltimo el mes de comparación.
  meses: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto"],

  // Estado de resultado integral, mes a mes (fuente: EEFF a 31 de agosto de 2026).
  er: {
    ingresosBrutos: [68143710, 73029822, 75134467, 79319558, 96131805, 109170837, 91410054, 68648289],
    devoluciones: [4982250, 2135250, 711750, 5800763, 4987233, 7003620, 16398000, 0],
    // Total costos de operación. Julio y agosto incluyen la mano de obra indirecta
    // (asesoría jurídica, intangibles, gastos de viaje: $5.885.521 y $7.087.256) que
    // el subtotal mensual del PDF omite pero que sí está en el acumulado ($163.689.638).
    costos: [19428771, 19301954, 19684195, 20258395, 22188262, 18151432, 21345703, 23330926],
    gastosAdmon: [28644316, 36206194, 19657780, 53519267, 23905588, 98058387, 37671013, 40769091],
    ingresosNoOp: [14479.12, 28512.09, 19642.15, 63242452.82, 26514219.89, 102633816.89, 32436521.91, 35543021.99],
    gastosNoOp: [735455, 603190, 726548, 16545842, 940840, 35490180, 21226129, 11035461],
    impuesto: [0, 13151000, 0, 27993000, 0, 28617000, 0, 22869000],
    utilidadAcumuladaReportada: 197347425, // control: debe coincidir con la suma de los meses
  },

  // Estado de situación financiera de los dos meses comparados.
  balance: {
    act: { ac: 1075010530, anc: 939078918, pc: 188850155, pnc: 5985894, pat: 1819253399, resultado: 197347425 },
    ant: { ac: 1040544452, anc: 932524991, pc: 154055635, pnc: 6310348, pat: 1812703461, resultado: 190797487 },
  },

  // Gastos de administración y ventas: los 6 rubros más grandes del mes [nombre, mes actual, mes anterior]
  gastosTop: [
    ["Gastos Unión Temporal", 16769915, 13261581],
    ["Gastos de personal", 5834065, 8491308],
    ["Honorarios", 3306279, 3306279],
    ["Contribuciones y afiliaciones", 2566854, 2636966],
    ["Otros gastos diversos", 2457524, 636761],
    ["Depreciaciones", 2234749, 2234749],
  ],
};

// ───────────────────────────── TEXTOS DE ANÁLISIS ─────────────────────────────
const TEXTOS = {
  historia: [
    "Los ingresos netos bajaron 8,5% en agosto y la utilidad neta cayó 77,3%, hasta $6,2 M.",
    "La causa principal es el pago bimestral del impuesto SIMPLE ($22,9 M), que julio no tuvo: antes de impuestos agosto ganó $29,1 M, 6,8% más que julio, gracias a un mejor resultado neto de la Unión Temporal.",
  ],
  puente: [
    ["Operación: −$8,3 M. ", "Caen los ingresos netos ($6,4 M) y sube el costo de operación ($2,0 M)."],
    ["Gastos admon.: −$3,1 M. ", "Sube la Unión Temporal (+$3,5 M) y bajan los gastos de personal ($2,7 M menos)."],
    ["No operacional: +$13,3 M. ", "Suben los ingresos ($3,1 M) y bajan los gastos no operacionales ($10,2 M)."],
    ["Impuesto SIMPLE: −$22,9 M. ", "Pago bimestral que julio no tuvo."],
  ],
  tendencia: "Ingresos netos acumulados de enero a agosto. Tras el pico de junio ($102,2 M), los ingresos bajaron en julio (−26,6%) y agosto (−8,5%).",
  composicion: "En agosto no hubo devoluciones (julio: $16,4 M, el mayor nivel del año), lo que amortiguó la caída del ingreso bruto. Actividades de consultoría · cifras en millones de pesos.",
  concentracion: [
    { text: "Los gastos de la Unión Temporal ($16,8 M) son el " },
    { text: "41% del gasto del mes", options: { bold: true, color: "navy" } },
    { text: " y subieron $3,5 M, más que el aumento total ($3,1 M); los gastos de personal bajaron $2,7 M." },
  ],
  factoresTitulo: "TRES FACTORES DETRÁS DEL RESULTADO DE AGOSTO",
  factores: [
    ["01", "Pago bimestral del régimen SIMPLE", "$0", "$22,9 M",
      "El SIMPLE se paga cada dos meses (febrero, abril, junio, agosto). Julio no tuvo pago; agosto absorbió el bimestre completo y explica la mayor parte de la caída."],
    ["02", "Menores ingresos de consultoría", "$91,4 M", "$68,6 M",
      "Los ingresos brutos cayeron 24,9%. Sin devoluciones en agosto, el ingreso neto bajó solo 8,5%, pero el margen bruto pasó de 71,5% a 66,0%."],
    ["03", "Unión Temporal EVB Abogados", "−$1,4 M neto", "+$8,2 M neto",
      "En agosto la UT aportó $35,5 M de ingreso no operacional frente a $27,3 M de gastos asociados ($16,8 M administrativos y $10,5 M no operacionales)."],
  ],
  lecturaClave: "antes de impuestos agosto ganó más que julio (+6,8%). La caída de la utilidad neta la explica el SIMPLE; el punto a vigilar es la operación: ingresos brutos −24,9% y resultado operacional de solo $4,5 M.",
  balance: "El activo corriente creció $34,5 M, impulsado por bancos (+$43,9 M); el pasivo corriente subió $34,8 M por el anticipo del SIMPLE por pagar ($44,6 M).",
  conclusiones: [
    ["Ingresos netos −8,5% en agosto ", "($68,6 M vs $75,0 M en julio): los ingresos brutos cayeron 24,9%, compensados en parte porque no hubo devoluciones. El acumulado del año suma $619,0 M."],
    ["Utilidad neta −77,3% ($6,2 M): ", "la explica el pago bimestral del SIMPLE ($22,9 M); antes de impuestos agosto ganó $29,1 M, 6,8% más que julio."],
    ["Resultado operacional de $4,5 M (margen 6,6%): ", "menor ingreso, costo de operación +9,3% y gastos de administración +8,2%, impulsados por la Unión Temporal."],
    ["Liquidez sólida pero menor: ", "razón corriente de 5,69 (vs 6,75) y endeudamiento de 9,7% (vs 8,1%), por el anticipo del SIMPLE por pagar ($44,6 M)."],
    ["Patrimonio +0,4% y activo total +2,1%: ", "caja y bancos suben a $146,9 M (+43%) y la utilidad acumulada del año llega a $197,3 M."],
  ],
};

// ───────────────────────────── CÁLCULOS ─────────────────────────────
const E = DATOS.er;
const N = DATOS.meses.length, iA = N - 1, iB = N - 2;
const MES_A = DATOS.meses[iA], MES_B = DATOS.meses[iB];
const mesA = MES_A.toLowerCase(), mesB = MES_B.toLowerCase();
const abrev = DATOS.meses.map((m) => m.slice(0, 3));
const serie = (f) => DATOS.meses.map((_, i) => f(i));
const R = {};
R.netos = serie((i) => E.ingresosBrutos[i] - E.devoluciones[i]);
R.bruta = serie((i) => R.netos[i] - E.costos[i]);
R.oper = serie((i) => R.bruta[i] - E.gastosAdmon[i]);
R.rai = serie((i) => R.oper[i] + E.ingresosNoOp[i] - E.gastosNoOp[i]);
R.neta = serie((i) => R.rai[i] - E.impuesto[i]);
R.mb = serie((i) => (R.bruta[i] / R.netos[i]) * 100);
R.mo = serie((i) => (R.oper[i] / R.netos[i]) * 100);
R.mn = serie((i) => (R.neta[i] / R.netos[i]) * 100);
const suma = (a) => a.reduce((x, y) => x + y, 0);
if (Math.abs(suma(R.neta) - E.utilidadAcumuladaReportada) > 5) {
  console.warn(`⚠ La suma de utilidades mensuales (${Math.round(suma(R.neta))}) no cuadra con el acumulado reportado (${E.utilidadAcumuladaReportada}).`);
}
const B = DATOS.balance;
const bal = (b) => Object.assign({}, b, { at: b.ac + b.anc, pt: b.pc + b.pnc });
const BA = bal(B.act), BB = bal(B.ant);
BA.ptp = BA.pt + BA.pat; BB.ptp = BB.pt + BB.pat;

// Formato colombiano
const num = (v, d = 1) => {
  const [e, f] = Math.abs(v).toFixed(d).split(".");
  return (v < 0 && +Math.abs(v).toFixed(d) !== 0 ? "-" : "") + e.replace(/\B(?=(\d{3})+(?!\d))/g, ".") + (f ? "," + f : "");
};
const mm = (v) => num(v / 1e6, 1);
const money = (v) => (v < 0 ? "-$" : "$") + num(Math.abs(v) / 1e6, 1) + " M";
const sgn = (v, d = 1) => (+v.toFixed(d) >= 0 ? "+" : "−") + num(Math.abs(v), d);
const varM = (a, b) => sgn((a - b) / 1e6);
const varP = (a, b) => (b === 0 ? "Nuevo" : sgn(((a - b) / Math.abs(b)) * 100) + "%");
const fav = (a, b, sentido) => (Math.abs(a - b) < 1 ? null : (a - b) * sentido > 0);
const flecha = (a, b) => (a >= b ? "▲ " : "▼ ");

const CORTE = `Corte: ${mesA} ${DATOS.anio}`;
const COMPARA = `${MES_B} vs ${mesA} ${DATOS.anio}`;

// ───────────────────────────── TEMA ─────────────────────────────
const THEME = {
  name: "ACONTIS",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A2233", lt1: "FFFFFF", dk2: "011E66", lt2: "F3F7FB",
    accent1: "011E66", // azul marino
    accent2: "00AFBC", // turquesa
    accent3: "50C1E4", // azul claro
    accent4: "056875", // verde petróleo
    accent5: "E8A13D", // ámbar
    accent6: "C4453B", // rojo
    hlink: "00AFBC", folHlink: "056875",
  },
};
const HEX = {
  navy: "011E66", teal: "00AFBC", sky: "50C1E4", sea: "056875", amber: "E8A13D", red: "C4453B",
  ink: "1A2233", muted: "5A6572", line: "DDE6EF", card: "F3F7FB", grid: "E6ECF2", pale: "7DD1EB",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13,33" × 7,5"
pres.author = "ACONTIS";
pres.company = "Asesorías Contables del Caribe S.A.S.";
pres.title = "Análisis Financiero · " + DATOS.empresaNombre;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

const C = pres.SchemeColor;
const COL = {
  navy: C.accent1, teal: C.accent2, sky: C.accent3, sea: C.accent4, amber: C.accent5, red: C.accent6,
  ink: C.text1, white: C.background1, card: C.background2, muted: "5A6572", line: "DDE6EF", pale: "7DD1EB",
};
const tono = (f) => (f === null || f === undefined ? COL.muted : f ? COL.teal : COL.red);

// ───────────────────────────── DISEÑOS (layouts) ─────────────────────────────
const fondoOscuro = () => [
  { image: { x: 0, y: 0, w: 13.333, h: 7.5, path: path.join(ASSETS, "fondo_oficina.jpg"), transparency: 65 } },
  { image: { x: 10.2, y: 6.2, w: 2.72, h: 0.94, path: path.join(ASSETS, "logo_acontis_fondo_oscuro.png") } },
];

pres.defineSlideMaster({
  title: "PORTADA",
  background: { color: HEX.navy },
  objects: fondoOscuro(),
});

pres.defineSlideMaster({
  title: "SEPARADOR",
  background: { color: HEX.navy },
  objects: [
    ...fondoOscuro(),
    { placeholder: { options: { name: "numero", type: "body", x: 0.9, y: 1.6, w: 3, h: 1.1, fontSize: 80, bold: true, color: COL.white, margin: 0, valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "antetitulo", type: "body", x: 0.9, y: 2.95, w: 10, h: 0.55, fontSize: 30, color: COL.sky, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 3.55, w: 11, h: 0.95, fontSize: 52, bold: true, color: COL.white, margin: 0, align: "left" }, text: "" } },
    { placeholder: { options: { name: "subtitulo", type: "body", x: 0.9, y: 4.6, w: 10, h: 0.45, fontSize: 22, bold: true, color: COL.sky, margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 11.55, y: 7.05, w: 1.2, h: 0.3, fontSize: 10, color: COL.sky, align: "right" },
});

pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "seccion", type: "body", x: 0.6, y: 0.33, w: 9.8, h: 0.3, fontSize: 12, bold: true, color: COL.sea, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 10.2, h: 0.62, fontSize: 30, bold: true, color: COL.navy, margin: 0, valign: "middle", align: "left" }, text: "" } },
    { image: { x: 0.55, y: 6.86, w: 1.3, h: 0.45, path: path.join(ASSETS, "logo_acontis_fondo_claro.png") } },
    { text: { text: "Página", options: { x: 10.9, y: 7.03, w: 1.35, h: 0.3, fontSize: 10, color: COL.muted, align: "right", valign: "middle", margin: 0 } } },
  ],
  slideNumber: { x: 12.27, y: 7.03, w: 0.45, h: 0.3, fontSize: 10, color: COL.muted, align: "left", valign: "middle", margin: [0, 0, 0, 0.05] },
});

// ───────────────────────────── AYUDANTES ─────────────────────────────
let nObj = 0;
const id = (n) => `${n} ${++nObj}`;

function contenido(seccion, seccionPptx, titulo) {
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: seccionPptx });
  s.addText(seccion, { placeholder: "seccion" });
  s.addText(titulo, { placeholder: "title" });
  // Píldora "Corte"
  s.addText(CORTE, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.21, objectName: id("Corte"),
    x: 10.85, y: 0.45, w: 1.87, h: 0.42, fill: { color: COL.card }, line: { color: COL.line, width: 1 },
    fontSize: 12, bold: true, color: COL.sea, align: "center", valign: "middle", margin: 0,
  });
  return s;
}

function tarjeta(s, x, y, w, h, opts = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, objectName: id(opts.name || "Tarjeta"),
    fill: { color: opts.fill || COL.card }, line: { color: opts.line || COL.line, width: 1 },
  });
}

// Evita que una cifra quede separada de su unidad ("63,6" / "M") al hacer salto de línea.
const nb = (t) => t.replace(/(\d) (M|veces|pp)\b/g, "$1 $2");
// Permite usar nombres de color ("navy", "teal"...) en los textos de TEXTOS.
const runs = (t) => t.map((r) => (r.options && COL[r.options.color] ? Object.assign({}, r, { options: Object.assign({}, r.options, { color: COL[r.options.color] }) }) : r));
function texto(s, t, o) {
  t = typeof t === "string" ? nb(t) : runs(t).map((r) => Object.assign({}, r, { text: nb(r.text) }));
  s.addText(t, Object.assign({ isTextBox: true, margin: 0, color: COL.ink, fontSize: 13, valign: "top" }, o, { objectName: id(o.name || "Texto") }));
}

function circulo(s, x, y, d, label, size = 18) {
  s.addText(label, {
    shape: pres.shapes.OVAL, x, y, w: d, h: d, fill: { color: COL.navy }, line: { color: COL.navy, width: 0 },
    fontSize: size, bold: true, color: COL.white, align: "center", valign: "middle", margin: 0, objectName: id("Número"),
  });
}

// Celdas de tabla con el estilo ACONTIS
const TH = (t, o = {}) => ({ text: t, options: Object.assign({ fill: { color: COL.navy }, color: COL.white, bold: true, fontSize: 11, align: "center", valign: "middle" }, o) });
const TD = (t, o = {}) => ({ text: t, options: Object.assign({ color: COL.ink, fontSize: 12, valign: "middle" }, o) });
const BORDE = () => ({ type: "solid", pt: 1, color: "FFFFFF" });
const zebra = (i) => ({ color: i % 2 ? "F3F7FB" : "FFFFFF" });

function banda(texto, n) {
  return [{ text: texto, options: { colspan: n, fill: { color: COL.sea }, color: COL.white, bold: true, fontSize: 11, valign: "middle" } }];
}

const ejeTexto = () => ({
  catAxisLabelColor: HEX.muted, valAxisLabelColor: HEX.muted, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" },
  catAxisLineShow: true, catAxisLineColor: HEX.line, valAxisLineShow: false,
  titleColor: HEX.navy, titleFontSize: 13, legendColor: HEX.ink, legendFontSize: 11,
});

// ───────────────────────────── 1. PORTADA ─────────────────────────────
pres.addSection({ title: "Portada" });
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Portada" });
  texto(s, "Análisis", { name: "Título 1", x: 0.9, y: 1.25, w: 10, h: 1.0, fontSize: 80, bold: true, color: COL.white, valign: "middle" });
  texto(s, "Financiero", { name: "Título 2", x: 0.9, y: 2.25, w: 10, h: 1.0, fontSize: 80, bold: true, color: COL.pale, valign: "middle" });
  texto(s, DATOS.empresaNombre, { name: "Empresa", x: 0.9, y: 3.6, w: 11, h: 0.5, fontSize: 32, bold: true, color: COL.white, valign: "middle" });
  texto(s, `NIT ${DATOS.nit}  -  Cifras en pesos colombianos`, { name: "NIT", x: 0.9, y: 4.18, w: 11, h: 0.35, fontSize: 18, bold: true, color: COL.pale, valign: "middle" });
  texto(s, `${DATOS.corteLargo} · Comparativo mensual ${mesB} – ${mesA} ${DATOS.anio}`, { name: "Corte", x: 0.9, y: 4.75, w: 11.5, h: 0.5, fontSize: 28, color: COL.white, valign: "middle" });
  texto(s, DATOS.preparadoPor, { name: "Preparado por", x: 0.9, y: 6.55, w: 8.5, h: 0.3, fontSize: 13, color: COL.pale, valign: "middle" });
}

// ───────────────────────────── 2. ÍNDICE ─────────────────────────────
{
  const s = contenido("ÍNDICE", "Portada", "CONTENIDO");
  const items = [
    ["Resumen ejecutivo", "Cifras clave y la historia del mes"],
    ["Estado de resultados", "Cuadro de mando, puente de utilidad y márgenes"],
    ["Ingresos y gastos", "De ingreso bruto a neto y estructura del gasto"],
    ["Análisis de causas", `Tres factores detrás del resultado de ${mesA}`],
    ["Situación financiera", "Balance comparativo e indicadores clave"],
    ["Conclusiones", "Hallazgos del mes y prioridades"],
  ];
  items.forEach(([t, sub], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.17, y = 1.6 + row * 1.7, w = 5.96, h = 1.4;
    tarjeta(s, x, y, w, h, { name: "Índice " + (i + 1) });
    circulo(s, x + 0.3, y + 0.39, 0.62, String(i + 1), 22);
    texto(s, t, { name: "Índice título", x: x + 1.15, y: y + 0.3, w: w - 1.35, h: 0.45, fontSize: 24, bold: true, color: COL.navy, valign: "middle" });
    texto(s, sub, { name: "Índice detalle", x: x + 1.15, y: y + 0.78, w: w - 1.35, h: 0.32, fontSize: 14, color: COL.muted, valign: "middle" });
  });
}

// ───────────────────────────── 3. RESUMEN EJECUTIVO ─────────────────────────────
pres.addSection({ title: "1. Resumen ejecutivo" });
{
  const s = contenido("1 · RESUMEN EJECUTIVO", "1. Resumen ejecutivo", `RESUMEN EJECUTIVO - ${COMPARA.toUpperCase()}`);
  const kpi = (t, a, b, sentido) => ({
    t, v: money(a), l1: `${MES_B}: ${money(b)}`, l2: `${flecha(a, b)}${varP(a, b)} vs. ${mesB}`, c2: tono(fav(a, b, sentido)),
  });
  const kpis = [
    kpi("INGRESOS NETOS", R.netos[iA], R.netos[iB], 1),
    kpi("ACTIVO TOTAL", BA.at, BB.at, 1),
    kpi("PATRIMONIO", BA.pat, BB.pat, 1),
    { t: "UTILIDAD NETA", v: money(R.neta[iA]), l1: `Margen neto ${num(R.mn[iA])}%`, l2: `${MES_B}: ${money(R.neta[iB])}  (${varP(R.neta[iA], R.neta[iB])})`, c2: COL.white, dark: true },
  ];
  kpis.forEach((k, i) => {
    const x = 0.6 + i * 3.08, y = 1.5, w = 2.88, h = 1.95;
    tarjeta(s, x, y, w, h, { name: "KPI " + k.t, fill: k.dark ? COL.navy : COL.card, line: k.dark ? COL.navy : COL.line });
    texto(s, k.t, { name: "KPI etiqueta", x: x + 0.25, y: y + 0.25, w: w - 0.4, h: 0.3, fontSize: 13, bold: true, color: k.dark ? COL.white : COL.muted });
    texto(s, k.v, { name: "KPI valor", x: x + 0.25, y: y + 0.58, w: w - 0.4, h: 0.6, fontSize: 32, bold: true, color: k.dark ? COL.white : COL.navy, valign: "middle" });
    texto(s, k.l1, { name: "KPI detalle", x: x + 0.25, y: y + 1.24, w: w - 0.4, h: 0.26, fontSize: 12, color: k.dark ? COL.sky : COL.muted });
    texto(s, k.l2, { name: "KPI variación", x: x + 0.25, y: y + 1.52, w: w - 0.4, h: 0.26, fontSize: 12, bold: true, color: k.c2 });
  });

  const cats = ["Ingresos netos", "Utilidad bruta", "Gastos admon. y ventas", "Utilidad neta"];
  const val = (i) => [R.netos[i], R.bruta[i], E.gastosAdmon[i], R.neta[i]].map((v) => +(v / 1e6).toFixed(1));
  s.addChart(pres.charts.BAR, [
    { name: `${MES_A} ${DATOS.anio}`, labels: cats, values: val(iA) },
    { name: `${MES_B} ${DATOS.anio}`, labels: cats, values: val(iB) },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 3.7, w: 7.6, h: 3.05, objectName: "Gráfico comparativo",
    barDir: "col", barGapWidthPct: 60, chartColors: [HEX.navy, HEX.teal],
    showTitle: true, title: `${MES_B} vs ${mesA}, en millones de pesos`,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"#,##0.0', dataLabelFontSize: 10, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    valAxisLabelFormatCode: "#,##0", valAxisHidden: true, valGridLine: { style: "none" },
    showLegend: true, legendPos: "b",
  }));

  tarjeta(s, 8.45, 3.7, 4.27, 3.05, { name: "Lectura del mes" });
  texto(s, "La historia del mes", { name: "Lectura título", x: 8.75, y: 3.95, w: 3.7, h: 0.38, fontSize: 18, bold: true, color: COL.sea });
  texto(s, TEXTOS.historia.map((t, i) => ({ text: t, options: { breakLine: i < TEXTOS.historia.length - 1, paraSpaceAfter: 8 } })),
    { name: "Lectura texto", x: 8.75, y: 4.42, w: 3.7, h: 2.2, fontSize: 13, color: COL.ink });
}

// ───────────────────────────── 4. SEPARADOR · ESTADO DE RESULTADOS ─────────────────────────────
function separador(num, ante, titulo, sub, seccionPptx) {
  pres.addSection({ title: seccionPptx });
  const s = pres.addSlide({ masterName: "SEPARADOR", sectionTitle: seccionPptx });
  s.addText(num, { placeholder: "numero" });
  s.addText(ante, { placeholder: "antetitulo" });
  s.addText(titulo, { placeholder: "title" });
  s.addText(sub, { placeholder: "subtitulo" });
  return s;
}
separador("2", "Análisis comparativo", "Estado de Resultados", `${COMPARA.toUpperCase()} · EVOLUCIÓN DEL AÑO`, "2. Estado de resultados");

// ───────────────────────────── 5. CUADRO DE MANDO ─────────────────────────────
{
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", "CUADRO DE MANDO - COMPARATIVO INTEGRAL");
  // [concepto, actual, anterior, sentido (+1 subir es favorable / -1 desfavorable), negrita]
  const er = [
    ["Ingresos netos", R.netos, 1, true],
    ["Utilidad bruta", R.bruta, 1, true],
    ["Gastos de administración y ventas", E.gastosAdmon, -1, false],
    ["Resultado operacional", R.oper, 1, true],
    ["Resultado antes de impuesto", R.rai, 1, false],
    ["Utilidad neta", R.neta, 1, true],
  ].map(([c, a, sen, neg]) => [c, a[iA], a[iB], sen, neg]);
  const esf = [
    ["Activo total", BA.at, BB.at, 1, true],
    ["Pasivo total", BA.pt, BB.pt, -1, true],
    ["Patrimonio", BA.pat, BB.pat, 1, true],
  ];
  const fila = ([concepto, a, b, sen, neg], i) => {
    const fill = zebra(i), f = fav(a, b, sen);
    return [
      TD(neg ? concepto.toUpperCase() : concepto, { fill, bold: neg, color: neg ? COL.navy : COL.ink }),
      TD(mm(a), { fill, bold: true, color: COL.navy, align: "right" }),
      TD(mm(b), { fill, color: COL.muted, align: "right" }),
      TD(varM(a, b), { fill, bold: true, color: tono(f), align: "right" }),
      TD(varP(a, b), { fill, bold: true, color: tono(f), align: "right" }),
    ];
  };
  const filas = [
    [TH("CONCEPTO", { align: "left" }), TH(`${MES_A.toUpperCase()} ${DATOS.anio}`), TH(`${MES_B.toUpperCase()} ${DATOS.anio}`, { fill: { color: COL.ink } }), TH("VAR. $"), TH("VAR. %")],
    banda("ESTADO DE RESULTADO INTEGRAL", 5),
    ...er.map(fila),
    banda("ESTADO DE SITUACIÓN FINANCIERA", 5),
    ...esf.map(fila),
  ];
  s.addTable(filas, {
    x: 0.6, y: 1.5, w: 12.13, colW: [4.53, 1.9, 1.9, 1.9, 1.9], rowH: 0.38,
    border: BORDE(), margin: [0, 0.1, 0, 0.1], objectName: "Tabla cuadro de mando",
  });
  texto(s, "Cifras en millones de pesos colombianos.   Código de color:  turquesa = favorable  ·  rojo = desfavorable", {
    name: "Nota de color", x: 0.6, y: 6.35, w: 12.13, h: 0.3, fontSize: 11, color: COL.muted, align: "right",
  });
}

// ───────────────────────────── 6. PUENTE DE UTILIDAD ─────────────────────────────
{
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", `PUENTE DE UTILIDAD - DE ${MES_B.toUpperCase()} A ${MES_A.toUpperCase()}`);
  // Pasos del puente: [rubro, efecto sobre la utilidad (pesos)]
  const pasos = [
    ["Ingresos netos", R.netos[iA] - R.netos[iB]],
    ["Costo de operación", -(E.costos[iA] - E.costos[iB])],
    ["Gastos admon. y ventas", -(E.gastosAdmon[iA] - E.gastosAdmon[iB])],
    ["Ingresos no operacionales", E.ingresosNoOp[iA] - E.ingresosNoOp[iB]],
    ["Gastos no operacionales", -(E.gastosNoOp[iA] - E.gastosNoOp[iB])],
    ["Impuesto SIMPLE", -(E.impuesto[iA] - E.impuesto[iB])],
  ];
  // Serie "Base" invisible + series visibles = cascada editable en Excel.
  const r1 = (v) => +(v / 1e6).toFixed(1);
  const cats = [`Utilidad ${mesB}`, ...pasos.map((p) => p[0]), `Utilidad ${mesA}`];
  const base = [0], total = [r1(R.neta[iB])], sube = [0], baja = [0];
  let acum = R.neta[iB];
  pasos.forEach(([_, d]) => {
    const fin = acum + d;
    base.push(r1(Math.min(acum, fin))); total.push(0);
    sube.push(d > 0 ? r1(d) : 0); baja.push(d < 0 ? r1(-d) : 0);
    acum = fin;
  });
  base.push(0); total.push(r1(R.neta[iA])); sube.push(0); baja.push(0);
  s.addChart(pres.charts.BAR, [
    { name: "Base", labels: cats, values: base },
    { name: "Utilidad", labels: cats, values: total },
    { name: "Aumenta", labels: cats, values: sube },
    { name: "Disminuye", labels: cats, values: baja },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 8.25, h: 5.25, objectName: "Gráfico puente de utilidad",
    barDir: "col", barGrouping: "stacked", barGapWidthPct: 45,
    chartColors: ["FFFFFF", HEX.navy, HEX.teal, HEX.red],
    showTitle: true, title: `Utilidad neta ${mesB} → ${mesA}, en millones de pesos`,
    showValue: true, dataLabelPosition: "inEnd", dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelColor: "FFFFFF",
    dataLabelFormatCode: '"$"#,##0.0;;', catAxisLabelFontSize: 10,
    valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false,
  }));

  tarjeta(s, 9.1, 1.45, 3.62, 5.25, { name: "Lectura del puente" });
  texto(s, "Lectura del puente", { name: "Puente título", x: 9.35, y: 1.7, w: 3.15, h: 0.38, fontSize: 18, bold: true, color: COL.sea });
  const puntos = TEXTOS.puente;
  texto(s, puntos.flatMap(([b, t], i) => [
    { text: b, options: { bold: true, color: COL.navy } },
    { text: t, options: { breakLine: i < puntos.length - 1, paraSpaceAfter: 10 } },
  ]), { name: "Puente texto", x: 9.35, y: 2.2, w: 3.15, h: 4.3, fontSize: 13 });
}

// ───────────────────────────── 7. EVOLUCIÓN DEL AÑO ─────────────────────────────
{
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", `EVOLUCIÓN MENSUAL - ENERO A ${MES_A.toUpperCase()} ${DATOS.anio}`);
  s.addChart(pres.charts.BAR, [
    { name: "Ingresos netos", labels: abrev, values: R.netos.map((v) => +(v / 1e6).toFixed(1)) },
    { name: "Utilidad neta", labels: abrev, values: R.neta.map((v) => +(v / 1e6).toFixed(1)) },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 7.6, h: 5.25, objectName: "Gráfico evolución mensual",
    barDir: "col", barGapWidthPct: 50, chartColors: [HEX.navy, HEX.teal],
    showTitle: true, title: "Ingresos y utilidad neta mensual (millones de pesos)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0.0", dataLabelFontSize: 9, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    valAxisLabelFormatCode: "#,##0", showLegend: true, legendPos: "b",
  }));

  const sumI = suma(R.netos), sumU = suma(R.neta);
  const filas = [[TH("MES"), TH("INGRESOS"), TH("UTILIDAD"), TH("MARGEN")]];
  abrev.forEach((m, i) => {
    const fill = zebra(i), mg = R.mn[i];
    filas.push([
      TD(m, { fill }), TD(mm(R.netos[i]), { fill, align: "right" }), TD(mm(R.neta[i]), { fill, align: "right" }),
      TD(num(mg) + "%", { fill, align: "right", bold: true, color: mg >= 20 ? COL.teal : COL.amber }),
    ]);
  });
  const tot = { fill: { color: "F3F7FB" }, bold: true, color: COL.navy };
  filas.push([
    TD("Total", tot), TD(mm(sumI), Object.assign({ align: "right" }, tot)), TD(mm(sumU), Object.assign({ align: "right" }, tot)),
    TD(num((sumU / sumI) * 100) + "%", Object.assign({ align: "right" }, tot)),
  ]);
  const rowH = 0.32;
  s.addTable(filas, { x: 8.45, y: 1.5, w: 4.27, colW: [0.85, 1.17, 1.17, 1.08], rowH, border: BORDE(), margin: [0, 0.08, 0, 0.08], objectName: "Tabla evolución" });

  const yT = 1.5 + rowH * filas.length + 0.3, hT = 6.7 - yT;
  tarjeta(s, 8.45, yT, 4.27, hT, { name: "Tendencia" });
  texto(s, money(sumI), { name: "Tendencia valor", x: 8.7, y: yT + 0.15, w: 3.8, h: 0.55, fontSize: 30, bold: true, color: COL.teal, valign: "middle" });
  texto(s, TEXTOS.tendencia, { name: "Tendencia texto", x: 8.7, y: yT + 0.75, w: 3.8, h: hT - 0.85, fontSize: 12, color: COL.ink });
}

// ───────────────────────────── 8. MÁRGENES ─────────────────────────────
{
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", "RENTABILIDAD - MÁRGENES DEL AÑO");
  const r1 = (a) => a.map((v) => +v.toFixed(1));
  const minV = Math.floor(Math.min(0, ...R.mo, ...R.mn) / 20) * 20;
  s.addChart(pres.charts.LINE, [
    { name: "Margen bruto", labels: abrev, values: r1(R.mb) },
    { name: "Margen operacional", labels: abrev, values: r1(R.mo) },
    { name: "Margen neto", labels: abrev, values: r1(R.mn) },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 8.6, h: 5.25, objectName: "Gráfico márgenes",
    chartColors: [HEX.navy, HEX.teal, HEX.amber], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 8,
    showTitle: true, title: `% sobre ingresos netos · enero a ${mesA} ${DATOS.anio}`,
    valAxisLabelFormatCode: '0"%"', valAxisMinVal: minV, valAxisMaxVal: 100, valAxisMajorUnit: 20,
    showLegend: true, legendPos: "b",
  }));
  const tarjetaMargen = (t, a, b) => {
    const d = a - b;
    const c = a < 0 ? COL.red : d >= 0 ? COL.teal : COL.amber;
    return [t, num(a) + "%", `${d >= 0 ? "▲" : "▼"} ${sgn(d)} pp  ·  ${mesB} ${num(b)}%`, c];
  };
  const tarjetas = [
    tarjetaMargen("MARGEN BRUTO", R.mb[iA], R.mb[iB]),
    tarjetaMargen("MARGEN OPERACIONAL", R.mo[iA], R.mo[iB]),
    tarjetaMargen("MARGEN NETO", R.mn[iA], R.mn[iB]),
  ];
  tarjetas.forEach(([t, v, d, c], i) => {
    const x = 9.45, y = 1.45 + i * 1.82, w = 3.27, h = 1.6;
    tarjeta(s, x, y, w, h, { name: "KPI " + t });
    texto(s, t, { name: "KPI etiqueta", x, y: y + 0.18, w, h: 0.3, fontSize: 13, bold: true, color: COL.muted, align: "center" });
    texto(s, v, { name: "KPI valor", x, y: y + 0.5, w, h: 0.62, fontSize: 36, bold: true, color: c, align: "center", valign: "middle" });
    texto(s, d, { name: "KPI detalle", x, y: y + 1.16, w, h: 0.28, fontSize: 12, color: COL.muted, align: "center" });
  });
}

// ───────────────────────────── 9. SEPARADOR · INGRESOS Y GASTOS ─────────────────────────────
separador("3", "Composición y estructura", "Ingresos y Gastos", "DE INGRESO BRUTO A NETO · GASTOS DE ADMINISTRACIÓN Y VENTAS", "3. Ingresos y gastos");

// ───────────────────────────── 10. COMPOSICIÓN DE INGRESOS ─────────────────────────────
{
  const s = contenido("3 · INGRESOS Y GASTOS", "3. Ingresos y gastos", "DE INGRESO BRUTO A INGRESO NETO");
  const dev = (v) => (v === 0 ? "$0,0 M" : "-" + money(v));
  const filas = [iB, iA].map((i) => [DATOS.meses[i].toUpperCase(), String(DATOS.anio), [money(E.ingresosBrutos[i]), dev(E.devoluciones[i]), money(R.netos[i])]]);
  const cols = [
    { x: 2.35, w: 2.95, t: "INGRESOS BRUTOS" },
    { x: 6.0, w: 2.95, t: "(−) DEVOLUCIONES" },
    { x: 9.65, w: 3.07, t: "INGRESOS NETOS", dark: true },
  ];
  filas.forEach(([mes, anio, vals], r) => {
    const y = 1.55 + r * 1.5, h = 1.25;
    texto(s, mes, { name: "Mes", x: 0.6, y: y + 0.18, w: 1.7, h: 0.5, fontSize: 28, bold: true, color: COL.navy, valign: "middle" });
    texto(s, anio, { name: "Año", x: 0.6, y: y + 0.68, w: 1.6, h: 0.35, fontSize: 18, color: COL.muted, valign: "middle" });
    cols.forEach((c, i) => {
      tarjeta(s, c.x, y, c.w, h, { name: c.t + " " + mes, fill: c.dark ? COL.navy : COL.card, line: c.dark ? COL.navy : COL.line });
      texto(s, c.t, { name: "Etiqueta", x: c.x + 0.25, y: y + 0.2, w: c.w - 0.4, h: 0.28, fontSize: 12, bold: true, color: c.dark ? COL.sky : COL.muted });
      texto(s, vals[i], { name: "Valor", x: c.x + 0.25, y: y + 0.52, w: c.w - 0.4, h: 0.55, fontSize: 30, bold: true, color: c.dark ? COL.white : i === 1 ? COL.red : COL.navy, valign: "middle" });
      if (i < 2) {
        s.addShape(pres.shapes.RIGHT_ARROW, { x: c.x + c.w + 0.17, y: y + 0.45, w: 0.38, h: 0.35, fill: { color: COL.teal }, line: { color: COL.teal, width: 0 }, objectName: id("Flecha") });
      }
    });
  });
  const variacion = (t, a, b, sentido) => {
    const d = a - b;
    return [t, `${flecha(a, b)}${d >= 0 ? "+" : "−"}${money(Math.abs(d))}`, `${varP(a, b)} vs. ${mesB}`, tono(fav(a, b, sentido))];
  };
  const vars = [
    variacion("VAR. INGRESOS BRUTOS", E.ingresosBrutos[iA], E.ingresosBrutos[iB], 1),
    variacion("VAR. DEVOLUCIONES", E.devoluciones[iA], E.devoluciones[iB], -1),
    variacion("VAR. INGRESOS NETOS", R.netos[iA], R.netos[iB], 1),
  ];
  vars.forEach(([t, v, d, c], i) => {
    const x = cols[i].x, w = cols[i].w, y = 4.65, h = 1.45;
    tarjeta(s, x, y, w, h, { name: t });
    texto(s, t, { name: "Etiqueta", x: x + 0.25, y: y + 0.18, w: w - 0.4, h: 0.28, fontSize: 12, bold: true, color: COL.muted });
    texto(s, v, { name: "Valor", x: x + 0.25, y: y + 0.48, w: w - 0.4, h: 0.5, fontSize: 26, bold: true, color: c, valign: "middle" });
    texto(s, d, { name: "Detalle", x: x + 0.25, y: y + 1.02, w: w - 0.4, h: 0.28, fontSize: 12, color: COL.muted });
  });
  texto(s, "VARIACIÓN", { name: "Var etiqueta", x: 0.6, y: 5.1, w: 1.6, h: 0.5, fontSize: 20, bold: true, color: COL.sea, valign: "middle" });
  texto(s, TEXTOS.composicion, { name: "Nota", x: 0.6, y: 6.3, w: 12.13, h: 0.35, fontSize: 12, italic: true, color: COL.muted });
}

// ───────────────────────────── 11. GASTOS TOP 6 ─────────────────────────────
{
  const s = contenido("3 · INGRESOS Y GASTOS", "3. Ingresos y gastos", "GASTOS DE ADMINISTRACIÓN Y VENTAS - TOP 6");
  const G = DATOS.gastosTop;
  s.addChart(pres.charts.BAR, [
    { name: `${MES_A} ${DATOS.anio}`, labels: G.map((g) => g[0]), values: G.map((g) => +(g[1] / 1e6).toFixed(1)) },
    { name: `${MES_B} ${DATOS.anio}`, labels: G.map((g) => g[0]), values: G.map((g) => +(g[2] / 1e6).toFixed(1)) },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 7.0, h: 5.25, objectName: "Gráfico gastos top 6",
    barDir: "bar", barGapWidthPct: 45, chartColors: [HEX.navy, HEX.teal], catAxisOrientation: "maxMin",
    showTitle: true, title: `${MES_B} vs ${mesA}, en millones de pesos`,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0.0", dataLabelFontSize: 10, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    valAxisHidden: true, valGridLine: { style: "none" }, showLegend: true, legendPos: "b",
  }));
  const filas = [[TH("CONCEPTO", { align: "left" }), TH(MES_A.toUpperCase()), TH(MES_B.toUpperCase(), { fill: { color: COL.ink } }), TH("VAR. %")]];
  G.forEach(([c, a, b], i) => {
    const fill = zebra(i);
    filas.push([
      TD(c, { fill, fontSize: 11 }), TD(mm(a), { fill, bold: true, color: COL.navy, align: "right" }),
      TD(mm(b), { fill, color: COL.muted, align: "right" }),
      TD(varP(a, b), { fill, bold: true, align: "right", color: tono(fav(a, b, -1)) }),
    ]);
  });
  const ta = E.gastosAdmon[iA], tb = E.gastosAdmon[iB], fl = { color: "F3F7FB" };
  filas.push([
    TD("TOTAL GASTOS ADMON. Y VENTAS", { fill: fl, bold: true, color: COL.navy, fontSize: 10 }),
    TD(mm(ta), { fill: fl, bold: true, color: COL.navy, align: "right" }),
    TD(mm(tb), { fill: fl, color: COL.muted, align: "right" }),
    TD(varP(ta, tb), { fill: fl, bold: true, color: tono(fav(ta, tb, -1)), align: "right" }),
  ]);
  s.addTable(filas, { x: 7.85, y: 1.5, w: 4.87, colW: [2.2, 0.85, 0.85, 0.97], rowH: 0.38, border: BORDE(), margin: [0, 0.08, 0, 0.08], objectName: "Tabla gastos" });

  tarjeta(s, 7.85, 4.75, 4.87, 1.6, { name: "Concentración" });
  texto(s, "Concentración", { name: "Concentración título", x: 8.1, y: 4.92, w: 4.4, h: 0.38, fontSize: 18, bold: true, color: COL.sea });
  texto(s, TEXTOS.concentracion, { name: "Concentración texto", x: 8.1, y: 5.38, w: 4.4, h: 0.9, fontSize: 13 });
  texto(s, "Código de color:  turquesa = favorable  ·  rojo = desfavorable", { name: "Nota de color", x: 7.85, y: 6.45, w: 4.87, h: 0.25, fontSize: 10, color: COL.muted, align: "right" });
}

// ───────────────────────────── 12. TRES FACTORES ─────────────────────────────
pres.addSection({ title: "4. Análisis de causas" });
{
  const s = contenido("4 · ANÁLISIS DE CAUSAS", "4. Análisis de causas", TEXTOS.factoresTitulo);
  TEXTOS.factores.forEach(([n, t, vb, va, d], i) => {
    const x = 0.6 + i * 4.115, y = 1.5, w = 3.9, h = 3.95;
    tarjeta(s, x, y, w, h, { name: "Factor " + n });
    circulo(s, x + 0.25, y + 0.25, 0.6, n, 18);
    texto(s, t.toUpperCase(), { name: "Factor título", x: x + 1.0, y: y + 0.2, w: w - 1.2, h: 0.7, fontSize: 14, bold: true, color: COL.navy, valign: "middle" });
    // Mini comparativo mes anterior / mes actual
    [[MES_B.toUpperCase(), vb, COL.muted], [MES_A.toUpperCase(), va, COL.navy]].forEach(([l, v, c], j) => {
      const bx = x + 0.25 + j * 1.75, by = y + 1.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w: 1.65, h: 0.85, rectRadius: 0.08, fill: { color: COL.white }, line: { color: COL.line, width: 1 }, objectName: id("Factor " + l) });
      texto(s, l, { name: "Factor mes", x: bx, y: by + 0.08, w: 1.65, h: 0.25, fontSize: 11, bold: true, color: COL.muted, align: "center" });
      texto(s, v, { name: "Factor valor", x: bx, y: by + 0.36, w: 1.65, h: 0.4, fontSize: v.length > 8 ? 16 : 20, bold: true, color: c, align: "center", valign: "middle" });
    });
    texto(s, d, { name: "Factor detalle", x: x + 0.25, y: y + 2.15, w: w - 0.5, h: 1.7, fontSize: 14, color: COL.ink });
  });
  tarjeta(s, 0.6, 5.65, 12.13, 1.05, { name: "Lectura clave", fill: COL.navy, line: COL.navy });
  texto(s, [
    { text: "Lectura clave: ", options: { bold: true, color: COL.sky } },
    { text: TEXTOS.lecturaClave },
  ], { name: "Lectura clave texto", x: 0.9, y: 5.65, w: 11.6, h: 1.05, fontSize: 15, color: COL.white, valign: "middle" });
}

// ───────────────────────────── 13. SEPARADOR · SITUACIÓN FINANCIERA ─────────────────────────────
separador("5", "Estado de Situación Financiera", "Balance e Indicadores", "LIQUIDEZ · ENDEUDAMIENTO · RENTABILIDAD", "5. Situación financiera");

// ───────────────────────────── 14. BALANCE COMPARATIVO ─────────────────────────────
{
  const s = contenido("5 · SITUACIÓN FINANCIERA", "5. Situación financiera", "BALANCE GENERAL COMPARATIVO");
  // [cuenta, clave, sentido, tipo]
  const filasBal = [
    ["ACTIVO", null, 0, "banda"],
    ["Activo corriente", "ac", 1, "fila"],
    ["Activo no corriente", "anc", 1, "fila"],
    ["TOTAL ACTIVO", "at", 1, "total"],
    ["PASIVO", null, 0, "banda"],
    ["Pasivo corriente", "pc", -1, "fila"],
    ["Pasivo no corriente", "pnc", -1, "fila"],
    ["TOTAL PASIVO", "pt", -1, "total"],
    ["PATRIMONIO", null, 0, "banda"],
    ["TOTAL PATRIMONIO", "pat", 1, "total"],
    ["TOTAL PASIVO + PATRIMONIO", "ptp", 1, "total"],
  ];
  const filas = [[TH("CUENTA", { align: "left" }), TH(`${MES_A.toUpperCase()} ${DATOS.anio}`), TH(`${MES_B.toUpperCase()} ${DATOS.anio}`, { fill: { color: COL.ink } }), TH("VAR. $"), TH("VAR. %")]];
  let k = 0;
  filasBal.forEach(([r, key, sen, tipo]) => {
    if (tipo === "banda") { filas.push(banda(r, 5)); k = 0; return; }
    const tot = tipo === "total", a = BA[key], b = BB[key], f = fav(a, b, sen);
    const fill = tot ? { color: "F3F7FB" } : zebra(k++);
    filas.push([
      TD(r, { fill, bold: tot, color: tot ? COL.navy : COL.ink, fontSize: tot ? 11 : 12 }),
      TD(mm(a), { fill, bold: true, color: COL.navy, align: "right" }),
      TD(mm(b), { fill, color: COL.muted, align: "right" }),
      TD(varM(a, b), { fill, bold: true, color: tono(f), align: "right" }),
      TD(varP(a, b), { fill, bold: true, color: tono(f), align: "right" }),
    ]);
  });
  s.addTable(filas, { x: 0.6, y: 1.5, w: 7.45, colW: [2.85, 1.15, 1.15, 1.1, 1.2], rowH: 0.37, border: BORDE(), margin: [0, 0.08, 0, 0.08], objectName: "Tabla balance" });

  tarjeta(s, 0.6, 6.05, 7.45, 0.65, { name: "Lectura balance" });
  texto(s, TEXTOS.balance, { name: "Lectura balance texto", x: 0.85, y: 6.05, w: 7.0, h: 0.65, fontSize: 12, color: COL.ink, valign: "middle" });

  const dona = (titulo, labels, values, y, colores, nombre) => s.addChart(pres.charts.DOUGHNUT, [{ name: titulo, labels, values: values.map((v) => +(v / 1e6).toFixed(1)) }], Object.assign(ejeTexto(), {
    x: 8.3, y, w: 4.42, h: 2.55, objectName: nombre, holeSize: 55, chartColors: colores,
    showTitle: true, title: titulo, titleFontSize: 12,
    showPercent: true, showValue: false, showLabel: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 10, dataLabelFontBold: true,
    showLegend: true, legendPos: "r", legendFontSize: 10, dataBorder: { pt: 1, color: "FFFFFF" },
  }));
  dona(`En qué está invertido el activo · ${mesA}`, ["Activo corriente", "Activo no corriente"], [BA.ac, BA.anc], 1.45, [HEX.navy, HEX.teal], "Gráfico estructura del activo");
  dona(`Con qué se financia la compañía · ${mesA}`, ["Pasivo corriente", "Pasivo no corriente", "Patrimonio"], [BA.pc, BA.pnc, BA.pat], 4.15, [HEX.amber, HEX.sky, HEX.teal], "Gráfico estructura de financiación");
}

// ───────────────────────────── 15. INDICADORES ─────────────────────────────
{
  const s = contenido("5 · SITUACIÓN FINANCIERA", "5. Situación financiera", "INDICADORES CLAVE - LIQUIDEZ Y ENDEUDAMIENTO");
  const rc = (b) => b.ac / b.pc, end = (b) => (b.pt / b.at) * 100, roe = (b) => (b.resultado / b.pat) * 100;
  // [indicador, fórmula, actual, anterior, formato, mayor es mejor]
  const ratios = [
    ["Razón corriente", "Activo corriente / Pasivo corriente", rc(BA), rc(BB), (v) => num(v, 2), true],
    ["Endeudamiento", "Pasivo total / Activo total", end(BA), end(BB), (v) => num(v) + "%", false],
    ["Margen bruto", "Utilidad bruta / Ingresos netos", R.mb[iA], R.mb[iB], (v) => num(v) + "%", true],
    ["Margen neto", "Utilidad neta / Ingresos netos", R.mn[iA], R.mn[iB], (v) => num(v) + "%", true],
    ["ROE acumulado", "Utilidad neta acum. / Patrimonio", roe(BA), roe(BB), (v) => num(v) + "%", true],
  ];
  const filas = [[TH("INDICADOR", { align: "left" }), TH("FÓRMULA", { align: "left" }), TH(`${MES_A.toUpperCase()} ${DATOS.anio}`), TH(`${MES_B.toUpperCase()} ${DATOS.anio}`, { fill: { color: COL.ink } }), TH("LECTURA")]];
  ratios.forEach(([ind, f, a, b, fmt, mayorMejor], i) => {
    const fill = zebra(i), ok = mayorMejor ? a >= b : a <= b;
    filas.push([
      TD(ind, { fill, bold: true, color: COL.navy, fontSize: 13 }), TD(f, { fill, color: COL.muted }),
      TD(fmt(a), { fill, bold: true, color: COL.navy, align: "center", fontSize: 14 }), TD(fmt(b), { fill, color: COL.muted, align: "center" }),
      TD(ok ? "▲ Mejora" : "▼ Baja", { fill: { color: ok ? "E3F6F7" : "FBE9E7" }, bold: true, color: ok ? COL.sea : COL.red, align: "center" }),
    ]);
  });
  s.addTable(filas, { x: 0.6, y: 1.5, w: 12.13, colW: [2.5, 4.43, 1.75, 1.75, 1.7], rowH: 0.45, border: BORDE(), margin: [0, 0.12, 0, 0.12], objectName: "Tabla indicadores" });

  const tarj = [
    ["LIQUIDEZ", `${num(rc(BA), 2)} veces`, `Por cada peso de pasivo corriente hay $${num(rc(BA), 2)} de activo corriente (${mesB}: ${num(rc(BB), 2)}).`],
    ["SOLIDEZ", `${num((BA.pat / BA.at) * 100)}%`, `Del activo está financiado con patrimonio; el pasivo es el ${num(end(BA))}%.`],
    ["RENTABILIDAD", `${num(roe(BA))}%`, `ROE acumulado: utilidad del año (${money(BA.resultado)}) sobre el patrimonio.`],
  ];
  tarj.forEach(([t, v, d], i) => {
    const x = 0.6 + i * 4.115, y = 4.45, w = 3.9, h = 2.2;
    tarjeta(s, x, y, w, h, { name: "Indicador " + t });
    texto(s, t, { name: "Indicador etiqueta", x, y: y + 0.2, w, h: 0.3, fontSize: 13, bold: true, color: COL.muted, align: "center" });
    texto(s, v, { name: "Indicador valor", x, y: y + 0.52, w, h: 0.65, fontSize: 36, bold: true, color: COL.teal, align: "center", valign: "middle" });
    texto(s, d, { name: "Indicador detalle", x: x + 0.3, y: y + 1.3, w: w - 0.6, h: 0.75, fontSize: 12, color: COL.ink, align: "center" });
  });
}

// ───────────────────────────── 16. CONCLUSIONES ─────────────────────────────
pres.addSection({ title: "6. Conclusiones" });
{
  const s = contenido("6 · SÍNTESIS", "6. Conclusiones", "CONCLUSIONES Y HALLAZGOS DEL MES");
  TEXTOS.conclusiones.forEach(([b, t], i) => {
    const y = 1.45 + i * 1.05, h = 0.92;
    tarjeta(s, 0.6, y, 12.13, h, { name: "Conclusión " + (i + 1) });
    circulo(s, 0.85, y + 0.18, 0.56, String(i + 1).padStart(2, "0"), 16);
    texto(s, [{ text: b, options: { bold: true, color: COL.navy } }, { text: t }], {
      name: "Conclusión texto", x: 1.65, y, w: 10.85, h, fontSize: 14, color: COL.ink, valign: "middle",
    });
  });
}

// ───────────────────────────── 17. CIERRE ─────────────────────────────
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "6. Conclusiones" });
  texto(s, "Gracias", { name: "Cierre título", x: 0.9, y: 2.0, w: 10, h: 1.1, fontSize: 80, bold: true, color: COL.white, valign: "middle" });
  texto(s, DATOS.empresa, { name: "Cierre empresa", x: 0.9, y: 3.3, w: 11.5, h: 0.5, fontSize: 28, bold: true, color: COL.pale, valign: "middle" });
  texto(s, "Análisis Financiero · " + DATOS.corteLargo, { name: "Cierre corte", x: 0.9, y: 3.9, w: 11, h: 0.4, fontSize: 20, color: COL.white, valign: "middle" });
  texto(s, DATOS.preparadoPor, { name: "Preparado por", x: 0.9, y: 6.55, w: 8.5, h: 0.3, fontSize: 13, color: COL.pale, valign: "middle" });
}

// ───────────────────────────── ESCRITURA + AJUSTES ─────────────────────────────
// 1) Aplica la paleta ACONTIS al tema (pptxgenjs no escribe colores de tema).
// 2) Cascada: serie "Base" sin relleno ni etiquetas; formatos de etiqueta por serie.
async function postProceso(archivo) {
  const zip = await JSZip.loadAsync(fs.readFileSync(archivo));
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/theme\/theme\d+\.xml$/.test(name)) {
      let x = await zip.file(name).async("string");
      const clr = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
        .map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join("");
      x = x.replace(/<a:clrScheme name="[^"]*">[\s\S]*?<\/a:clrScheme>/, `<a:clrScheme name="${THEME.name}">${clr}</a:clrScheme>`);
      x = x.replace(/(<a:majorFont>\s*<a:latin typeface=")[^"]*"/, `$1${THEME.headFontFace}"`);
      x = x.replace(/(<a:minorFont>\s*<a:latin typeface=")[^"]*"/, `$1${THEME.bodyFontFace}"`);
      zip.file(name, x);
    }
    if (/^ppt\/charts\/chart\d+\.xml$/.test(name)) {
      let x = await zip.file(name).async("string");
      if (!x.includes("<c:v>Disminuye</c:v>")) continue;
      const formatos = {
        Utilidad: ['"$"#,##0.0" M";;', "FFFFFF"],
        Aumenta: ['"+$"#,##0.0" M";;', HEX.navy],
        Disminuye: ['"−$"#,##0.0" M";;', "FFFFFF"],
      };
      x = x.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (ser) => {
        const nombre = (ser.match(/<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/) || [])[1];
        if (nombre === "Base") {
          ser = ser.replace(/<c:spPr>[\s\S]*?<\/c:spPr>/, "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>");
          ser = ser.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, "<c:dLbls><c:delete val=\"1\"/></c:dLbls>");
          return ser;
        }
        const f = formatos[nombre];
        if (!f) return ser;
        ser = ser.replace(/(<c:dLbls>[\s\S]*?<c:numFmt formatCode=")[^"]*(")/, (m, a, b) => a + f[0].replace(/"/g, "&quot;") + b);
        ser = ser.replace(/(<c:dLbls>[\s\S]*?<c:txPr>[\s\S]*?<a:solidFill>)[\s\S]*?(<\/a:solidFill>)/, `$1<a:srgbClr val="${f[1]}"/>$2`);
        return ser;
      });
      zip.file(name, x);
    }
  }
  fs.writeFileSync(archivo, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

(async () => {
  const salida = path.join(__dirname, DATOS.archivo);
  await pres.writeFile({ fileName: salida });
  await postProceso(salida);
  console.log("Presentación generada:", salida);
})();
