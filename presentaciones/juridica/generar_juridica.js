// Genera el Análisis Financiero de JURIDICA ABOGADOS Y CONSULTORES S.A.S.
// con el formato gráfico ACONTIS (mismo estilo del informe Santa Vicenta).
//
// Uso:  npm run juridica      (o)   node presentaciones/juridica/generar_juridica.js
//
// Todo lo que cambia mes a mes (cifras, textos, gráficos y tablas) está en el
// bloque DATOS. Los gráficos son nativos de PowerPoint y las tablas son tablas
// reales, así que también se pueden editar directamente en el .pptx.

const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");
const fs = require("fs");

const ROOT = path.resolve(__dirname, "..", "..");
const ASSETS = path.join(ROOT, "assets", "acontis");
const SALIDA = path.join(__dirname, "Analisis_Financiero_Juridica_Junio_2026.pptx");

// ───────────────────────────── DATOS ─────────────────────────────
const DATOS = {
  empresa: "JURIDICA ABOGADOS Y CONSULTORES S.A.S.",
  empresaNombre: "Juridica Abogados y Consultores S.A.S.",
  nit: "900.944.440-3",
  corte: "Corte: junio 2026",
  corteLargo: "Corte a 30 de junio de 2026",
  comparativo: "Comparativo mensual mayo – junio 2026",
  preparadoPor: "Preparado por Asesorías Contables del Caribe S.A.S. (ACONTIS)",
  mesAct: "Junio",
  mesAnt: "Mayo",

  // Evolución del semestre (millones de COP)
  meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun"],
  ingresosMes: [63.16, 70.89, 74.42, 73.52, 91.14, 102.17],
  utilidadMes: [14.37, 1.66, 34.37, 18.44, 70.62, 28.07],
  margenBruto: [69.2, 72.8, 73.6, 72.4, 75.7, 79.5],
  margenOperacional: [23.9, 21.7, 47.1, -0.4, 49.4, -10.2],
  margenNeto: [22.7, 2.3, 46.2, 25.1, 77.5, 27.5],

  // Cuadro de mando: [concepto, junio, mayo, var $, var %, favorable?, negrita?]
  cuadroER: [
    ["Ingresos netos", "102,2", "91,1", "+11,0", "+12,1%", true, true],
    ["Utilidad bruta", "81,2", "69,0", "+12,3", "+17,8%", true, true],
    ["Gastos de administración y ventas", "91,7", "23,9", "+67,8", "+283,5%", false, false],
    ["Resultado operacional", "-10,5", "45,1", "-55,5", "−123,2%", false, true],
    ["Resultado antes de impuesto", "56,7", "70,6", "-13,9", "−19,7%", false, false],
    ["Utilidad neta", "28,1", "70,6", "-42,5", "−60,2%", false, true],
  ],
  cuadroESF: [
    ["Activo total", "2.004,0", "1.982,6", "+21,4", "+1,1%", true, true],
    ["Pasivo total", "214,5", "221,2", "-6,7", "−3,0%", true, true],
    ["Patrimonio", "1.789,5", "1.761,4", "+28,1", "+1,6%", true, true],
  ],

  // Comparativo resumen (gráfico de columnas, millones)
  resumenCats: ["Ingresos netos", "Utilidad bruta", "Gastos admon. y ventas", "Utilidad neta"],
  resumenAct: [102.2, 81.2, 91.7, 28.1],
  resumenAnt: [91.1, 69.0, 23.9, 70.6],

  // Puente de utilidad mayo → junio (millones). tipo: total | sube | baja
  puente: [
    ["Utilidad mayo", 70.6, "total"],
    ["Ingresos netos", 11.0, "sube"],
    ["Costo de operación", 1.2, "sube"],
    ["Gastos admon. y ventas", 67.8, "baja"],
    ["Ingresos no operacionales", 76.1, "sube"],
    ["Gastos no operacionales", 34.5, "baja"],
    ["Impuesto SIMPLE", 28.6, "baja"],
    ["Utilidad junio", 28.1, "total"],
  ],

  // Gastos de administración y ventas · top 6 (millones)
  gastosCats: ["Gastos Unión Temporal", "Honorarios", "Servicios", "Contribuciones y afiliaciones", "Depreciaciones", "Amortizaciones"],
  gastosAct: [63.6, 4.7, 3.3, 3.2, 2.2, 0.8],
  gastosAnt: [0.0, 3.3, 0.9, 2.8, 2.4, 0.4],
  gastosVar: ["Nuevo", "+43,0%", "+281,4%", "+15,0%", "−6,4%", "+110,1%"],

  // Balance (millones)
  balance: [
    // [rubro, junio, mayo, var $, var %, favorable?, tipo]  tipo: banda | fila | total
    ["ACTIVO", "", "", "", "", null, "banda"],
    ["Activo corriente", "1.067,4", "1.041,9", "+25,5", "+2,4%", true, "fila"],
    ["Activo no corriente", "936,6", "940,7", "-4,1", "−0,4%", false, "fila"],
    ["TOTAL ACTIVO", "2.004,0", "1.982,6", "+21,4", "+1,1%", true, "total"],
    ["PASIVO", "", "", "", "", null, "banda"],
    ["Pasivo corriente", "208,2", "214,9", "-6,7", "−3,1%", true, "fila"],
    ["Pasivo no corriente", "6,3", "6,3", "0,0", "+0,0%", null, "fila"],
    ["TOTAL PASIVO", "214,5", "221,2", "-6,7", "−3,0%", true, "total"],
    ["PATRIMONIO", "", "", "", "", null, "banda"],
    ["TOTAL PATRIMONIO", "1.789,5", "1.761,4", "+28,1", "+1,6%", true, "total"],
    ["TOTAL PASIVO + PATRIMONIO", "2.004,0", "1.982,6", "+21,4", "+1,1%", true, "total"],
  ],
  estructuraActivo: { labels: ["Activo corriente", "Activo no corriente"], values: [1067.4, 936.6] },
  estructuraFinanciacion: { labels: ["Pasivo corriente", "Pasivo no corriente", "Patrimonio"], values: [208.2, 6.3, 1789.5] },

  // Indicadores: [indicador, fórmula, junio, mayo, lectura, mejora?]
  ratios: [
    ["Razón corriente", "Activo corriente / Pasivo corriente", "5,13", "4,85", "Mejora", true],
    ["Endeudamiento", "Pasivo total / Activo total", "10,7%", "11,2%", "Mejora", true],
    ["Margen bruto", "Utilidad bruta / Ingresos netos", "79,5%", "75,7%", "Mejora", true],
    ["Margen neto", "Utilidad neta / Ingresos netos", "27,5%", "77,5%", "Baja", false],
    ["ROE acumulado", "Utilidad neta acum. / Patrimonio", "9,4%", "7,9%", "Mejora", true],
  ],
};

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
const tono = (fav) => (fav === null || fav === undefined ? COL.muted : fav ? COL.teal : COL.red);

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
  s.addText(DATOS.corte, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.21, objectName: id("Corte"),
    x: 10.95, y: 0.45, w: 1.77, h: 0.42, fill: { color: COL.card }, line: { color: COL.line, width: 1 },
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
const nb = (t) => t.replace(/(\d) (M|veces|pp)\b/g, "$1\u00A0$2");
function texto(s, t, o) {
  t = typeof t === "string" ? nb(t) : t.map((r) => Object.assign({}, r, { text: nb(r.text) }));
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
  texto(s, DATOS.corteLargo + " · " + DATOS.comparativo, { name: "Corte", x: 0.9, y: 4.75, w: 11.5, h: 0.5, fontSize: 28, color: COL.white, valign: "middle" });
  texto(s, DATOS.preparadoPor, { name: "Preparado por", x: 0.9, y: 6.55, w: 8.5, h: 0.3, fontSize: 13, color: COL.pale, valign: "middle" });
}

// ───────────────────────────── 2. ÍNDICE ─────────────────────────────
{
  const s = contenido("ÍNDICE", "Portada", "CONTENIDO");
  const items = [
    ["Resumen ejecutivo", "Cifras clave y la historia del mes"],
    ["Estado de resultados", "Cuadro de mando, puente de utilidad y márgenes"],
    ["Ingresos y gastos", "De ingreso bruto a neto y estructura del gasto"],
    ["Análisis de causas", "Tres factores detrás de la caída de junio"],
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
  const s = contenido("1 · RESUMEN EJECUTIVO", "1. Resumen ejecutivo", "RESUMEN EJECUTIVO - MAYO VS JUNIO 2026");
  const kpis = [
    { t: "INGRESOS NETOS", v: "$102,2 M", l1: "Mayo: $91,1 M", l2: "▲ +12,1% vs. mayo", c2: COL.teal },
    { t: "ACTIVO TOTAL", v: "$2.004,0 M", l1: "Mayo: $1.982,6 M", l2: "▲ +1,1% vs. mayo", c2: COL.teal },
    { t: "PATRIMONIO", v: "$1.789,5 M", l1: "Mayo: $1.761,4 M", l2: "▲ +1,6% vs. mayo", c2: COL.teal },
    { t: "UTILIDAD NETA", v: "$28,1 M", l1: "Margen neto 27,5%", l2: "Mayo: $70,6 M  (−60,2%)", c2: COL.white, dark: true },
  ];
  kpis.forEach((k, i) => {
    const x = 0.6 + i * 3.08, y = 1.5, w = 2.88, h = 1.95;
    tarjeta(s, x, y, w, h, { name: "KPI " + k.t, fill: k.dark ? COL.navy : COL.card, line: k.dark ? COL.navy : COL.line });
    texto(s, k.t, { name: "KPI etiqueta", x: x + 0.25, y: y + 0.25, w: w - 0.4, h: 0.3, fontSize: 13, bold: true, color: k.dark ? COL.white : COL.muted });
    texto(s, k.v, { name: "KPI valor", x: x + 0.25, y: y + 0.58, w: w - 0.4, h: 0.6, fontSize: 32, bold: true, color: k.dark ? COL.white : COL.navy, valign: "middle" });
    texto(s, k.l1, { name: "KPI detalle", x: x + 0.25, y: y + 1.24, w: w - 0.4, h: 0.26, fontSize: 12, color: k.dark ? COL.sky : COL.muted });
    texto(s, k.l2, { name: "KPI variación", x: x + 0.25, y: y + 1.52, w: w - 0.4, h: 0.26, fontSize: 12, bold: true, color: k.c2 });
  });

  s.addChart(pres.charts.BAR, [
    { name: "Junio 2026", labels: DATOS.resumenCats, values: DATOS.resumenAct },
    { name: "Mayo 2026", labels: DATOS.resumenCats, values: DATOS.resumenAnt },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 3.7, w: 7.6, h: 3.05, objectName: "Gráfico comparativo",
    barDir: "col", barGapWidthPct: 60, chartColors: [HEX.navy, HEX.teal],
    showTitle: true, title: "Mayo vs junio, en millones de pesos",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"#,##0.0', dataLabelFontSize: 10, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    valAxisLabelFormatCode: "#,##0", valAxisHidden: true, valGridLine: { style: "none" },
    showLegend: true, legendPos: "b",
  }));

  tarjeta(s, 8.45, 3.7, 4.27, 3.05, { name: "Lectura del mes" });
  texto(s, "La historia del mes", { name: "Lectura título", x: 8.75, y: 3.95, w: 3.7, h: 0.38, fontSize: 18, bold: true, color: COL.sea });
  texto(s, [
    { text: "Los ingresos netos crecieron 12,1% en junio, pero la utilidad neta cayó 60,2%.", options: { breakLine: true, paraSpaceAfter: 8 } },
    { text: "Tres factores no operativos explican la diferencia: el fin del ingreso financiero extraordinario de mayo, la liquidación bimestral de la Unión Temporal EVB Abogados y el pago del régimen SIMPLE." },
  ], { name: "Lectura texto", x: 8.75, y: 4.42, w: 3.7, h: 2.15, fontSize: 13, color: COL.ink });
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
separador("2", "Análisis comparativo", "Estado de Resultados", "MAYO VS JUNIO 2026 · EVOLUCIÓN DEL SEMESTRE", "2. Estado de resultados");

// ───────────────────────────── 5. CUADRO DE MANDO ─────────────────────────────
{
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", "CUADRO DE MANDO - COMPARATIVO INTEGRAL");
  const n = 5;
  const fila = (r, i) => {
    const [concepto, act, ant, vd, vp, fav, neg] = r;
    const fill = { color: i % 2 ? "F3F7FB" : "FFFFFF" };
    const c = neg ? COL.navy : COL.ink;
    return [
      TD(neg ? concepto.toUpperCase() : concepto, { fill, bold: neg, color: c }),
      TD(act, { fill, bold: true, color: COL.navy, align: "right" }),
      TD(ant, { fill, color: COL.muted, align: "right" }),
      TD(vd, { fill, bold: true, color: tono(fav), align: "right" }),
      TD(vp, { fill, bold: true, color: tono(fav), align: "right" }),
    ];
  };
  const filas = [
    [TH("CONCEPTO", { align: "left" }), TH("JUNIO 2026"), TH("MAYO 2026", { fill: { color: COL.ink } }), TH("VAR. $"), TH("VAR. %")],
    banda("ESTADO DE RESULTADO INTEGRAL", n),
    ...DATOS.cuadroER.map(fila),
    banda("ESTADO DE SITUACIÓN FINANCIERA", n),
    ...DATOS.cuadroESF.map(fila),
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
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", "PUENTE DE UTILIDAD - DE MAYO A JUNIO");
  // Serie "Base" invisible + series visibles = cascada editable en Excel.
  const cats = DATOS.puente.map((p) => p[0]);
  const base = [], total = [], sube = [], baja = [];
  let acum = 0;
  DATOS.puente.forEach(([_, v, t]) => {
    if (t === "total") { base.push(0); total.push(v); sube.push(0); baja.push(0); acum = v; }
    else if (t === "sube") { base.push(+acum.toFixed(1)); total.push(0); sube.push(v); baja.push(0); acum += v; }
    else { acum -= v; base.push(+acum.toFixed(1)); total.push(0); sube.push(0); baja.push(v); }
  });
  s.addChart(pres.charts.BAR, [
    { name: "Base", labels: cats, values: base },
    { name: "Utilidad", labels: cats, values: total },
    { name: "Aumenta", labels: cats, values: sube },
    { name: "Disminuye", labels: cats, values: baja },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 8.25, h: 5.25, objectName: "Gráfico puente de utilidad",
    barDir: "col", barGrouping: "stacked", barGapWidthPct: 45,
    chartColors: ["FFFFFF", HEX.navy, HEX.teal, HEX.red],
    showTitle: true, title: "Utilidad neta mayo → junio, en millones de pesos",
    showValue: true, dataLabelPosition: "inEnd", dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelColor: "FFFFFF",
    dataLabelFormatCode: '"$"#,##0.0;;', catAxisLabelFontSize: 10,
    valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false,
  }));

  tarjeta(s, 9.1, 1.45, 3.62, 5.25, { name: "Lectura del puente" });
  texto(s, "Lectura del puente", { name: "Puente título", x: 9.35, y: 1.7, w: 3.15, h: 0.38, fontSize: 18, bold: true, color: COL.sea });
  const puntos = [
    ["Operación: +$12,2 M. ", "Crecen los ingresos netos en $11,0 M y baja el costo de operación en $1,2 M."],
    ["Gastos admon.: −$67,8 M. ", "De ellos, $63,6 M son gastos de la Unión Temporal."],
    ["No operacional: +$41,6 M neto. ", "Ingresos de $76,1 M frente a gastos de $34,5 M."],
    ["Impuesto SIMPLE: −$28,6 M. ", "Pago bimestral que mayo no tuvo."],
  ];
  texto(s, puntos.flatMap(([b, t], i) => [
    { text: b, options: { bold: true, color: COL.navy } },
    { text: t, options: { breakLine: i < puntos.length - 1, paraSpaceAfter: 10 } },
  ]), { name: "Puente texto", x: 9.35, y: 2.2, w: 3.15, h: 4.3, fontSize: 13 });
}

// ───────────────────────────── 7. EVOLUCIÓN DEL SEMESTRE ─────────────────────────────
{
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", "EVOLUCIÓN MENSUAL - ENERO A JUNIO 2026");
  s.addChart(pres.charts.BAR, [
    { name: "Ingresos netos", labels: DATOS.meses, values: DATOS.ingresosMes },
    { name: "Utilidad neta", labels: DATOS.meses, values: DATOS.utilidadMes },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 7.6, h: 5.25, objectName: "Gráfico evolución mensual",
    barDir: "col", barGapWidthPct: 55, chartColors: [HEX.navy, HEX.teal],
    showTitle: true, title: "Ingresos y utilidad neta mensual (millones de pesos)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0.0", dataLabelFontSize: 10, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    valAxisLabelFormatCode: "#,##0", showLegend: true, legendPos: "b",
  }));

  const fmt = (v, d = 1) => v.toFixed(d).replace(".", ",");
  const sumI = DATOS.ingresosMes.reduce((a, b) => a + b, 0), sumU = DATOS.utilidadMes.reduce((a, b) => a + b, 0);
  const filas = [[TH("MES"), TH("INGRESOS"), TH("UTILIDAD"), TH("MARGEN")]];
  DATOS.meses.forEach((m, i) => {
    const fill = { color: i % 2 ? "F3F7FB" : "FFFFFF" };
    const mg = DATOS.margenNeto[i];
    filas.push([
      TD(m, { fill }), TD(fmt(DATOS.ingresosMes[i]), { fill, align: "right" }), TD(fmt(DATOS.utilidadMes[i]), { fill, align: "right" }),
      TD(fmt(mg) + "%", { fill, align: "right", bold: true, color: mg >= 20 ? COL.teal : COL.amber }),
    ]);
  });
  filas.push([
    TD("Total", { fill: { color: "F3F7FB" }, bold: true, color: COL.navy }),
    TD(fmt(sumI), { fill: { color: "F3F7FB" }, bold: true, color: COL.navy, align: "right" }),
    TD(fmt(sumU), { fill: { color: "F3F7FB" }, bold: true, color: COL.navy, align: "right" }),
    TD(fmt((sumU / sumI) * 100) + "%", { fill: { color: "F3F7FB" }, bold: true, color: COL.navy, align: "right" }),
  ]);
  s.addTable(filas, { x: 8.45, y: 1.5, w: 4.27, colW: [0.85, 1.17, 1.17, 1.08], rowH: 0.34, border: BORDE(), margin: [0, 0.08, 0, 0.08], objectName: "Tabla evolución" });

  tarjeta(s, 8.45, 4.75, 4.27, 1.95, { name: "Tendencia" });
  texto(s, "+61,8%", { name: "Tendencia valor", x: 8.7, y: 4.92, w: 3.8, h: 0.6, fontSize: 32, bold: true, color: COL.teal, valign: "middle" });
  texto(s, "Crecimiento de los ingresos netos de enero a junio. La utilidad neta fue más volátil por factores no operativos puntuales en cada mes.", {
    name: "Tendencia texto", x: 8.7, y: 5.55, w: 3.8, h: 1.0, fontSize: 12, color: COL.ink,
  });
}

// ───────────────────────────── 8. MÁRGENES ─────────────────────────────
{
  const s = contenido("2 · ESTADO DE RESULTADOS", "2. Estado de resultados", "RENTABILIDAD - MÁRGENES DEL SEMESTRE");
  s.addChart(pres.charts.LINE, [
    { name: "Margen bruto", labels: DATOS.meses, values: DATOS.margenBruto },
    { name: "Margen operacional", labels: DATOS.meses, values: DATOS.margenOperacional },
    { name: "Margen neto", labels: DATOS.meses, values: DATOS.margenNeto },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 8.6, h: 5.25, objectName: "Gráfico márgenes",
    chartColors: [HEX.navy, HEX.teal, HEX.amber], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 8,
    showTitle: true, title: "% sobre ingresos netos · enero a junio 2026",
    valAxisLabelFormatCode: '0"%"', valAxisMinVal: -20, valAxisMaxVal: 100, valAxisMajorUnit: 20,
    showLegend: true, legendPos: "b",
  }));
  const tarjetas = [
    ["MARGEN BRUTO", "79,5%", "▲ +3,8 pp  ·  mayo 75,7%", COL.teal],
    ["MARGEN OPERACIONAL", "-10,2%", "▼ −59,7 pp  ·  mayo 49,4%", COL.red],
    ["MARGEN NETO", "27,5%", "▼ −50,0 pp  ·  mayo 77,5%", COL.amber],
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
  const filas = [
    ["MAYO", "2026", ["$96,1 M", "-$5,0 M", "$91,1 M"]],
    ["JUNIO", "2026", ["$109,2 M", "-$7,0 M", "$102,2 M"]],
  ];
  const cols = [
    { x: 2.35, w: 2.95, t: "INGRESOS BRUTOS" },
    { x: 6.0, w: 2.95, t: "(−) DEVOLUCIONES" },
    { x: 9.65, w: 3.07, t: "INGRESOS NETOS", dark: true },
  ];
  filas.forEach(([mes, anio, vals], r) => {
    const y = 1.55 + r * 1.5, h = 1.25;
    texto(s, mes, { name: "Mes", x: 0.6, y: y + 0.18, w: 1.6, h: 0.5, fontSize: 28, bold: true, color: COL.navy, valign: "middle" });
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
  const vars = [
    ["VAR. INGRESOS BRUTOS", "▲ +$13,0 M", "+13,6% vs. mayo", COL.teal],
    ["VAR. DEVOLUCIONES", "▲ +$2,0 M", "+40,4% vs. mayo", COL.red],
    ["VAR. INGRESOS NETOS", "▲ +$11,0 M", "+12,1% vs. mayo", COL.teal],
  ];
  vars.forEach(([t, v, d, c], i) => {
    const x = cols[i].x, w = cols[i].w, y = 4.65, h = 1.45;
    tarjeta(s, x, y, w, h, { name: t });
    texto(s, t, { name: "Etiqueta", x: x + 0.25, y: y + 0.18, w: w - 0.4, h: 0.28, fontSize: 12, bold: true, color: COL.muted });
    texto(s, v, { name: "Valor", x: x + 0.25, y: y + 0.48, w: w - 0.4, h: 0.5, fontSize: 26, bold: true, color: c, valign: "middle" });
    texto(s, d, { name: "Detalle", x: x + 0.25, y: y + 1.02, w: w - 0.4, h: 0.28, fontSize: 12, color: COL.muted });
  });
  texto(s, "VARIACIÓN", { name: "Var etiqueta", x: 0.6, y: 5.1, w: 1.6, h: 0.5, fontSize: 20, bold: true, color: COL.sea, valign: "middle" });
  texto(s, "Junio concentró el mayor nivel de devoluciones del semestre, sin afectar el crecimiento neto de los ingresos. Actividades de consultoría · cifras en millones de pesos.", {
    name: "Nota", x: 0.6, y: 6.3, w: 12.13, h: 0.35, fontSize: 12, italic: true, color: COL.muted,
  });
}

// ───────────────────────────── 11. GASTOS TOP 6 ─────────────────────────────
{
  const s = contenido("3 · INGRESOS Y GASTOS", "3. Ingresos y gastos", "GASTOS DE ADMINISTRACIÓN Y VENTAS - TOP 6");
  s.addChart(pres.charts.BAR, [
    { name: "Junio 2026", labels: DATOS.gastosCats, values: DATOS.gastosAct },
    { name: "Mayo 2026", labels: DATOS.gastosCats, values: DATOS.gastosAnt },
  ], Object.assign(ejeTexto(), {
    x: 0.6, y: 1.45, w: 7.0, h: 5.25, objectName: "Gráfico gastos top 6",
    barDir: "bar", barGapWidthPct: 45, chartColors: [HEX.navy, HEX.teal], catAxisOrientation: "maxMin",
    showTitle: true, title: "Mayo vs junio, en millones de pesos",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0.0", dataLabelFontSize: 10, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    valAxisHidden: true, valGridLine: { style: "none" }, showLegend: true, legendPos: "b",
  }));
  const fmt = (v) => v.toFixed(1).replace(".", ",");
  const filas = [[TH("CONCEPTO", { align: "left" }), TH("JUNIO"), TH("MAYO", { fill: { color: COL.ink } }), TH("VAR. %")]];
  DATOS.gastosCats.forEach((c, i) => {
    const fill = { color: i % 2 ? "F3F7FB" : "FFFFFF" };
    const v = DATOS.gastosVar[i];
    filas.push([
      TD(c, { fill, fontSize: 11 }), TD(fmt(DATOS.gastosAct[i]), { fill, bold: true, color: COL.navy, align: "right" }),
      TD(fmt(DATOS.gastosAnt[i]), { fill, color: COL.muted, align: "right" }),
      TD(v, { fill, bold: true, align: "right", color: v.startsWith("−") ? COL.teal : COL.red }),
    ]);
  });
  filas.push([
    TD("TOTAL GASTOS ADMON. Y VENTAS", { fill: { color: "F3F7FB" }, bold: true, color: COL.navy, fontSize: 10 }),
    TD("91,7", { fill: { color: "F3F7FB" }, bold: true, color: COL.navy, align: "right" }),
    TD("23,9", { fill: { color: "F3F7FB" }, color: COL.muted, align: "right" }),
    TD("+283,5%", { fill: { color: "F3F7FB" }, bold: true, color: COL.red, align: "right" }),
  ]);
  s.addTable(filas, { x: 7.85, y: 1.5, w: 4.87, colW: [2.2, 0.85, 0.85, 0.97], rowH: 0.38, border: BORDE(), margin: [0, 0.08, 0, 0.08], objectName: "Tabla gastos" });

  tarjeta(s, 7.85, 4.75, 4.87, 1.6, { name: "Concentración" });
  texto(s, "Concentración", { name: "Concentración título", x: 8.1, y: 4.92, w: 4.4, h: 0.38, fontSize: 18, bold: true, color: COL.sea });
  texto(s, [
    { text: "Los gastos de la Unión Temporal ($63,6 M) explican el " },
    { text: "94% del incremento", options: { bold: true, color: COL.navy } },
    { text: " total en gastos de administración y ventas de junio." },
  ], { name: "Concentración texto", x: 8.1, y: 5.38, w: 4.4, h: 0.9, fontSize: 13 });
  texto(s, "Código de color:  turquesa = favorable  ·  rojo = desfavorable", { name: "Nota de color", x: 7.85, y: 6.45, w: 4.87, h: 0.25, fontSize: 10, color: COL.muted, align: "right" });
}

// ───────────────────────────── 12. TRES FACTORES ─────────────────────────────
pres.addSection({ title: "4. Análisis de causas" });
{
  const s = contenido("4 · ANÁLISIS DE CAUSAS", "4. Análisis de causas", "TRES FACTORES DETRÁS DE LA CAÍDA DE JUNIO");
  const f = [
    ["01", "Rendimientos financieros extraordinarios", "$25,6 M", "$0,4 M",
      "Mayo incluyó un ingreso financiero puntual por rendimientos de CDT que no se repitió en junio: una caída no recurrente de $25,2 M."],
    ["02", "Liquidación Unión Temporal EVB Abogados", "$0", "+$3,8 M neto",
      "La UT liquida cada dos meses (abril y junio). En junio aportó $102,2 M de ingreso no operacional y $98,4 M de costos asociados: efecto neto positivo, pero de gran volumen contable."],
    ["03", "Pago bimestral del régimen SIMPLE", "$0", "$28,6 M",
      "El impuesto SIMPLE se paga cada dos meses (febrero, abril, junio). Mayo no tuvo pago; junio absorbió el bimestre completo, reduciendo la utilidad neta."],
  ];
  f.forEach(([n, t, may, jun, d], i) => {
    const x = 0.6 + i * 4.115, y = 1.5, w = 3.9, h = 3.95;
    tarjeta(s, x, y, w, h, { name: "Factor " + n });
    circulo(s, x + 0.25, y + 0.25, 0.6, n, 18);
    texto(s, t.toUpperCase(), { name: "Factor título", x: x + 1.0, y: y + 0.2, w: w - 1.2, h: 0.7, fontSize: 14, bold: true, color: COL.navy, valign: "middle" });
    // Mini comparativo mayo / junio
    [["MAYO", may, COL.muted], ["JUNIO", jun, COL.navy]].forEach(([l, v, c], j) => {
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
    { text: "sin estos tres efectos puntuales, el resultado operativo del negocio principal (consultoría) siguió una trayectoria estable y creciente en el semestre." },
  ], { name: "Lectura clave texto", x: 0.9, y: 5.65, w: 11.6, h: 1.05, fontSize: 15, color: COL.white, valign: "middle" });
}

// ───────────────────────────── 13. SEPARADOR · SITUACIÓN FINANCIERA ─────────────────────────────
separador("5", "Estado de Situación Financiera", "Balance e Indicadores", "LIQUIDEZ · ENDEUDAMIENTO · RENTABILIDAD", "5. Situación financiera");

// ───────────────────────────── 14. BALANCE COMPARATIVO ─────────────────────────────
{
  const s = contenido("5 · SITUACIÓN FINANCIERA", "5. Situación financiera", "BALANCE GENERAL COMPARATIVO");
  const filas = [[TH("CUENTA", { align: "left" }), TH("JUNIO 2026"), TH("MAYO 2026", { fill: { color: COL.ink } }), TH("VAR. $"), TH("VAR. %")]];
  let k = 0;
  DATOS.balance.forEach(([r, a, b, vd, vp, fav, tipo]) => {
    if (tipo === "banda") { filas.push(banda(r, 5)); k = 0; return; }
    const tot = tipo === "total";
    const fill = { color: tot ? "F3F7FB" : k++ % 2 ? "F3F7FB" : "FFFFFF" };
    filas.push([
      TD(r, { fill, bold: tot, color: tot ? COL.navy : COL.ink, fontSize: tot ? 11 : 12 }),
      TD(a, { fill, bold: true, color: COL.navy, align: "right" }),
      TD(b, { fill, color: COL.muted, align: "right" }),
      TD(vd, { fill, bold: true, color: tono(fav), align: "right" }),
      TD(vp, { fill, bold: true, color: tono(fav), align: "right" }),
    ]);
  });
  s.addTable(filas, { x: 0.6, y: 1.5, w: 7.45, colW: [2.85, 1.15, 1.15, 1.1, 1.2], rowH: 0.37, border: BORDE(), margin: [0, 0.08, 0, 0.08], objectName: "Tabla balance" });

  tarjeta(s, 0.6, 6.0, 7.45, 0.7, { name: "Lectura balance" });
  texto(s, "La compañía redujo el pasivo corriente y fortaleció el activo corriente, mejorando su liquidez frente a mayo.", {
    name: "Lectura balance texto", x: 0.85, y: 6.0, w: 7.0, h: 0.7, fontSize: 13, color: COL.ink, valign: "middle",
  });

  const dona = (titulo, datos, y, colores, nombre) => s.addChart(pres.charts.DOUGHNUT, [{ name: titulo, labels: datos.labels, values: datos.values }], Object.assign(ejeTexto(), {
    x: 8.3, y, w: 4.42, h: 2.55, objectName: nombre, holeSize: 55, chartColors: colores,
    showTitle: true, title: titulo, titleFontSize: 12,
    showPercent: true, showValue: false, showLabel: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 10, dataLabelFontBold: true,
    showLegend: true, legendPos: "r", legendFontSize: 10, dataBorder: { pt: 1, color: "FFFFFF" },
  }));
  dona("En qué está invertido el activo · junio", DATOS.estructuraActivo, 1.45, [HEX.navy, HEX.teal], "Gráfico estructura del activo");
  dona("Con qué se financia la compañía · junio", DATOS.estructuraFinanciacion, 4.15, [HEX.amber, HEX.sky, HEX.teal], "Gráfico estructura de financiación");
}

// ───────────────────────────── 15. INDICADORES ─────────────────────────────
{
  const s = contenido("5 · SITUACIÓN FINANCIERA", "5. Situación financiera", "INDICADORES CLAVE - LIQUIDEZ Y ENDEUDAMIENTO");
  const filas = [[TH("INDICADOR", { align: "left" }), TH("FÓRMULA", { align: "left" }), TH("JUNIO 2026"), TH("MAYO 2026", { fill: { color: COL.ink } }), TH("LECTURA")]];
  DATOS.ratios.forEach(([ind, f, a, b, l, ok], i) => {
    const fill = { color: i % 2 ? "F3F7FB" : "FFFFFF" };
    filas.push([
      TD(ind, { fill, bold: true, color: COL.navy, fontSize: 13 }), TD(f, { fill, color: COL.muted }),
      TD(a, { fill, bold: true, color: COL.navy, align: "center", fontSize: 14 }), TD(b, { fill, color: COL.muted, align: "center" }),
      TD((ok ? "▲ " : "▼ ") + l, { fill: { color: ok ? "E3F6F7" : "FBE9E7" }, bold: true, color: ok ? COL.sea : COL.red, align: "center" }),
    ]);
  });
  s.addTable(filas, { x: 0.6, y: 1.5, w: 12.13, colW: [2.5, 4.43, 1.75, 1.75, 1.7], rowH: 0.45, border: BORDE(), margin: [0, 0.12, 0, 0.12], objectName: "Tabla indicadores" });

  const tarj = [
    ["LIQUIDEZ", "5,13 veces", "Por cada peso de pasivo corriente hay $5,13 de activo corriente (mayo: 4,85)."],
    ["SOLIDEZ", "89,3%", "Del activo está financiado con patrimonio; el pasivo es solo el 10,7%."],
    ["RENTABILIDAD", "9,4%", "ROE acumulado: utilidad del semestre ($167,5 M) sobre el patrimonio."],
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
  const items = [
    ["Ingresos netos +12,1% en junio ", "($102,2 M vs $91,1 M en mayo): sexta expansión mensual consecutiva y crecimiento acumulado del 61,8% desde enero."],
    ["Utilidad neta −60,2%: ", "explicada por tres factores puntuales — fin del rendimiento financiero extraordinario de mayo, liquidación bimestral de la Unión Temporal EVB Abogados y pago del régimen SIMPLE ($28,6 M) — y no por el negocio operativo."],
    ["Margen operacional negativo (−10,2%): ", "el alza de 283,5% en gastos de administración y ventas, concentrada en un 94% en los gastos de la Unión Temporal, explica el giro."],
    ["Liquidez y endeudamiento mejoran: ", "la razón corriente sube a 5,13 (vs 4,85) y el endeudamiento baja a 10,7% del activo total."],
    ["Patrimonio +1,6% y activo total +1,1%: ", "impulsados por la utilidad acumulada del semestre ($167,5 M), una base financiera sólida pese a la volatilidad mensual."],
  ];
  items.forEach(([b, t], i) => {
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
      const clr = Object.entries({
        dk1: THEME.colors.dk1, lt1: THEME.colors.lt1, dk2: THEME.colors.dk2, lt2: THEME.colors.lt2,
        accent1: THEME.colors.accent1, accent2: THEME.colors.accent2, accent3: THEME.colors.accent3,
        accent4: THEME.colors.accent4, accent5: THEME.colors.accent5, accent6: THEME.colors.accent6,
        hlink: THEME.colors.hlink, folHlink: THEME.colors.folHlink,
      }).map(([k, v]) => `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`).join("");
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
  await pres.writeFile({ fileName: SALIDA });
  await postProceso(SALIDA);
  console.log("Presentación generada:", SALIDA);
})();
