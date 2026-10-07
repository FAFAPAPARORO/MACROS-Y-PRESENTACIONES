// Propuesta comercial y económica ACONTIS para el Ministerio de Agricultura y
// Desarrollo Rural (MADR): consultoría forense sobre los recursos del Fondo
// Nacional del Arroz. Formato gráfico ACONTIS.
//
// Uso:  npm run madr      (o)   node presentaciones/madr/generar_madr.js
//
// Contenido: documento Word «MINISTERIO DE AGRICULTURA Y DESARROLLO RURAL – MADR».
// Los valores económicos se calculan a partir de los valores antes de IVA del bloque
// VALORES, así que totales, IVA y forma de pago siempre cuadran.

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");

const ROOT = path.resolve(__dirname, "..", "..");
const ASSETS = path.join(ROOT, "assets", "acontis");
const SALIDA = path.join(__dirname, "Propuesta_MADR_Fondo_Nacional_del_Arroz.pptx");

// ───────────────────────────── DATOS ─────────────────────────────
const PILDORA = "Propuesta · octubre 2026";
const IVA = 0.19;
const VALORES = {
  entregables: [
    // [No., nombre, valor antes de IVA, % de pago]
    ["1", "Plan de Auditoría Detallado", 55000000, 0.10],
    ["2", "Informe preliminar – alertas cuantificadas y casos priorizados", 220000000, 0.40],
    ["3", "Informe final y expedientes de soporte", 275000000, 0.50],
  ],
  dictamen: 85000000, // componente pericial condicional, por dictamen
};

const SECCIONES = [
  ["01", "Antecedentes y entendimiento"],
  ["02", "Objetivo de la consultoría"],
  ["03", "Alcance"],
  ["04", "Metodología y plan de trabajo"],
  ["05", "Entregables y cronograma"],
  ["06", "Equipo y aseguramiento de calidad"],
  ["07", "Valor de la propuesta"],
];

// ───────────────────────────── TEMA ─────────────────────────────
const THEME = {
  name: "ACONTIS",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A2233", lt1: "FFFFFF", dk2: "011E66", lt2: "F3F7FB",
    accent1: "011E66", accent2: "00AFBC", accent3: "50C1E4",
    accent4: "056875", accent5: "E8A13D", accent6: "C4453B",
    hlink: "00AFBC", folHlink: "056875",
  },
};
const HEX = { navy: "011E66", teal: "00AFBC", sky: "50C1E4", sea: "056875", amber: "E8A13D", red: "C4453B", ink: "1A2233", muted: "5A6572", line: "DDE6EF", card: "F3F7FB", pale: "7DD1EB" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "ACONTIS";
pres.company = "Asesorías Contables del Caribe S.A.S.";
pres.title = "Propuesta comercial y económica · MADR · Fondo Nacional del Arroz";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

const C = pres.SchemeColor;
const COL = {
  navy: C.accent1, teal: C.accent2, sky: C.accent3, sea: C.accent4, amber: C.accent5, red: C.accent6,
  ink: C.text1, white: C.background1, card: C.background2, muted: "5A6572", line: "DDE6EF", pale: "7DD1EB", tint: "E3F6F7",
};

// Formato colombiano
const pesos = (v) => "$" + Math.round(v).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const pct = (v) => Math.round(v * 100) + " %";

// ───────────────────────────── DISEÑOS ─────────────────────────────
const fondoOscuro = () => [
  { image: { x: 0, y: 0, w: 13.333, h: 7.5, path: path.join(ASSETS, "fondo_oficina.jpg"), transparency: 65 } },
  { image: { x: 10.2, y: 6.2, w: 2.72, h: 0.94, path: path.join(ASSETS, "logo_acontis_fondo_oscuro.png") } },
];
pres.defineSlideMaster({ title: "PORTADA", background: { color: HEX.navy }, objects: fondoOscuro() });
pres.defineSlideMaster({
  title: "SEPARADOR",
  background: { color: HEX.navy },
  objects: [
    ...fondoOscuro(),
    { placeholder: { options: { name: "numero", type: "body", x: 0.9, y: 1.5, w: 4, h: 1.3, fontSize: 96, bold: true, color: COL.pale, margin: 0, valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 3.0, w: 11, h: 1.0, fontSize: 52, bold: true, color: COL.white, margin: 0, align: "left" }, text: "" } },
    { placeholder: { options: { name: "subtitulo", type: "body", x: 0.9, y: 4.15, w: 10.5, h: 0.8, fontSize: 20, color: COL.sky, margin: 0, valign: "top" }, text: "" } },
  ],
  slideNumber: { x: 11.55, y: 7.05, w: 1.2, h: 0.3, fontSize: 10, color: COL.sky, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "seccion", type: "body", x: 0.6, y: 0.33, w: 9.8, h: 0.3, fontSize: 12, bold: true, color: COL.sea, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 10.1, h: 0.62, fontSize: 30, bold: true, color: COL.navy, margin: 0, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "bajada", type: "body", x: 0.6, y: 1.28, w: 12.13, h: 0.38, fontSize: 15, color: COL.muted, margin: 0, valign: "middle" }, text: "" } },
    { image: { x: 0.55, y: 6.86, w: 1.3, h: 0.45, path: path.join(ASSETS, "logo_acontis_fondo_claro.png") } },
    { text: { text: "Página", options: { x: 10.9, y: 7.03, w: 1.35, h: 0.3, fontSize: 10, color: COL.muted, align: "right", valign: "middle", margin: 0 } } },
  ],
  slideNumber: { x: 12.27, y: 7.03, w: 0.45, h: 0.3, fontSize: 10, color: COL.muted, align: "left", valign: "middle", margin: [0, 0, 0, 0.05] },
});

// ───────────────────────────── AYUDANTES ─────────────────────────────
let nObj = 0;
const id = (n) => `${n} ${++nObj}`;
let seccionActual = "Portada";

function contenido(numSeccion, titulo, bajada) {
  const nombre = SECCIONES.find((s) => s[0] === numSeccion);
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: seccionActual });
  s.addText(nombre ? `${nombre[0]} · ${nombre[1].toUpperCase()}` : numSeccion, { placeholder: "seccion" });
  s.addText(titulo, { placeholder: "title" });
  if (bajada) s.addText(bajada, { placeholder: "bajada" });
  s.addText(PILDORA, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.21, objectName: id("Píldora"),
    x: 10.6, y: 0.45, w: 2.12, h: 0.42, fill: { color: COL.card }, line: { color: COL.line, width: 1 },
    fontSize: 12, bold: true, color: COL.sea, align: "center", valign: "middle", margin: 0,
  });
  return s;
}

function separador(num, titulo, sub) {
  seccionActual = `${num}. ${titulo}`;
  pres.addSection({ title: seccionActual });
  const s = pres.addSlide({ masterName: "SEPARADOR", sectionTitle: seccionActual });
  s.addText(num, { placeholder: "numero" });
  s.addText(titulo, { placeholder: "title" });
  s.addText(sub, { placeholder: "subtitulo" });
  PAGINAS[num] = pres.slides.length;
  return s;
}
const PAGINAS = {};

function tarjeta(s, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, objectName: id(o.name || "Tarjeta"),
    fill: { color: o.fill || COL.card }, line: { color: o.line || COL.line, width: 1 },
  });
}
function texto(s, t, o) {
  s.addText(t, Object.assign({ isTextBox: true, margin: 0, color: COL.ink, fontSize: 13, valign: "top" }, o, { objectName: id(o.name || "Texto") }));
}
function vinetas(s, items, o) {
  const runs = items.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: o.espacio === undefined ? 5 : o.espacio } }));
  texto(s, runs, Object.assign({ fontSize: 12 }, o));
}
function circulo(s, x, y, d, label, size = 16, color = COL.navy) {
  s.addText(label, {
    shape: pres.shapes.OVAL, x, y, w: d, h: d, fill: { color }, line: { color, width: 0 },
    fontSize: size, bold: true, color: COL.white, align: "center", valign: "middle", margin: 0, objectName: id("Número"),
  });
}
function mensaje(s, t, y = 6.1) {
  tarjeta(s, 0.6, y, 12.13, 0.62, { name: "Mensaje clave", fill: COL.navy, line: COL.navy });
  texto(s, t, { name: "Mensaje clave texto", x: 0.9, y, w: 11.6, h: 0.62, fontSize: 14, bold: true, color: COL.white, valign: "middle" });
}
const TH = (t, o = {}) => ({ text: t, options: Object.assign({ fill: { color: COL.navy }, color: COL.white, bold: true, fontSize: 12, align: "left", valign: "middle" }, o) });
const TD = (t, o = {}) => ({ text: t, options: Object.assign({ color: COL.ink, fontSize: 12, valign: "middle" }, o) });
const BORDE = () => ({ type: "solid", pt: 1, color: "FFFFFF" });
const zebra = (i) => ({ color: i % 2 ? "F3F7FB" : "FFFFFF" });

// ───────────────────────────── PORTADA ─────────────────────────────
pres.addSection({ title: "Portada" });
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Portada" });
  texto(s, "Propuesta", { name: "Título 1", x: 0.9, y: 0.95, w: 11, h: 1.0, fontSize: 72, bold: true, color: COL.white, valign: "middle" });
  texto(s, "Consultoría Forense", { name: "Título 2", x: 0.9, y: 1.9, w: 11, h: 1.0, fontSize: 72, bold: true, color: COL.pale, valign: "middle" });
  texto(s, "Ministerio de Agricultura y Desarrollo Rural – MADR", { name: "Cliente", x: 0.9, y: 3.2, w: 11.5, h: 0.5, fontSize: 30, bold: true, color: COL.white, valign: "middle" });
  texto(s, "Propuesta comercial y económica · Consultoría especializada con enfoque forense para la evaluación de los recursos del Fondo Nacional del Arroz",
    { name: "Objeto", x: 0.9, y: 3.8, w: 10.8, h: 0.75, fontSize: 18, bold: true, color: COL.pale });
  texto(s, "Bogotá D.C., octubre de 2026", { name: "Fecha", x: 0.9, y: 4.8, w: 10, h: 0.45, fontSize: 24, color: COL.white, valign: "middle" });
  texto(s, "Preparado por Asesorías Contables del Caribe S.A.S. (ACONTIS)", { name: "Preparado por", x: 0.9, y: 6.55, w: 8.5, h: 0.3, fontSize: 13, color: COL.pale, valign: "middle" });
}

// ───────────────────────────── CONTENIDO ─────────────────────────────
const indice = contenido("CONTENIDO", "CONTENIDO", "Una propuesta construida a partir de los requerimientos técnicos definidos por el Ministerio.");
// (las filas del índice se agregan al final, cuando ya se conocen los números de página)

// ───────────────────────────── 01 ANTECEDENTES Y ENTENDIMIENTO ─────────────────────────────
separador("01", "Antecedentes y entendimiento", "Lo que el Ministerio necesita y cómo lo entendemos");
{
  const s = contenido("01", "ANTECEDENTES Y ENTENDIMIENTO",
    "El MADR requiere una revisión independiente para fortalecer su vigilancia administrativa sobre los recursos del Fondo.");
  const cols = [
    ["NATURALEZA DE LOS RECURSOS", [
      "Recursos de naturaleza parafiscal del Fondo Nacional del Arroz.",
      "Administrados bajo el Contrato de Administración de la Cuota de Fomento Arrocero No. 198 de 1996.",
      "Vigencias objeto de revisión: del 1 de enero de 2024 al 31 de diciembre de 2025.",
    ]],
    ["LO QUE REQUIERE EL MADR", [
      "Revisar el universo de información contable, financiera, contractual, presupuestal y de recaudo.",
      "Diferenciar los hechos acreditados de los que requieren comprobaciones adicionales.",
      "Hallazgos soportados con evidencia organizada y trazable.",
    ]],
    ["CÓMO LO ABORDAMOS", [
      "Auditoría financiera y contable, y revisión jurídica y contractual.",
      "Análisis patrimonial, investigación forense y analítica de datos.",
      "Revisión documental, verificación de operaciones, partes vinculadas, visitas de campo y expedientes de evidencia.",
    ]],
  ];
  cols.forEach(([t, items], i) => {
    const x = 0.6 + i * 4.115, y = 1.85, w = 3.9, h = 4.0;
    tarjeta(s, x, y, w, h, { name: t, fill: i === 2 ? COL.navy : COL.card, line: i === 2 ? COL.navy : COL.line });
    texto(s, t, { name: "Columna título", x: x + 0.3, y: y + 0.28, w: w - 0.6, h: 0.32, fontSize: 15, bold: true, color: i === 2 ? COL.pale : COL.sea });
    vinetas(s, items, { name: "Columna detalle", x: x + 0.3, y: y + 0.8, w: w - 0.55, h: h - 1.0, fontSize: 14, espacio: 10, color: i === 2 ? COL.white : COL.ink });
  });
  mensaje(s, "Nuestro propósito: establecer hechos verificables con su criterio, evidencia, incidencia y, cuando la información lo permita, su cuantificación.");
}
{
  const s = contenido("01", "CUATRO PREGUNTAS QUE GUÍAN LA CONSULTORÍA",
    "La metodología se orienta a responder, para cada situación, cuatro preguntas fundamentales.");
  const q = [
    ["1", "HECHO", "¿Qué ocurrió?", "La situación verificable, el período y las operaciones involucradas."],
    ["2", "CRITERIO", "¿Cuál era el criterio aplicable?", "El marco contractual, legal, financiero o administrativo contra el que se evalúa."],
    ["3", "EVIDENCIA", "¿Qué evidencia lo demuestra?", "Soportes suficientes, pertinentes y trazables, organizados por situación."],
    ["4", "EFECTO", "¿Cuál es su efecto y cómo se cuantifica?", "La posible incidencia y la cuantificación de la desviación o del posible perjuicio."],
  ];
  q.forEach(([n, et, preg, d], i) => {
    const x = 0.6 + i * 3.08, y = 1.85, w = 2.88, h = 4.0;
    tarjeta(s, x, y, w, h, { name: "Pregunta " + n });
    circulo(s, x + 0.3, y + 0.3, 0.7, n, 24);
    texto(s, et, { name: "Pregunta etiqueta", x: x + 1.15, y: y + 0.42, w: w - 1.3, h: 0.45, fontSize: 16, bold: true, color: COL.sea, valign: "middle" });
    texto(s, preg, { name: "Pregunta", x: x + 0.3, y: y + 1.25, w: w - 0.55, h: 1.05, fontSize: 20, bold: true, color: COL.navy });
    texto(s, d, { name: "Pregunta detalle", x: x + 0.3, y: y + 2.4, w: w - 0.55, h: 1.4, fontSize: 14, color: COL.ink });
  });
  mensaje(s, "No basta con identificar diferencias: cada situación se sustenta en procedimientos técnicos y evidencia verificable.");
}

// ───────────────────────────── 02 OBJETIVO ─────────────────────────────
separador("02", "Objetivo de la consultoría", "Una evaluación integral con enfoque forense de las vigencias 2024 y 2025");
{
  const s = contenido("02", "OBJETIVO DE LA CONSULTORÍA",
    "Una evaluación diseñada para entregar al Ministerio evidencia propia sobre el manejo de los recursos.");
  tarjeta(s, 0.6, 1.85, 12.13, 1.0, { name: "Objetivo", fill: COL.navy, line: COL.navy });
  texto(s, "Realizar una evaluación integral con enfoque forense sobre la administración, recaudo, manejo, ejecución y aplicación de los recursos del Fondo Nacional del Arroz durante las vigencias 2024 y 2025, mediante auditoría, análisis financiero, revisión contractual, analítica de datos y verificación documental y física.",
    { name: "Objetivo texto", x: 0.9, y: 1.85, w: 11.55, h: 1.0, fontSize: 14, color: COL.white, valign: "middle" });
  const pilares = [
    ["RECAUDO", ["Integridad del recaudo de la Cuota de Fomento Arrocero.", "Diferencias, omisiones, evasión o elusión identificables."]],
    ["APLICACIÓN DE LOS RECURSOS", ["Destinación legal y correspondencia con planes, programas y proyectos.", "Razonabilidad de operaciones financieras y contables.", "Contraprestación de administración y reservas."]],
    ["GOBIERNO Y SEPARACIÓN", ["Separación patrimonial y contable entre el Fondo y el administrador.", "Operaciones con terceros y partes relacionadas.", "Trazabilidad de las decisiones de los órganos de dirección."]],
    ["HALLAZGOS Y EVIDENCIA", ["Situaciones susceptibles de constituir hallazgos.", "Cuantificación de desviaciones o posibles perjuicios.", "Organización de la evidencia de cada situación."]],
  ];
  pilares.forEach(([t, items], i) => {
    const x = 0.6 + i * 3.08, y = 3.05, w = 2.88, h = 2.85;
    tarjeta(s, x, y, w, h, { name: "Pilar " + t });
    circulo(s, x + 0.25, y + 0.25, 0.5, String(i + 1), 16, COL.teal);
    texto(s, t, { name: "Pilar título", x: x + 0.9, y: y + 0.22, w: w - 1.05, h: 0.56, fontSize: 13, bold: true, color: COL.navy, valign: "middle" });
    vinetas(s, items, { name: "Pilar detalle", x: x + 0.25, y: y + 0.95, w: w - 0.45, h: h - 1.1, fontSize: 12, espacio: 6 });
  });
  mensaje(s, "El resultado es evidencia independiente y trazable que el MADR puede usar como soporte de sus actuaciones administrativas.");
}

// ───────────────────────────── 03 ALCANCE ─────────────────────────────
separador("03", "Alcance", "Seis componentes que cubren el ciclo completo de los recursos");
{
  const s = contenido("03", "ALCANCE",
    "Vigencias 2024 y 2025: del recaudo de la cuota a los mecanismos de gobierno y toma de decisiones.");
  const filas = [
    ["Recaudo de la Cuota de Fomento Arrocero", "Liquidación, retención, consignación y recaudo; conciliaciones; gestión de cartera y cobro; saldos e intereses; agentes retenedores omisos o inexactos y cruces de bases de información."],
    ["Gastos de funcionamiento e inversión", "Ejecución presupuestal, acuerdos de gasto, modificaciones y traslados; planes, programas y proyectos; contratos, pagos, proveedores y partes relacionadas; ejecución financiera frente a física."],
    ["Contraprestación de administración", "Base de cálculo, porcentajes aplicables, topes, liquidación, reconocimiento, pago y conciliación con los registros financieros."],
    ["Reservas del Fondo", "Constitución, saldos, movimientos, rendimientos, aplicación, destinación y soportes, frente a las decisiones adoptadas."],
    ["Separación patrimonial y contable", "Separación de recursos, operaciones cruzadas, costos compartidos, registros contables, transferencias, cuentas bancarias, conciliaciones y asignación de gastos."],
    ["Gobierno y administración", "Actas, decisiones, autorizaciones, acuerdos, informes y seguimiento; cumplimiento contractual, partes relacionadas, mecanismos de control y reportes al MADR."],
  ];
  const t = [[TH("COMPONENTE"), TH("QUÉ SE VERIFICA")]];
  filas.forEach(([c, d], i) => t.push([TD(c, { fill: zebra(i), bold: true, color: COL.navy, fontSize: 13 }), TD(d, { fill: zebra(i), fontSize: 12 })]));
  s.addTable(t, { x: 0.6, y: 1.85, w: 12.13, colW: [3.2, 8.93], rowH: [0.4, 0.62, 0.62, 0.55, 0.55, 0.55, 0.62], border: BORDE(), margin: [0.04, 0.12, 0.04, 0.12], objectName: "Tabla alcance" });
  mensaje(s, "Límite del encargo: la consultoría no sustituye el control fiscal ni imputa responsabilidades fiscales, disciplinarias o penales.");
}

// ───────────────────────────── 04 METODOLOGÍA ─────────────────────────────
separador("04", "Metodología y plan de trabajo", "Un modelo integrado de auditoría forense en cinco frentes");
{
  const s = contenido("04", "METODOLOGÍA · CINCO FRENTES DE TRABAJO",
    "Cada señal que detecta la analítica se corrobora con evidencia antes de convertirse en hallazgo.");
  const f = [
    ["Planeación y conocimiento", ["Marco normativo y Contrato No. 198 de 1996", "Estructura, procesos y flujos financieros", "Sistemas de información y controles", "Matriz inicial de riesgos y plan de procedimientos"]],
    ["Integridad y estructuración de información", ["Recepción e inventario de bases", "Validación de integridad y depuración", "Homologación e integración", "Conciliaciones y cruces entre fuentes"]],
    ["Analítica y detección de señales", ["Operaciones inusuales y pagos atípicos", "Duplicidades y concentraciones", "Proveedores recurrentes y terceros relacionados", "Diferencias entre fuentes"]],
    ["Investigación y corroboración", ["Revisión documental y contractual", "Confirmaciones y entrevistas", "Verificación física y visitas de campo", "Cruces externos y trazabilidad de pagos"]],
    ["Consolidación probatoria y cierre", ["Hecho, criterio y evidencia", "Análisis, cuantificación, riesgo e incidencia", "Respuesta del administrador y su evaluación", "Conclusión"]],
  ];
  f.forEach(([t, items], i) => {
    const x = 0.6 + i * 2.453, y = 1.85, w = 2.3, h = 4.05;
    tarjeta(s, x, y, w, h, { name: "Frente " + (i + 1), fill: i === 4 ? COL.navy : COL.card, line: i === 4 ? COL.navy : COL.line });
    circulo(s, x + 0.22, y + 0.22, 0.55, String(i + 1), 18, i === 4 ? COL.teal : COL.navy);
    texto(s, "FRENTE " + (i + 1), { name: "Frente número", x: x + 0.9, y: y + 0.3, w: w - 1.0, h: 0.4, fontSize: 12, bold: true, color: i === 4 ? COL.pale : COL.sea, valign: "middle" });
    texto(s, t, { name: "Frente título", x: x + 0.22, y: y + 0.95, w: w - 0.4, h: 0.75, fontSize: 15, bold: true, color: i === 4 ? COL.white : COL.navy });
    vinetas(s, items, { name: "Frente detalle", x: x + 0.22, y: y + 1.8, w: w - 0.37, h: 2.1, fontSize: 12, espacio: 6, color: i === 4 ? COL.white : COL.ink });
  });
  mensaje(s, "El informe final diferencia claramente las situaciones comprobadas de los asuntos que requieren actuaciones adicionales.");
}
{
  const s = contenido("04", "PLAN DE TRABAJO POR ETAPAS",
    "Cinco etapas sucesivas, desde la planeación hasta la verificación en campo.");
  const e = [
    ["Planeación", ["Reunión de inicio y canales de comunicación", "Revisión normativa y contractual", "Riesgos y diseño de procedimientos", "Protocolo de custodia y cronograma"], "Plan de Auditoría Detallado"],
    ["Recolección y estructuración", ["Recepción e inventario de bases", "Validación de integridad", "Conciliaciones e integración de fuentes", "Vacíos de información"], "Bases depuradas e integradas"],
    ["Investigación forense", ["Recaudo, cartera, gastos e inversión", "Contratación, pagos y bancos", "Contabilidad, patrimonio y reservas", "Contraprestación, gobierno y partes relacionadas"], "Resultados por componente"],
    ["Analítica de datos", ["Procesamiento del universo disponible", "Priorización de situaciones", "Resultados por componente", "Clasificación por nivel de criticidad"], "Alertas priorizadas"],
    ["Visitas de campo", ["Mínimo 4 visitas regionales incluidas", "Verificación de existencia y ejecución", "Inspección física y entrevistas", "Validación de beneficiarios y soportes"], "Informes de visita"],
  ];
  // línea de tiempo
  s.addShape(pres.shapes.LINE, { x: 1.2, y: 2.18, w: 10.95, h: 0, line: { color: COL.line, width: 3 }, objectName: id("Línea de tiempo") });
  e.forEach(([t, items, prod], i) => {
    const x = 0.6 + i * 2.453, w = 2.3;
    circulo(s, x + w / 2 - 0.33, 1.85, 0.66, String(i + 1), 22, i === 4 ? COL.teal : COL.navy);
    texto(s, "ETAPA " + (i + 1), { name: "Etapa número", x, y: 2.6, w, h: 0.28, fontSize: 12, bold: true, color: COL.sea, align: "center" });
    texto(s, t, { name: "Etapa título", x, y: 2.88, w, h: 0.62, fontSize: 15, bold: true, color: COL.navy, align: "center" });
    tarjeta(s, x, 3.55, w, 2.35, { name: "Etapa " + (i + 1) });
    vinetas(s, items, { name: "Etapa detalle", x: x + 0.18, y: 3.7, w: w - 0.3, h: 1.6, fontSize: 11, espacio: 4 });
    texto(s, [{ text: "Producto: ", options: { bold: true, color: COL.sea } }, { text: prod, options: { bold: true, color: COL.navy } }],
      { name: "Etapa producto", x: x + 0.18, y: 5.35, w: w - 0.3, h: 0.45, fontSize: 11, valign: "bottom" });
  });
  mensaje(s, "Incluye como mínimo cuatro (4) visitas regionales a plantas, sedes seccionales o proyectos del Fondo, dentro del valor de la propuesta.");
}
{
  const s = contenido("04", "DE LA SEÑAL AL HALLAZGO",
    "Cada situación relevante se documenta con la misma estructura, de modo que pueda verificarse de punta a punta.");
  const campos = [
    ["Hecho", "Qué ocurrió y qué operaciones involucra."],
    ["Criterio", "Norma, cláusula o parámetro aplicable."],
    ["Evidencia", "Soportes que demuestran el hecho."],
    ["Período", "Vigencia y fechas de la situación."],
    ["Análisis", "Evaluación técnica del hecho."],
    ["Cuantificación", "Efecto económico, cuando proceda."],
    ["Riesgo", "Nivel de criticidad de la situación."],
    ["Incidencia", "Posible efecto sobre los recursos."],
    ["Respuesta del administrador", "Su derecho de contradicción."],
    ["Evaluación de la respuesta", "Análisis de lo aportado."],
    ["Conclusión", "Situación comprobada o que requiere actuaciones adicionales."],
  ];
  campos.forEach(([t, d], i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.6 + col * 3.08, y = 1.85 + row * 1.42, w = 2.88, h = 1.27;
    tarjeta(s, x, y, w, h, { name: "Campo " + t });
    circulo(s, x + 0.2, y + 0.22, 0.46, String(i + 1), 14, i === 10 ? COL.teal : COL.navy);
    texto(s, t, { name: "Campo título", x: x + 0.8, y: y + 0.17, w: w - 0.95, h: 0.56, fontSize: 14, bold: true, color: COL.navy, valign: "middle" });
    texto(s, d, { name: "Campo detalle", x: x + 0.2, y: y + 0.78, w: w - 0.35, h: 0.45, fontSize: 12, color: COL.muted });
  });
  const x = 0.6 + 3 * 3.08, y = 1.85 + 2 * 1.42;
  tarjeta(s, x, y, 2.88, 1.27, { name: "Control", fill: COL.navy, line: COL.navy });
  texto(s, "Ningún hallazgo es definitivo sin evidencia, criterio, trazabilidad, cuantificación y análisis de la contradicción.",
    { name: "Control texto", x: x + 0.2, y, w: 2.5, h: 1.27, fontSize: 12, bold: true, color: COL.white, valign: "middle" });
}

// ───────────────────────────── 05 ENTREGABLES ─────────────────────────────
separador("05", "Entregables y cronograma", "Tres entregables base y un componente pericial condicional");
const ent = VALORES.entregables.map(([n, t, v, p]) => ({ n, t, v, iva: v * IVA, total: v * (1 + IVA), p }));
const dict = { v: VALORES.dictamen, iva: VALORES.dictamen * IVA, total: VALORES.dictamen * (1 + IVA) };
const base = { v: ent.reduce((a, e) => a + e.v, 0) };
base.iva = base.v * IVA; base.total = base.v * (1 + IVA);
{
  const s = contenido("05", "ENTREGABLES DE LA CONSULTORÍA",
    "Cada entregable es acumulativo y trazable; el informe final es autocontenido.");
  const filas = [
    [ent[0], "Metodología, cronograma semanal, responsables, matriz de riesgos, procedimientos por componente, protocolo de custodia, plan de visitas y matriz de información requerida.", "Hasta 8 días calendario tras el perfeccionamiento y ejecución del contrato"],
    [ent[1], "Resultados de análisis; cruces financieros, contables y contractuales; alertas y operaciones inusuales; cuantificaciones preliminares; matriz de criticidad y priorización de visitas.", "Hasta 20 días calendario tras la aprobación del Entregable 1"],
    [ent[2], "Informe ejecutivo, informe técnico por componente, matriz consolidada de hallazgos y expedientes de soporte con registro de custodia.", "Según el cronograma de la consultoría"],
  ];
  const t = [[TH("No.", { align: "center" }), TH("ENTREGABLE"), TH("CONTENIDO PRINCIPAL"), TH("PLAZO"), TH("VALOR TOTAL (IVA INCL.)", { align: "right" })]];
  filas.forEach(([e, c, pl], i) => t.push([
    TD(e.n, { fill: zebra(i), bold: true, color: COL.navy, align: "center", fontSize: 16 }),
    TD(e.t, { fill: zebra(i), bold: true, color: COL.navy }),
    TD(c, { fill: zebra(i), fontSize: 11 }),
    TD(pl, { fill: zebra(i), fontSize: 11, color: COL.muted }),
    TD(pesos(e.total), { fill: zebra(i), bold: true, color: COL.navy, align: "right" }),
  ]));
  t.push([
    TD("4", { fill: { color: "FDF3E3" }, bold: true, color: COL.amber, align: "center", fontSize: 16 }),
    TD("Dictamen pericial (condicional)", { fill: { color: "FDF3E3" }, bold: true, color: COL.navy }),
    TD("Solo con orden escrita del MADR: estudio del caso, evaluación de la evidencia, metodología, cuantificación, dictamen suscrito y sustentación ante la autoridad.", { fill: { color: "FDF3E3" }, fontSize: 11 }),
    TD("A solicitud del MADR", { fill: { color: "FDF3E3" }, fontSize: 11, color: COL.muted }),
    TD(pesos(dict.total) + " por dictamen", { fill: { color: "FDF3E3" }, bold: true, color: COL.navy, align: "right" }),
  ]);
  s.addTable(t, { x: 0.6, y: 1.85, w: 12.13, colW: [0.6, 2.6, 5.23, 1.95, 1.75], rowH: [0.4, 0.92, 0.92, 0.82, 0.92], border: BORDE(), margin: [0.05, 0.1, 0.05, 0.1], objectName: "Tabla entregables" });
  mensaje(s, `Valor base de la consultoría: ${pesos(base.total)} IVA incluido. El dictamen pericial solo se factura si el MADR lo ordena por escrito.`);
}
{
  const s = contenido("05", "ENTREGABLE 3 · INFORME FINAL Y EXPEDIENTES",
    "El producto principal de la consultoría: autocontenido y soportado en expedientes por situación.");
  const bloques = [
    { t: "INFORME EJECUTIVO", x: 0.6, y: 1.85, w: 3.9, h: 1.95, items: ["Resumen de resultados y principales situaciones", "Cuantificaciones y riesgos", "Conclusiones y recomendaciones"] },
    { t: "EXPEDIENTES DE SOPORTE", x: 0.6, y: 3.95, w: 3.9, h: 1.95, items: ["Identificación, fuente, fecha y descripción", "Relación con el hallazgo", "Registro de custodia y trazabilidad de acceso"] },
    { t: "INFORME TÉCNICO", x: 4.715, y: 1.85, w: 3.9, h: 4.05, items: ["Hallazgos por componente", "Hechos comprobados y criterio aplicable", "Evidencia y período", "Análisis y cuantificación", "Riesgo y presunta incidencia", "Respuesta a la contradicción y su evaluación"] },
  ];
  bloques.forEach((b) => {
    tarjeta(s, b.x, b.y, b.w, b.h, { name: b.t });
    texto(s, b.t, { name: "Bloque título", x: b.x + 0.3, y: b.y + 0.22, w: b.w - 0.5, h: 0.32, fontSize: 15, bold: true, color: COL.sea });
    vinetas(s, b.items, { name: "Bloque detalle", x: b.x + 0.3, y: b.y + 0.68, w: b.w - 0.5, h: b.h - 0.85, fontSize: 13, espacio: 6 });
  });
  const x = 8.83, y = 1.85, w = 3.9, h = 4.05;
  tarjeta(s, x, y, w, h, { name: "Matriz de hallazgos", fill: COL.navy, line: COL.navy });
  texto(s, "MATRIZ CONSOLIDADA DE HALLAZGOS", { name: "Matriz título", x: x + 0.3, y: y + 0.22, w: w - 0.5, h: 0.32, fontSize: 14, bold: true, color: COL.pale });
  const campos = ["Hallazgo", "Componente", "Fecha o período", "Hecho", "Criterio", "Evidencia", "Cuantificación", "Riesgo", "Norma relacionada", "Análisis de fraude", "Incidencia", "Respuesta del administrador", "Evaluación de la respuesta", "Recomendación"];
  [campos.slice(0, 7), campos.slice(7)].forEach((col, c) => {
    texto(s, col.map((t, i) => ({ text: `${c * 7 + i + 1}. ${t}`, options: { breakLine: i < col.length - 1, paraSpaceAfter: 6 } })),
      { name: "Matriz campos", x: x + 0.3 + c * 1.75, y: y + 0.75, w: 1.75, h: 3.1, fontSize: 11, color: COL.white });
  });
  mensaje(s, "Cada conclusión queda vinculada con la evidencia que la sustenta.");
}
{
  const s = contenido("05", "CRONOGRAMA DE ALTO NIVEL",
    "Catorce fases, del inicio a la sustentación; los entregables marcan los hitos de pago.");
  const fases = [
    ["Inicio, diagnóstico y planeación", "Entregable 1"], ["Recolección y estructuración", "Bases depuradas"], ["Análisis financiero y contable", "Resultados analíticos"],
    ["Revisión contractual y jurídica", "Matriz contractual"], ["Analítica forense", "Alertas"], ["Investigación y corroboración", "Casos priorizados"],
    ["Visitas de campo", "Informes de visita"], ["Informe preliminar", "Entregable 2"], ["Contradicción", "Matriz de respuestas"],
    ["Profundización", "Expedientes"], ["Consolidación", "Matriz definitiva"], ["Informe final", "Entregable 3"],
    ["Sustentación", "Presentación al MADR"], ["Dictamen", "Entregable 4, si se ordena"],
  ];
  [fases.slice(0, 7), fases.slice(7)].forEach((grupo, g) => {
    const t = [[TH("FASE", { align: "center" }), TH("ACTIVIDAD"), TH("PRODUCTO")]];
    grupo.forEach(([a, p], i) => {
      const n = g * 7 + i + 1, hito = /^Entregable [123]$/.test(p), cond = n === 14;
      const fill = hito ? { color: COL.tint } : cond ? { color: "FDF3E3" } : zebra(i);
      t.push([
        TD(String(n), { fill, bold: true, color: COL.navy, align: "center" }),
        TD(a, { fill, bold: hito }),
        TD(p, { fill, bold: hito, color: hito ? COL.sea : COL.ink }),
      ]);
    });
    s.addTable(t, { x: 0.6 + g * 6.165, y: 1.85, w: 5.965, colW: [0.7, 2.9, 2.365], rowH: 0.5, border: BORDE(), margin: [0, 0.1, 0, 0.1], objectName: "Tabla cronograma " + (g + 1) });
  });
  mensaje(s, `Hitos de pago: Entregable 1 (${pct(ent[0].p)}), Entregable 2 (${pct(ent[1].p)}) y Entregable 3 (${pct(ent[2].p)}).`);
}

// ───────────────────────────── 06 EQUIPO Y CALIDAD ─────────────────────────────
separador("06", "Equipo y aseguramiento de calidad", "Un equipo multidisciplinario y tres niveles de revisión");
{
  const s = contenido("06", "EQUIPO DE TRABAJO PROPUESTO",
    "Un equipo multidisciplinario alineado con los perfiles mínimos solicitados.");
  const eq = [
    ["Director de Consultoría Forense / Coord. General", "100 %", "Dirección integral y supervisión; aprobación de procedimientos, hallazgos e informes; interlocución con el MADR y control de calidad."],
    ["Coordinador de Analítica de Datos", "100 %", "Estructuración, integración y cruces de bases; identificación de anomalías; modelos de priorización y perfilamiento de riesgos."],
    ["Coordinador Jurídico y Contractual", "100 %", "Revisión contractual y normativa; obligaciones e incumplimientos; partes relacionadas; conceptos jurídicos asociados a los hallazgos."],
    ["Investigadores (2 profesionales)", "100 %", "Investigación documental, verificaciones, visitas y entrevistas; recolección de evidencia; análisis de terceros; apoyo en expedientes."],
    ["Auditor Financiero y Contable", "50 %", "Revisión contable, bancos, pagos, ingresos y gastos; conciliaciones y cuantificaciones."],
    ["Especialista Informático contable-forense", "50 %", "Información digital, integridad de datos, trazabilidad, custodia y evidencia digital; soporte tecnológico."],
    ["Especialista Legal", "50 %", "Análisis jurídico de la evidencia, la contratación y las actuaciones; apoyo en expedientes."],
  ];
  eq.forEach(([rol, ded, d], i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.6 + col * 3.08, y = 1.85 + row * 2.12, w = 2.88, h = 2.0;
    tarjeta(s, x, y, w, h, { name: "Rol " + rol });
    texto(s, rol, { name: "Rol", x: x + 0.22, y: y + 0.55, w: w - 0.4, h: 0.5, fontSize: 13, bold: true, color: COL.navy, valign: "middle" });
    s.addText(ded, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.13, x: x + 0.22, y: y + 0.18, w: 0.7, h: 0.3, fill: { color: ded === "100 %" ? COL.navy : COL.tint }, line: { color: ded === "100 %" ? COL.navy : COL.tint, width: 0 }, fontSize: 12, bold: true, color: ded === "100 %" ? COL.white : COL.sea, align: "center", valign: "middle", margin: 0, objectName: id("Dedicación") });
    texto(s, d, { name: "Rol detalle", x: x + 0.22, y: y + 1.12, w: w - 0.4, h: 0.82, fontSize: 11, color: COL.ink });
  });
  const x = 0.6 + 3 * 3.08, y = 1.85 + 2.12;
  tarjeta(s, x, y, 2.88, 2.0, { name: "Resumen equipo", fill: COL.navy, line: COL.navy });
  texto(s, "8", { name: "Equipo número", x: x + 0.22, y: y + 0.15, w: 2.5, h: 0.85, fontSize: 48, bold: true, color: COL.pale, valign: "middle" });
  texto(s, "profesionales: 5 con dedicación del 100 % y 3 con dedicación del 50 %.", { name: "Equipo texto", x: x + 0.22, y: y + 1.0, w: 2.5, h: 0.8, fontSize: 12, color: COL.white });
  texto(s, "Perfil del director: Contaduría Pública, Ingeniería de Sistemas, Estadística, Economía o afines, con posgrado y experiencia en auditoría, control, investigación forense y fraude.",
    { name: "Perfil director", x: 0.6, y: 6.1, w: 12.13, h: 0.55, fontSize: 12, italic: true, color: COL.muted, valign: "middle" });
}
{
  const s = contenido("06", "CONTROL DE CALIDAD Y CADENA DE CUSTODIA",
    "Tres niveles de revisión y una evidencia identificada y trazable de principio a fin.");
  texto(s, "SISTEMA DE CONTROL DE CALIDAD", { name: "Calidad título", x: 0.6, y: 1.85, w: 6.5, h: 0.35, fontSize: 15, bold: true, color: COL.sea });
  const niveles = [
    ["Nivel 1 · Equipo ejecutor", "Validación de papeles de trabajo y evidencia."],
    ["Nivel 2 · Líder de componente", "Revisión técnica de resultados."],
    ["Nivel 3 · Dirección de consultoría", "Revisión integral antes de comunicar resultados al MADR."],
  ];
  niveles.forEach(([t, d], i) => {
    const y = 2.3 + i * 0.95;
    tarjeta(s, 0.6, y, 6.5, 0.82, { name: t });
    circulo(s, 0.8, y + 0.15, 0.52, String(i + 1), 16, i === 2 ? COL.teal : COL.navy);
    texto(s, t, { name: "Nivel título", x: 1.5, y: y + 0.1, w: 5.4, h: 0.32, fontSize: 14, bold: true, color: COL.navy, valign: "middle" });
    texto(s, d, { name: "Nivel detalle", x: 1.5, y: y + 0.44, w: 5.4, h: 0.3, fontSize: 12, color: COL.muted, valign: "middle" });
  });
  tarjeta(s, 0.6, 5.2, 6.5, 0.72, { name: "Regla de calidad", fill: COL.tint, line: COL.tint });
  texto(s, "Ningún hallazgo es definitivo sin verificar evidencia, criterio, trazabilidad, cuantificación, respuesta del administrador y análisis de contradicción.",
    { name: "Regla de calidad texto", x: 0.85, y: 5.2, w: 6.05, h: 0.72, fontSize: 12, bold: true, color: COL.sea, valign: "middle" });

  const x = 7.45, w = 5.28;
  tarjeta(s, x, 1.85, w, 4.07, { name: "Cadena de custodia", fill: COL.navy, line: COL.navy });
  texto(s, "CADENA DE CUSTODIA", { name: "Custodia título", x: x + 0.35, y: 2.05, w: w - 0.6, h: 0.35, fontSize: 15, bold: true, color: COL.pale });
  texto(s, "Cada evidencia relevante contará con:", { name: "Custodia intro", x: x + 0.35, y: 2.45, w: w - 0.6, h: 0.3, fontSize: 12, color: COL.white });
  const attrs = ["Identificación", "Fuente", "Fecha de recepción", "Responsable", "Ubicación", "Relación con el hallazgo", "Registro de acceso", "Control de versiones"];
  [attrs.slice(0, 4), attrs.slice(4)].forEach((col, c) => {
    texto(s, col.map((t, i) => ({ text: "✓  " + t, options: { breakLine: i < col.length - 1, paraSpaceAfter: 9 } })),
      { name: "Custodia atributos", x: x + 0.35 + c * 2.45, y: 2.95, w: 2.4, h: 1.9, fontSize: 13, bold: true, color: COL.white });
  });
  texto(s, "La evidencia se organiza para asociarla directamente con el hallazgo o la situación que soporta.",
    { name: "Custodia nota", x: x + 0.35, y: 5.0, w: w - 0.6, h: 0.75, fontSize: 12, color: COL.sky });
  mensaje(s, "Resultados estructurados, trazables y técnicamente sustentados.");
}
{
  const s = contenido("06", "VALOR AGREGADO DE LA PROPUESTA",
    "Elementos que diferencian el servicio de una revisión financiera tradicional.");
  const v = [
    ["Cobertura integral", "Finanzas + contabilidad + contratación + jurídica + analítica + investigación + evidencia."],
    ["Enfoque 100 % orientado a datos", "La analítica procesa grandes volúmenes y detecta patrones que una revisión documental no mostraría."],
    ["Trazabilidad de hallazgos", "Cada conclusión queda vinculada con la evidencia que la sustenta."],
    ["Cuantificación", "Cuando la información lo permite, se cuantifica económicamente cada diferencia o desviación."],
    ["Visión preventiva y correctiva", "Recomendaciones para fortalecer los mecanismos de vigilancia administrativa."],
    ["Preparación para actuaciones", "Expedientes organizados que facilitan al MADR las actuaciones que considere procedentes."],
  ];
  v.forEach(([t, d], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.115, y = 1.85 + row * 2.1, w = 3.9, h = 1.95;
    tarjeta(s, x, y, w, h, { name: "Valor " + t, fill: i === 0 ? COL.navy : COL.card, line: i === 0 ? COL.navy : COL.line });
    circulo(s, x + 0.25, y + 0.25, 0.52, String(i + 1), 16, i === 0 ? COL.teal : COL.navy);
    texto(s, t, { name: "Valor título", x: x + 0.95, y: y + 0.22, w: w - 1.1, h: 0.58, fontSize: 16, bold: true, color: i === 0 ? COL.white : COL.navy, valign: "middle" });
    texto(s, d, { name: "Valor detalle", x: x + 0.25, y: y + 0.95, w: w - 0.45, h: 0.9, fontSize: 13, color: i === 0 ? COL.white : COL.ink });
  });
}

// ───────────────────────────── 07 VALOR ─────────────────────────────
separador("07", "Valor de la propuesta", "Valores en pesos colombianos · IVA del 19 %");
{
  const s = contenido("07", "VALOR DE LA PROPUESTA",
    "Incluye todos los recursos profesionales, técnicos, tecnológicos, logísticos y de campo necesarios.");
  const t = [[TH("No.", { align: "center" }), TH("ENTREGABLE"), TH("VALOR ANTES DE IVA", { align: "right" }), TH("IVA 19 %", { align: "right" }), TH("VALOR TOTAL", { align: "right" })]];
  ent.forEach((e, i) => t.push([
    TD(e.n, { fill: zebra(i), bold: true, color: COL.navy, align: "center" }), TD(e.t, { fill: zebra(i) }),
    TD(pesos(e.v), { fill: zebra(i), align: "right" }), TD(pesos(e.iva), { fill: zebra(i), align: "right", color: COL.muted }),
    TD(pesos(e.total), { fill: zebra(i), align: "right", bold: true, color: COL.navy }),
  ]));
  const tot = { fill: { color: COL.sea }, bold: true, color: COL.white };
  t.push([TD("", tot), TD("TOTAL CONSULTORÍA", tot), TD(pesos(base.v), Object.assign({ align: "right" }, tot)), TD(pesos(base.iva), Object.assign({ align: "right" }, tot)), TD(pesos(base.total), Object.assign({ align: "right" }, tot))]);
  const co = { fill: { color: "FDF3E3" } };
  t.push([TD("4", Object.assign({ bold: true, color: COL.amber, align: "center" }, co)), TD("Dictamen pericial – condicional (por dictamen)", co),
    TD(pesos(dict.v), Object.assign({ align: "right" }, co)), TD(pesos(dict.iva), Object.assign({ align: "right", color: COL.muted }, co)), TD(pesos(dict.total), Object.assign({ align: "right", bold: true, color: COL.navy }, co))]);
  s.addTable(t, { x: 0.6, y: 1.85, w: 12.13, colW: [0.7, 5.33, 2.1, 1.9, 2.1], rowH: 0.42, border: BORDE(), margin: [0, 0.12, 0, 0.12], objectName: "Tabla valor" });

  const cards = [
    { t: "VALOR BASE DE LA CONSULTORÍA", v: pesos(base.total), d: `IVA incluido · ${pesos(base.v)} antes de IVA`, dark: true },
    { t: "DICTAMEN PERICIAL (CONDICIONAL)", v: pesos(dict.total), d: "Por dictamen, IVA incluido · solo con orden escrita del MADR" },
    { t: "VALOR MÁXIMO EVENTUAL", v: pesos(base.total + dict.total), d: "IVA incluido · consultoría más un dictamen pericial" },
  ];
  cards.forEach((c, i) => {
    const x = 0.6 + i * 4.115, y = 4.6, w = 3.9, h = 1.62;
    tarjeta(s, x, y, w, h, { name: c.t, fill: c.dark ? COL.navy : COL.card, line: c.dark ? COL.navy : COL.line });
    texto(s, c.t, { name: "Valor etiqueta", x: x + 0.25, y: y + 0.17, w: w - 0.4, h: 0.3, fontSize: 12, bold: true, color: c.dark ? COL.pale : COL.muted });
    texto(s, c.v, { name: "Valor cifra", x: x + 0.25, y: y + 0.5, w: w - 0.4, h: 0.6, fontSize: 30, bold: true, color: c.dark ? COL.white : COL.navy, valign: "middle" });
    texto(s, c.d, { name: "Valor detalle", x: x + 0.25, y: y + 1.15, w: w - 0.4, h: 0.35, fontSize: 11, color: c.dark ? COL.sky : COL.muted });
  });
  texto(s, `Si el formulario solicita el valor total de la consultoría sin el componente condicional, el valor a registrar es ${pesos(base.total)} IVA incluido.`,
    { name: "Nota formulario", x: 0.6, y: 6.35, w: 12.13, h: 0.35, fontSize: 12, italic: true, color: COL.muted, valign: "middle" });
}
{
  const s = contenido("07", "FORMA DE PAGO Y CONDICIONES",
    "El esquema económico base mantiene la lógica de pago definida por el MADR.");
  // un pago por entregable
  const nombres = ["Primer pago", "Segundo pago", "Tercer pago"];
  ent.forEach((e, i) => {
    const x = 0.6 + i * 4.115, y = 1.85, w = 3.9, h = 1.25;
    const dark = i === 2;
    tarjeta(s, x, y, w, h, { name: "Pago " + (i + 1), fill: dark ? COL.navy : COL.card, line: dark ? COL.navy : COL.line });
    texto(s, pct(e.p), { name: "Pago porcentaje", x: x + 0.25, y, w: 1.3, h, fontSize: 34, bold: true, color: dark ? COL.pale : COL.teal, valign: "middle" });
    texto(s, [
      { text: `${nombres[i]} · Entregable ${e.n}`, options: { fontSize: 12, bold: true, color: dark ? COL.sky : COL.sea, breakLine: true } },
      { text: pesos(e.total), options: { fontSize: 22, bold: true, color: dark ? COL.white : COL.navy, breakLine: true } },
      { text: "IVA incluido", options: { fontSize: 11, color: dark ? COL.sky : COL.muted } },
    ], { name: "Pago detalle", x: x + 1.6, y, w: w - 1.75, h, valign: "middle" });
  });
  const bloques = [
    ["GASTOS INCLUIDOS", ["Honorarios, personal técnico y analítica de datos", "Herramientas tecnológicas, insumos y papeles de trabajo", "Desplazamientos, viáticos y las 4 visitas regionales", "Informes, presentación de resultados y custodia documental", "Costos administrativos, impuestos, tasas y contribuciones"]],
    ["SUPUESTOS PARA LA EJECUCIÓN", ["Entrega oportuna de la información por el MADR y el administrador", "Formatos procesables y acceso a los documentos", "Facilidades para las verificaciones requeridas", "Visitas coordinadas con la supervisión", "Pericial solo con orden escrita del MADR"]],
    ["LIMITACIONES DEL SERVICIO", ["Actuación dentro de las competencias del contratista", "No asume funciones de la Contraloría ni de autoridades disciplinarias o penales", "La determinación de responsabilidades corresponde a las autoridades competentes"]],
  ];
  bloques.forEach(([t, items], i) => {
    const bx = 0.6 + i * 4.115, y = 3.3, w = 3.9, h = 2.65;
    tarjeta(s, bx, y, w, h, { name: t });
    texto(s, t, { name: "Condición título", x: bx + 0.25, y: y + 0.18, w: w - 0.4, h: 0.32, fontSize: 14, bold: true, color: COL.sea });
    vinetas(s, items, { name: "Condición detalle", x: bx + 0.25, y: y + 0.6, w: w - 0.4, h: h - 0.7, fontSize: 11, espacio: 4 });
  });
  mensaje(s, "Vigencia de la propuesta: noventa (90) días calendario desde su presentación, salvo que el proceso establezca otro término.");
}

// ───────────────────────────── CIERRE ─────────────────────────────
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: seccionActual });
  texto(s, "¡Gracias!", { name: "Cierre título", x: 0.9, y: 1.4, w: 10, h: 1.1, fontSize: 80, bold: true, color: COL.white, valign: "middle" });
  texto(s, "Un instrumento técnico, no solo un informe de auditoría", { name: "Cierre subtítulo", x: 0.9, y: 2.75, w: 11, h: 0.5, fontSize: 26, bold: true, color: COL.pale, valign: "middle" });
  texto(s, "ACONTIS pone a disposición del Ministerio de Agricultura y Desarrollo Rural un equipo multidisciplinario para desarrollar una revisión integral, independiente y orientada a evidencia: comprender el comportamiento de los recursos, identificar situaciones relevantes, cuantificar sus efectos y fortalecer los mecanismos de vigilancia administrativa.",
    { name: "Cierre texto", x: 0.9, y: 3.45, w: 10.6, h: 1.5, fontSize: 17, color: COL.white });
  texto(s, "ASESORÍAS CONTABLES DEL CARIBE S.A.S. – ACONTIS", { name: "Cierre firma", x: 0.9, y: 5.3, w: 9, h: 0.4, fontSize: 18, bold: true, color: COL.pale, valign: "middle" });
}

// ───────────────────────────── ÍNDICE (se completa al final) ─────────────────────────────
{
  const s = indice;
  SECCIONES.forEach(([n, t], i) => {
    const y = 1.85 + i * 0.6, h = 0.5;
    tarjeta(s, 0.6, y, 7.6, h, { name: "Índice " + n });
    texto(s, n, { name: "Índice número", x: 0.85, y, w: 0.7, h, fontSize: 18, bold: true, color: COL.teal, valign: "middle" });
    texto(s, t, { name: "Índice título", x: 1.65, y, w: 5.5, h, fontSize: 16, bold: true, color: COL.navy, valign: "middle" });
    texto(s, String(PAGINAS[n]), { name: "Índice página", x: 7.25, y, w: 0.75, h, fontSize: 14, color: COL.muted, align: "right", valign: "middle" });
  });
  const x = 8.55, w = 4.18;
  tarjeta(s, x, 1.85, w, 4.1, { name: "Datos clave", fill: COL.navy, line: COL.navy });
  texto(s, "EN RESUMEN", { name: "Resumen título", x: x + 0.3, y: 2.05, w: w - 0.5, h: 0.3, fontSize: 13, bold: true, color: COL.pale });
  const datos = [
    ["Vigencias", "2024 y 2025"],
    ["Componentes del alcance", "6"],
    ["Entregables base", "3 + dictamen condicional"],
    ["Valor base IVA incluido", pesos(base.total)],
  ];
  datos.forEach(([l, v], i) => {
    const y = 2.5 + i * 0.85;
    texto(s, l, { name: "Resumen etiqueta", x: x + 0.3, y, w: w - 0.5, h: 0.28, fontSize: 12, color: COL.sky });
    texto(s, v, { name: "Resumen valor", x: x + 0.3, y: y + 0.28, w: w - 0.5, h: 0.45, fontSize: 22, bold: true, color: COL.white, valign: "middle" });
  });
}

// ───────────────────────────── ESCRITURA ─────────────────────────────
async function aplicarTema(archivo) {
  const zip = await JSZip.loadAsync(fs.readFileSync(archivo));
  for (const name of Object.keys(zip.files)) {
    if (!/^ppt\/theme\/theme\d+\.xml$/.test(name)) continue;
    let x = await zip.file(name).async("string");
    const clr = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
      .map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join("");
    x = x.replace(/<a:clrScheme name="[^"]*">[\s\S]*?<\/a:clrScheme>/, `<a:clrScheme name="${THEME.name}">${clr}</a:clrScheme>`);
    x = x.replace(/(<a:majorFont>\s*<a:latin typeface=")[^"]*"/, `$1${THEME.headFontFace}"`);
    x = x.replace(/(<a:minorFont>\s*<a:latin typeface=")[^"]*"/, `$1${THEME.bodyFontFace}"`);
    zip.file(name, x);
  }
  fs.writeFileSync(archivo, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

(async () => {
  await pres.writeFile({ fileName: SALIDA });
  await aplicarTema(SALIDA);
  console.log("Presentación generada:", SALIDA);
})();
