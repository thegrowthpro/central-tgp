const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'Repaso SDR Junior — TGP';

// ─── PALETA ───────────────────────────────────────────────────────────────────
const C = {
  dark:     "1C1C1C",
  white:    "FFFFFF",
  card:     "F5F5F5",
  cardBdr:  "E0E0E0",
  green:    "00C853",
  greenDk:  "00953D",
  yellow:   "FFD600",
  gray:     "9E9E9E",
  darkGray: "424242",
  textDark: "231F20",
};

// Etiquetas de correo (colores reales)
const ETIQUETAS = {
  NO_INTERESADO: "C62828",
  PROCESO:       "F9A825",
  AGENDADO:      "2E7D32",
  REAGENDAR:     "1565C0",
  NO_APLICA:     "E65100",
};

// Imágenes extraídas del PPTX original
const IMG = {
  capturas0:      "/tmp/tgp_slide_imgs/capturas_1_image3.png",
  capturas1:      "/tmp/tgp_slide_imgs/capturas_2_image2.png",
  capturas2:      "/tmp/tgp_slide_imgs/capturas_3_image4.png",
  barrido:        "/tmp/tgp_slide_imgs/barrido_1_image12.png",
  aprobacion:     "/tmp/tgp_slide_imgs/aprobacion_0_image8.png",  // slack channel screenshot
  aprobSlack:     "/tmp/tgp_slide_imgs/aprobacion_1_image13.png",
  aprobEmoji:     "/tmp/tgp_slide_imgs/aprobacion_2_image14.png",
  detalleWA:      "/tmp/tgp_slide_imgs/detalle_1_image16.png",
  detalleSlack:   "/tmp/tgp_slide_imgs/detalle_2_image15.png",
  confirmacion:   "/tmp/tgp_slide_imgs/confirmacion_1_image18.png",
  recordatorio1:  "/tmp/tgp_slide_imgs/recordatorios_1_image19.png",
  recordatorio2:  "/tmp/tgp_slide_imgs/recordatorios_2_image20.png",
  recordatorio3:  "/tmp/tgp_slide_imgs/recordatorios_3_image21.png",
  etiqueta1:      "/tmp/tgp_slide_imgs/etiqueta_1_image23.png",
  etiqueta2:      "/tmp/tgp_slide_imgs/etiqueta_2_image24.png",
  etiqueta3:      "/tmp/tgp_slide_imgs/etiqueta_3_image22.png",
  etiqueta4:      "/tmp/tgp_slide_imgs/etiqueta_4_image26.png",
  noshowIA1:      "/tmp/tgp_slide_imgs/noshow_1_image25.png",
  noshowIA2:      "/tmp/tgp_slide_imgs/noshow_2_image27.png",
  noshowIA3:      "/tmp/tgp_slide_imgs/noshow_3_image28.png",
  diagCorreos:    "/tmp/diagrama_correos-1.jpg",
  diagAgend:      "/tmp/diagrama_agendamiento-1.jpg",
};

const mkShadow     = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.10 });
const mkCardShadow = () => ({ type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.08 });

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function logoLight(slide) {
  const LOGO = "/Users/belenforti/Desktop/Escritorio - MacBook Air de Belen/Logos TGP/Diseño sin título (3).svg";
  slide.addImage({ path: LOGO, x: 0.35, y: 0.18, w: 0.7, h: 0.23 });
}

function logoDark(slide) {
  slide.addText("TGP", {
    x: 0.35, y: 0.18, w: 0.7, h: 0.28,
    fontSize: 18, bold: true, color: C.white, fontFace: "Arial Black", margin: 0
  });
}

function slideNum(slide, n, dark = false) {
  slide.addText(`${n}`, {
    x: 9.6, y: 5.28, w: 0.3, h: 0.2,
    fontSize: 9, color: dark ? "555555" : "BBBBBB",
    align: "right", margin: 0
  });
}

function lightSlide(n) {
  const sl = pres.addSlide();
  sl.background = { color: C.white };
  logoLight(sl);
  slideNum(sl, n);
  return sl;
}

function darkSlide(n) {
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  slideNum(sl, n, true);
  return sl;
}

function sectionDivider(n, icon, title, sub) {
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.green }, line: { color: C.green } });
  logoDark(sl);
  sl.addText(icon, { x: 0.5, y: 1.7, w: 1.5, h: 1.2, fontSize: 56, align: "center", margin: 0 });
  sl.addText(title, { x: 0.5, y: 3.0, w: 9.0, h: 0.9, fontSize: 36, bold: true, color: C.white, fontFace: "Arial Black", margin: 0 });
  if (sub) sl.addText(sub, { x: 0.5, y: 3.95, w: 8.5, h: 0.45, fontSize: 15, color: C.gray, margin: 0 });
  slideNum(sl, n, true);
  return sl;
}

// Pre-quiz slide: dark, yellow accent, questions listed
function preQuiz(n, title, questions) {
  const sl = pres.addSlide();
  sl.background = { color: "111111" };
  logoDark(sl);
  // Yellow top bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.yellow }, line: { color: C.yellow } });
  sl.addText("❓  ANTES DE VER:", { x: 0.5, y: 0.15, w: 6, h: 0.38, fontSize: 11, bold: true, color: C.yellow, charSpacing: 2, margin: 0 });
  sl.addText(title, { x: 0.5, y: 0.6, w: 9.1, h: 0.65, fontSize: 26, bold: true, color: C.white, fontFace: "Arial Black", margin: 0, wrap: true });
  const items = questions.map((q, i) => ({
    text: `${i + 1}.  ${q}`,
    options: { breakLine: true, paraSpaceAfter: 10, fontSize: 16, color: "DDDDDD" }
  }));
  items[items.length - 1].options.breakLine = false;
  sl.addText(items, { x: 0.5, y: 1.42, w: 9.1, h: 3.9, valign: "top", margin: 0 });
  sl.addText("Ronda oral — todos participan", { x: 0.5, y: 5.2, w: 9, h: 0.3, fontSize: 11, color: "555555", margin: 0 });
  slideNum(sl, n, true);
  return sl;
}

// Regular quiz (for final section)
function quizSlide(n, questions) {
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("🏆  QUIZ FINAL", { x: 0.5, y: 0.15, w: 5, h: 0.38, fontSize: 11, bold: true, color: C.green, charSpacing: 3, margin: 0 });
  const items = questions.map((q, i) => ({
    text: `${i + 1}.  ${q}`,
    options: { breakLine: true, paraSpaceAfter: 14, fontSize: 17, color: C.white }
  }));
  items[items.length - 1].options.breakLine = false;
  sl.addText(items, { x: 0.5, y: 0.65, w: 9.1, h: 4.7, valign: "top", margin: 0 });
  slideNum(sl, n, true);
  return sl;
}

function card(slide, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: C.card }, line: { color: C.cardBdr, width: 0.5 }, shadow: mkCardShadow()
  });
}

function accentCard(slide, x, y, w, h, color) {
  card(slide, x, y, w, h);
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color }, line: { color } });
}

// Dark image slide: full-width image with title
function imageSlide(n, title, imgPath, imgX, imgY, imgW, imgH) {
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText(title, { x: 0.5, y: 0.5, w: 9.1, h: 0.65, fontSize: 28, bold: true, color: C.white, fontFace: "Arial Black", margin: 0 });
  sl.addImage({ path: imgPath, x: imgX, y: imgY, w: imgW, h: imgH });
  slideNum(sl, n, true);
  return sl;
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDES
// ─────────────────────────────────────────────────────────────────────────────

// ─── 1: PORTADA ──────────────────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.22, w: 10, h: 0.4, fill: { color: C.green }, line: { color: C.green } });
  sl.addText("TGP", { x: 0.5, y: 0.45, w: 4, h: 1.1, fontSize: 64, bold: true, color: C.white, fontFace: "Arial Black", margin: 0 });
  sl.addText("Ya llegaste.", { x: 0.5, y: 1.55, w: 9, h: 0.85, fontSize: 46, bold: true, color: C.white, fontFace: "Arial Black", margin: 0 });
  sl.addText("Ahora empieza lo bueno.", { x: 0.5, y: 2.35, w: 9, h: 0.85, fontSize: 46, bold: true, color: C.green, fontFace: "Arial Black", margin: 0 });
  sl.addText("Repaso SDR Junior · Sesión de la tarde", { x: 0.5, y: 3.3, w: 9, h: 0.4, fontSize: 15, color: C.gray, margin: 0 });
  slideNum(sl, 1, true);
}

// ─── 2: AGENDA ───────────────────────────────────────────────────────────────
{
  const sl = lightSlide(2);
  sl.addText("¿Qué hacemos hoy?", { x: 0.5, y: 0.55, w: 9, h: 0.65, fontSize: 30, bold: true, color: C.textDark, fontFace: "Arial Black", margin: 0 });
  const items = [
    { n:"01", icon:"🏛️", label:"Pilares TGP",           t:"10 min" },
    { n:"02", icon:"📖", label:"Glosario clave",         t:"10 min" },
    { n:"03", icon:"🔄", label:"Proceso de prospección", t:"10 min" },
    { n:"04", icon:"📧", label:"Bandejas + Paso a paso", t:"25 min" },
    { n:"05", icon:"📅", label:"Agendamiento",           t:"20 min" },
    { n:"06", icon:"🏆", label:"IA + Quiz + Accesos",    t:"15 min" },
  ];
  const cw=2.95, ch=1.5, gx=0.18, gy=0.22, startX=0.5, startY=1.35;
  items.forEach((item, i) => {
    const x = startX + (i%3)*(cw+gx), y = startY + Math.floor(i/3)*(ch+gy);
    card(sl, x, y, cw, ch);
    sl.addText(item.n, { x:x+0.15, y:y+0.12, w:0.5, h:0.3, fontSize:11, bold:true, color:C.green, margin:0 });
    sl.addText(item.icon, { x:x+cw-0.55, y:y+0.08, w:0.45, h:0.35, fontSize:20, align:"center", margin:0 });
    sl.addText(item.label, { x:x+0.15, y:y+0.52, w:cw-0.3, h:0.55, fontSize:15, bold:true, color:C.textDark, margin:0, wrap:true });
    sl.addText(item.t, { x:x+0.15, y:y+1.15, w:cw-0.3, h:0.25, fontSize:11, color:C.gray, margin:0 });
  });
}

// ─── 3: CHECK-IN ─────────────────────────────────────────────────────────────
{
  const sl = darkSlide(3);
  sl.addText("💬", { x:4.3, y:0.8, w:1.5, h:1.2, fontSize:60, align:"center", margin:0 });
  sl.addText("¿Qué fue lo que más\nles llamó la atención\ndel onboarding?", {
    x:0.9, y:1.95, w:8.3, h:2.2, fontSize:32, bold:true, color:C.white, fontFace:"Arial Black", align:"center", margin:0
  });
  sl.addText("Ronda abierta — todos participan", { x:2, y:4.25, w:6, h:0.4, fontSize:13, color:C.gray, align:"center", margin:0 });
}

// ════════════════════════════════════════════════════════════════════════════
// SECCIÓN 1: PILARES
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(4, "🏛️", "Los Pilares TGP", null);

// ─── 5: PRE-QUIZ PILARES ─────────────────────────────────────────────────────
preQuiz(5, "¿Cuáles son los 5 pilares TGP?", [
  "¿Cuántos son y cómo se llaman?",
  "¿Qué significa el pilar de Comunicación en el día a día?",
  "¿Cada cuánto hay que hacer push a cada prospecto?",
  "¿Cuántos días mínimo hay que ir a la oficina?",
]);

// ─── 6: PILARES ──────────────────────────────────────────────────────────────
{
  const sl = lightSlide(6);
  sl.addText("Los 5 Pilares", { x:0.5, y:0.55, w:9, h:0.6, fontSize:30, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const pilares = [
    { icon:"💬", name:"Comunicación",   desc:"Slack al día, sin mensajes sin leer. Grupos de clientes, equipo, RRHH y Manager." },
    { icon:"📅", name:"Organización",   desc:"Actividades y horarios claros. Google Calendar actualizado siempre." },
    { icon:"🏢", name:"Presencialidad", desc:"Asistencia a la oficina mínimo 2 veces a la semana." },
    { icon:"🔄", name:"Constancia",     desc:"Bandejas al día. Push mínimo 2 veces por semana a cada prospecto." },
    { icon:"🎯", name:"Responsabilidad",desc:"Todas las reuniones, ausencias a tiempo, cumplir metas semanales." },
  ];
  const cw=1.7, ch=3.9, gap=0.1, startX=0.38, y=1.1;
  pilares.forEach((p, i) => {
    const x = startX + i*(cw+gap);
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w:cw, h:ch, fill:{color: i===0?C.green:C.card}, line:{color: i===0?C.greenDk:C.cardBdr, width:0.5}, shadow:mkCardShadow() });
    sl.addShape(pres.shapes.OVAL, { x:x+cw/2-0.32, y:y+0.22, w:0.64, h:0.64, fill:{color: i===0?"FFFFFF":C.green}, line:{color: i===0?"FFFFFF":C.greenDk} });
    sl.addText(p.icon, { x:x+cw/2-0.32, y:y+0.22, w:0.64, h:0.64, fontSize:22, align:"center", valign:"middle", margin:0 });
    sl.addText(p.name, { x:x+0.1, y:y+1.0, w:cw-0.2, h:0.55, fontSize:13, bold:true, color:i===0?C.white:C.textDark, align:"center", margin:0, wrap:true });
    sl.addText(p.desc, { x:x+0.12, y:y+1.6, w:cw-0.24, h:2.15, fontSize:10.5, color:i===0?"D0FFE8":C.darkGray, align:"center", margin:0, wrap:true });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SECCIÓN 2: GLOSARIO
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(7, "📖", "Glosario Clave", null);

// ─── 8: PRE-QUIZ GLOSARIO ────────────────────────────────────────────────────
preQuiz(8, "¿Sabes qué significa cada uno?", [
  "¿Qué es un Prospecto? ¿Y un Cliente?",
  "¿Qué es el ICP y para qué sirve?",
  "¿Quién es el KDM?",
  "¿Qué es un Push?",
  "¿Qué es un No Show?",
]);

// ─── 9: GLOSARIO ─────────────────────────────────────────────────────────────
{
  const sl = lightSlide(9);
  sl.addText("Términos clave", { x:0.5, y:0.55, w:9, h:0.6, fontSize:30, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const terms = [
    { t:"Prospecto",      d:"Potencial cliente al que contactamos para agendar reunión" },
    { t:"Cliente",        d:"Empresa que nos contrata y dispone de ejecutivos para reuniones" },
    { t:"ICP",            d:"Ideal Customer Profile — Perfil de cliente ideal" },
    { t:"KDM",            d:"Key Decision Maker — Tomador de decisión en la empresa" },
    { t:"B2B",            d:"Business to Business — Empresa que vende a otras empresas" },
    { t:"Push",           d:"Correo/llamado/WhatsApp de insistencia a un prospecto" },
    { t:"Call to Action", d:"Frase que indica el siguiente paso: agendar, dar info, etc." },
    { t:"No Show",        d:"Prospecto que no se presenta a la reunión agendada" },
  ];
  const cw=4.4, ch=0.82, gap=0.08, col2X=5.1;
  terms.forEach((term, i) => {
    const x = i%2===0?0.5:col2X, y = 1.35+Math.floor(i/2)*(ch+gap);
    accentCard(sl, x, y, cw, ch, C.green);
    sl.addText(term.t, { x:x+0.18, y:y+0.06, w:cw-0.28, h:0.3, fontSize:12, bold:true, color:C.textDark, margin:0 });
    sl.addText(term.d, { x:x+0.18, y:y+0.4, w:cw-0.28, h:0.38, fontSize:10, color:C.darkGray, margin:0, wrap:true });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SECCIÓN 3: PROCESO
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(10, "🔄", "El Proceso de Prospección", null);

// ─── 11: PRE-QUIZ PROCESO ────────────────────────────────────────────────────
preQuiz(11, "¿Cómo funciona el proceso de prospección?", [
  "¿Quién define el ICP/KDM/Foco?",
  "¿Quién arma la base de datos?",
  "¿Quién envía las secuencias de correo?",
  "¿Cuándo entra el SDR al proceso?",
  "¿Qué hace el SDR Senior? ¿Y el CSR?",
]);

// ─── 12: PROCESO ─────────────────────────────────────────────────────────────
{
  const sl = lightSlide(12);
  sl.addText("El proceso de prospección", { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const nodes = [
    { label:"Cliente",          icon:"🏢", color:C.dark       },
    { label:"Célula / Pod",     icon:"🎯", color:"1565C0"      },
    { label:"Base de datos",    icon:"🗃️", color:"6A1B9A"      },
    { label:"Secuencia correo", icon:"📧", color:C.greenDk     },
    { label:"Respuesta",        icon:"💬", color:"E65100"      },
    { label:"SDR (tú) ⭐",      icon:"⭐", color:C.green       },
    { label:"Manager",          icon:"👤", color:"455A64"      },
    { label:"IA",               icon:"🤖", color:"880E4F"      },
    { label:"CSR",              icon:"🤝", color:"1B5E20"      },
  ];
  const boxW=0.9, boxH=0.72, gap=0.1, total=nodes.length;
  const totalW=total*boxW+(total-1)*gap, startX=(10-totalW)/2, y=1.6;
  nodes.forEach((nd, i) => {
    const x = startX+i*(boxW+gap);
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w:boxW, h:boxH, fill:{color:nd.color}, line:{color:nd.color}, shadow:mkCardShadow() });
    sl.addText(nd.icon, { x, y:y+0.05, w:boxW, h:0.3, fontSize:16, align:"center", margin:0 });
    sl.addText(nd.label, { x, y:y+0.38, w:boxW, h:0.3, fontSize:7.5, bold:true, color:C.white, align:"center", margin:0, wrap:true });
    if (i<total-1) sl.addShape(pres.shapes.LINE, { x:x+boxW, y:y+boxH/2, w:gap, h:0, line:{color:C.gray, width:1} });
  });
  // Highlight SDR
  const sdrX = startX+5*(boxW+gap);
  sl.addShape(pres.shapes.RECTANGLE, { x:sdrX-0.06, y:y-0.06, w:boxW+0.12, h:boxH+0.12, fill:{color:"FFFFFF",transparency:100}, line:{color:C.yellow, width:2.5} });
  sl.addText("Tu punto de entrada en el proceso", { x:0.5, y:3.3, w:9, h:0.4, fontSize:13, color:C.darkGray, align:"center", italic:true, margin:0 });
  sl.addText("Tus 2 herramientas principales:", { x:0.5, y:3.78, w:9, h:0.32, fontSize:12, bold:true, color:C.textDark, align:"center", margin:0 });
  const tools = ["📧  Bandejas de correo electrónico", "📞  Base de datos con teléfonos"];
  tools.forEach((tool, i) => {
    card(sl, 1.5+i*3.8, 4.18, 3.3, 0.6);
    sl.addText(tool, { x:1.5+i*3.8+0.15, y:4.18, w:3.0, h:0.6, fontSize:12, bold:true, color:C.textDark, valign:"middle", margin:0 });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SECCIÓN 4: BANDEJAS + PASO A PASO
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(13, "📧", "Bandejas + Paso a Paso", "El corazón del trabajo diario.");

// ────────────────────────────────── SECUENCIAS ───────────────────────────────

preQuiz(14, "Secuencias de correo", [
  "¿Qué son las secuencias de correo?",
  "¿Quién las envía?",
  "¿A qué hora se mandan?",
  "¿Qué hago si una secuencia no tiene respuesta?",
  "¿Cuándo respondo o hago push?",
]);

// ─── 15: SECUENCIAS ──────────────────────────────────────────────────────────
{
  const sl = lightSlide(15);
  sl.addText("Secuencias de correo electrónico", { x:0.5, y:0.55, w:9, h:0.6, fontSize:28, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const rules = [
    { icon:"⏰", text:"Se envían automáticamente todos los días en diferentes horarios." },
    { icon:"👀", text:"Solo se responden cuando hay respuesta del prospecto." },
    { icon:"🔄", text:"Solo se hace push cuando hay respuesta del prospecto." },
    { icon:"📁", text:"Sin respuesta → Archivar. No etiquetar, no dejar en Recibidos.", highlight:true },
  ];
  rules.forEach((r, i) => {
    const y = 1.38+i*0.95;
    sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:9.1, h:0.78, fill:{color: r.highlight?"FFF8E1":C.card}, line:{color: r.highlight?C.yellow:C.cardBdr, width: r.highlight?1.5:0.5} });
    sl.addText(r.icon, { x:0.7, y, w:0.55, h:0.78, fontSize:24, valign:"middle", align:"center", margin:0 });
    sl.addText(r.text, { x:1.4, y, w:8.0, h:0.78, fontSize:14, color:C.textDark, valign:"middle", margin:0 });
    if (r.highlight) {
      // Archive button visual indicator
      sl.addShape(pres.shapes.RECTANGLE, { x:8.5, y:y+0.18, w:0.65, h:0.42, fill:{color:C.yellow}, line:{color:"E6C000"} });
      sl.addText("⬆", { x:8.5, y:y+0.18, w:0.65, h:0.42, fontSize:18, align:"center", valign:"middle", margin:0 });
    }
  });
  sl.addText("El botón de archivar ⬆ (resaltado) mueve el correo fuera de Recibidos sin etiqueta.", {
    x:0.5, y:5.15, w:9.1, h:0.3, fontSize:11, color:C.gray, margin:0
  });
}

// ─── 16: CAPTURAS SECUENCIAS ─────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Ejemplos de secuencias", { x:0.5, y:0.5, w:9.1, h:0.6, fontSize:26, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  // Show 3 example images side by side
  const imgs = [IMG.capturas0, IMG.capturas1, IMG.capturas2];
  const iw=2.9, ih=3.8, gap=0.18, startX=0.5, y=1.25;
  imgs.forEach((img, i) => {
    sl.addImage({ path:img, x:startX+i*(iw+gap), y, w:iw, h:ih });
  });
  slideNum(sl, 16, true);
}

// ────────────────────────────────── ETIQUETAS ────────────────────────────────

preQuiz(17, "Las etiquetas de correo", [
  "¿Cuántas etiquetas existen?",
  "¿Cuándo usas 'PROCESO'?",
  "¿Cuándo usas 'NO INTERESADO'?",
  "¿Qué diferencia hay entre 'REAGENDAR' y 'PROCESO'?",
  "¿Qué significa 'NO APLICA'?",
]);

// ─── 18: ETIQUETAS ───────────────────────────────────────────────────────────
{
  const sl = lightSlide(18);
  sl.addText("Las 5 etiquetas", { x:0.5, y:0.55, w:9, h:0.6, fontSize:30, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const labels = [
    { name:"NO INTERESADO", color:ETIQUETAS.NO_INTERESADO, desc:"Prospecto perdido, no quiere saber más, imposible agendar." },
    { name:"PROCESO",       color:ETIQUETAS.PROCESO,       desc:"Muestra interés. Hace dudas, \"no soy el encargado\", \"escríbeme la próxima semana\"." },
    { name:"AGENDADO",      color:ETIQUETAS.AGENDADO,      desc:"Se completaron los pasos del proceso de agendamiento." },
    { name:"REAGENDAR",     color:ETIQUETAS.REAGENDAR,     desc:"No Show. La reunión no se realizó y hay que buscar otro horario." },
    { name:"NO APLICA",     color:ETIQUETAS.NO_APLICA,     desc:"El prospecto/empresa no cumple los parámetros del cliente. No se puede agendar." },
  ];
  labels.forEach((lb, i) => {
    const y = 1.32+i*0.77;
    sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:9.1, h:0.65, fill:{color:C.card}, line:{color:C.cardBdr, width:0.5} });
    sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:2.3, h:0.65, fill:{color:lb.color}, line:{color:lb.color} });
    sl.addText(lb.name, { x:0.55, y, w:2.2, h:0.65, fontSize:13, bold:true, color:C.white, valign:"middle", align:"center", margin:0 });
    sl.addText(lb.desc, { x:2.95, y, w:6.5, h:0.65, fontSize:12, color:C.darkGray, valign:"middle", margin:0, wrap:true });
  });
}

// ────────────────────────────────── PASO A PASO ──────────────────────────────

preQuiz(19, "¿Qué hago cuando un prospecto responde?", [
  "¿Cuál es el primer paso al recibir una respuesta?",
  "¿Qué hago si el prospecto pone dudas del servicio?",
  "¿Qué hago si dice 'no soy el encargado, habla con X'?",
  "¿Qué pasa si dice que está de vacaciones?",
  "¿Cuándo y cómo hago push?",
]);

// ─── 20: DIAGRAMA PASO A PASO ────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Paso a paso — diagrama completo", { x:0.5, y:0.5, w:9.1, h:0.58, fontSize:26, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addImage({ path:IMG.diagCorreos, x:0.5, y:1.18, w:9.1, h:4.22 });
  slideNum(sl, 20, true);
}

// ─── 21: BARRIDO DE BANDEJA ──────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Barrido de la bandeja", { x:0.5, y:0.5, w:9.1, h:0.58, fontSize:26, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("Del primer correo al último — push uno por uno, sin saltear.", {
    x:0.5, y:1.1, w:9.1, h:0.38, fontSize:14, color:C.gray, margin:0
  });
  sl.addImage({ path:IMG.barrido, x:1.5, y:1.55, w:7.1, h:3.85 });
  slideNum(sl, 21, true);
}

// ────────────────────────────────── PUSH ─────────────────────────────────────

preQuiz(22, "Estructura del push", [
  "¿Cuáles son las partes de un push?",
  "¿Qué va en el cuerpo del mensaje?",
  "¿Cómo termina el push?",
  "¿Cada cuánto se hace?",
]);

// ─── 23: ESTRUCTURA PUSH ─────────────────────────────────────────────────────
{
  const sl = lightSlide(23);
  sl.addText("Estructura del push", { x:0.5, y:0.55, w:9, h:0.6, fontSize:28, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const parts = [
    { n:"1", title:"Saludo",          desc:"Breve y natural. Nombra al prospecto." },
    { n:"2", title:"Frase amable",    desc:"Toma algo de lo que dijo el prospecto en su último correo." },
    { n:"3", title:"Cuerpo",          desc:"Lo que el prospecto quiere escuchar: beneficios en su empresa, solución a su problema." },
    { n:"4", title:"Call to Action",  desc:'Frase directa que indica el siguiente paso. Ej: "¿Prefieres el martes o miércoles entre 10 y 13h?"' },
    { n:"5", title:"Despedida",       desc:"Corta y amable." },
  ];
  parts.forEach((p, i) => {
    const y = 1.3+i*0.77;
    accentCard(sl, 0.5, y, 9.1, 0.65, C.green);
    sl.addShape(pres.shapes.OVAL, { x:0.72, y:y+0.1, w:0.45, h:0.45, fill:{color:C.green}, line:{color:C.green} });
    sl.addText(p.n, { x:0.72, y:y+0.1, w:0.45, h:0.45, fontSize:14, bold:true, color:C.white, align:"center", valign:"middle", margin:0 });
    sl.addText(p.title, { x:1.3, y, w:2.2, h:0.65, fontSize:13, bold:true, color:C.textDark, valign:"middle", margin:0 });
    sl.addText(p.desc, { x:3.6, y, w:5.8, h:0.65, fontSize:12, color:C.darkGray, valign:"middle", margin:0, wrap:true });
  });
}

// ────────────────────────────────── LLAMADO EN TIBIO ─────────────────────────

preQuiz(24, "Llamado en tibio", [
  "¿Qué es un llamado en tibio?",
  "¿En qué se diferencia del llamado en frío?",
  "¿Cuándo se hace?",
]);

// ─── 25: LLAMADO EN TIBIO ────────────────────────────────────────────────────
{
  const sl = lightSlide(25);
  sl.addText("Llamado en tibio 📞", { x:0.5, y:0.55, w:9, h:0.6, fontSize:28, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  accentCard(sl, 0.5, 1.28, 9.1, 0.82, C.green);
  sl.addText("¿Qué es?", { x:0.72, y:1.28, w:2, h:0.82, fontSize:12, bold:true, color:C.textDark, valign:"middle", margin:0 });
  sl.addText("Contactar por teléfono a un prospecto con quien ya tuvimos conversación por correo. Ya nos conocen — no es en frío.", {
    x:2.85, y:1.28, w:6.6, h:0.82, fontSize:13, color:C.darkGray, valign:"middle", margin:0, wrap:true
  });
  sl.addText("¿Cuándo hacer llamado en tibio?", { x:0.5, y:2.28, w:9.1, h:0.42, fontSize:14, bold:true, color:C.textDark, margin:0 });
  const cuando = [
    "Mostró interés y dejó de responder",
    "Quieres agendar más rápido",
    "Ha costado coordinar una fecha",
    "Quiere más información del servicio",
  ];
  cuando.forEach((item, i) => {
    const x=0.5+(i%2)*4.65, y=2.82+Math.floor(i/2)*0.9;
    card(sl, x, y, 4.45, 0.72);
    sl.addShape(pres.shapes.OVAL, { x:x+0.15, y:y+0.14, w:0.44, h:0.44, fill:{color:C.green}, line:{color:C.green} });
    sl.addText("✓", { x:x+0.15, y:y+0.14, w:0.44, h:0.44, fontSize:13, bold:true, color:C.white, align:"center", valign:"middle", margin:0 });
    sl.addText(item, { x:x+0.75, y, w:3.55, h:0.72, fontSize:12, color:C.textDark, valign:"middle", margin:0 });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:4.72, w:9.1, h:0.38, fill:{color:C.yellow}, line:{color:C.yellow} });
  sl.addText("💡  SIEMPRE que aplique una de estas situaciones — no esperes.", {
    x:0.65, y:4.72, w:8.8, h:0.38, fontSize:12, bold:true, color:C.textDark, valign:"middle", margin:0
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SECCIÓN 5: AGENDAMIENTO
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(26, "📅", "Agendamiento", "9 pasos. Todos importan.");

// ─── 27: PRE-QUIZ AGENDAMIENTO ───────────────────────────────────────────────
preQuiz(27, "Proceso de agendamiento", [
  "¿Cuáles son los pasos del agendamiento? ¿Cuántos son?",
  "¿A través de qué plataforma confirmas SIEMPRE la reunión?",
  "¿Quién aprueba el agendamiento y cómo lo indica?",
  "¿Cuántos recordatorios se programan y cuándo?",
]);

// ─── 28: DIAGRAMA AGENDAMIENTO ───────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Proceso de agendamiento — diagrama completo", { x:0.5, y:0.5, w:9.1, h:0.55, fontSize:24, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addImage({ path:IMG.diagAgend, x:0.4, y:1.1, w:9.2, h:4.3 });
  slideNum(sl, 28, true);
}

// ─── 29: PASO 3-4 — CANAL SLACK + APROBACIÓN ─────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Paso 3 — Agendar en App TGP + Canal Slack", { x:0.5, y:0.5, w:9.1, h:0.55, fontSize:22, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("Paso 4 — Aprobación del Manager", { x:0.5, y:1.05, w:9.1, h:0.38, fontSize:16, color:C.green, bold:true, margin:0 });
  sl.addImage({ path:IMG.aprobSlack, x:0.4, y:1.5, w:5.8, h:3.9 });
  // Side note
  sl.addShape(pres.shapes.RECTANGLE, { x:6.45, y:1.5, w:3.2, h:1.7, fill:{color:"2A2A2A"}, line:{color:"3A3A3A"} });
  sl.addText("👍", { x:6.5, y:1.55, w:0.6, h:0.6, fontSize:28, align:"center", margin:0 });
  sl.addText("El Manager aprueba reaccionando con un emoji en el mensaje del canal.\n\nSin esa reacción, la reunión no está confirmada.", {
    x:6.5, y:2.05, w:3.05, h:1.1, fontSize:12, color:"CCCCCC", margin:0, wrap:true
  });
  sl.addShape(pres.shapes.RECTANGLE, { x:6.45, y:3.3, w:3.2, h:1.1, fill:{color:"2A2A2A"}, line:{color:"3A3A3A"} });
  sl.addText("Formato del mensaje al canal:", { x:6.5, y:3.35, w:3.1, h:0.3, fontSize:11, bold:true, color:C.yellow, margin:0 });
  sl.addText("Empresa · Mail · Contacto · Cargo\nTeléfono · Día y Hora · SDR\n¿Es reagendamiento?", {
    x:6.5, y:3.65, w:3.1, h:0.72, fontSize:10, color:"CCCCCC", margin:0
  });
  slideNum(sl, 29, true);
}

// ─── 30: PASO 5 — DETALLE AL CLIENTE ─────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Paso 5 — Detalle y contexto al cliente", { x:0.5, y:0.5, w:9.1, h:0.55, fontSize:24, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("WhatsApp del grupo de clientes o canal de Slack del pod", { x:0.5, y:1.05, w:9.1, h:0.35, fontSize:14, color:C.gray, margin:0 });
  sl.addImage({ path:IMG.detalleWA,    x:0.4, y:1.48, w:4.45, h:3.9 });
  sl.addImage({ path:IMG.detalleSlack, x:5.1, y:1.48, w:4.5,  h:3.9 });
  slideNum(sl, 30, true);
}

// ─── 31: PASO 6 — CONFIRMACIÓN AL PROSPECTO ──────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Paso 6 — Confirmación al prospecto", { x:0.5, y:0.5, w:9.1, h:0.55, fontSize:24, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("Correo de confirmación con la fecha, hora y link de la reunión.", { x:0.5, y:1.05, w:9.1, h:0.35, fontSize:14, color:C.gray, margin:0 });
  sl.addImage({ path:IMG.confirmacion, x:1.8, y:1.48, w:6.5, h:3.9 });
  slideNum(sl, 31, true);
}

// ─── 32: PASO 7 — RECORDATORIOS ──────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Paso 7 — Programar 2 recordatorios", { x:0.5, y:0.5, w:9.1, h:0.55, fontSize:24, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("Recordatorio 1: día anterior  ·  Recordatorio 2: mismo día de la reunión", { x:0.5, y:1.05, w:9.1, h:0.35, fontSize:14, color:C.green, bold:true, margin:0 });
  sl.addImage({ path:IMG.recordatorio1, x:0.4, y:1.48, w:4.45, h:3.9 });
  sl.addImage({ path:IMG.recordatorio3, x:5.1, y:1.48, w:4.5,  h:3.9 });
  slideNum(sl, 32, true);
}

// ─── 33: PASO 8 — ACTUALIZAR ETIQUETA ────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Paso 8 — Actualizar etiqueta → AGENDADO", { x:0.5, y:0.5, w:9.1, h:0.55, fontSize:24, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("Seleccionar correo → ícono etiqueta → marcar AGENDADO → Aplicar", { x:0.5, y:1.05, w:9.1, h:0.35, fontSize:14, color:C.gray, margin:0 });
  // 2x2 grid of screenshots
  const imgs = [IMG.etiqueta3, IMG.etiqueta1, IMG.etiqueta2, IMG.etiqueta4];
  const iw=4.4, ih=2.0, gap=0.15;
  imgs.forEach((img, i) => {
    const x=0.4+(i%2)*(iw+gap), y=1.48+Math.floor(i/2)*(ih+0.1);
    sl.addImage({ path:img, x, y, w:iw, h:ih });
  });
  slideNum(sl, 33, true);
}

// ─── 34: PASO 9 — NO SHOW + IA ───────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  logoDark(sl);
  sl.addText("Paso 9 — No Show + IA", { x:0.5, y:0.5, w:9.1, h:0.55, fontSize:26, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });

  // Canal explanation
  sl.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.12, w:4.2, h:4.22, fill:{color:"1A1A1A"}, line:{color:"333333"} });
  sl.addText("🔔  Canal por pod", { x:0.55, y:1.18, w:3.9, h:0.4, fontSize:13, bold:true, color:C.green, margin:0 });
  sl.addText("pod-nombre-noshow", { x:0.55, y:1.6, w:3.9, h:0.38, fontSize:14, bold:true, color:C.white, fontFace:"Consolas", margin:0 });

  const steps = [
    { t:"A la hora exacta", d:"La IA notifica al canal con: empresa, SDR, hora, link de Meet, teléfono." },
    { t:"+4 min",           d:"Primera verificación. Si hay No Show: menciona @SDR y deja pantallazos de la sala de espera de Meet." },
    { t:"+8 min",           d:"Segunda verificación con pantallazos." },
    { t:"+15 min",          d:"Verificación final. Todo queda registrado en App TGP (columna No Show)." },
  ];
  steps.forEach((s, i) => {
    const y = 2.1+i*0.82;
    sl.addShape(pres.shapes.OVAL, { x:0.55, y:y+0.1, w:0.38, h:0.38, fill:{color:C.green}, line:{color:C.green} });
    sl.addText(`${i+1}`, { x:0.55, y:y+0.1, w:0.38, h:0.38, fontSize:11, bold:true, color:C.white, align:"center", valign:"middle", margin:0 });
    sl.addText(s.t, { x:1.05, y, w:1.2, h:0.35, fontSize:10, bold:true, color:C.yellow, margin:0 });
    sl.addText(s.d, { x:1.05, y:y+0.32, w:3.4, h:0.45, fontSize:10, color:"BBBBBB", margin:0, wrap:true });
  });

  // WhatsApp bonus
  sl.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.0, w:4.2, h:0.38, fill:{color:C.greenDk}, line:{color:C.greenDk} });
  sl.addText("📱  WhatsApp al prospecto 1h antes si tiene número.", { x:0.55, y:5.0, w:4.0, h:0.38, fontSize:11, bold:true, color:C.white, valign:"middle", margin:0 });

  // Images from original
  sl.addImage({ path:IMG.noshowIA1, x:4.78, y:1.12, w:2.3, h:4.28 });
  sl.addImage({ path:IMG.noshowIA2, x:7.25, y:1.12, w:2.35, h:4.28 });

  slideNum(sl, 34, true);
}

// ════════════════════════════════════════════════════════════════════════════
// SECCIÓN 6: IA EN TGP
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(35, "🤖", "La IA en TGP", null);

// ─── 36: PRE-QUIZ IA ─────────────────────────────────────────────────────────
preQuiz(36, "Etiquetas de IA en los correos", [
  "¿Qué etiquetas de IA pueden aparecer en un correo?",
  "¿Qué indica 'AI_Thread'? ¿Qué hago cuando lo veo?",
  "¿Qué hace la IA cuando etiqueta un correo como 'TGP Processed'?",
  "¿Qué significa 'TGP Closed'?",
]);

// ─── 37: ETIQUETAS IA ────────────────────────────────────────────────────────
{
  const sl = lightSlide(37);
  sl.addText("Las 3 etiquetas de IA", { x:0.5, y:0.55, w:9, h:0.6, fontSize:28, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const ia_tags = [
    { tag:"TGP Processed", color:"1565C0", desc:"La IA procesó este correo. A veces genera borradores — revisarlos antes de enviar.", rule:"Revisar borrador si existe" },
    { tag:"AI_Thread",     color:"880E4F", desc:"La IA está trabajando activamente en esta conversación ahora mismo.", rule:"NO INTERFERIR" },
    { tag:"TGP Closed",    color:C.greenDk, desc:"La IA cerró la conversación — se agendó o fue rechazada definitivamente.", rule:"Verificar en App TGP" },
  ];
  ia_tags.forEach((tag, i) => {
    const y = 1.35+i*1.28;
    sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:9.1, h:1.12, fill:{color:C.card}, line:{color:C.cardBdr, width:0.5} });
    sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:0.07, h:1.12, fill:{color:tag.color}, line:{color:tag.color} });
    sl.addText(tag.tag, { x:0.72, y:y+0.1, w:3.0, h:0.38, fontSize:16, bold:true, color:tag.color, margin:0, fontFace:"Consolas" });
    sl.addText(tag.desc, { x:0.72, y:y+0.52, w:5.8, h:0.55, fontSize:12, color:C.darkGray, margin:0, wrap:true });
    sl.addShape(pres.shapes.RECTANGLE, { x:6.8, y:y+0.26, w:2.65, h:0.6, fill:{color:tag.color}, line:{color:tag.color} });
    sl.addText(tag.rule, { x:6.8, y:y+0.26, w:2.65, h:0.6, fontSize:11, bold:true, color:C.white, align:"center", valign:"middle", margin:0, wrap:true });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// QUIZ FINAL
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(38, "🏆", "Quiz Final", "¿Lo tienen? Vamos a comprobarlo.");

quizSlide(39, [
  "¿Cuál es la diferencia entre un cliente y un prospecto?",
  "¿Cuál es la regla principal de Slack?",
  "¿Qué hago con las secuencias que no tienen respuesta?",
  "¿Cuál es el rol del SDR Senior?",
  "¿Cuál es el rol del CSR?",
]);

quizSlide(40, [
  "¿Qué hacer cuando llega un correo automático de 'fuera de la oficina'?",
  "¿A través de qué plataforma confirmas SIEMPRE la reunión?",
  "¿Qué pasa si no haces los 2 recordatorios?",
  "¿Qué etiqueta de IA indica que NO debes tocar el correo?",
  "¿Cuál es la importancia de agendar reuniones válidas?",
]);

// ════════════════════════════════════════════════════════════════════════════
// HERRAMIENTAS + ACCESOS
// ════════════════════════════════════════════════════════════════════════════
sectionDivider(41, "⚡", "Tus Herramientas", "Todo en un solo lugar.");

// ─── 42: DASHBOARD + APP TGP ─────────────────────────────────────────────────
{
  const sl = lightSlide(42);
  sl.addText("Tus 2 plataformas de trabajo", { x:0.5, y:0.55, w:9, h:0.6, fontSize:28, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  const tools = [
    { icon:"📊", title:"Dashboard Llamados", sub:"Data histórica de agendamiento.\nSeguimiento de tu actividad y resultados.", color:C.dark },
    { icon:"📱", title:"App TGP",            sub:"Plataforma de agendamiento.\nAcá confirmas y registras todas las reuniones.", color:"1A3A6A" },
  ];
  tools.forEach((t, i) => {
    const x=0.4+i*4.85;
    sl.addShape(pres.shapes.RECTANGLE, { x, y:1.32, w:4.65, h:3.8, fill:{color:t.color}, line:{color:t.color}, shadow:mkShadow() });
    sl.addText(t.icon, { x, y:1.45, w:4.65, h:1.0, fontSize:42, align:"center", margin:0 });
    sl.addText(t.title, { x:x+0.15, y:2.5, w:4.35, h:0.75, fontSize:20, bold:true, color:C.white, fontFace:"Arial Black", align:"center", margin:0 });
    sl.addText(t.sub, { x:x+0.15, y:3.3, w:4.35, h:0.9, fontSize:12, color:"AAAAAA", align:"center", margin:0, wrap:true });
    sl.addShape(pres.shapes.RECTANGLE, { x:x+0.15, y:4.35, w:4.35, h:0.55, fill:{color:C.green}, line:{color:C.green} });
    sl.addText("Acceso desde Central TGP → Recursos", { x:x+0.15, y:4.35, w:4.35, h:0.55, fontSize:11, bold:true, color:C.white, align:"center", valign:"middle", margin:0 });
  });
}

// ─── 43: CENTRAL TGP ─────────────────────────────────────────────────────────
{
  const sl = lightSlide(43);
  sl.addText("Tu Central TGP — todo está ahí", { x:0.5, y:0.55, w:9, h:0.6, fontSize:28, bold:true, color:C.textDark, fontFace:"Arial Black", margin:0 });
  sl.addText("thegrowthpro.github.io/central-tgp  ·  Entra con tu correo @thegrowth.pro", {
    x:0.5, y:1.2, w:9.1, h:0.38, fontSize:12, color:C.gray, align:"center", margin:0
  });
  const accesos = [
    { icon:"📋", section:"Onboarding",   desc:"Tu proceso completo, día 1 a día 3" },
    { icon:"📚", section:"Recursos",      desc:"Bandejas · Llamado · Agendamiento · Apps y Kits" },
    { icon:"👥", section:"RRHH",          desc:"Ausencias, vacaciones, beneficios" },
    { icon:"📊", section:"Dashboard",     desc:"Métricas y seguimiento de actividad" },
    { icon:"📱", section:"App TGP",       desc:"Agendamiento y registro de reuniones" },
    { icon:"💬", section:"Slack TGP",     desc:"Canal del equipo, novedades y contacto" },
  ];
  const cw=2.85, ch=1.1, gx=0.18, gy=0.2, startX=0.5, startY=1.72;
  accesos.forEach((ac, i) => {
    const x=startX+(i%3)*(cw+gx), y=startY+Math.floor(i/3)*(ch+gy);
    card(sl, x, y, cw, ch);
    sl.addText(ac.icon, { x:x+0.12, y, w:0.6, h:ch, fontSize:26, valign:"middle", align:"center", margin:0 });
    sl.addText(ac.section, { x:x+0.78, y:y+0.14, w:cw-0.9, h:0.36, fontSize:13, bold:true, color:C.textDark, margin:0 });
    sl.addText(ac.desc, { x:x+0.78, y:y+0.52, w:cw-0.9, h:0.5, fontSize:10.5, color:C.darkGray, margin:0, wrap:true });
  });
}

// ─── 44: PRÓXIMOS PASOS ──────────────────────────────────────────────────────
{
  const sl = darkSlide(44);
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:5.22, w:10, h:0.4, fill:{color:C.green}, line:{color:C.green} });
  sl.addText("¿Qué sigue?", { x:0.5, y:0.55, w:9, h:0.6, fontSize:30, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  const next = [
    { icon:"📅", text:"Reunión con tu SDR Senior esta semana" },
    { icon:"📧", text:"Primeras bandejas de correo asignadas" },
    { icon:"📊", text:"Primera semana de actividad real" },
    { icon:"🎯", text:"Metas semanales — las coordinás con tu Manager" },
  ];
  next.forEach((item, i) => {
    const y = 1.32+i*0.88;
    sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:9.1, h:0.72, fill:{color:"2A2A2A"}, line:{color:"3A3A3A", width:0.5} });
    sl.addText(item.icon, { x:0.7, y, w:0.55, h:0.72, fontSize:24, valign:"middle", align:"center", margin:0 });
    sl.addText(item.text, { x:1.38, y, w:7.8, h:0.72, fontSize:15, bold:true, color:C.white, valign:"middle", margin:0 });
  });
}

// ─── 45: CIERRE ──────────────────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  sl.background = { color: C.dark };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.08, h:5.625, fill:{color:C.green}, line:{color:C.green} });
  sl.addText("TGP", { x:0.5, y:0.5, w:5, h:1.1, fontSize:64, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("Bienvenido al equipo.", { x:0.5, y:1.65, w:9, h:0.75, fontSize:38, bold:true, color:C.white, fontFace:"Arial Black", margin:0 });
  sl.addText("Ahora a ejecutar.", { x:0.5, y:2.4, w:9, h:0.75, fontSize:38, bold:true, color:C.green, fontFace:"Arial Black", margin:0 });
  sl.addText("¿Preguntas finales?", { x:0.5, y:3.35, w:9, h:0.5, fontSize:18, color:C.gray, margin:0 });
  slideNum(sl, 45, true);
}

// ─── WRITE ────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/Users/belenforti/Documents/Claude/Projects/Onboarding/Repaso SDR Junior — TGP.pptx" })
  .then(() => console.log("✅ Presentación creada — 45 slides."))
  .catch(e => console.error("❌", e));
