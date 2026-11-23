/**
 * gigES · Taller de validación de la evaluación — respaldo en Google Drive
 * CliO Consulting · Evaluación final ES-T1341
 *
 * Qué hace:
 *  - Recibe votos, notas de las mesas y síntesis de moderación desde la web del taller
 *    (https://farjonaclio.github.io/gigES-tablero/taller/) y los guarda en una hoja de Drive.
 *  - Entrega los resultados agregados para el plenario (promedios por rol y notas por mesa).
 *  - Lleva un registro de cada envío (hoja «registro»), como respaldo histórico.
 *  - Si alguien abre la URL de esta aplicación, le muestra la misma web del taller.
 *
 * Instalación (una vez): pegar este archivo, ejecutar «configurar» e implementar como
 * aplicación web (Ejecutar como: yo · Acceso: cualquier persona).
 */

var PAGINA = "https://farjonaclio.github.io/gigES-tablero/taller/";
var CARPETA_ID = "1wIar0HCEOfjoBQQse8mdUrc803PQ2_mn";   // carpeta «Taller de validación · web y respuestas»
var NOMBRE_HOJA = "gigES · Taller de validación · respuestas";

var DIMS = ["eficacia", "implementacion", "impacto", "sostenibilidad", "genero"];
var NH = 5, NR = 5;
var HALL_T = ["Productos superados", "Acceso real menor al reportado", "Usuarios mejoran más", "Cada solución mejora donde actúa", "Sin acuerdos de continuidad"];
var REC_T = ["Pagar por uso", "Financiar la activación", "Empezar por el ingreso", "Continuidad desde el diseño", "Género por segmento"];
var NOMBRE_DIM = {
  eficacia: "eficacia", implementacion: "implementación", impacto: "impacto",
  sostenibilidad: "sostenibilidad", genero: "género"
};
function rango_(p, n) { var o = []; for (var i = 0; i < n; i++) o.push(p + i); return o; }

var COLS = {
  votos: ["pid", "actualizado", "mesa", "actor", "rol", "solucion"].concat(DIMS.map(function (d) { return "v_" + d; })),
  notas: ["pid", "actualizado", "mesa", "actor", "rol", "solucion"]
    .concat(DIMS.map(function (d) { return "d_" + d; }))
    .concat(DIMS.map(function (d) { return "c_" + d; }))
    .concat(DIMS.map(function (d) { return "cp_" + d; }))
    .concat(rango_("h_", NH)).concat(rango_("hm_", NH)).concat(["hallazgo"])
    .concat(rango_("r_", NR)).concat(rango_("rm_", NR))
    .concat(["rec1", "rec1p", "rec2", "rec2p"]),
  sintesis: ["pid", "actualizado", "mesa", "actor", "rol", "s_comp", "s_npers", "s_otros", "s_porque"]
    .concat(DIMS.map(function (d) { return "sn_" + d; }))
    .concat(["s_hallazgos", "s_recs"]),
  registro: ["recibido", "pid", "tipo", "mesa", "actor", "json"]
};

/* ---------------- configuración ---------------- */
function configurar() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("SHEET_ID");
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create(NOMBRE_HOJA);
    props.setProperty("SHEET_ID", ss.getId());
    if (CARPETA_ID) DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(CARPETA_ID));
  }
  Object.keys(COLS).forEach(function (n) { hoja_(ss, n); });
  var sobra = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (sobra && ss.getSheets().length > 1) ss.deleteSheet(sobra);
  Logger.log("Hoja lista: " + ss.getUrl());
  return ss.getUrl();
}

function libro_() {
  var id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!id) { configurar(); id = PropertiesService.getScriptProperties().getProperty("SHEET_ID"); }
  return SpreadsheetApp.openById(id);
}

function hoja_(ss, nombre) {
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, COLS[nombre].length).setValues([COLS[nombre]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------------- web ---------------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.accion === "resultados") {
    return ContentService.createTextOutput(JSON.stringify(resultados()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var html = UrlFetchApp.fetch(PAGINA, { muteHttpExceptions: true }).getContentText();
  return HtmlService.createHtmlOutput(html)
    .setTitle("Taller de validación · gigES")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  guardar(d);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- guardar (una fila por persona y tipo; se reemplaza al reenviar) ---------------- */
function guardar(d) {
  var tipos = { voto: "votos", notas: "notas", sintesis: "sintesis" };
  var nombre = tipos[d.tipo];
  if (!nombre || !d.pid) throw new Error("Envío incompleto");

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = libro_();
    var ahora = new Date();
    hoja_(ss, "registro").appendRow([ahora, d.pid, d.tipo, d.mesa || "", d.actor || "", JSON.stringify(d)]);

    var sh = hoja_(ss, nombre);
    var cols = COLS[nombre];
    var fila = cols.map(function (c) {
      if (c === "actualizado") return ahora;
      var v = d[c];
      return v == null ? "" : String(v).slice(0, 5000);
    });
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
    var donde = -1;
    for (var i = 0; i < ids.length; i++) { if (ids[i][0] === d.pid) { donde = i + 2; break; } }
    if (donde > 0) sh.getRange(donde, 1, 1, cols.length).setValues([fila]);
    else sh.appendRow(fila);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ---------------- resultados agregados ---------------- */
function resultados() {
  var ss = libro_();
  var out = { n: 0, promedios: {}, notas: [], mesas: { A: 0, B: 0 }, cambios: {}, hall: [], recs: [] };
  DIMS.forEach(function (k) { out.cambios[k] = { Subir: 0, Mantener: 0, Bajar: 0 }; });
  for (var i = 0; i < NH; i++) out.hall.push({ "De acuerdo": 0, "Con matices": 0, "En desacuerdo": 0 });
  for (var j = 0; j < NR; j++) out.recs.push({ Alta: 0, Media: 0, Baja: 0 });

  // votos individuales: promedio por rol y del taller (sin el equipo evaluador)
  var acc = {};
  function suma(g, r) {
    if (!acc[g]) acc[g] = { n: 0, s: DIMS.map(function () { return 0; }), c: DIMS.map(function () { return 0; }) };
    acc[g].n++;
    DIMS.forEach(function (k, i) { var v = parseFloat(r["v_" + k]); if (!isNaN(v)) { acc[g].s[i] += v; acc[g].c[i]++; } });
  }
  filas_(ss, "votos").forEach(function (r) {
    if (r.actor === "Equipo CliO") return;
    out.n++; suma("Taller (todos)", r); suma((r.actor === "BID Lab" || r.actor === "Unión Europea") ? "Financiadores" : r.actor, r);
  });
  Object.keys(acc).forEach(function (g) {
    out.promedios[g] = { n: acc[g].n, v: acc[g].s.map(function (s, i) { return acc[g].c[i] ? Math.round(s / acc[g].c[i] * 100) / 100 : null; }) };
  });

  // notas de las mesas
  function nota(r, tema, texto) {
    out.notas.push({ mesa: r.mesa, actor: r.actor, tema: tema, texto: texto, t: +new Date(r.actualizado) });
    out.mesas[r.mesa]++;
  }
  filas_(ss, "notas").forEach(function (r) {
    if (r.mesa !== "A" && r.mesa !== "B") return;
    DIMS.forEach(function (k) {
      if (r["d_" + k]) nota(r, "Por qué vemos distinto · " + NOMBRE_DIM[k], r["d_" + k]);
      var c = r["c_" + k];
      if (out.cambios[k][c] !== undefined) out.cambios[k][c]++;
      if (c && c !== "Mantener") nota(r, c + " la nota de " + NOMBRE_DIM[k], r["cp_" + k] || "(sin comentario)");
    });
    for (var i = 0; i < NH; i++) {
      var h = r["h_" + i];
      if (out.hall[i][h] !== undefined) out.hall[i][h]++;
      if (r["hm_" + i]) nota(r, "Hallazgo «" + HALL_T[i] + "» · " + (h || "matiz"), r["hm_" + i]);
    }
    if (r.hallazgo) nota(r, "Hallazgo que falta", r.hallazgo);
    for (var j = 0; j < NR; j++) {
      var p = r["r_" + j];
      if (out.recs[j][p] !== undefined) out.recs[j][p]++;
      if (r["rm_" + j]) nota(r, "Recomendación «" + REC_T[j] + "» · " + (p || "comentario"), r["rm_" + j]);
    }
    ["rec1", "rec2"].forEach(function (k) { if (r[k]) nota(r, "Nueva recomendación" + (r[k + "p"] ? " para " + r[k + "p"] : ""), r[k]); });
  });

  // síntesis de moderación (arriba en cada mesa)
  var TS = { s_porque: "Por qué vemos distinto", s_hallazgos: "Hallazgos", s_recs: "Recomendaciones" };
  filas_(ss, "sintesis").forEach(function (r) {
    if (r.mesa !== "A" && r.mesa !== "B") return;
    var t = +new Date(r.actualizado) + 1e12;
    if (r.s_comp || r.s_npers) out.notas.push({ mesa: r.mesa, actor: "Moderación", tema: "Composición de la mesa", texto: [r.s_comp, r.s_otros].filter(String).join(", ") + (r.s_npers ? " · " + r.s_npers + " personas" : ""), sintesis: true, t: t + 1 });
    Object.keys(TS).forEach(function (k) { if (r[k]) out.notas.push({ mesa: r.mesa, actor: "Moderación", tema: TS[k], texto: r[k], sintesis: true, t: t }); });
    var props = DIMS.filter(function (k) { return r["sn_" + k]; }).map(function (k) { return NOMBRE_DIM[k] + " " + r["sn_" + k]; });
    if (props.length) out.notas.push({ mesa: r.mesa, actor: "Moderación", tema: "Calificación que propone la mesa", texto: props.join(" · "), sintesis: true, t: t });
  });

  out.notas.sort(function (a, b) { return b.t - a.t; });
  out.notas.forEach(function (x) { delete x.t; });
  return out;
}

function filas_(ss, nombre) {
  var sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var cab = v.shift();
  return v.map(function (fila) {
    var o = {};
    cab.forEach(function (c, i) { o[c] = fila[i]; });
    return o;
  });
}
