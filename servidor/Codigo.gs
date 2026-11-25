/**
 * ReCéTaMe · servidor de recetas (Google Apps Script)
 * ----------------------------------------------------
 * Este archivo es una copia de seguridad. El que funciona es el pegado en
 * la hoja «ReCéTaMe datos» → Extensiones → Apps Script.
 *
 * Para actualizarlo:
 *   1. Borra todo el código del editor y pega este archivo entero.
 *   2. Vuelve a escribir tu clave en CLAVE_ADMIN.
 *   3. Implementar → Gestionar implementaciones → ✏️ (editar) →
 *      Versión: «Nueva versión» → Implementar.
 *      (Así la URL /exec no cambia.)
 */

// ⚠️ Escribe aquí tu clave de administrador.
const CLAVE_ADMIN = 'cambia-esta-clave';

const HOJA = 'Recetas';
const CARPETA_FOTOS = 'ReCéTaMe fotos';
const CAMPOS = [
  'id', 'titulo', 'icono', 'subtitulo', 'categorias', 'raciones', 'tiempo',
  'macros', 'ingredientes', 'pasos', 'consejos', 'foto', 'origen',
  'creada', 'actualizada'
];
// Se guardan como JSON en la celda (las edita la app, no hace falta tocarlas a mano).
const CAMPOS_JSON = ['macros', 'ingredientes', 'pasos', 'consejos'];
// Lista separada por « | » (cómoda de leer y corregir en la hoja).
const CAMPOS_LISTA = ['categorias'];


/* ============ Puesta en marcha ============ */

/** Ejecútala una vez desde el editor para crear la hoja, la carpeta y dar permisos. */
function preparar() {
  const h = hoja_();
  // Si la hoja tenía la cabecera antigua, la actualiza.
  h.getRange(1, 1, 1, CAMPOS.length).setValues([CAMPOS]).setFontWeight('bold');
  carpeta_();
  marcarVersion_();
  Logger.log('Listo. Ya puedes implementar la aplicación web.');
}


/* ============ Lectura pública (GET) ============ */

function doGet(e) {
  const accion = (e && e.parameter && e.parameter.accion) || 'listar';
  try {
    if (accion === 'listar') {
      return json_({ ok: true, version: version_(), recetas: leerRecetas_() });
    }
    if (accion === 'version') {
      return json_({ ok: true, version: version_(), total: contar_() });
    }
    return json_({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}


/* ============ Escritura (POST) ============ */

function doPost(e) {
  let datos;
  try {
    datos = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Datos no válidos' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Carga inicial: solo se permite sin clave mientras la hoja está vacía.
    if (datos && datos.accion === 'importar' && contar_() === 0 && datos.clave === undefined) {
      return json_({ ok: true, importadas: importar_(datos.recetas || []) });
    }
    if (!datos || datos.clave !== CLAVE_ADMIN) {
      return json_({ ok: false, error: 'Clave incorrecta' });
    }
    switch (datos.accion) {
      case 'comprobar':
        return json_({ ok: true });
      case 'guardar':
        return json_({ ok: true, receta: guardarReceta_(datos.receta || {}, datos.foto, datos.quitarFoto), version: version_() });
      case 'borrar':
        borrarReceta_(datos.id);
        return json_({ ok: true, version: version_() });
      case 'importar':
        return json_({ ok: true, importadas: importar_(datos.recetas || []) });
      case 'fotosSites':
        return json_({ ok: true, resultado: fotosSites_(datos.categoria, datos.items || []), version: version_() });
      default:
        return json_({ ok: false, error: 'Acción desconocida' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}


/* ============ Recetas ============ */

function leerRecetas_() {
  const h = hoja_();
  const n = h.getLastRow();
  if (n < 2) return [];
  return h.getRange(2, 1, n - 1, CAMPOS.length).getValues()
    .filter(f => String(f[1]).trim() !== '')
    .map(filaAObjeto_);
}

function contar_() {
  const h = hoja_();
  const n = h.getLastRow();
  if (n < 2) return 0;
  return h.getRange(2, 2, n - 1, 1).getValues().filter(f => String(f[0]).trim() !== '').length;
}

function guardarReceta_(r, foto, quitarFoto) {
  if (!r.titulo || !String(r.titulo).trim()) throw new Error('La receta necesita un título');
  const h = hoja_();
  const ahora = new Date().toISOString();
  const fila = buscarFila_(r.id);
  const anterior = fila ? filaAObjeto_(h.getRange(fila, 1, 1, CAMPOS.length).getValues()[0]) : null;

  const receta = Object.assign({}, anterior || {}, r);
  receta.id = (anterior && anterior.id) || nuevoId_();
  receta.creada = (anterior && anterior.creada) || ahora;
  receta.actualizada = ahora;

  if (foto) {
    if (anterior && anterior.foto) borrarFoto_(anterior.foto);
    receta.foto = subirFoto_(foto, receta.titulo);
  } else if (quitarFoto) {
    if (anterior && anterior.foto) borrarFoto_(anterior.foto);
    receta.foto = '';
  }

  const valores = [objetoAFila_(receta)];
  if (fila) h.getRange(fila, 1, 1, CAMPOS.length).setValues(valores);
  else h.getRange(h.getLastRow() + 1, 1, 1, CAMPOS.length).setValues(valores);
  marcarVersion_();
  return filaAObjeto_(valores[0]);
}

function borrarReceta_(id) {
  const fila = buscarFila_(id);
  if (!fila) throw new Error('No existe esa receta');
  const h = hoja_();
  const receta = filaAObjeto_(h.getRange(fila, 1, 1, CAMPOS.length).getValues()[0]);
  if (receta.foto) borrarFoto_(receta.foto);
  h.deleteRow(fila);
  marcarVersion_();
}

/** Carga masiva. Omite títulos repetidos. */
function importar_(lista) {
  const h = hoja_();
  const existentes = new Set(leerRecetas_().map(r => normalizar_(r.titulo)));
  const ahora = new Date().toISOString();
  const nuevas = [];
  lista.forEach(r => {
    const clave = normalizar_(r.titulo);
    if (!clave || existentes.has(clave)) return;
    existentes.add(clave);
    nuevas.push(objetoAFila_(Object.assign({}, r, {
      id: r.id || nuevoId_(),
      creada: r.creada || ahora,
      actualizada: ahora
    })));
  });
  if (nuevas.length) {
    h.getRange(h.getLastRow() + 1, 1, nuevas.length, CAMPOS.length).setValues(nuevas);
    marcarVersion_();
  }
  return nuevas.length;
}


/* ============ Fotos de la antigua web de Google Sites ============ */

const WEB_SITES = 'https://sites.google.com/g.educaand.es/recetaweb/';

/**
 * Copia a Drive las fotos de una página de categoría de la web antigua.
 * items: [{ id, indice }] — indice = posición de la imagen en la página.
 * Solo pone foto a las recetas que todavía no tienen.
 */
function fotosSites_(categoria, items) {
  const html = UrlFetchApp.fetch(WEB_SITES + encodeURIComponent(categoria), { muteHttpExceptions: true }).getContentText();
  const fuentes = [];
  const re = /<img[^>]*\ssrc="([^"]+)"/g;
  let m;
  while ((m = re.exec(html))) fuentes.push(m[1].replace(/&amp;/g, '&').split('=')[0]);

  const h = hoja_();
  const res = [];
  items.forEach(it => {
    try {
      const fila = buscarFila_(it.id);
      if (!fila) { res.push({ id: it.id, ok: false, error: 'Receta no encontrada' }); return; }
      const receta = filaAObjeto_(h.getRange(fila, 1, 1, CAMPOS.length).getValues()[0]);
      if (receta.foto) { res.push({ id: it.id, ok: true, omitida: true }); return; }
      const src = fuentes[it.indice];
      if (!src) { res.push({ id: it.id, ok: false, error: 'Imagen no encontrada' }); return; }
      const r = UrlFetchApp.fetch(src + '=w1600', { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) { res.push({ id: it.id, ok: false, error: 'HTTP ' + r.getResponseCode() }); return; }
      const blob = r.getBlob();
      const tipo = blob.getContentType() || 'image/jpeg';
      const ext = (tipo.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
      blob.setName(normalizar_(receta.titulo).replace(/\s+/g, '-') + '.' + ext);
      const archivo = carpeta_().createFile(blob);
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      receta.foto = 'https://lh3.googleusercontent.com/d/' + archivo.getId();
      receta.actualizada = new Date().toISOString();
      h.getRange(fila, 1, 1, CAMPOS.length).setValues([objetoAFila_(receta)]);
      res.push({ id: it.id, ok: true, foto: receta.foto });
    } catch (err) {
      res.push({ id: it.id, ok: false, error: String(err) });
    }
  });
  marcarVersion_();
  return res;
}


/**
 * Ejecútala desde el editor (▶ traerFotosDeSites) para copiar a Drive las fotos
 * de la antigua web. Solo toca recetas sin foto; se puede repetir sin problema.
 */
function traerFotosDeSites() {
  const MAPA = {"ensalada-templada-de-pollo-manzana-y-mango": ["ensaladas-y-platos-fríos", 2], "salpicón-de-melón-pulpo-langostinos-y-mejillones": ["ensaladas-y-platos-fríos", 3], "pipirrana": ["ensaladas-y-platos-fríos", 4], "aderezo-de-lima-perejil-y-cilantro": ["ensaladas-y-platos-fríos", 5], "gazpacho-de-sandía": ["ensaladas-y-platos-fríos", 6], "gazpacho-de-melón": ["ensaladas-y-platos-fríos", 7], "tartar-de-jamón-y-tomate-con-crema-de-melón": ["ensaladas-y-platos-fríos", 8], "atún-con-verduras": ["airfryer-horno", 2], "pastel-de-berenjena-jamón-y-queso": ["airfryer-horno", 3], "lasaña-de-calabacín-y-atún": ["airfryer-horno", 4], "burger-pavo-y-espinacas": ["carnes-y-salteados", 2], "caramel-beef-burguebab": ["carnes-y-salteados", 3], "burger-de-zanahoria": ["carnes-y-salteados", 4], "pollo-al-ajillo-gourmet": ["carnes-y-salteados", 5], "solomillo-de-pollo-glaseado-con-verduras": ["carnes-y-salteados", 6], "pollo-al-ajillo-cremoso-de-queso": ["carnes-y-salteados", 7], "stronoff-de-pavo-con-setas": ["carnes-y-salteados", 8], "pollo-en-salsa-cremosa-de-queso-y-zanahoria": ["carnes-y-salteados", 9], "pollo-a-la-cerveza": ["carnes-y-salteados", 10], "carne-con-tomate": ["carnes-y-salteados", 11], "pollo-thai-cremoso": ["carnes-y-salteados", 12], "habitas-baby-con-jamón-y-huevo-gourmet": ["carnes-y-salteados", 13], "solomillo-de-cerdo-con-salsa-de-almendras": ["carnes-y-salteados", 14], "solomillo-de-pavo-con-salsa-de-castañas": ["carnes-y-salteados", 15], "cabeza-de-lomo-con-puré-de-castañas": ["carnes-y-salteados", 16], "atún-encebollao": ["pescado", 2], "salmón-soja-miel-y-limón": ["pescado", 3], "dorada-al-vapor-con-salsa-de-almendras": ["pescado", 4], "combo-fideuá-y-salmón-al-vapor": ["pescado", 5], "sepia-con-marisco": ["pescado", 6], "merluza-con-crema-de-calabaza-y-langostinos": ["pescado", 7], "merluza-en-salsa-de-guisantes": ["pescado", 8], "bacalao-al-pil-pil-con-almejas": ["pescado", 9], "salmón-con-calamaritos-en-sofrito-de-tomate-y-pimiento": ["pescado", 10], "merluza-al-vapor-con-salsa": ["pescado", 11], "merluza-al-horno-con-salsa-cremosa-de-zanahoria-y-queso": ["pescado", 12], "pisto-en-mc": ["guisos", 2], "garbanzos-con-calabaza-y-espinacas": ["guisos", 3], "garbanzos-con-berenjenas": ["guisos", 4], "estofado-de-carne": ["guisos", 5], "lentejas": ["guisos", 6], "alubias-con-crema-de-anacardos": ["guisos", 7], "estofado-de-conejo": ["guisos", 8], "curry-de-champiniones-y-garbanzos": ["guisos", 9], "alubias-con-langostinos": ["guisos", 10], "calamares-con-garbanzos": ["guisos", 14], "potaje-de-habichuelas": ["guisos", 15], "garbanzos-a-la-carbonara": ["guisos", 16], "garbanzos-a-la-marinera": ["guisos", 17], "salsa-de-pimiento-rojo-y-calabacín": ["cremas-y-sopas", 2], "puré-de-patata-con-zanahoria": ["cremas-y-sopas", 3], "crema-de-calabacín": ["cremas-y-sopas", 4], "mayonesa-de-zanahoria": ["cremas-y-sopas", 5], "crema-de-castañas-con-setas": ["cremas-y-sopas", 6], "crema-de-calabaza": ["cremas-y-sopas", 7], "sopa-castellana-sopa-de-ajo": ["cremas-y-sopas", 8], "crema-de-zanahoria": ["cremas-y-sopas", 9], "crema-de-lombarda-manzana-y-nueces": ["cremas-y-sopas", 10], "crema-fría-de-remolacha-con-yogur-y-cebollino": ["cremas-y-sopas", 12], "paella-mixta": ["arroces", 2], "risotto-de-setas-y-jamón": ["arroces", 3], "paella-marinera": ["arroces", 4], "arroz-meloso-con-setas-y-manzana": ["arroces", 5], "arroz-meloso-con-pollo-y-champiniones": ["arroces", 6], "kalimotxo": ["bebidas-y-elixires", 2], "omelette-mediterránea": ["omelettes-y-huevos", 2], "omelette-champis-cebolla-y-espinacas": ["omelettes-y-huevos", 3], "revuelto-de-huevos": ["omelettes-y-huevos", 4], "lomo-embuchado-casero": ["aperitivos", 2], "rollitos-salados-de-espinacas-sal": ["aperitivos", 3]};
  const porCat = {};
  leerRecetas_().forEach(r => {
    const m = MAPA[r.origen];
    if (m && !r.foto) (porCat[m[0]] = porCat[m[0]] || []).push({ id: r.id, indice: m[1] });
  });
  let bien = 0, mal = 0;
  Object.keys(porCat).forEach(c => {
    fotosSites_(c, porCat[c]).forEach(x => { if (x.ok) bien++; else { mal++; Logger.log(c + ' · ' + x.id + ' · ' + x.error); } });
  });
  Logger.log('Fotos copiadas: ' + bien + ' · con error: ' + mal);
}


/* ============ Fotos en Drive ============ */

function subirFoto_(dataUrl, titulo) {
  const m = String(dataUrl).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) throw new Error('Formato de foto no válido');
  const extension = m[1].split('/')[1].replace('jpeg', 'jpg');
  const nombre = normalizar_(titulo).replace(/\s+/g, '-') + '-' + Date.now() + '.' + extension;
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nombre);
  const archivo = carpeta_().createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + archivo.getId();
}

function borrarFoto_(url) {
  const m = String(url).match(/googleusercontent\.com\/d\/([\w-]+)/);
  if (!m) return;
  try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (err) { /* ya no existe */ }
}

function carpeta_() {
  const it = DriveApp.getFoldersByName(CARPETA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FOTOS);
}


/* ============ Utilidades ============ */

function hoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA);
  if (!h) {
    h = libro.insertSheet(HOJA);
    h.getRange(1, 1, 1, CAMPOS.length).setValues([CAMPOS]).setFontWeight('bold');
    h.setFrozenRows(1);
    h.getRange('A:Z').setVerticalAlignment('top');
  }
  return h;
}

function buscarFila_(id) {
  if (!id) return 0;
  const h = hoja_();
  const n = h.getLastRow();
  if (n < 2) return 0;
  const ids = h.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function filaAObjeto_(fila) {
  const o = {};
  CAMPOS.forEach((c, i) => {
    let v = fila[i];
    if (v instanceof Date) v = v.toISOString();
    if (CAMPOS_JSON.includes(c)) {
      try { v = v === '' ? (c === 'macros' ? null : []) : JSON.parse(v); }
      catch (err) { v = c === 'macros' ? null : []; }
    } else if (CAMPOS_LISTA.includes(c)) {
      v = String(v || '').split(/\s*\|\s*/).map(s => s.trim()).filter(Boolean);
    } else if (c === 'raciones') {
      v = Number(v) || 1;
    } else {
      v = v === null ? '' : String(v);
    }
    o[c] = v;
  });
  return o;
}

function objetoAFila_(o) {
  return CAMPOS.map(c => {
    let v = o[c];
    if (CAMPOS_JSON.includes(c)) return v == null ? '' : "'" + JSON.stringify(v);
    if (CAMPOS_LISTA.includes(c) && Array.isArray(v)) v = v.join(' | ');
    if (v === undefined || v === null) v = '';
    v = String(v);
    return v.startsWith('=') ? "'" + v : v; // Evita que la hoja lo trate como fórmula.
  });
}

function normalizar_(t) {
  return String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9ñ ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function nuevoId_() {
  return Utilities.getUuid().slice(0, 8);
}

function version_() {
  return PropertiesService.getScriptProperties().getProperty('version') || '0';
}

function marcarVersion_() {
  PropertiesService.getScriptProperties().setProperty('version', String(Date.now()));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
