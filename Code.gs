/**
 * VINCI Expedition - Sistema de Votacion
 * Septiembre 2026
 *
 * INSTRUCCIONES DE CONFIGURACION:
 *
 * OPCION A - Manual:
 * 1. Crear un nuevo proyecto en Google Apps Script (script.google.com)
 * 2. Copiar este archivo como Code.gs
 * 3. Copiar Index.html como archivo HTML en el proyecto
 * 4. Ejecutar la funcion inicializarSistema() UNA VEZ desde el editor
 * 5. Desplegar > Nueva implementacion > Aplicacion web
 *
 * OPCION B - Con clasp (recomendada):
 * 1. npm install -g @google/clasp
 * 2. clasp login
 * 3. clasp create --type webapp --title "VINCI Expedition"
 *    (o clasp clone <scriptId> si ya existe el proyecto)
 * 4. clasp push
 * 5. clasp deploy
 * 6. Ejecutar inicializarSistema() una vez desde el editor
 *
 * DESPLIEGUE:
 * - Ejecutar como: Tu cuenta
 * - Acceso: "Cualquier persona de tu organizacion" (auto-detecta email)
 *           O "Cualquier persona" (requiere login manual)
 */

const CONFIG = {
  ADMIN_PASSWORD: 'gvarguru26',
  PESO_PONDERADO: 0.80,
  PESO_REGULAR: 0.20,
  CACHE_EMAILS_VOTARON: 300,
  CACHE_RESULTADOS: 15,
  CACHE_PONDERADOS: 1500,
  CACHE_IMAGENES: 1500,
  SLIDES_ID: '1Mkj7SD3SOSkDjyAjGjIIJob-TqvHu9rrrAAjsBvGNXI',
  CARPETA_IMAGENES: 'VINCI Expedition - Imagenes',

};

/**
 * Cada votante elige un pais por categoria. Un pais pertenece a una sola
 * categoria, y los que no figuran en ninguna no se votan: hoy Chile, Mexico
 * y Turquia. Para sumar un pais, agregalo a la lista de su categoria y volve
 * a correr limpiarImagenes() + exportarImagenesSlides().
 *
 * 'columna' es el encabezado que lleva su voto en la hoja Votos.
 */
const CATEGORIAS = [
  {
    id: 'mejora_continua',
    nombre: 'Mejora Continua',
    columna: 'VotoMejoraContinua',
    paises: ['Argentina', 'CIB', 'Peru']
  },
  {
    id: 'ia_modelos',
    nombre: 'IA / Modelos',
    columna: 'VotoIAModelos',
    paises: ['Colombia', 'Holding', 'Venezuela']
  },
  {
    id: 'proyecto_sda',
    nombre: 'Proyecto SDA',
    columna: 'VotoProyectoSDA',
    paises: ['Espana', 'Uruguay']
  }
];

const PAISES = [
  'Argentina',
  'Chile',
  'CIB',
  'Colombia',
  'Espana',
  'Holding',
  'Mexico',
  'Peru',
  'Turquia',
  'Uruguay',
  'Venezuela'
];

function categoriaDe_(pais) {
  for (var i = 0; i < CATEGORIAS.length; i++) {
    if (CATEGORIAS[i].paises.indexOf(pais) !== -1) return CATEGORIAS[i];
  }
  return null;
}

function categoriaPorId_(id) {
  for (var i = 0; i < CATEGORIAS.length; i++) {
    if (CATEGORIAS[i].id === id) return CATEGORIAS[i];
  }
  return null;
}

function paisEnVotacion_(pais) {
  return categoriaDe_(pais) !== null;
}

function encabezadosVotos_() {
  var h = ['Timestamp', 'Email', 'PaisVotante', 'EsPonderado'];
  CATEGORIAS.forEach(function(c) { h.push(c.columna); });
  return h;
}

const HOJAS = {
  VOTOS: 'Votos',
  LISTA_PONDERADA: 'Lista Ponderada',
  IMAGENES: 'Imagenes'
};

// =============================================
// PUNTO DE ENTRADA
// =============================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('VINCI Expedition')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =============================================
// INICIALIZACION (ejecutar una sola vez)
// =============================================

function inicializarSistema() {
  var ss = obtenerHoja_();
  var nombres = ss.getSheets().map(function(s) { return s.getName(); });

  if (nombres.indexOf(HOJAS.VOTOS) === -1) {
    var sv = ss.insertSheet(HOJAS.VOTOS);
    var encabezados = encabezadosVotos_();
    sv.appendRow(encabezados);
    sv.getRange(1, 1, 1, encabezados.length).setFontWeight('bold');
  }

  if (nombres.indexOf(HOJAS.LISTA_PONDERADA) === -1) {
    var sp = ss.insertSheet(HOJAS.LISTA_PONDERADA);
    sp.appendRow(['Email', 'Nombre', 'Pais', 'Cargo']);
    sp.getRange('A1:D1').setFontWeight('bold');
    sp.appendRow(['director.innovacion@bbva.com', 'Ejemplo Director', 'Espana', 'Director de Innovacion']);
  }

  try {
    var h1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
    if (h1 && ss.getSheets().length > 1) ss.deleteSheet(h1);
  } catch(e) {}

  Logger.log('Sistema listo. Hoja: ' + ss.getUrl());
  return { url: ss.getUrl(), id: ss.getId() };
}

// =============================================
// UTILIDADES PRIVADAS
// =============================================

function obtenerHoja_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  var ss = SpreadsheetApp.create('VINCI Expedition - Votacion');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function leerHoja_(nombre) {
  var sheet = obtenerHoja_().getSheetByName(nombre);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function obtenerEmailsVotaron_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('emails_votaron');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
  if (!sheet || sheet.getLastRow() <= 1) {
    cache.put('emails_votaron', '[]', CONFIG.CACHE_EMAILS_VOTARON);
    return [];
  }

  var emails = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return r[0] ? r[0].toString().toLowerCase().trim() : ''; })
    .filter(function(e) { return e; });

  guardarEnCache_('emails_votaron', emails, CONFIG.CACHE_EMAILS_VOTARON);
  return emails;
}

/**
 * CacheService rechaza valores de mas de 100KB. Con ~1000 votantes la lista
 * de emails ronda los 30KB, pero si se pasa preferimos no cachear a romper.
 */
function guardarEnCache_(clave, valor, ttl) {
  var json = JSON.stringify(valor);
  if (json.length > 90000) return false;
  try {
    CacheService.getScriptCache().put(clave, json, ttl);
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * Suma un email a la lista cacheada en vez de invalidarla. Invalidar obligaria
 * al proximo votante a releer toda la hoja, que es justo lo que no queremos
 * cuando entran todos juntos.
 */
function registrarEmailEnCache_(email) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('emails_votaron');
  if (!cached) return;
  try {
    var emails = JSON.parse(cached);
    if (emails.indexOf(email) === -1) {
      emails.push(email);
      guardarEnCache_('emails_votaron', emails, CONFIG.CACHE_EMAILS_VOTARON);
    }
  } catch(e) {}
}

function obtenerSetPonderados_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('ponderados');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var sheet = obtenerHoja_().getSheetByName(HOJAS.LISTA_PONDERADA);
  if (!sheet || sheet.getLastRow() <= 1) {
    guardarEnCache_('ponderados', [], CONFIG.CACHE_PONDERADOS);
    return [];
  }

  var lista = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return r[0] ? r[0].toString().toLowerCase().trim() : ''; })
    .filter(function(e) { return e; });

  guardarEnCache_('ponderados', lista, CONFIG.CACHE_PONDERADOS);
  return lista;
}

function invalidarCachePonderados_() {
  CacheService.getScriptCache().removeAll(['ponderados', 'resultados']);
}

/**
 * Vacia todos los caches. Correr desde el editor despues de editar a mano la
 * hoja "Lista Ponderada" o "Imagenes": si no, los cambios pueden tardar hasta
 * CACHE_PONDERADOS / CACHE_IMAGENES segundos en verse.
 */
function refrescarCaches() {
  CacheService.getScriptCache().removeAll([
    'emails_votaron', 'ponderados', 'imagenes', 'resultados'
  ]);
  Logger.log('Caches vaciados.');
  return { ok: true, msg: 'Caches vaciados.' };
}

function invalidarCacheVotos_() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['emails_votaron', 'resultados']);
}

function invalidarCacheResultados_() {
  CacheService.getScriptCache().remove('resultados');
}

// =============================================
// API PUBLICA - DETECCION DE EMAIL
// =============================================

/**
 * Email de quien esta usando el web app. Funciona porque el despliegue es
 * de dominio: la persona ya viene logueada con su cuenta BBVA. Si por lo que
 * sea no se puede leer, el front cae al campo de email escrito a mano.
 */
function obtenerEmailUsuario() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email && email.indexOf('@') !== -1) {
      return { ok: true, email: email };
    }
  } catch(e) {}
  return { ok: false, email: '' };
}

// =============================================
// API PUBLICA - REGISTRO
// =============================================

function obtenerPaises() {
  return PAISES;
}

/**
 * URL publica del web app, para armar el QR desde el panel admin.
 * Devuelve la del despliegue activo, asi que hay que llamarla desde el
 * web app publicado y no desde el editor.
 */
function obtenerUrlWebApp() {
  try {
    return { ok: true, url: ScriptApp.getService().getUrl() };
  } catch(e) {
    return { ok: false, url: '', msg: e.message };
  }
}

function validarUsuario(email, pais) {
  email = email.toLowerCase().trim();
  if (!email || email.indexOf('@') === -1) {
    return { ok: false, msg: 'Ingresa un email valido.' };
  }
  if (PAISES.indexOf(pais) === -1) {
    return { ok: false, msg: 'Selecciona un pais valido.' };
  }

  var votaron = obtenerEmailsVotaron_();
  if (votaron.indexOf(email) !== -1) {
    return { ok: false, msg: 'Ya registraste tu voto. Solo se permite un voto por persona.', yaVoto: true };
  }

  var ponderados = obtenerSetPonderados_();
  var esPonderado = ponderados.indexOf(email) !== -1;

  return { ok: true, email: email, pais: pais, esPonderado: esPonderado };
}

// =============================================
// API PUBLICA - VOTACION
// =============================================

/**
 * Devuelve las categorias con los paises que este votante puede elegir.
 * Se saca su propio pais, y si una categoria queda vacia no se ofrece.
 */
function obtenerCategoriasParaVotar(paisUsuario) {
  return CATEGORIAS.map(function(c) {
    return {
      id: c.id,
      nombre: c.nombre,
      paises: c.paises.filter(function(p) { return p !== paisUsuario; })
    };
  }).filter(function(c) {
    return c.paises.length > 0;
  });
}

/**
 * Devuelve los votos ya registrados de un email como { idCategoria: pais }.
 * Solo se llama en el camino del duplicado, que es raro, asi que la lectura
 * de hoja no pesa.
 */
function buscarVotosDe_(email) {
  var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
  if (!sheet || sheet.getLastRow() <= 1) return {};

  var ancho = encabezadosVotos_().length;
  var filas = sheet.getRange(2, 1, sheet.getLastRow() - 1, ancho).getValues();

  for (var i = 0; i < filas.length; i++) {
    if (filas[i][1] && filas[i][1].toString().toLowerCase().trim() === email) {
      var votos = {};
      CATEGORIAS.forEach(function(c, j) {
        // Las columnas de voto arrancan despues de las cuatro fijas.
        votos[c.id] = filas[i][4 + j] || '';
      });
      return votos;
    }
  }
  return {};
}

/**
 * Registra los votos de una persona: uno por categoria, todos en una sola
 * fila. Una fila por votante mantiene el append en una unica escritura, que
 * es lo que sostiene el ritmo cuando entran todos juntos.
 *
 * votos = { idCategoria: pais, ... }
 */
function registrarVoto(email, votos, paisVotante) {
  email = (email || '').toString().toLowerCase().trim();
  votos = votos || {};

  if (!email || email.indexOf('@') === -1) {
    return { ok: false, msg: 'Email invalido.' };
  }
  if (PAISES.indexOf(paisVotante) === -1) {
    return { ok: false, msg: 'Pais no valido.' };
  }

  var categorias = obtenerCategoriasParaVotar(paisVotante);
  var elegidos = {};

  for (var i = 0; i < categorias.length; i++) {
    var cat = categorias[i];
    var elegido = votos[cat.id];

    if (!elegido) {
      return { ok: false, msg: 'Falta elegir un pais en ' + cat.nombre + '.' };
    }
    if (cat.paises.indexOf(elegido) === -1) {
      return { ok: false, msg: 'La opcion elegida en ' + cat.nombre + ' no es valida.' };
    }
    elegidos[cat.id] = elegido;
  }

  // Chequeo de duplicado contra la lista cacheada, fuera del lock. No es
  // autoritativo: dos pestanas simultaneas del mismo email podrian colar dos
  // filas. El conteo se queda con el primer voto de cada email, asi que un
  // duplicado no altera el resultado.
  if (obtenerEmailsVotaron_().indexOf(email) !== -1) {
    return {
      ok: false,
      yaVoto: true,
      votos: buscarVotosDe_(email),
      msg: 'Ya registraste tu voto anteriormente.'
    };
  }

  var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
  if (!sheet) {
    return { ok: false, msg: 'Error de configuracion: no se encontro la hoja de votos. Ejecuta inicializarSistema().' };
  }

  // La lista de ponderados esta cacheada, asi que esto no toca la hoja.
  var esPond = obtenerSetPonderados_().indexOf(email) !== -1;
  var fila = [new Date(), email, paisVotante, esPond ? 'SI' : 'NO'];
  CATEGORIAS.forEach(function(c) { fila.push(elegidos[c.id] || ''); });

  // Unico tramo serializado: el append. Todo lo caro quedo afuera, asi que la
  // seccion critica es un solo viaje a Sheets y la cola avanza rapido.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    sheet.appendRow(fila);
  } catch(e) {
    return { ok: false, msg: 'Hubo mucha demanda y no se pudo registrar. Reintenta en unos segundos.' };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }

  registrarEmailEnCache_(email);
  invalidarCacheResultados_();

  return {
    ok: true,
    msg: 'Votos registrados exitosamente!',
    votos: elegidos
  };
}

// =============================================
// API PUBLICA - ADMIN
// =============================================

function loginAdmin(password) {
  return password === CONFIG.ADMIN_PASSWORD;
}

function obtenerResultadosAdmin(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };

  var cache = CacheService.getScriptCache();
  var cached = cache.get('resultados');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  // Sin lock en la escritura pueden colarse filas duplicadas del mismo email:
  // vale la primera fila, el resto se descarta al contar.
  var vistos = {};
  var votantes = leerHoja_(HOJAS.VOTOS).filter(function(v) {
    var e = v.Email ? v.Email.toString().toLowerCase().trim() : '';
    if (!e || vistos[e]) return false;
    vistos[e] = true;
    return true;
  });

  var totalPond = 0;
  var totalReg = 0;
  votantes.forEach(function(v) {
    if (v.EsPonderado === 'SI') totalPond++; else totalReg++;
  });

  // El 80/20 se calcula dentro de cada categoria: el denominador es la gente
  // que efectivamente voto en esa categoria, no el total de votantes.
  var categorias = CATEGORIAS.map(function(cat) {
    var conteosPond = {};
    var conteosReg = {};
    cat.paises.forEach(function(p) { conteosPond[p] = 0; conteosReg[p] = 0; });

    var catPond = 0;
    var catReg = 0;

    votantes.forEach(function(v) {
      var elegido = v[cat.columna];
      if (!elegido || conteosPond[elegido] === undefined) return;

      if (v.EsPonderado === 'SI') {
        catPond++;
        conteosPond[elegido]++;
      } else {
        catReg++;
        conteosReg[elegido]++;
      }
    });

    var ranking = cat.paises.map(function(pais) {
      var vp = conteosPond[pais];
      var vr = conteosReg[pais];
      var sp = catPond > 0 ? (vp / catPond) * CONFIG.PESO_PONDERADO * 100 : 0;
      var sr = catReg > 0 ? (vr / catReg) * CONFIG.PESO_REGULAR * 100 : 0;

      return {
        pais: pais,
        votosPond: vp, votosReg: vr, totalVotos: vp + vr,
        scorePond: Math.round(sp * 100) / 100,
        scoreReg:  Math.round(sr * 100) / 100,
        scoreTotal: Math.round((sp + sr) * 100) / 100
      };
    });

    ranking.sort(function(a, b) { return b.scoreTotal - a.scoreTotal; });

    return {
      id: cat.id,
      nombre: cat.nombre,
      ranking: ranking,
      ponderados: catPond,
      regulares: catReg
    };
  });

  var detallePond = votantes
    .filter(function(v) { return v.EsPonderado === 'SI'; })
    .map(function(v) {
      var elecciones = CATEGORIAS.map(function(c) {
        return { categoria: c.nombre, pais: v[c.columna] || '-' };
      });
      return {
        email: v.Email,
        paisVotante: v.PaisVotante,
        elecciones: elecciones,
        fecha: v.Timestamp
      };
    });

  var resultado = {
    ok: true,
    categorias: categorias,
    stats: {
      total: votantes.length,
      ponderados: totalPond,
      regulares: totalReg,
      paises: CATEGORIAS.reduce(function(n, c) { return n + c.paises.length; }, 0)
    },
    detallePond: detallePond,
    actualizado: new Date().toISOString()
  };

  try { cache.put('resultados', JSON.stringify(resultado), CONFIG.CACHE_RESULTADOS); } catch(e) {}
  return resultado;
}

function obtenerListaPonderada(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  return { ok: true, lista: leerHoja_(HOJAS.LISTA_PONDERADA) };
}

function agregarUsuarioPonderado(password, datos) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };

  var ponderados = obtenerSetPonderados_();
  if (ponderados.indexOf(datos.email.toLowerCase()) !== -1) {
    return { ok: false, msg: 'El usuario ya esta en la lista.' };
  }

  obtenerHoja_().getSheetByName(HOJAS.LISTA_PONDERADA).appendRow([
    datos.email, datos.nombre || '', datos.pais || '', datos.cargo || ''
  ]);
  invalidarCachePonderados_();
  return { ok: true };
}

function eliminarUsuarioPonderado(password, email) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var sheet = obtenerHoja_().getSheetByName(HOJAS.LISTA_PONDERADA);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      invalidarCachePonderados_();
      return { ok: true };
    }
  }
  return { ok: false, msg: 'Usuario no encontrado.' };
}

function resetearVotos(password) {
  if (password !== CONFIG.ADMIN_PASSWORD) return { ok: false };
  var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  invalidarCacheVotos_();
  return { ok: true };
}

// =============================================
// IMAGENES DESDE GOOGLE SLIDES
// =============================================

function obtenerCarpetaImagenes_(crear) {
  var folders = DriveApp.getFoldersByName(CONFIG.CARPETA_IMAGENES);
  if (folders.hasNext()) return folders.next();
  return crear ? DriveApp.createFolder(CONFIG.CARPETA_IMAGENES) : null;
}

/**
 * Borra todo lo exportado: manda a papelera los PNG de la carpeta
 * y vacia la hoja "Imagenes". Correr antes de reexportar cuando
 * cambio el orden o la cantidad de slides.
 */
function limpiarImagenes() {
  var borrados = 0;
  var folder = obtenerCarpetaImagenes_(false);
  if (folder) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      files.next().setTrashed(true);
      borrados++;
    }
  }

  var sheet = obtenerHoja_().getSheetByName(HOJAS.IMAGENES);
  if (sheet) {
    sheet.clear();
    sheet.appendRow(['Pais', 'ImagenURL']);
    sheet.getRange('A1:B1').setFontWeight('bold');
  }

  invalidarCacheImagenes_();

  var msg = 'Limpieza lista: ' + borrados + ' archivos a papelera y hoja "' +
            HOJAS.IMAGENES + '" vaciada.';
  Logger.log(msg);
  return { ok: true, borrados: borrados, msg: msg };
}

/**
 * Ejecutar UNA VEZ (o cuando cambien las slides).
 * Exporta cada slide como PNG a una carpeta de Drive
 * y guarda el mapeo Pais -> URL en la hoja "Imagenes".
 *
 * PREREQUISITO: En el editor de Apps Script:
 *   1. Servicios (+) > Google Slides API > Agregar
 *   2. Esto habilita automaticamente la API en el Cloud Project
 *
 * Despues de ejecutar, revisar la hoja "Imagenes" y
 * corregir el mapeo si el orden de slides no coincide con PAISES.
 */
function exportarImagenesSlides() {
  var presId = CONFIG.SLIDES_ID;
  if (!presId) {
    return { ok: false, msg: 'Configura SLIDES_ID en CONFIG.' };
  }

  var token = ScriptApp.getOAuthToken();

  var presUrl = 'https://slides.googleapis.com/v1/presentations/' + presId;
  var presResp = UrlFetchApp.fetch(presUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (presResp.getResponseCode() !== 200) {
    var errorMsg = presResp.getContentText();
    Logger.log('Error accediendo a la presentacion: ' + errorMsg);
    if (errorMsg.indexOf('not enabled') !== -1 || errorMsg.indexOf('accessNotConfigured') !== -1) {
      return { ok: false, msg: 'La API de Slides no esta habilitada. En el editor de Apps Script: Servicios (+) > Google Slides API > Agregar.' };
    }
    return { ok: false, msg: 'No se pudo acceder a la presentacion. Verifica el ID y permisos. Codigo: ' + presResp.getResponseCode() };
  }

  var presData = JSON.parse(presResp.getContentText());
  var slides = presData.slides;
  if (!slides || slides.length === 0) {
    return { ok: false, msg: 'La presentacion no tiene slides.' };
  }

  var folder = obtenerCarpetaImagenes_(true);

  try {
    folder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    Logger.log('Compartir carpeta automaticamente no disponible: ' + e.message);
  }

  var ss = obtenerHoja_();
  var sheet = ss.getSheetByName(HOJAS.IMAGENES);
  if (!sheet) {
    sheet = ss.insertSheet(HOJAS.IMAGENES);
  } else {
    sheet.clear();
  }
  sheet.appendRow(['Pais', 'ImagenURL']);
  sheet.getRange('A1:B1').setFontWeight('bold');

  var total = 0;
  var errores = [];

  var omitidos = 0;

  for (var i = 0; i < slides.length; i++) {
    var pageId = slides[i].objectId;
    var pais = i < PAISES.length ? PAISES[i] : 'Slide ' + (i + 1);

    // El mapeo es posicional: el slide i corresponde a PAISES[i]. Los paises
    // sin categoria se saltean pero siguen ocupando su indice, asi que el
    // orden de la presentacion tiene que seguir siendo el de PAISES.
    if (!paisEnVotacion_(pais)) {
      omitidos++;
      continue;
    }

    try {
      var thumbUrl = 'https://slides.googleapis.com/v1/presentations/' + presId +
                     '/pages/' + pageId + '/thumbnail?thumbnailProperties.thumbnailSize=MEDIUM';
      var thumbResp = UrlFetchApp.fetch(thumbUrl, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      if (thumbResp.getResponseCode() !== 200) {
        errores.push('Slide ' + (i + 1) + ': HTTP ' + thumbResp.getResponseCode());
        continue;
      }

      var thumbData = JSON.parse(thumbResp.getContentText());
      var imageResp = UrlFetchApp.fetch(thumbData.contentUrl, { muteHttpExceptions: true });

      if (imageResp.getResponseCode() !== 200) {
        errores.push('Slide ' + (i + 1) + ': Error descargando imagen');
        continue;
      }

      var fileName = 'slide_' + (i + 1) + '.png';
      var imageBlob = imageResp.getBlob().setName(fileName);

      var existing = folder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);

      var file = folder.createFile(imageBlob);

      try {
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {}

      sheet.appendRow([pais, 'https://drive.google.com/uc?export=view&id=' + file.getId()]);
      total++;

    } catch(e) {
      errores.push('Slide ' + (i + 1) + ': ' + e.message);
    }
  }

  invalidarCacheImagenes_();

  var msg = 'Exportadas ' + total + ' de ' + slides.length + ' imagenes.';
  if (omitidos > 0) {
    msg += ' Omitidas ' + omitidos + ' (paises sin categoria).';
  }
  if (errores.length > 0) {
    msg += ' Errores: ' + errores.join('; ');
    Logger.log('Errores: ' + errores.join('; '));
  }
  msg += ' Revisa la hoja "Imagenes" y ajusta el mapeo Pais si es necesario.';

  Logger.log(msg);
  return { ok: total > 0, total: total, msg: msg };
}

// La llama cada visita al abrir la pagina, asi que va cacheada: si no, son
// mil lecturas de la hoja para devolver siempre lo mismo.
function obtenerImagenesPaises() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('imagenes');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var mapping = {};
  var sheet = obtenerHoja_().getSheetByName(HOJAS.IMAGENES);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(row) {
      if (row[0] && row[1]) mapping[row[0]] = row[1];
    });
  }

  guardarEnCache_('imagenes', mapping, CONFIG.CACHE_IMAGENES);
  return mapping;
}

function invalidarCacheImagenes_() {
  CacheService.getScriptCache().remove('imagenes');
}
