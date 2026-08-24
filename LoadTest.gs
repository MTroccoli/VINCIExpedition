/**
 * PRUEBAS DE CARGA - VINCI Expedition
 *
 * Este archivo NO forma parte del sistema de votacion. Es todo el aparato de
 * testing junto: borra el archivo del proyecto y no queda ningun rastro,
 * incluido el endpoint doPost de mas abajo.
 *
 * ANTES DE LA GALA: borra este archivo, o al menos deja LOADTEST.TOKEN vacio.
 * Con el token vacio el endpoint rechaza todo.
 *
 * COMO USARLO
 *   1. simularVotos(100)              -> mide la latencia por voto
 *   2. limpiarVotosDePrueba()         -> borra las filas que dejo el test
 *   3. verResumenPrueba()             -> imprime el ultimo resultado guardado
 *
 * Para medir concurrencia real (opcional, requiere el web app desplegado):
 *   1. Poner LOADTEST.TOKEN y LOADTEST.WEBAPP_URL
 *   2. Volver a desplegar (el doPost tiene que estar en la version publicada)
 *   3. simularVotosConcurrentes(100)
 *   4. limpiarVotosDePrueba()
 */

var LOADTEST = {
  PREFIJO: 'loadtest-',
  DOMINIO: 'loadtest.invalid',

  // Pais desde el que "vota" el test. Conviene uno sin categoria (Chile,
  // Mexico o Turquia): asi no se le descuenta ninguna opcion y el test elige
  // en las tres categorias, como la mayoria de la gente.
  PAIS_VOTANTE: 'Chile',

  // Corta la corrida antes del limite de 6 minutos de Apps Script.
  MAX_MS: 4 * 60 * 1000,

  // Solo para simularVotosConcurrentes. Dejar ambos vacios desactiva el doPost.
  TOKEN: '',
  WEBAPP_URL: '',
  LOTE: 25
};

// =============================================
// TEST 1: LATENCIA (secuencial)
// =============================================

/**
 * Vota n veces seguidas midiendo cada llamada.
 *
 * Que mide: cuanto tarda un voto de punta a punta en el servidor. Como el
 * ScriptLock serializa la escritura, el throughput maximo del sistema es
 * 1 / latencia, asi que este numero es el techo real de votos por segundo.
 *
 * Que NO mide: la contencion del lock ni el cupo de ejecuciones simultaneas
 * de Apps Script. Para eso esta simularVotosConcurrentes.
 */
function simularVotos(n) {
  n = n || 100;

  var votos = votosDePrueba_();
  var tiempos = [];
  var ok = 0;
  var fallos = {};
  var cortado = false;
  var inicio = Date.now();

  for (var i = 1; i <= n; i++) {
    if (Date.now() - inicio > LOADTEST.MAX_MS) {
      cortado = true;
      n = i - 1;
      break;
    }

    var t0 = Date.now();
    var res = registrarVoto(emailDePrueba_(i), votos, LOADTEST.PAIS_VOTANTE);
    tiempos.push(Date.now() - t0);

    if (res && res.ok) {
      ok++;
    } else {
      var msg = (res && res.msg) || 'sin respuesta';
      fallos[msg] = (fallos[msg] || 0) + 1;
    }
  }

  var totalMs = Date.now() - inicio;
  var resumen = armarResumen_('secuencial', n, ok, fallos, tiempos, totalMs, cortado);
  guardarResumen_(resumen);
  Logger.log(formatearResumen_(resumen));
  return resumen;
}

// =============================================
// TEST 2: CONCURRENCIA (contra el web app)
// =============================================

/**
 * Dispara n votos en paralelo contra la URL /exec con UrlFetchApp.fetchAll,
 * en lotes de LOADTEST.LOTE. Cada request abre su propia ejecucion del web
 * app, asi que esto si golpea el lock y el cupo de ejecuciones simultaneas.
 */
function simularVotosConcurrentes(n) {
  n = n || 100;

  if (!LOADTEST.TOKEN) {
    return { ok: false, msg: 'Configura LOADTEST.TOKEN (y volve a desplegar) antes de correr esto.' };
  }
  if (!LOADTEST.WEBAPP_URL) {
    return { ok: false, msg: 'Configura LOADTEST.WEBAPP_URL con la URL /exec del web app.' };
  }

  var token = ScriptApp.getOAuthToken();
  var votosJson = JSON.stringify(votosDePrueba_());
  var tiempos = [];
  var ok = 0;
  var fallos = {};
  var cortado = false;
  var inicio = Date.now();
  var hechos = 0;

  for (var desde = 1; desde <= n; desde += LOADTEST.LOTE) {
    if (Date.now() - inicio > LOADTEST.MAX_MS) {
      cortado = true;
      break;
    }

    var hasta = Math.min(desde + LOADTEST.LOTE - 1, n);
    var requests = [];

    for (var i = desde; i <= hasta; i++) {
      requests.push({
        url: LOADTEST.WEBAPP_URL,
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: {
          token: LOADTEST.TOKEN,
          email: emailDePrueba_(i),
          votos: votosJson,
          paisVotante: LOADTEST.PAIS_VOTANTE
        },
        muteHttpExceptions: true,
        followRedirects: true
      });
    }

    var t0 = Date.now();
    var respuestas;
    try {
      respuestas = UrlFetchApp.fetchAll(requests);
    } catch(e) {
      fallos['fetchAll fallo: ' + e.message] = (fallos['fetchAll fallo: ' + e.message] || 0) + requests.length;
      hechos += requests.length;
      continue;
    }
    var loteMs = Date.now() - t0;

    // fetchAll devuelve cuando termino todo el lote, asi que el tiempo por
    // voto es una estimacion: el lote entero repartido entre sus requests.
    for (var j = 0; j < respuestas.length; j++) {
      tiempos.push(Math.round(loteMs / respuestas.length));
      var r = respuestas[j];

      if (r.getResponseCode() !== 200) {
        var claveHttp = 'HTTP ' + r.getResponseCode();
        fallos[claveHttp] = (fallos[claveHttp] || 0) + 1;
        continue;
      }

      var cuerpo;
      try {
        cuerpo = JSON.parse(r.getContentText());
      } catch(e) {
        // Si devuelve HTML en vez de JSON casi siempre es la pantalla de login:
        // el web app no esta accesible con este token.
        fallos['respuesta no JSON (revisar acceso del web app)'] =
          (fallos['respuesta no JSON (revisar acceso del web app)'] || 0) + 1;
        continue;
      }

      if (cuerpo.ok) {
        ok++;
      } else {
        var m = cuerpo.msg || 'sin mensaje';
        fallos[m] = (fallos[m] || 0) + 1;
      }
    }

    hechos += respuestas.length;
  }

  var totalMs = Date.now() - inicio;
  var resumen = armarResumen_('concurrente x' + LOADTEST.LOTE, hechos, ok, fallos, tiempos, totalMs, cortado);
  guardarResumen_(resumen);
  Logger.log(formatearResumen_(resumen));
  return resumen;
}

/**
 * Endpoint para el test concurrente. Deshabilitado mientras LOADTEST.TOKEN
 * este vacio. Borrar este archivo lo elimina por completo.
 */
function doPost(e) {
  var responder = function(obj) {
    return ContentService
      .createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
  };

  if (!LOADTEST.TOKEN) {
    return responder({ ok: false, msg: 'Endpoint deshabilitado.' });
  }

  var p = (e && e.parameter) || {};
  if (p.token !== LOADTEST.TOKEN) {
    return responder({ ok: false, msg: 'Token invalido.' });
  }

  var votos;
  try {
    votos = JSON.parse(p.votos || '{}');
  } catch(err) {
    return responder({ ok: false, msg: 'votos no es JSON valido.' });
  }

  return responder(registrarVoto(p.email, votos, p.paisVotante));
}

// =============================================
// LIMPIEZA
// =============================================

/**
 * Borra de la hoja Votos todas las filas cuyo email arranque con el prefijo
 * de prueba. No toca los votos reales.
 */
function limpiarVotosDePrueba() {
  var sheet = obtenerHoja_().getSheetByName(HOJAS.VOTOS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { ok: true, borrados: 0, msg: 'No hay votos que borrar.' };
  }

  var emails = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  var borrados = 0;

  // De abajo hacia arriba: borrar de arriba corre los indices de las de abajo.
  for (var i = emails.length - 1; i >= 0; i--) {
    var e = emails[i][0] ? emails[i][0].toString().toLowerCase() : '';
    if (e.indexOf(LOADTEST.PREFIJO) === 0) {
      sheet.deleteRow(i + 2);
      borrados++;
    }
  }

  invalidarCacheVotos_();

  var msg = 'Borradas ' + borrados + ' filas de prueba.';
  Logger.log(msg);
  return { ok: true, borrados: borrados, msg: msg };
}

// =============================================
// AUXILIARES
// =============================================

function emailDePrueba_(i) {
  return LOADTEST.PREFIJO + ('0000' + i).slice(-4) + '@' + LOADTEST.DOMINIO;
}

/**
 * Arma un voto valido: el primer pais disponible de cada categoria para este
 * votante. No reparte los votos, no interesa el ranking sino el tiempo.
 */
function votosDePrueba_() {
  var votos = {};
  obtenerCategoriasParaVotar(LOADTEST.PAIS_VOTANTE).forEach(function(cat) {
    votos[cat.id] = cat.paises[0];
  });
  return votos;
}

function percentil_(ordenados, p) {
  if (!ordenados.length) return 0;
  var idx = Math.min(ordenados.length - 1, Math.floor(ordenados.length * p));
  return ordenados[idx];
}

function armarResumen_(modo, n, ok, fallos, tiempos, totalMs, cortado) {
  var ordenados = tiempos.slice().sort(function(a, b) { return a - b; });
  var suma = 0;
  for (var i = 0; i < tiempos.length; i++) suma += tiempos[i];

  var segundos = totalMs / 1000;
  var porSegundo = segundos > 0 ? n / segundos : 0;

  return {
    modo: modo,
    votos: n,
    ok: ok,
    fallidos: n - ok,
    fallos: fallos,
    cortadoPorTiempo: cortado,
    totalMs: totalMs,
    promedioMs: tiempos.length ? Math.round(suma / tiempos.length) : 0,
    minMs: ordenados.length ? ordenados[0] : 0,
    p50Ms: percentil_(ordenados, 0.50),
    p95Ms: percentil_(ordenados, 0.95),
    maxMs: ordenados.length ? ordenados[ordenados.length - 1] : 0,
    votosPorSegundo: Math.round(porSegundo * 100) / 100,
    // A 1000 votantes, cuanto tardaria la votacion entera a este ritmo.
    minutosPara1000: porSegundo > 0 ? Math.round((1000 / porSegundo / 60) * 10) / 10 : null,
    cuando: new Date().toISOString()
  };
}

function formatearResumen_(r) {
  var lineas = [
    'PRUEBA DE CARGA (' + r.modo + ')',
    '  votos enviados : ' + r.votos + (r.cortadoPorTiempo ? ' (cortado por limite de tiempo)' : ''),
    '  exitosos       : ' + r.ok,
    '  fallidos       : ' + r.fallidos,
    '  tiempo total   : ' + (Math.round(r.totalMs / 100) / 10) + ' s',
    '  latencia ms    : min ' + r.minMs + ' | p50 ' + r.p50Ms + ' | p95 ' + r.p95Ms + ' | max ' + r.maxMs,
    '  promedio ms    : ' + r.promedioMs,
    '  votos/segundo  : ' + r.votosPorSegundo,
    '  1000 votos en  : ' + (r.minutosPara1000 !== null ? r.minutosPara1000 + ' min' : 'n/d'),
    '  objetivo       : 3.33 votos/seg (1000 en 5 min)'
  ];

  var claves = Object.keys(r.fallos);
  if (claves.length) {
    lineas.push('  detalle de fallos:');
    claves.forEach(function(k) {
      lineas.push('    ' + r.fallos[k] + 'x ' + k);
    });
  }

  lineas.push(r.votosPorSegundo >= 3.33
    ? '  RESULTADO: entra en la ventana de 5 minutos.'
    : '  RESULTADO: NO entra en 5 minutos a este ritmo.');

  return lineas.join('\n');
}

function guardarResumen_(resumen) {
  try {
    PropertiesService.getScriptProperties()
      .setProperty('ULTIMA_PRUEBA_CARGA', JSON.stringify(resumen));
  } catch(e) {}
}

function verResumenPrueba() {
  var raw = PropertiesService.getScriptProperties().getProperty('ULTIMA_PRUEBA_CARGA');
  if (!raw) {
    Logger.log('Todavia no corriste ninguna prueba.');
    return null;
  }
  var resumen = JSON.parse(raw);
  Logger.log(formatearResumen_(resumen));
  return resumen;
}
