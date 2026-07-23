/**
 * Proxy de DISPONIBILIDAD Calendly — Dashboard COO Testo Boost — Google Apps Script
 *
 * Un mismo archivo para los dos equipos: cada implementación elige su equipo con
 * la propiedad de script EQUIPO ('2' o '3'). El token NUNCA va en el código.
 *
 * Qué resuelve respecto de la versión anterior:
 *  1. SEGURIDAD — token y organización salen de Propiedades del script.
 *  2. CUOTA — caché de 10 minutos: N cargas del dashboard dentro de esa ventana
 *     cuestan 0 llamadas a Calendly (antes eran 5 por profesional por carga).
 *     Además cachea el event_type de cada profesional por 6 h (casi nunca cambia),
 *     lo que elimina una llamada por profesional por refresco.
 *  3. NO QUEDARSE SIN DATOS — si Calendly o la cuota fallan, devuelve la última
 *     respuesta buena guardada, marcada como stale, en vez de un error pelado.
 *     Así la pestaña Disponibilidad nunca queda vacía.
 *
 * INSTALACIÓN (en cada uno de los dos proyectos de Apps Script):
 *  1. Pegar este archivo completo (reemplaza todo el código anterior).
 *  2. Configuración del proyecto (engranaje) → Propiedades del script → agregar:
 *       CALENDLY_TOKEN = <Personal Access Token de esa cuenta de Calendly>
 *       CALENDLY_ORG   = https://api.calendly.com/organizations/<uuid de esa cuenta>
 *       EQUIPO         = 2   (o 3, según el proyecto)
 *  3. Implementar → Administrar implementaciones → ✏️ → Versión: "Nueva versión"
 *     → Implementar. La URL /exec no cambia.
 *
 * Parámetros opcionales de doGet:
 *   ?nocache=1  fuerza consultar Calendly ignorando el caché (para depurar)
 *   ?debug=1    incluye la traza de llamadas en la respuesta
 */

const RANGE_DAYS = 15;          // ventana de disponibilidad a traer
const CACHE_MIN = 10;           // minutos que se sirve la respuesta cacheada
const CACHE_ET_HORAS = 6;       // horas que se recuerda el event_type de cada profesional
const MAX_SNAPSHOT_CHARS = 8000; // límite prudente para Propiedades del script (~9KB)

// Mapa de profesionales por equipo. Para agregar a alguien, sumar su URI de usuario.
const EQUIPOS = {
  '2': {
    etiqueta: 'Equipo 2',
    pros: {
      'https://api.calendly.com/users/72d3dbf1-ea9a-44b3-bc35-88838f2db017': { nombre: 'Martín Asís (Tino)', rol: 'Coach', slug: 'sesion-ontologica-martin' },
      'https://api.calendly.com/users/a3a01f0f-457c-44fd-b6e2-37361f23ad45': { nombre: 'Majo', rol: 'Nutricionista', slug: 'sesion-1' },
      'https://api.calendly.com/users/c4a624a6-487b-426f-b855-f70d70d6b8c7': { nombre: 'Doc Alí', rol: 'Médico', slug: 'sesion-endocrinologica-doctor-ali' },
      'https://api.calendly.com/users/7abeadfc-6406-4136-be7d-8bb3258b8753': { nombre: 'Luciana Labarthe', rol: 'Profesional', slug: 'luciana-labarthe-sesion-de-entrenamiento' },
    },
  },
  '3': {
    etiqueta: 'Equipo 3',
    pros: {
      'https://api.calendly.com/users/be8fcfb5-c19b-4f61-9e74-293cdc7ba818': { nombre: 'Dra. Liliana Bargi', rol: 'Médico', slug: '30min' },
      'https://api.calendly.com/users/8b26b698-f564-4100-ad65-16ec59a94d5a': { nombre: 'Tomas Bellatti', rol: 'Coach', slug: '30min' },
      'https://api.calendly.com/users/c25515f6-6765-4e14-8732-69abb23e47ed': { nombre: 'Juliana Selvaggi', rol: 'Nutricionista', slug: 'nueva-reunion' },
      'https://api.calendly.com/users/98416f98-90d8-4892-acce-4ba13b74d886': { nombre: 'Mario Danieli', rol: 'Coach', slug: '30min' },
      'https://api.calendly.com/users/99234c0c-75ea-4dbd-b270-24ec34f0ea0e': { nombre: 'Martín Alves', rol: 'Entrenador', slug: '60min' },
    },
  },
};

/* ══════════════ Entrada web ══════════════ */

function doGet(e) {
  const params = (e && e.parameter) || {};
  const salida = function (obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
  };

  let cfg;
  try {
    cfg = leerConfig_();
  } catch (err) {
    return salida({ ok: false, error: err.message });
  }

  const cache = CacheService.getScriptCache();
  const claveCache = 'disp_eq' + cfg.equipo;

  // 1) Respuesta fresca en caché → 0 llamadas a Calendly
  if (params.nocache !== '1') {
    const enCache = cache.get(claveCache);
    if (enCache) {
      const payload = JSON.parse(enCache);
      payload.cached = true;
      return salida(payload);
    }
  }

  // 2) Consultar Calendly
  try {
    const resultado = traerDisponibilidad_(cfg);
    if (!resultado.rows.length) throw new Error('Calendly no devolvió disponibilidad para ningún profesional');

    const payload = {
      ok: true,
      rows: resultado.rows,
      equipo: cfg.etiqueta,
      generado: new Date().toISOString(),
      stale: false,
      cached: false,
    };
    if (params.debug === '1') payload.debug = resultado.debug;

    cache.put(claveCache, JSON.stringify(payload), CACHE_MIN * 60);
    guardarSnapshot_(cfg.equipo, resultado.rows);
    return salida(payload);

  } catch (err) {
    // 3) Falló Calendly (cuota, token, caída) → servir la última respuesta buena
    const snap = leerSnapshot_(cfg.equipo);
    if (snap) {
      return salida({
        ok: true,
        rows: snap.rows,
        equipo: cfg.etiqueta,
        generado: snap.generado,
        stale: true,
        edad_minutos: Math.round((Date.now() - new Date(snap.generado).getTime()) / 60000),
        aviso: 'Datos de la última sincronización exitosa. Calendly no respondió: ' + err.message,
      });
    }
    return salida({ ok: false, error: err.message });
  }
}

/* ══════════════ Configuración ══════════════ */

function leerConfig_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('CALENDLY_TOKEN');
  const org = props.getProperty('CALENDLY_ORG');
  const equipo = String(props.getProperty('EQUIPO') || '').trim();

  if (!token) throw new Error('Falta la propiedad de script CALENDLY_TOKEN.');
  if (!org) throw new Error('Falta la propiedad de script CALENDLY_ORG.');
  if (!EQUIPOS[equipo]) throw new Error('La propiedad de script EQUIPO debe ser "2" o "3" (valor actual: "' + equipo + '").');

  return { token: token, org: org, equipo: equipo, etiqueta: EQUIPOS[equipo].etiqueta, pros: EQUIPOS[equipo].pros };
}

/* ══════════════ Calendly ══════════════ */

function traerDisponibilidad_(cfg) {
  const headers = { Authorization: 'Bearer ' + cfg.token };
  const cache = CacheService.getScriptCache();
  const rows = [];
  const debug = [];

  const ahora = new Date();
  const medianoche = new Date(ahora); medianoche.setHours(0, 0, 0, 0);
  const evStart = medianoche.toISOString();
  const evEnd = new Date(ahora.getTime() + RANGE_DAYS * 86400000).toISOString();

  // event_type_available_times acepta ventanas de máximo 7 días
  const tramos = [];
  let ini = new Date(ahora.getTime() + 60000); // +1 min: la ventana debe ser futura
  const fin = new Date(ahora.getTime() + RANGE_DAYS * 86400000);
  while (ini < fin) {
    let corte = new Date(ini.getTime() + 7 * 86400000);
    if (corte > fin) corte = fin;
    tramos.push({ start: ini.toISOString(), end: corte.toISOString() });
    ini = corte;
  }
  debug.push('tramos de disponibilidad: ' + tramos.length);

  let fallos = 0;
  Object.keys(cfg.pros).forEach(function (userUri) {
    const pro = cfg.pros[userUri];
    try {
      const etUri = resolverEventType_(userUri, pro, headers, cache, debug);
      if (!etUri) { fallos++; return; }

      // Agendadas: una sola llamada para toda la ventana
      const agPorDia = {};
      const evUrl = 'https://api.calendly.com/scheduled_events?user=' + encodeURIComponent(userUri) +
        '&organization=' + encodeURIComponent(cfg.org) +
        '&status=active&min_start_time=' + encodeURIComponent(evStart) +
        '&max_start_time=' + encodeURIComponent(evEnd) + '&count=100';
      const evResp = UrlFetchApp.fetch(evUrl, { headers: headers, muteHttpExceptions: true });
      if (evResp.getResponseCode() === 200) {
        (JSON.parse(evResp.getContentText()).collection || []).forEach(function (ev) {
          const d = fechaArg_(ev.start_time);
          agPorDia[d] = (agPorDia[d] || 0) + 1;
        });
      } else {
        debug.push(pro.nombre + ': scheduled_events HTTP ' + evResp.getResponseCode());
      }

      // Slots libres: una llamada por tramo
      const libPorDia = {};
      tramos.forEach(function (t) {
        const avUrl = 'https://api.calendly.com/event_type_available_times' +
          '?event_type=' + encodeURIComponent(etUri) +
          '&start_time=' + t.start + '&end_time=' + t.end;
        const avResp = UrlFetchApp.fetch(avUrl, { headers: headers, muteHttpExceptions: true });
        if (avResp.getResponseCode() !== 200) {
          debug.push(pro.nombre + ': available_times HTTP ' + avResp.getResponseCode() +
            ' [' + t.start.slice(0, 10) + '..' + t.end.slice(0, 10) + '] ' + avResp.getContentText().substring(0, 160));
          return;
        }
        (JSON.parse(avResp.getContentText()).collection || []).forEach(function (slot) {
          const d = fechaArg_(slot.start_time);
          libPorDia[d] = (libPorDia[d] || 0) + 1;
        });
      });

      const dias = {};
      Object.keys(agPorDia).forEach(function (d) { dias[d] = true; });
      Object.keys(libPorDia).forEach(function (d) { dias[d] = true; });
      Object.keys(dias).forEach(function (dia) {
        rows.push({
          PROFESIONAL: pro.nombre,
          ROL: pro.rol,
          EQUIPO: cfg.etiqueta,
          FECHA: dia,
          AGENDADAS: agPorDia[dia] || 0,
          LIBRES: libPorDia[dia] || 0,
        });
      });
    } catch (err) {
      fallos++;
      debug.push(pro.nombre + ': ' + err.message);
    }
  });

  debug.push('profesionales con problemas: ' + fallos + '/' + Object.keys(cfg.pros).length);
  return { rows: rows, debug: debug };
}

// El event_type de un profesional casi nunca cambia: se cachea 6 h para ahorrar
// una llamada por profesional en cada refresco del dashboard.
function resolverEventType_(userUri, pro, headers, cache, debug) {
  const clave = 'et_' + userUri.split('/').pop();
  const cacheado = cache.get(clave);
  if (cacheado) return cacheado;

  const etUrl = 'https://api.calendly.com/event_types?user=' + encodeURIComponent(userUri) + '&active=true&count=20';
  const resp = UrlFetchApp.fetch(etUrl, { headers: headers, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    debug.push(pro.nombre + ': event_types HTTP ' + resp.getResponseCode());
    return null;
  }

  const ets = JSON.parse(resp.getContentText()).collection || [];
  const elegido = ets.filter(function (et) {
    return et.scheduling_url && et.scheduling_url.indexOf('/' + pro.slug) !== -1;
  })[0] || ets[0];

  if (!elegido) {
    debug.push(pro.nombre + ': sin event types activos (slug buscado: ' + pro.slug + ')');
    return null;
  }
  if (elegido.scheduling_url.indexOf('/' + pro.slug) === -1) {
    debug.push(pro.nombre + ': no se encontró el slug "' + pro.slug + '", se usa "' + elegido.slug + '"');
  }

  cache.put(clave, elegido.uri, CACHE_ET_HORAS * 3600);
  return elegido.uri;
}

/* ══════════════ Última respuesta buena ══════════════ */
// Se guarda compacta (arrays en vez de objetos) porque Propiedades del script
// admite ~9KB por valor.

function guardarSnapshot_(equipo, rows) {
  try {
    const compacto = JSON.stringify({
      g: new Date().toISOString(),
      r: rows.map(function (r) { return [r.PROFESIONAL, r.ROL, r.EQUIPO, r.FECHA, r.AGENDADAS, r.LIBRES]; }),
    });
    if (compacto.length > MAX_SNAPSHOT_CHARS) return; // demasiado grande: se omite
    PropertiesService.getScriptProperties().setProperty('SNAPSHOT_EQ' + equipo, compacto);
  } catch (err) {
    // guardar el snapshot nunca debe romper la respuesta
  }
}

function leerSnapshot_(equipo) {
  try {
    const crudo = PropertiesService.getScriptProperties().getProperty('SNAPSHOT_EQ' + equipo);
    if (!crudo) return null;
    const datos = JSON.parse(crudo);
    return {
      generado: datos.g,
      rows: datos.r.map(function (a) {
        return { PROFESIONAL: a[0], ROL: a[1], EQUIPO: a[2], FECHA: a[3], AGENDADAS: a[4], LIBRES: a[5] };
      }),
    };
  } catch (err) {
    return null;
  }
}

/* ══════════════ Utilidades ══════════════ */

// Calendly devuelve UTC; el dashboard agrupa por día en horario argentino (UTC-3).
function fechaArg_(iso) {
  return new Date(new Date(iso).getTime() - 3 * 3600000).toISOString().slice(0, 10);
}

/* ══════════════ Diagnóstico ══════════════ */
// Ejecutar a mano desde el editor (botón ▶) para ver qué está pasando.

function probar() {
  const cfg = leerConfig_();
  const res = traerDisponibilidad_(cfg);
  Logger.log('Equipo: ' + cfg.etiqueta);
  Logger.log('Filas: ' + res.rows.length);
  Logger.log(res.debug.join('\n'));
}

// Borra el caché para forzar una consulta fresca en el próximo doGet.
function limpiarCache() {
  const cfg = leerConfig_();
  CacheService.getScriptCache().remove('disp_eq' + cfg.equipo);
  Logger.log('Caché borrado para ' + cfg.etiqueta);
}
