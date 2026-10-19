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

// Mapa de profesionales por equipo. Para agregar a alguien hay dos formas:
//  - Con su URI de usuario de Calendly (https://api.calendly.com/users/<uuid>), o
//  - Con '@' + el nombre de su perfil de Calendly (lo que va después de calendly.com/),
//    por ejemplo '@amparo-testoboost': el script busca el URI solo.
// Campos: slug = evento de SESIÓN; control = evento de CONTROL (opcional). Si no se
// indica "control", se busca solo un evento activo cuyo nombre o slug diga "control".
const EQUIPOS = {
  '2': {
    etiqueta: 'Equipo 2',
    pros: {
      'https://api.calendly.com/users/72d3dbf1-ea9a-44b3-bc35-88838f2db017': { nombre: 'Martín Asís (Tino)', rol: 'Coach', slug: 'sesion-ontologica-martin' },
      'https://api.calendly.com/users/a3a01f0f-457c-44fd-b6e2-37361f23ad45': { nombre: 'Majo', rol: 'Nutricionista', slug: 'sesion-1' },
      'https://api.calendly.com/users/c4a624a6-487b-426f-b855-f70d70d6b8c7': { nombre: 'Doc Alí', rol: 'Médico', slug: 'sesion-endocrinologica-doctor-ali' },
      'https://api.calendly.com/users/7abeadfc-6406-4136-be7d-8bb3258b8753': { nombre: 'Luciana Labarthe', rol: 'Profesional', slug: 'luciana-labarthe-sesion-de-entrenamiento' },
      '@santiagonavarro-testoboost': { nombre: 'Santiago Navarro', rol: 'Entrenador', slug: '60min', control: 'sesion-de-control-santiago-navarro' },
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
      'https://api.calendly.com/users/becb68d0-fe18-4fda-af7e-439c5837e1f8': { nombre: 'Juan Cruz Primus', rol: 'Entrenador', slug: '30min' },
      '@amparo-testoboost': { nombre: 'Amparo García', rol: 'Nutricionista', slug: 'sesion-nutricional-amparo-garcia', control: 'new-meeting' },
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

  const pros = resolverPerfiles_(cfg, headers, cache, debug);
  const contar = function (o) { return Object.keys(o || {}).length; };
  const unir = function (a, b) {
    const u = {};
    Object.keys(a || {}).forEach(function (k) { u[k] = true; });
    Object.keys(b || {}).forEach(function (k) { u[k] = true; });
    return Object.keys(u).length;
  };

  let fallos = 0;
  Object.keys(pros).forEach(function (userUri) {
    const pro = pros[userUri];
    try {
      const tipos = resolverEventTypes_(userUri, pro, headers, cache, debug);
      if (!tipos || !tipos.sesion) { fallos++; return; }
      const controlUri = tipos.control;

      // Agendadas: una sola llamada para toda la ventana, separadas por tipo de evento
      const agSesion = {}, agControl = {};
      const evUrl = 'https://api.calendly.com/scheduled_events?user=' + encodeURIComponent(userUri) +
        '&organization=' + encodeURIComponent(cfg.org) +
        '&status=active&min_start_time=' + encodeURIComponent(evStart) +
        '&max_start_time=' + encodeURIComponent(evEnd) + '&count=100';
      const evResp = UrlFetchApp.fetch(evUrl, { headers: headers, muteHttpExceptions: true });
      if (evResp.getResponseCode() === 200) {
        (JSON.parse(evResp.getContentText()).collection || []).forEach(function (ev) {
          const d = fechaArg_(ev.start_time);
          const destino = (controlUri && ev.event_type === controlUri) ? agControl : agSesion;
          destino[d] = (destino[d] || 0) + 1;
        });
      } else {
        debug.push(pro.nombre + ': scheduled_events HTTP ' + evResp.getResponseCode());
      }

      // Slots libres por tipo de evento (cada horario se guarda por su hora de inicio
      // para poder unir sesión + control sin contar dos veces el mismo horario)
      const libSesion = slotsPorDia_(tipos.sesion, tramos, headers, debug, pro.nombre);
      const libControl = controlUri ? slotsPorDia_(controlUri, tramos, headers, debug, pro.nombre) : {};

      const dias = {};
      [agSesion, agControl, libSesion, libControl].forEach(function (m) {
        Object.keys(m).forEach(function (d) { dias[d] = true; });
      });
      Object.keys(dias).forEach(function (dia) {
        const aS = agSesion[dia] || 0, aC = agControl[dia] || 0;
        rows.push({
          PROFESIONAL: pro.nombre,
          ROL: pro.rol,
          EQUIPO: cfg.etiqueta,
          FECHA: dia,
          AGENDADAS: aS + aC,
          LIBRES: unir(libSesion[dia], libControl[dia]),  // horarios distintos, sin repetir
          AGENDADAS_SESION: aS,
          AGENDADAS_CONTROL: aC,
          LIBRES_SESION: contar(libSesion[dia]),
          LIBRES_CONTROL: contar(libControl[dia]),
          TIENE_CONTROL: controlUri ? 1 : 0,
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

// Horarios libres de un tipo de evento, agrupados por día: { 'YYYY-MM-DD': { <hora de inicio>: true } }
function slotsPorDia_(etUri, tramos, headers, debug, nombre) {
  const porDia = {};
  tramos.forEach(function (t) {
    const avUrl = 'https://api.calendly.com/event_type_available_times' +
      '?event_type=' + encodeURIComponent(etUri) +
      '&start_time=' + t.start + '&end_time=' + t.end;
    const avResp = UrlFetchApp.fetch(avUrl, { headers: headers, muteHttpExceptions: true });
    if (avResp.getResponseCode() !== 200) {
      debug.push(nombre + ': available_times HTTP ' + avResp.getResponseCode() +
        ' [' + t.start.slice(0, 10) + '..' + t.end.slice(0, 10) + '] ' + avResp.getContentText().substring(0, 160));
      return;
    }
    (JSON.parse(avResp.getContentText()).collection || []).forEach(function (slot) {
      const d = fechaArg_(slot.start_time);
      if (!porDia[d]) porDia[d] = {};
      porDia[d][slot.start_time] = true;
    });
  });
  return porDia;
}

// Los profesionales cargados como '@perfil' se resuelven a su URI de usuario con la
// lista de miembros de la cuenta (una llamada, cacheada 6 h).
function resolverPerfiles_(cfg, headers, cache, debug) {
  const pros = {};
  const pendientes = [];
  Object.keys(cfg.pros).forEach(function (k) {
    if (k.charAt(0) === '@') pendientes.push(k); else pros[k] = cfg.pros[k];
  });
  if (!pendientes.length) return pros;

  const clave = 'perfiles_eq' + cfg.equipo;
  let porPerfil = null;
  const enCache = cache.get(clave);
  if (enCache) {
    porPerfil = JSON.parse(enCache);
  } else {
    const resp = UrlFetchApp.fetch(
      'https://api.calendly.com/organization_memberships?organization=' + encodeURIComponent(cfg.org) + '&count=100',
      { headers: headers, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) {
      debug.push('organization_memberships HTTP ' + resp.getResponseCode() + ': no se pudieron resolver los perfiles ' + pendientes.join(', '));
      return pros;
    }
    porPerfil = {};
    (JSON.parse(resp.getContentText()).collection || []).forEach(function (m) {
      porPerfil[m.user.slug] = m.user.uri;
    });
    // Solo se cachea si se encontraron todos, para que un alta reciente se vea enseguida
    const todos = pendientes.every(function (k) { return porPerfil[k.slice(1)]; });
    if (todos) cache.put(clave, JSON.stringify(porPerfil), CACHE_ET_HORAS * 3600);
  }

  pendientes.forEach(function (k) {
    const uri = porPerfil[k.slice(1)];
    if (uri) pros[uri] = cfg.pros[k];
    else debug.push(cfg.pros[k].nombre + ': no se encontró el perfil "' + k.slice(1) + '" en la cuenta de Calendly');
  });
  return pros;
}

// Los event_types de un profesional casi nunca cambian: se cachean 6 h para ahorrar
// una llamada por profesional en cada refresco del dashboard.
// Devuelve { sesion: <uri>, control: <uri|null> }.
function resolverEventTypes_(userUri, pro, headers, cache, debug) {
  const clave = 'et2_' + userUri.split('/').pop();
  const cacheado = cache.get(clave);
  if (cacheado) return JSON.parse(cacheado);

  const etUrl = 'https://api.calendly.com/event_types?user=' + encodeURIComponent(userUri) + '&active=true&count=20';
  const resp = UrlFetchApp.fetch(etUrl, { headers: headers, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    debug.push(pro.nombre + ': event_types HTTP ' + resp.getResponseCode());
    return null;
  }

  const ets = JSON.parse(resp.getContentText()).collection || [];
  const ultimo = function (et) { return String(et.scheduling_url || '').split('?')[0].split('/').pop(); };
  const porSlug = function (slug) {
    return ets.filter(function (et) { return ultimo(et) === slug; })[0];
  };

  // Sesión: slug exacto; si no, coincidencia parcial como antes
  const sesionElegida = porSlug(pro.slug) || ets.filter(function (et) {
    return et.scheduling_url && et.scheduling_url.indexOf('/' + pro.slug) !== -1;
  })[0];

  // Control: el slug indicado, o (sin indicar) un evento activo que diga "control"
  const control = pro.control
    ? porSlug(pro.control)
    : ets.filter(function (et) {
        return et !== sesionElegida && /control/i.test((et.slug || '') + ' ' + (et.name || ''));
      })[0];
  if (pro.control && !control) debug.push(pro.nombre + ': no se encontró el evento de control "' + pro.control + '"');
  if (!pro.control && !control) debug.push(pro.nombre + ': sin evento de control detectado (se cuenta solo sesión)');

  const sesion = sesionElegida || ets.filter(function (et) { return et !== control; })[0];
  if (!sesion) {
    debug.push(pro.nombre + ': sin event types activos (slug buscado: ' + pro.slug + ')');
    return null;
  }
  if (!sesionElegida) debug.push(pro.nombre + ': no se encontró el slug "' + pro.slug + '", se usa "' + sesion.slug + '"');

  const res = { sesion: sesion.uri, control: control ? control.uri : null };
  cache.put(clave, JSON.stringify(res), CACHE_ET_HORAS * 3600);
  return res;
}

/* ══════════════ Última respuesta buena ══════════════ */
// Se guarda compacta (tabla de profesionales + filas con índices) porque Propiedades
// del script admite ~9KB por valor.

function guardarSnapshot_(equipo, rows) {
  try {
    const profs = [];
    const idx = {};
    const filas = rows.map(function (r) {
      const k = r.PROFESIONAL + '|' + r.ROL + '|' + r.EQUIPO;
      if (idx[k] === undefined) { idx[k] = profs.length; profs.push([r.PROFESIONAL, r.ROL, r.EQUIPO]); }
      return [idx[k], r.FECHA, r.AGENDADAS, r.LIBRES, r.AGENDADAS_SESION || 0, r.AGENDADAS_CONTROL || 0,
        r.LIBRES_SESION || 0, r.LIBRES_CONTROL || 0, r.TIENE_CONTROL || 0];
    });
    const compacto = JSON.stringify({ g: new Date().toISOString(), p: profs, r: filas });
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
        if (!datos.p) { // formato anterior (sin desglose sesión/control)
          return { PROFESIONAL: a[0], ROL: a[1], EQUIPO: a[2], FECHA: a[3], AGENDADAS: a[4], LIBRES: a[5] };
        }
        const p = datos.p[a[0]];
        return {
          PROFESIONAL: p[0], ROL: p[1], EQUIPO: p[2], FECHA: a[1], AGENDADAS: a[2], LIBRES: a[3],
          AGENDADAS_SESION: a[4], AGENDADAS_CONTROL: a[5], LIBRES_SESION: a[6], LIBRES_CONTROL: a[7], TIENE_CONTROL: a[8],
        };
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

// Lista todos los miembros de la cuenta de Calendly con su URI, para sumar
// profesionales nuevos a EQUIPOS. Ejecutar a mano desde el editor (botón ▶).
function listarIntegrantes() {
  const cfg = leerConfig_();
  const resp = UrlFetchApp.fetch(
    'https://api.calendly.com/organization_memberships?organization=' + encodeURIComponent(cfg.org) + '&count=100',
    { headers: { Authorization: 'Bearer ' + cfg.token }, muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) {
    Logger.log('Error HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 200));
    return;
  }
  (JSON.parse(resp.getContentText()).collection || []).forEach(function (m) {
    const registrado = (cfg.pros[m.user.uri] || cfg.pros['@' + m.user.slug]) ? '  [YA ESTÁ]' : '  [NUEVO]';
    Logger.log(m.user.name + ' | ' + m.user.scheduling_url + ' | ' + m.user.uri + registrado);
  });
}
