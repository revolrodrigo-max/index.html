/**
 * Proxy de Agendas Calendly — Dashboard COO Testo Boost — Google Apps Script
 *
 * Qué hace:
 *  1. syncAgendas() — consulta la API de Calendly de las dos cuentas, baja todos
 *     los agendamientos (paciente, profesional/host, evento, fecha, estado,
 *     cancelaciones, no-shows) y los guarda en una planilla de Google Sheets
 *     que funciona como REGISTRO HISTÓRICO permanente (nunca borra filas viejas).
 *  2. doGet() — expone ese registro como JSON para la pestaña "Agendas" del dashboard.
 *  3. instalarTrigger() — deja el sync corriendo automáticamente cada hora.
 *
 * INSTALACIÓN (una sola vez):
 *  1. En CADA cuenta de Calendly: Integrations & apps → API & webhooks →
 *     Personal Access Token → generar y copiar el token.
 *  2. https://script.google.com → Nuevo proyecto → pegar este archivo completo.
 *  3. Configuración del proyecto (engranaje) → Propiedades del script → agregar:
 *       CALENDLY_TOKEN_EQ2 = <token de la cuenta del Equipo 2>
 *       CALENDLY_TOKEN_EQ3 = <token de la cuenta del Equipo 3>
 *  4. Ejecutar una vez syncAgendas (botón ▶) y autorizar permisos.
 *     La planilla "Agendas Calendly — Testo Boost" se crea sola en tu Drive.
 *  5. Ejecutar una vez instalarTrigger → queda sincronizando cada hora.
 *  6. Implementar → Nueva implementación → Aplicación web →
 *     "Ejecutar como: yo" y "Acceso: cualquier persona" → Implementar.
 *  7. Copiar la URL que termina en /exec y pegarla en la constante
 *     CALENDLY_AGENDAS_URL del index.html del dashboard.
 */

// Para agregar otra cuenta: sumar una entrada acá y su token en Propiedades del script.
// "pros" mapea el UUID de usuario de Calendly → nombre/rol que usa el dashboard
// (mismo mapeo que los proxies de disponibilidad). Es necesario porque algunos
// eventos figuran hosteados por la cuenta madre ("Testo Boost Agendas" /
// "testo-equipo3") y no por el profesional real.
const CUENTAS = [
  {
    etiqueta: 'Equipo 2',
    propToken: 'CALENDLY_TOKEN_EQ2',
    pros: {
      '72d3dbf1-ea9a-44b3-bc35-88838f2db017': { nombre: 'Martín Asís (Tino)', rol: 'Coach' },
      'a3a01f0f-457c-44fd-b6e2-37361f23ad45': { nombre: 'Majo', rol: 'Nutricionista' },
      'c4a624a6-487b-426f-b855-f70d70d6b8c7': { nombre: 'Doc Alí', rol: 'Médico' },
      '7abeadfc-6406-4136-be7d-8bb3258b8753': { nombre: 'Luciana Labarthe', rol: 'Profesional' },
    },
  },
  {
    etiqueta: 'Equipo 3',
    propToken: 'CALENDLY_TOKEN_EQ3',
    pros: {
      'be8fcfb5-c19b-4f61-9e74-293cdc7ba818': { nombre: 'Dra. Liliana Bargi', rol: 'Médico' },
      '8b26b698-f564-4100-ad65-16ec59a94d5a': { nombre: 'Tomas Bellatti', rol: 'Coach' },
      'c25515f6-6765-4e14-8732-69abb23e47ed': { nombre: 'Juliana Selvaggi', rol: 'Nutricionista' },
      '98416f98-90d8-4892-acce-4ba13b74d886': { nombre: 'Mario Danieli', rol: 'Coach' },
      '99234c0c-75ea-4dbd-b270-24ec34f0ea0e': { nombre: 'Martín Alves', rol: 'Entrenador' },
    },
  },
];

const DIAS_ATRAS = 60;     // ventana de sincronización hacia atrás
const DIAS_ADELANTE = 60;  // y hacia adelante (lo anterior ya sincronizado se conserva)
const TZ = 'America/Argentina/Buenos_Aires';
const NOMBRE_PLANILLA = 'Agendas Calendly — Testo Boost';
const NOMBRE_HOJA = 'AGENDAS';
const HEADERS = [
  'CUENTA', 'PROFESIONAL', 'ROL', 'EVENTO', 'PACIENTE', 'NOMBRE', 'APELLIDO', 'EMAIL',
  'FECHA', 'HORA', 'ESTADO', 'NO SHOW', 'CREADO', 'MOTIVO CANCELACIÓN',
  'EVENT_URI', 'INVITEE_URI', 'ULTIMO_SYNC',
];

/* ══════════════ API Calendly ══════════════ */

function apiGet_(url, token) {
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const data = JSON.parse(resp.getContentText() || '{}');
  if (code !== 200) {
    throw new Error('Calendly HTTP ' + code + ': ' + ((data.message || data.title || '') + ' — ' + url));
  }
  return data;
}

// Lista eventos agendados con paginación. Intenta a nivel organización (trae los
// eventos de todos los miembros); si el token no tiene permiso, cae a nivel usuario.
function listarEventos_(token) {
  const me = apiGet_('https://api.calendly.com/users/me', token).resource;
  const min = new Date(Date.now() - DIAS_ATRAS * 86400000).toISOString();
  const max = new Date(Date.now() + DIAS_ADELANTE * 86400000).toISOString();
  const base = 'https://api.calendly.com/scheduled_events?count=100&sort=start_time:asc'
    + '&min_start_time=' + encodeURIComponent(min)
    + '&max_start_time=' + encodeURIComponent(max);

  const alcances = [
    base + '&organization=' + encodeURIComponent(me.current_organization),
    base + '&user=' + encodeURIComponent(me.uri),
  ];
  for (let a = 0; a < alcances.length; a++) {
    try {
      const eventos = [];
      let url = alcances[a];
      for (let pag = 0; pag < 30 && url; pag++) {
        const data = apiGet_(url, token);
        eventos.push.apply(eventos, data.collection || []);
        url = data.pagination && data.pagination.next_page ? data.pagination.next_page : null;
      }
      return eventos;
    } catch (e) {
      if (a === alcances.length - 1) throw e;
      // Sin permiso de organización → probar alcance usuario
    }
  }
  return [];
}

// Trae los invitados de todos los eventos en paralelo (lotes de 40).
function listarInvitados_(eventos, token) {
  const porEvento = {}; // event_uri → invitees[]
  for (let i = 0; i < eventos.length; i += 40) {
    const lote = eventos.slice(i, i + 40);
    const respuestas = UrlFetchApp.fetchAll(lote.map(function (ev) {
      return {
        url: ev.uri + '/invitees?count=100',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      };
    }));
    respuestas.forEach(function (resp, j) {
      if (resp.getResponseCode() !== 200) return;
      const data = JSON.parse(resp.getContentText() || '{}');
      porEvento[lote[j].uri] = data.collection || [];
    });
  }
  return porEvento;
}

// El profesional real: primer host del evento que figure en el mapeo "pros".
// Si ninguno está mapeado (profesional nuevo), usa el nombre que reporta Calendly
// para no perder el dato — conviene agregarlo al mapeo cuando aparezca.
function resolverProfesional_(ev, pros) {
  const ms = ev.event_memberships || [];
  for (let i = 0; i < ms.length; i++) {
    const uuid = String(ms[i].user || '').split('/').pop();
    if (pros[uuid]) return pros[uuid];
  }
  const fallback = ms.length ? (ms[0].user_name || ms[0].user_email || '') : '';
  return { nombre: fallback, rol: '' };
}

/* ══════════════ SYNC (correr con trigger horario) ══════════════ */

function syncAgendas() {
  const props = PropertiesService.getScriptProperties();
  const ahora = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  const nuevas = {}; // INVITEE_URI → fila

  CUENTAS.forEach(function (cuenta) {
    const token = props.getProperty(cuenta.propToken);
    if (!token) {
      console.warn('Falta ' + cuenta.propToken + ' en Propiedades del script — cuenta salteada.');
      return;
    }
    const eventos = listarEventos_(token);
    const invitados = listarInvitados_(eventos, token);

    eventos.forEach(function (ev) {
      const prof = resolverProfesional_(ev, cuenta.pros || {});
      const inicio = new Date(ev.start_time);
      (invitados[ev.uri] || []).forEach(function (inv) {
        const cancel = inv.cancellation || null;
        nuevas[inv.uri] = [
          cuenta.etiqueta,
          prof.nombre,
          prof.rol,
          ev.name || '',
          inv.name || '',
          inv.first_name || '',
          inv.last_name || '',
          inv.email || '',
          Utilities.formatDate(inicio, TZ, 'yyyy-MM-dd'),
          Utilities.formatDate(inicio, TZ, 'HH:mm'),
          inv.status === 'canceled' ? 'Cancelada' : 'Activa',
          inv.no_show ? 'SI' : '',
          inv.created_at ? Utilities.formatDate(new Date(inv.created_at), TZ, 'yyyy-MM-dd HH:mm') : '',
          cancel ? ((cancel.reason || 'Sin motivo') + ' (' + (cancel.canceled_by || '') + ')') : '',
          ev.uri,
          inv.uri,
          ahora,
        ];
      });
    });
  });

  // Merge con lo ya registrado: se actualizan las filas re-sincronizadas y se
  // conserva todo lo histórico (fuera de la ventana) → registro permanente.
  const hoja = obtenerHoja_();
  const existentes = hoja.getLastRow() > 1
    ? hoja.getRange(2, 1, hoja.getLastRow() - 1, HEADERS.length).getValues()
    : [];
  const idxInvitee = HEADERS.indexOf('INVITEE_URI');
  const mapa = {};
  existentes.forEach(function (fila) { if (fila[idxInvitee]) mapa[fila[idxInvitee]] = fila; });
  Object.keys(nuevas).forEach(function (k) { mapa[k] = nuevas[k]; });

  const idxFecha = HEADERS.indexOf('FECHA');
  const idxHora = HEADERS.indexOf('HORA');
  const filas = Object.keys(mapa).map(function (k) { return mapa[k]; })
    .sort(function (a, b) {
      return String(b[idxFecha] + ' ' + b[idxHora]).localeCompare(String(a[idxFecha] + ' ' + a[idxHora]));
    });

  hoja.clearContents();
  hoja.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (filas.length) hoja.getRange(2, 1, filas.length, HEADERS.length).setValues(filas);
  props.setProperty('ULTIMO_SYNC', ahora);
  console.log('Sync OK — ' + Object.keys(nuevas).length + ' agendas en ventana, ' + filas.length + ' totales en registro.');
}

function obtenerHoja_() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const id = props.getProperty('SHEET_ID');
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create(NOMBRE_PLANILLA);
    props.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].setName(NOMBRE_HOJA);
  }
  return ss.getSheetByName(NOMBRE_HOJA) || ss.insertSheet(NOMBRE_HOJA);
}

function instalarTrigger() {
  // Evitar duplicados si se corre más de una vez
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAgendas') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAgendas').timeBased().everyHours(1).create();
  console.log('Trigger horario instalado.');
}

/* ══════════════ WEB APP (JSON para el dashboard) ══════════════ */

function doGet() {
  const out = function (obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
  };
  try {
    const hoja = obtenerHoja_();
    const props = PropertiesService.getScriptProperties();
    if (hoja.getLastRow() < 2) return out({ ok: true, rows: [], updated: props.getProperty('ULTIMO_SYNC') || '' });
    const valores = hoja.getRange(1, 1, hoja.getLastRow(), HEADERS.length).getDisplayValues();
    const headers = valores[0];
    const rows = valores.slice(1).map(function (v) {
      const obj = {};
      headers.forEach(function (h, i) { if (h) obj[h] = v[i]; });
      return obj;
    });
    return out({ ok: true, rows: rows, updated: props.getProperty('ULTIMO_SYNC') || '' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
