/**
 * Proxy de IA para el Dashboard COO Testo Boost — Google Apps Script
 *
 * Llama a la API gratuita de Gemini manteniendo la clave del lado del servidor.
 * La clave NUNCA va en el dashboard (que es público en GitHub Pages).
 *
 * INSTALACIÓN:
 *  1. Crear clave gratuita en https://aistudio.google.com (botón "Get API key").
 *  2. En https://script.google.com → Nuevo proyecto → pegar este archivo completo.
 *  3. Configuración del proyecto (engranaje) → Propiedades del script →
 *     agregar propiedad: GEMINI_API_KEY = <tu clave>.
 *  4. Implementar → Nueva implementación → Aplicación web →
 *     "Ejecutar como: yo" y "Acceso: cualquier persona" → Implementar.
 *  5. Copiar la URL que termina en /exec y pegarla en el panel de IA del dashboard.
 */

// Se prueban en orden y se recuerda el primero que funcione (Google rota los
// modelos gratuitos disponibles; "gemini-flash-latest" es el alias que apunta
// siempre al flash vigente).
const GEMINI_MODELS = ['gemini-flash-latest', 'gemini-3-flash', 'gemini-2.5-flash'];
const LIMITE_DIARIO = 300;                 // tope de consultas por día (anti-abuso)
const MAX_TOKENS_RESPUESTA = 1024;

function doPost(e) {
  const out = function (obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
  };
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('GEMINI_API_KEY');
    if (!apiKey) return out({ ok: false, error: 'Falta GEMINI_API_KEY en Propiedades del script.' });

    // Límite diario simple para que nadie externo queme la cuota
    const hoy = new Date().toISOString().slice(0, 10);
    const claveUso = 'uso_' + hoy;
    const usadas = parseInt(props.getProperty(claveUso) || '0', 10);
    if (usadas >= LIMITE_DIARIO) return out({ ok: false, error: 'Límite diario de consultas alcanzado.' });

    const body = JSON.parse(e.postData.contents);
    const mensajes = (body.messages || []).slice(-12); // últimas 12 vueltas de conversación
    const system = String(body.system || '').slice(0, 30000);

    const contents = mensajes.map(function (m) {
      return {
        role: m.role === 'assistant' ? 'model' : 'user',
        parts: [{ text: String(m.text || '').slice(0, 4000) }],
      };
    });

    // Probar modelos en orden; el que funcione queda cacheado para las próximas
    const payload = JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: contents,
      generationConfig: { maxOutputTokens: MAX_TOKENS_RESPUESTA, temperature: 0.3 },
    });
    const cacheado = props.getProperty('MODELO_OK');
    const candidatos = cacheado
      ? [cacheado].concat(GEMINI_MODELS.filter(function (m) { return m !== cacheado; }))
      : GEMINI_MODELS.slice();

    let ultimoError = 'sin modelos disponibles';
    for (let i = 0; i < candidatos.length; i++) {
      const modelo = candidatos[i];
      const resp = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent?key=' + apiKey,
        { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: payload }
      );
      const data = JSON.parse(resp.getContentText());

      if (resp.getResponseCode() === 200) {
        const texto = data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts
          ? data.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('')
          : '';
        props.setProperty('MODELO_OK', modelo);
        props.setProperty(claveUso, String(usadas + 1));
        return out({ ok: true, text: texto || 'No obtuve respuesta del modelo.', model: modelo });
      }

      ultimoError = (data.error && data.error.message) || ('HTTP ' + resp.getResponseCode());
      // Si el error es de modelo no disponible/inexistente, probamos el siguiente; si es otro (clave inválida, cuota), cortamos
      const esErrorDeModelo = resp.getResponseCode() === 404 || /model|available|found/i.test(ultimoError);
      if (!esErrorDeModelo) break;
      if (modelo === cacheado) props.deleteProperty('MODELO_OK');
    }
    return out({ ok: false, error: ultimoError });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
