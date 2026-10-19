// Node.js function. Credentials exist only in the server runtime environment.
const MAX_REQUEST_BYTES = 1_100_000;
const UPSTREAM_TIMEOUT_MS = 25_000;

function json(body, status = 200, headers = {}) {
  return Response.json(body, { status, headers: { "Cache-Control": "no-store", ...headers } });
}

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

export function createHandler({ env = process.env, fetchBackend = fetch, timeoutMs = UPSTREAM_TIMEOUT_MS } = {}) {
  return {
    async fetch(request) {
      if (request.method !== "POST") return json({ detail: "Método no permitido" }, 405, { Allow: "POST" });
      if (request.headers.get("origin") !== new URL(request.url).origin) {
        return json({ detail: "Origen no permitido" }, 403);
      }
      if (!/^application\/json(?:\s*;|$)/i.test(request.headers.get("content-type") || "")) {
        return json({ detail: "Se requiere JSON" }, 415);
      }

      let backend;
      const key = env.API_KEY;
      try { backend = new URL((env.API_BASE_URL || "").trim()); } catch {
        return json({ detail: "Servicio de análisis no configurado" }, 503);
      }
      if (!key || !key.trim() || backend.protocol !== "https:" || backend.username || backend.password ||
          backend.search || backend.hash || backend.pathname !== "/" || backend.origin === new URL(request.url).origin) {
        return json({ detail: "Servicio de análisis no configurado" }, 503);
      }

      const contentLength = Number(request.headers.get("content-length"));
      if (Number.isFinite(contentLength) && contentLength > MAX_REQUEST_BYTES) {
        return json({ detail: "El correo es demasiado grande para analizarlo" }, 413);
      }
      const reader = request.body?.getReader();
      if (!reader) return json({ detail: "El correo no tiene contenido válido" }, 400);
      let raw = "";
      let size = 0;
      const decoder = new TextDecoder("utf-8", { fatal: true });
      let payload;
      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          size += value.byteLength;
          if (size > MAX_REQUEST_BYTES) {
            await reader.cancel();
            return json({ detail: "El correo es demasiado grande para analizarlo" }, 413);
          }
          raw += decoder.decode(value, { stream: true });
        }
        raw += decoder.decode();
        payload = JSON.parse(raw);
      } catch {
        return json({ detail: "El correo no tiene contenido válido" }, 400);
      } finally {
        reader.releaseLock();
      }
      if (!isObject(payload) || !isObject(payload.metadata) || typeof payload.contenido !== "string" ||
          (payload.security_features != null && !isObject(payload.security_features))) {
        return json({ detail: "El correo no tiene contenido válido" }, 400);
      }

      const controller = new AbortController();
      const abort = () => controller.abort();
      request.signal.addEventListener("abort", abort, { once: true });
      if (request.signal.aborted) controller.abort();
      const timer = setTimeout(abort, timeoutMs);
      try {
        const upstream = await fetchBackend(new URL("/api/v1/analyze", backend), {
          method: "POST",
          headers: { "Content-Type": "application/json", "X-API-Key": key },
          body: JSON.stringify({
            metadata: payload.metadata, contenido: payload.contenido,
            security_features: payload.security_features ?? null,
          }),
          redirect: "error",
          signal: controller.signal,
        });
        // Never expose upstream diagnostics, credentials or infrastructure details.
        if (!upstream.ok) {
          if (upstream.status === 429) return json({ detail: "Demasiadas solicitudes. Intentá nuevamente en unos momentos." }, 429);
          if (upstream.status === 422) return json({ detail: "El servicio no pudo procesar los datos del correo" }, 422);
          return json({ detail: "El servicio de análisis no está disponible. Intentá nuevamente." }, 502);
        }
        const data = await upstream.json();
        if (!isObject(data) || typeof data.is_phishing !== "boolean" ||
            typeof data.risk_score !== "number" || !Number.isFinite(data.risk_score) ||
            data.risk_score < 0 || data.risk_score > 1 || typeof data.reason !== "string") {
          return json({ detail: "El servicio devolvió una respuesta incompatible" }, 502);
        }
        return json({
          is_phishing: data.is_phishing, risk_score: data.risk_score, reason: data.reason,
          intent: data.intent, slm_explanation: data.slm_explanation,
          explanation_source: data.explanation_source, content_signals: data.content_signals,
        });
      } catch {
        return json({ detail: controller.signal.aborted
          ? "El análisis tardó demasiado. Intentá nuevamente."
          : "No se pudo conectar con el servicio de análisis. Intentá nuevamente." }, controller.signal.aborted ? 504 : 502);
      } finally {
        clearTimeout(timer);
        request.signal.removeEventListener("abort", abort);
      }
    },
  };
}

export default createHandler();
