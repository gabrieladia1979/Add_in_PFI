import { createHmac } from "node:crypto";
import { checkRateLimit } from "@vercel/firewall";

const MAX_REQUEST_BYTES = 256;
const TOKEN_SECONDS = 90;

function jsonResponse(body, status = 200) {
  return Response.json(body, { status, headers: { "Cache-Control": "no-store" } });
}

export function createHandler(rateLimit = checkRateLimit) {
  return {
    async fetch(request) {
      if (request.method !== "POST") return jsonResponse({ detail: "Método no permitido" }, 405);
      if (request.headers.get("origin") !== new URL(request.url).origin) {
        return jsonResponse({ detail: "Origen no permitido" }, 403);
      }
      if (!/^application\/json(?:\s*;|$)/i.test(request.headers.get("content-type") || "")) {
        return jsonResponse({ detail: "Se requiere JSON" }, 415);
      }
      const size = Number(request.headers.get("content-length"));
      if (Number.isFinite(size) && size > MAX_REQUEST_BYTES) {
        return jsonResponse({ detail: "Solicitud demasiado grande" }, 413);
      }

      const key = process.env.PHISHARG_TOKEN_KEY;
      let backend;
      try { backend = new URL(process.env.API_BASE_URL); } catch {
        return jsonResponse({ detail: "Servicio no configurado" }, 503);
      }
      if (!key || key.length < 32 || backend.protocol !== "https:" ||
          backend.username || backend.password) {
        return jsonResponse({ detail: "Servicio no configurado" }, 503);
      }

      if (process.env.VERCEL_ENV) {
        const rule = process.env.PHISHARG_RATE_LIMIT_ID;
        if (!rule) return jsonResponse({ detail: "Control de tasa no configurado" }, 503);
        try {
          const { rateLimited } = await rateLimit(rule, { request });
          if (rateLimited) return jsonResponse({ detail: "Demasiadas solicitudes" }, 429);
        } catch {
          return jsonResponse({ detail: "Control de tasa no disponible" }, 503);
        }
      }

      let raw = "";
      let bytes = 0;
      const reader = request.body?.getReader();
      if (!reader) return jsonResponse({ detail: "JSON inválido" }, 400);
      const decoder = new TextDecoder("utf-8", { fatal: true });
      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          bytes += value.byteLength;
          if (bytes > MAX_REQUEST_BYTES) {
            await reader.cancel();
            return jsonResponse({ detail: "Solicitud demasiado grande" }, 413);
          }
          raw += decoder.decode(value, { stream: true });
        }
        raw += decoder.decode();
      } catch {
        return jsonResponse({ detail: "JSON inválido" }, 400);
      }
      let input;
      try { input = JSON.parse(raw); } catch {
        return jsonResponse({ detail: "JSON inválido" }, 400);
      }
      if (!input || !/^[a-f0-9]{32}$/.test(input.nonce) ||
          !/^[a-f0-9]{64}$/.test(input.digest)) {
        return jsonResponse({ detail: "Huella inválida" }, 400);
      }

      const now = Math.floor(Date.now() / 1000);
      const payload = Buffer.from(JSON.stringify({
        aud: "phisharg-hybrid", iat: now, exp: now + TOKEN_SECONDS,
        nonce: input.nonce, digest: input.digest,
      })).toString("base64url");
      const signature = createHmac("sha256", key).update(payload).digest("base64url");
      return jsonResponse({ token: `${payload}.${signature}`, api_base_url: backend.origin });
    },
  };
}

export default createHandler();
