import assert from "node:assert/strict";
import { test } from "node:test";
import { createHandler } from "../api/v1/analyze.mjs";

const env = { API_BASE_URL: "https://backend.example.invalid", API_KEY: "server-only-test-secret" };
const payload = { metadata: { asunto: "Reunión", remitente_email: "test@example.invalid" }, contenido: "Correo completo", security_features: null };
const result = { is_phishing: false, risk_score: 0.1, reason: "Legítimo", slm_explanation: "Resumen" };
function request(body = JSON.stringify(payload), headers = {}) {
  return new Request("https://addin.example.invalid/api/v1/analyze", {
    method: "POST", headers: { Origin: "https://addin.example.invalid", "Content-Type": "application/json", ...headers }, body,
  });
}

test("reenvía el correo al endpoint de AWS incluyendo la clave en el servidor", async () => {
  const handler = createHandler({ env, fetchBackend: async (url, init) => {
    assert.equal(String(url), "https://backend.example.invalid/api/v1/analyze");
    assert.equal(init.headers["X-API-Key"], env.API_KEY);
    assert.equal(init.redirect, "error");
    assert.deepEqual(JSON.parse(init.body), payload);
    return Response.json({ ...result, internal_debug: env.API_KEY });
  } });
  const response = await handler.fetch(request(undefined, { "X-API-Key": "attacker-supplied" }));
  assert.equal(response.status, 200);
  assert.equal(response.headers.get("cache-control"), "no-store");
  assert.deepEqual(await response.json(), result);
});

test("rechaza métodos no permitidos, solicitudes de otro origen o formato inválido", async () => {
  const handler = createHandler({ env, fetchBackend: () => { assert.fail("No se debe invocar al backend"); } });
  assert.equal((await handler.fetch(new Request("https://addin.example.invalid/api/v1/analyze"))).status, 405);
  assert.equal((await handler.fetch(request(undefined, { Origin: "https://other.example.invalid" }))).status, 403);
  assert.equal((await handler.fetch(request(undefined, { "Content-Type": "text/plain" }))).status, 415);
  assert.equal((await handler.fetch(request("{invalid"))).status, 400);
  assert.equal((await handler.fetch(request("{}"))).status, 400);
});

test("rechaza cuerpos de solicitud que superan el tamaño máximo", async () => {
  const handler = createHandler({ env, fetchBackend: () => { assert.fail("No se debe invocar al backend"); } });
  const body = JSON.stringify({ ...payload, contenido: "a".repeat(1_100_001) });
  assert.equal((await handler.fetch(request(body))).status, 413);
});

test("responde 503 si faltan credenciales o configuración del servicio", async () => {
  for (const configuration of [{ ...env, API_KEY: "" }, { ...env, API_BASE_URL: "http://backend.example.invalid" },
    { ...env, API_BASE_URL: "https://user:password@backend.example.invalid" }, { ...env, API_BASE_URL: "https://addin.example.invalid" }]) {
    const handler = createHandler({ env: configuration, fetchBackend: () => { assert.fail("No se debe invocar al backend"); } });
    assert.equal((await handler.fetch(request())).status, 503);
  }
});

test("no expone credenciales ni errores internos del backend", async () => {
  const handler = createHandler({ env, fetchBackend: async () => new Response(`secret=${env.API_KEY}`, { status: 401 }) });
  const response = await handler.fetch(request());
  assert.equal(response.status, 502);
  assert.ok(!(await response.text()).includes(env.API_KEY));
});

test("valida la respuesta del servicio y preserva códigos de límite", async () => {
  const invalid = createHandler({ env, fetchBackend: async () => Response.json({ is_phishing: false, risk_score: 100, reason: "invalid" }) });
  assert.equal((await invalid.fetch(request())).status, 502);
  const limited = createHandler({ env, fetchBackend: async () => new Response("internal", { status: 429 }) });
  assert.equal((await limited.fetch(request())).status, 429);
});

test("cancela peticiones lentas y devuelve código de timeout", async () => {
  const handler = createHandler({ env, timeoutMs: 5, fetchBackend: (_, { signal }) => new Promise((resolve, reject) => {
    signal.addEventListener("abort", () => reject(new Error("aborted")), { once: true });
  }) });
  assert.equal((await handler.fetch(request())).status, 504);
});
