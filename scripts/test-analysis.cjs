// Uses Node's built-in runner and the project's existing TypeScript compiler.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const { createHash, webcrypto } = require('node:crypto');
const ts = require('typescript');

function loadService(name, globals = {}) {
  const file = path.join(__dirname, '../src/services', name + '.ts');
  const source = ts.transpileModule(fs.readFileSync(file, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
  }).outputText;
  const exports = {};
  vm.runInNewContext(source, { exports, __LOCAL_DEV__: true, ...globals }, { filename: file });
  return exports;
}

test('uses the backend verdict even when score is above 50%', () => {
  const { buildAnalysisResult } = loadService('analysisResult');
  const result = buildAnalysisResult({ is_phishing: false, confidence_score: 0.7373, raw_label: 'LEGÍTIMO' });
  assert.equal(result.isPhishing, false);
  assert.equal(result.threatLevel, 'Bajo');
  assert.equal(result.phishScore, 74);
});

test('preserves a phishing verdict and content evidence', () => {
  const { buildAnalysisResult } = loadService('analysisResult');
  const signals = [{ rule: 'credential_disclosure_request', description: 'Pedido de credenciales' }];
  const result = buildAnalysisResult({ is_phishing: true, confidence_score: 0.95, raw_label: 'PHISHING', content_signals: signals });
  assert.equal(result.isPhishing, true);
  assert.equal(result.threatLevel, 'Peligro Critico');
  assert.equal(result.contentSignals, signals);
});

test('maps backend evidence without requiring SLM JSON and uses configured backend', async () => {
  let requestedUrl;
  let requestedPayload;
  const { analyzeEmail } = loadService('analysisService', {
    fetch: async (url, init) => {
      requestedUrl = url;
      requestedPayload = JSON.parse(init.body);
      assert.equal(init.headers['X-API-Key'], undefined);
      return { ok: true, json: async () => ({
        is_phishing: true, risk_score: 0.95, reason: 'Pedido de claves',
        slm_explanation: 'Fallback de texto',
        content_signals: [{ rule: 'credential_disclosure_request', description: 'Pedido de credenciales' }],
        hybrid_xai: {
          scope: 'raw_hybrid_model_score', supports_final_verdict: true,
          technical_status: 'available',
          technical_attributions: [{ feature: 'secret_count', value: 1, contribution_log_odds: 0.8 }],
          lime_status: 'unreliable', lime_word_attributions: [{ word: 'clave', weight: 0.2 }],
        },
      }) };
    },
  });
  const result = await analyzeEmail('Asunto', 'Cuerpo', 'sender@example.invalid', null);
  assert.equal(requestedUrl, '/api/v1/analyze/hybrid');
  assert.equal(requestedPayload.metadata.remitente_email, 'sender@example.invalid');
  assert.equal(result.slm_parsed, null);
  assert.equal(result.content_signals[0].rule, 'credential_disclosure_request');
  assert.equal(result.is_phishing, true);
  assert.equal(result.hybrid_xai.technical_attributions[0].feature, 'secret_count');
  assert.equal(result.hybrid_xai.lime_word_attributions.length, 0);
});

test('quick request uses the fast verdict route and keeps the response contract', async () => {
  let requestedUrl;
  const { analyzeEmail } = loadService('analysisService', {
    fetch: async (url) => {
      requestedUrl = url;
      return { ok: true, json: async () => ({
        is_phishing: true, risk_score: 0.95, reason: 'PHISHING DETECTADO',
        slm_explanation: JSON.stringify({ summary: 'PHISHING DETECTADO', reasons: [], recommended_actions: [] }),
        hybrid_xai: { scope: 'raw_hybrid_model_score', supports_final_verdict: true,
          technical_status: 'skipped', technical_attributions: [],
          lime_status: 'skipped', lime_word_attributions: [] },
      }) };
    },
  });
  const result = await analyzeEmail('Asunto', 'Cuerpo', 'sender@example.invalid', null, true);
  assert.equal(requestedUrl, '/api/v1/analyze/hybrid?quick=true');
  assert.equal(result.slm_parsed.summary, 'PHISHING DETECTADO');
  assert.equal(result.hybrid_xai.technical_status, 'skipped');
});

test('rejects a backend that lacks the Hybrid XAI contract', async () => {
  const { analyzeEmail } = loadService('analysisService', {
    fetch: async () => ({ ok: true, json: async () => ({
      is_phishing: false, risk_score: 0.1, reason: 'LEGÍTIMO',
    }) }),
  });
  await assert.rejects(analyzeEmail('Asunto', 'Cuerpo', 'sender@example.invalid', null),
    /contrato XAI de Hybrid/);
});

test('reports a missing Hybrid route instead of silently using the old classifier', async () => {
  const { analyzeEmail } = loadService('analysisService', {
    fetch: async () => ({ ok: false, status: 404 }),
  });
  await assert.rejects(analyzeEmail('Asunto', 'Cuerpo', 'sender@example.invalid', null),
    /no tiene la ruta Hybrid/);
});

test('production sends only a salted digest to Vercel and the email directly to AWS', async () => {
  const calls = [];
  const { analyzeEmail } = loadService('analysisService', {
    __LOCAL_DEV__: false,
    crypto: webcrypto,
    TextEncoder,
    URL,
    fetch: async (url, init) => {
      calls.push({ url, init });
      if (calls.length === 1) {
        assert.equal(url, '/api/v1/analyze/token');
        assert.equal(init.body.includes('mensaje privado'), false);
        return { ok: true, status: 200, json: async () => ({
          token: 'signed.token', api_base_url: 'https://backend.example.invalid',
        }) };
      }
      assert.equal(url, 'https://backend.example.invalid/api/v1/analyze/hybrid');
      assert.equal(init.headers.Authorization, 'Bearer signed.token');
      assert.equal(init.headers['X-API-Key'], undefined);
      assert.equal(JSON.parse(init.body).contenido, 'mensaje privado');
      const { nonce, digest } = JSON.parse(calls[0].init.body);
      assert.equal(digest, createHash('sha256').update(`${nonce}.${init.body}`).digest('hex'));
      return { ok: true, json: async () => ({
        is_phishing: false, risk_score: 0.1, reason: 'LEGÍTIMO',
        hybrid_xai: { scope: 'raw_hybrid_model_score', supports_final_verdict: true,
          technical_status: 'skipped', technical_attributions: [],
          lime_status: 'skipped', lime_word_attributions: [] },
      }) };
    },
  });
  await analyzeEmail('Asunto', 'mensaje privado', 'sender@example.invalid', null);
  assert.equal(calls.length, 2);
});

test('two-phase production analysis requests a fresh token for each identical email body', async () => {
  const tokenRequests = [];
  const backendBodies = [];
  const { analyzeEmail } = loadService('analysisService', {
    __LOCAL_DEV__: false,
    crypto: webcrypto,
    TextEncoder,
    URL,
    fetch: async (url, init) => {
      if (url === '/api/v1/analyze/token') {
        tokenRequests.push(JSON.parse(init.body));
        return { ok: true, json: async () => ({
          token: `token-${tokenRequests.length}`, api_base_url: 'https://backend.example.invalid',
        }) };
      }
      backendBodies.push(init.body);
      assert.equal(init.headers.Authorization, `Bearer token-${backendBodies.length}`);
      return { ok: true, json: async () => ({
        is_phishing: true, risk_score: 0.95, reason: 'PHISHING DETECTADO',
        hybrid_xai: { scope: 'raw_hybrid_model_score', supports_final_verdict: true,
          technical_status: 'skipped', technical_attributions: [],
          lime_status: 'skipped', lime_word_attributions: [] },
      }) };
    },
  });
  await analyzeEmail('Asunto', 'mensaje privado', 'sender@example.invalid', null, true);
  await analyzeEmail('Asunto', 'mensaje privado', 'sender@example.invalid', null);
  assert.equal(tokenRequests.length, 2);
  assert.notEqual(tokenRequests[0].nonce, tokenRequests[1].nonce);
  assert.equal(backendBodies[0], backendBodies[1]);
  for (let i = 0; i < 2; i += 1) {
    assert.equal(tokenRequests[i].digest,
      createHash('sha256').update(`${tokenRequests[i].nonce}.${backendBodies[i]}`).digest('hex'));
  }
});

test('cancelling an analysis also cancels token authorization before email content is sent', async () => {
  const controller = new AbortController();
  let backendCalls = 0;
  const { analyzeEmail } = loadService('analysisService', {
    __LOCAL_DEV__: false,
    crypto: webcrypto,
    TextEncoder,
    fetch: async (url, init) => {
      if (url !== '/api/v1/analyze/token') backendCalls += 1;
      assert.equal(init.signal, controller.signal);
      controller.abort();
      throw Object.assign(new Error('aborted'), { name: 'AbortError' });
    },
  });
  await assert.rejects(
    analyzeEmail('Asunto', 'mensaje privado', 'sender@example.invalid', null, false, controller.signal),
    { name: 'AbortError' },
  );
  assert.equal(backendCalls, 0);
});

test('the backend request receives the same cancellation signal as authorization', async () => {
  const controller = new AbortController();
  const seen = [];
  const { analyzeEmail } = loadService('analysisService', {
    __LOCAL_DEV__: false,
    crypto: webcrypto,
    TextEncoder,
    URL,
    fetch: async (url, init) => {
      seen.push({ url, signal: init.signal });
      if (seen.length === 1) {
        return { ok: true, json: async () => ({
          token: 'signed.token', api_base_url: 'https://backend.example.invalid',
        }) };
      }
      controller.abort();
      throw Object.assign(new Error('aborted'), { name: 'AbortError' });
    },
  });
  await assert.rejects(
    analyzeEmail('Asunto', 'mensaje privado', 'sender@example.invalid', null, false, controller.signal),
    { name: 'AbortError' },
  );
  assert.equal(seen.length, 2);
  assert.equal(seen[0].signal, controller.signal);
  assert.equal(seen[1].signal, controller.signal);
});

test('rejects oversized email before asking Vercel for a token', async () => {
  let calls = 0;
  const { analyzeEmail } = loadService('analysisService', {
    __LOCAL_DEV__: false,
    TextEncoder,
    fetch: async () => { calls += 1; throw Error('unexpected request'); },
  });
  await assert.rejects(
    analyzeEmail('Asunto', 'x'.repeat(1_100_001), 'sender@example.invalid', null),
    /demasiado grande/,
  );
  assert.equal(calls, 0);
});
