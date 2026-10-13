// Uses Node's built-in runner and the project's existing TypeScript compiler.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const ts = require('typescript');

function loadService(name, globals = {}) {
  const file = path.join(__dirname, '../src/services', name + '.ts');
  const source = ts.transpileModule(fs.readFileSync(file, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
  }).outputText;
  const exports = {};
  vm.runInNewContext(source, { exports, process: { env: {} }, ...globals }, { filename: file });
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
    process: { env: { API_BASE_URL: 'https://backend.example.invalid' } },
    fetch: async (url, init) => {
      requestedUrl = url;
      requestedPayload = JSON.parse(init.body);
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
  assert.equal(requestedUrl, 'https://backend.example.invalid/api/v1/analyze/hybrid');
  assert.equal(requestedPayload.metadata.remitente_email, 'sender@example.invalid');
  assert.equal(result.slm_parsed, null);
  assert.equal(result.content_signals[0].rule, 'credential_disclosure_request');
  assert.equal(result.is_phishing, true);
  assert.equal(result.hybrid_xai.technical_attributions[0].feature, 'secret_count');
  assert.equal(result.hybrid_xai.lime_word_attributions.length, 0);
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
