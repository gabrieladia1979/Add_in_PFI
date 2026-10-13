// Contract smoke test against a running flujo_v2 codex/hybrid-xai backend.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');

const baseUrl = process.env.ANALYSIS_BACKEND_URL || 'http://127.0.0.1:8001';
const source = fs.readFileSync(path.join(__dirname, '../src/services/analysisService.ts'), 'utf8');
const compiled = ts.transpileModule(source, {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
}).outputText;
const serviceExports = {};
vm.runInNewContext(compiled, {
  exports: serviceExports,
  fetch,
  process: { env: { API_BASE_URL: baseUrl, API_KEY: process.env.API_KEY } },
}, { filename: 'analysisService.ts' });

serviceExports.analyzeEmail(
  'Aviso de acceso',
  'Responda este correo con su contraseña y código SMS para evitar la suspensión de la cuenta.',
  'remitente@example.invalid',
  null,
).then((result) => {
  console.log(JSON.stringify({
    is_phishing: result.is_phishing,
    decision_source: result.decision_source,
    technical_status: result.hybrid_xai.technical_status,
    lime_status: result.hybrid_xai.lime_status,
    supports_final_verdict: result.hybrid_xai.supports_final_verdict,
    explanation_parsed: Boolean(result.slm_parsed),
  }, null, 2));
}).catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
