const assert = require('node:assert/strict');
const { test } = require('node:test');
const { createHmac } = require('node:crypto');
const handlerPromise = import('../api/v1/analyze/token.mjs');

const route = 'https://add-in-pfi.vercel.app/api/v1/analyze/token';
const key = 'a'.repeat(32);
const body = JSON.stringify({ nonce: 'a'.repeat(32), digest: 'b'.repeat(64) });

function request(origin = 'https://add-in-pfi.vercel.app', value = body) {
  return new Request(route, {
    method: 'POST',
    headers: { Origin: origin, 'Content-Type': 'application/json' },
    body: value,
  });
}

async function withConfig(run) {
  const saved = Object.fromEntries(['API_BASE_URL', 'PHISHARG_TOKEN_KEY',
    'VERCEL_ENV', 'PHISHARG_RATE_LIMIT_ID'].map((name) => [name, process.env[name]]));
  process.env.API_BASE_URL = 'https://backend.example.invalid';
  process.env.PHISHARG_TOKEN_KEY = key;
  delete process.env.VERCEL_ENV;
  delete process.env.PHISHARG_RATE_LIMIT_ID;
  try { await run(); } finally {
    for (const [name, value] of Object.entries(saved)) {
      if (value === undefined) delete process.env[name]; else process.env[name] = value;
    }
  }
}

test('issues a short-lived, signed token bound to a digest, without receiving email text', async () => {
  await withConfig(async () => {
    const { default: handler } = await handlerPromise;
    const response = await handler.fetch(request());
    assert.equal(response.status, 200);
    assert.equal(response.headers.get('Cache-Control'), 'no-store');
    const issued = await response.json();
    assert.equal(issued.api_base_url, 'https://backend.example.invalid');
    const [payload, signature] = issued.token.split('.');
    assert.equal(signature, createHmac('sha256', key).update(payload).digest('base64url'));
    const claims = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    assert.equal(claims.nonce, 'a'.repeat(32));
    assert.equal(claims.digest, 'b'.repeat(64));
    assert.equal(claims.exp - claims.iat, 90);
    assert.equal(JSON.stringify(claims).includes('correo'), false);
  });
});

test('rejects foreign origins and malformed fingerprints', async () => {
  await withConfig(async () => {
    const { default: handler } = await handlerPromise;
    assert.equal((await handler.fetch(request('https://other.example.invalid'))).status, 403);
    assert.equal((await handler.fetch(request(undefined, '{"nonce":"x"}'))).status, 400);
  });
});

test('stops oversized token requests even without Content-Length', async () => {
  await withConfig(async () => {
    const { default: handler } = await handlerPromise;
    const oversized = request(undefined, 'x'.repeat(300));
    oversized.headers.delete('content-length');
    assert.equal((await handler.fetch(oversized)).status, 413);
  });
});

test('fails closed when the backend or distributed rate limit is not configured', async () => {
  await withConfig(async () => {
    const { createHandler } = await handlerPromise;
    process.env.API_BASE_URL = 'http://backend.example.invalid';
    assert.equal((await createHandler().fetch(request())).status, 503);
    process.env.API_BASE_URL = 'https://backend.example.invalid';
    process.env.VERCEL_ENV = 'production';
    assert.equal((await createHandler().fetch(request())).status, 503);
    process.env.PHISHARG_RATE_LIMIT_ID = 'phisharg-hybrid-token';
    assert.equal((await createHandler(async () => ({ rateLimited: true })).fetch(request())).status, 429);
    assert.equal((await createHandler(async () => { throw Error('WAF unavailable'); }).fetch(request())).status, 503);
  });
});
