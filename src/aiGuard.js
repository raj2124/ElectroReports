// A provider that is down, out of quota or not paid for must not slow every request or be hammered (Session 6I).
// The guard wraps each provider call: it logs the call (counts only, no content) to ai_usage, and after three
// failures in a row of the kind that will not fix themselves (quota, auth, timeout, upstream) it stops calling that
// provider for a few minutes and says why. Features fall back to their template text; nothing here blocks.

class ProviderError extends Error {
  constructor(code, message, { retryable = true } = {}) {
    super(message);
    this.name = 'ProviderError';
    this.code = code;       // QUOTA | AUTH | TIMEOUT | UPSTREAM | BAD_OUTPUT | NOT_CONFIGURED | COOLING
    this.retryable = retryable;
  }
}

const TRIPS = new Set(['QUOTA', 'AUTH', 'TIMEOUT', 'UPSTREAM']);
const REASON = {
  QUOTA: 'the provider quota or billing limit was reached',
  AUTH: 'the provider rejected the key',
  TIMEOUT: 'the provider did not answer in time',
  UPSTREAM: 'the provider returned an error',
  BAD_OUTPUT: 'the provider returned something unusable',
  NOT_CONFIGURED: 'the provider is not configured',
  COOLING: 'the provider is paused after repeated failures'
};

// Maps whatever an SDK threw to a ProviderError (never carries the message text of the SDK: it can hold a key).
function classify(error) {
  if (error instanceof ProviderError) return error;
  const status = Number((error && (error.status || (error.response && error.response.status))) || 0);
  const text = String((error && error.message) || '');
  if (status === 429 || /quota|rate.?limit|billing|exceeded/i.test(text)) return new ProviderError('QUOTA', REASON.QUOTA);
  if (status === 401 || status === 403 || /api key|unauthori[sz]ed|permission/i.test(text)) return new ProviderError('AUTH', REASON.AUTH);
  if (/timeout|timed out|ETIMEDOUT|ECONNABORTED/i.test(text)) return new ProviderError('TIMEOUT', REASON.TIMEOUT);
  return new ProviderError('UPSTREAM', REASON.UPSTREAM);
}

function createAiGuard({ store = null, now = () => Date.now(), threshold = 3, cooldownMs = 5 * 60 * 1000 } = {}) {
  const state = new Map(); // provider -> { failures, until, lastCode }
  const of = (provider) => {
    if (!state.has(provider)) state.set(provider, { failures: 0, until: 0, lastCode: '' });
    return state.get(provider);
  };

  // run('gemini', 'ocr', 'model-name', async () => result, { units })
  async function run(provider, kind, model, work, { units = 1 } = {}) {
    const entry = of(provider);
    if (entry.until > now()) {
      throw new ProviderError('COOLING', `${REASON.COOLING} (${REASON[entry.lastCode] || 'repeated failures'}); trying again after ${new Date(entry.until).toISOString().slice(11, 16)} UTC`, { retryable: true });
    }
    const started = now();
    try {
      const result = await work();
      entry.failures = 0;
      if (store) store.addAiUsage({ provider, model, kind, ok: true, code: '', ms: now() - started, units });
      return result;
    } catch (raw) {
      const error = classify(raw);
      if (TRIPS.has(error.code)) {
        entry.failures += 1;
        entry.lastCode = error.code;
        if (entry.failures >= threshold) entry.until = now() + cooldownMs;
      }
      if (store) store.addAiUsage({ provider, model, kind, ok: false, code: error.code, ms: now() - started, units });
      throw error;
    }
  }

  const status = (provider) => {
    const entry = of(provider);
    return entry.until > now() ? { state: 'cooling', until: new Date(entry.until).toISOString(), reason: REASON[entry.lastCode] || '' } : { state: 'ok', until: '', reason: '' };
  };
  const reset = (provider) => state.delete(provider);

  return { run, status, reset };
}

module.exports = { ProviderError, createAiGuard, classify, REASON };
