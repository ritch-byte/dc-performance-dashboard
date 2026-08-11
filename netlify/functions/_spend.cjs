/**
 * Shared spend guard for the AI proxies in this project.
 *
 * WHY THIS EXISTS
 * Every proxy used the same ANTHROPIC_API_KEY, so console.anthropic.com could only
 * ever group spend by model. That tells you Haiku vs Sonnet and nothing about which
 * app burned it, which is why the Aug 6 and Aug 10 surges could not be traced to a
 * source. Each proxy now reads its own key and falls back to the shared one, so this
 * is safe to deploy before the new keys exist and attribution starts the moment
 * each one is added.
 *
 * The leading underscore keeps Netlify from publishing this as its own endpoint.
 */

// Published rates in cents per million tokens, [input, output].
const RATES = {
  'claude-haiku-4-5-20251001': [100, 500],
  'claude-haiku-4-5': [100, 500],
  'claude-sonnet-4-6': [300, 1500],
  'claude-sonnet-5': [300, 1500],
  'claude-opus-5': [500, 2500],
}

/** What a wide-open proxy is allowed to spend on. */
const ALLOWED_MODELS = ['claude-haiku-4-5-20251001', 'claude-sonnet-4-6']
const MAX_TOKENS_CEILING = 4096

const envName = app => 'ANTHROPIC_KEY_' + String(app).toUpperCase()

/**
 * This app's key, falling back to the shared one.
 * The fallback is deliberate: it means adding a key per app is a rollout, not a
 * cutover, and a missing key degrades to today's behaviour rather than an outage.
 */
function keyFor(app) {
  return process.env[envName(app)] || process.env.ANTHROPIC_API_KEY || ''
}

/** True once this app has its own key, so the logs show whether the split is live. */
const usingOwnKey = app => Boolean(process.env[envName(app)])

/**
 * Clamp a client-supplied request body.
 *
 * These proxies answer `Access-Control-Allow-Origin: *` from a public static site,
 * so the request body is untrusted: anything on the internet can post to them. The
 * cap is what bounds the damage. It cannot stop someone calling the endpoint, but it
 * does stop one call costing dollars instead of fractions of a cent.
 */
function clamp(body) {
  const out = { ...body }

  if (!ALLOWED_MODELS.includes(out.model)) {
    console.warn(JSON.stringify({ clamped: 'model', got: out.model || null }))
    out.model = ALLOWED_MODELS[0]
  }

  const asked = Number(out.max_tokens) || MAX_TOKENS_CEILING
  if (asked > MAX_TOKENS_CEILING) {
    console.warn(JSON.stringify({ clamped: 'max_tokens', got: asked }))
  }
  out.max_tokens = Math.min(asked, MAX_TOKENS_CEILING)

  // Server-side tools bill separately, web search at $10 per 1,000 requests. No
  // caller here uses them, so an open proxy should not be a way to buy them.
  if (out.tools) {
    console.warn(JSON.stringify({ clamped: 'tools', got: out.tools.length }))
    delete out.tools
  }

  return out
}

/**
 * One structured line per call, so Netlify's function log is a per-app bill.
 * Costs are computed from the response's own usage numbers, not estimated.
 */
function logUsage(app, data) {
  const usage = (data && data.usage) || {}
  const inTok = usage.input_tokens || 0
  const outTok = usage.output_tokens || 0
  const model = (data && data.model) || 'unknown'
  const rate = RATES[model] || RATES['claude-haiku-4-5-20251001']
  const cents = (inTok * rate[0] + outTok * rate[1]) / 1e6

  console.log(JSON.stringify({
    spend: app,
    model,
    in: inTok,
    out: outTok,
    cents: Number(cents.toFixed(4)),
    own_key: usingOwnKey(app),
  }))
}

module.exports = { keyFor, usingOwnKey, clamp, logUsage, ALLOWED_MODELS, MAX_TOKENS_CEILING }