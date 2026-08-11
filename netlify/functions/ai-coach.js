// AI proxy for the coaching-log generator (Quick / Weekly / Coaching Insights).
// Proxies to the Anthropic API using the ANTHROPIC_API_KEY env var set in Netlify.
// CORS-enabled so it can be called from the GitHub Pages-hosted dashboard.
//
// This forwarded the client's body verbatim with no model allowlist and no token
// ceiling, so one call could ask for the priciest model at full output length. The
// body is clamped before it reaches Anthropic.
const { keyFor, clamp, logUsage } = require('./_spend.cjs');

const APP = 'dccoach';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: CORS_HEADERS, body: '' };
  }
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers: CORS_HEADERS, body: 'Method Not Allowed' };
  }

  const apiKey = keyFor(APP);
  if (!apiKey) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
      body: JSON.stringify({ error: 'ANTHROPIC_API_KEY not configured' }),
    };
  }

  try {
    const body = clamp(JSON.parse(event.body));
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    logUsage(APP, data);
    return {
      statusCode: res.status,
      headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
      body: JSON.stringify(data),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
      body: JSON.stringify({ error: err.message || 'Generation failed' }),
    };
  }
};
