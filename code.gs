/***** ========== CONFIG & CONSTANTS (SANITIZED) ========== *****/
const CFG = {
  // === Google Sheet (replace with your own)
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
  SHEET_NAME: 'Consolidated - Balance Sheet', // e.g., your tab name containing orders
  BASE_SHEET_NAME: 'Balance Sheet',           // optional/unused here; keep or remove

  // === Time window for Gmail scan
  LOOKBACK_DAYS_EMAILS: 10,

  // === Alert emails (no tracking after 72h)
  ALERT_RECIPIENTS: ['ops@example.com', 'support@example.com'],

  // === Shopify stores (tokens are stored via OAuth in ScriptProperties; DO NOT hardcode)
  // Replace with your stores. Keys must match STORE_NAMES keys.
  BRANDS: {
    'STORE_A': { storeUrl: 'your-store-a.myshopify.com', publicDomain: 'store-a.example.com' },
    'STORE_B': { storeUrl: 'your-store-b.myshopify.com', publicDomain: 'store-b.example.com' },
    'STORE_C': { storeUrl: 'your-store-c.myshopify.com', publicDomain: 'store-c.example.com' }
    // Add more stores as needed…
  },
  // Full, human-readable store names (must match what's in your Google Sheet column "From store?")
  STORE_NAMES: {
    'STORE_A': 'Example Store A',
    'STORE_B': 'Example Store B',
    'STORE_C': 'Example Store C'
  },

  // === AliExpress (REST + TOP) — placeholders; DO NOT publish real keys/secrets
  AE: {
    HOST_REST: 'https://api-sg.aliexpress.com/rest', // System & Business (GOP) host
    HOST_TOP_SYNC: 'https://api-sg.aliexpress.com/sync',
    APP_KEY: 'YOUR_AE_APP_KEY',
    APP_SECRET: 'YOUR_AE_APP_SECRET',
    AUTH_CODE: '', // one-shot, then leave empty (use AliAuth_ExchangeCodeOnce)
    // Tokens — prefer ScriptProperties (see AliAuth_ExchangeCodeOnce / AliAuth_Refresh)
    ACCESS_TOKEN: '',
    REFRESH_TOKEN: '',
    EXPIRES_AT_MS: 0,
    ALI_RATE: { MAX_RETRIES: 10, JITTER_MS: 150 }
  },

  // === Misc / Safety
  DRY_RUN: false,        // true: no Sheet writes, no Shopify push, no Gmail deletions
  CELL_SEP: ' | ',       // multi-value separator for "AE ID(s)" / "Tracking ID(s)"
  DEBUG_VERBOSE_MATCH: false // verbose address-match logs
};

/***** ========== GOOGLE SHEETS SAFE-WRITE (PROTECT FORMULAS) ========== *****/
// If a cell or its column contains any formula, writing is blocked
// to avoid breaking ARRAYFORMULA/spilled ranges.
const SHEET_SAFETY = {
  PROTECT_FORMULA_COLUMNS: true, // scan entire column for formulas
  ABORT_ON_PROTECTED_WRITE: true, // true: throw Error; false: log + skip
  LOG_PREFIX: '[SAFEWRITE]'
};

function _columnHasAnyFormula_(sh, col) {
  const lr = Math.max(sh.getLastRow(), 1);
  const formulas = sh.getRange(1, col, lr, 1).getFormulas();
  for (let i = 0; i < formulas.length; i++) {
    if (formulas[i][0]) return true;
  }
  return false;
}

function _isProtectedCell_(sh, row, col) {
  const rng = sh.getRange(row, col, 1, 1);
  const directFormula = !!rng.getFormula(); // top-left of arrays, or simple formula
  if (directFormula) return true;
  if (SHEET_SAFETY.PROTECT_FORMULA_COLUMNS) {
    // Conservative: if the column has any formula, assume it's spill-prone
    if (_columnHasAnyFormula_(sh, col)) return true;
  }
  return false;
}

function _safeSetValueText_(sh, row, col, text) {
  if (_isProtectedCell_(sh, row, col)) {
    const a1 = sh.getRange(row, col, 1, 1).getA1Notation();
    const msg = `${SHEET_SAFETY.LOG_PREFIX} WRITE BLOCKED at ${a1} (col ${col}) — formula detected in cell/column.`;
    if (SHEET_SAFETY.ABORT_ON_PROTECTED_WRITE) throw new Error(msg);
    console.warn(msg + ' Skipped.');
    return false;
  }
  sh.getRange(row, col, 1, 1).setValue(text);
  return true;
}

function _safeSetHeaderIfNew_(sh, r, c, headerText) {
  // Don't overwrite if a formula already sits on the header cell
  const rng = sh.getRange(r, c);
  if (!rng.getValue() && !rng.getFormula()) {
    rng.setValue(headerText);
  }
}

/***** ========== SHOPIFY APP (OAuth) — integrated in the same Apps Script project ========== *****/
// If you redeploy and your Web App URL changes, update Shopify Partner Dashboard
// (App URL + Allowed redirection URL(s)) or install will fail.
const SHOPIFY_APP = {
  CLIENT_ID:    'YOUR_SHOPIFY_CLIENT_ID',
  CLIENT_SECRET:'YOUR_SHOPIFY_CLIENT_SECRET',
  // Scopes required to read PII and push tracking:
  // - read_all_orders, read_orders: order history + shipping address
  // - read_customers: PII
  // - read_fulfillments / write_fulfillments: push tracking
  SCOPES: 'read_orders,read_all_orders,read_customers,read_fulfillments,write_fulfillments,read_merchant_managed_fulfillment_orders,write_merchant_managed_fulfillment_orders,read_assigned_fulfillment_orders,write_assigned_fulfillment_orders,read_third_party_fulfillment_orders,write_third_party_fulfillment_orders',
  REDIRECT_URI: ScriptApp.getService().getUrl() // current /exec of your deployment
};

// HMAC helpers for Shopify
function _toHexLower(bytes) {
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
function _shopifyHmacValid(params, providedHmac) {
  const secret = SHOPIFY_APP.CLIENT_SECRET;
  const entries = Object.keys(params)
    .filter(k => k !== 'hmac' && k !== 'signature')
    .sort()
    .map(k => `${k}=${params[k]}`);
  const message = entries.join('&');
  const digest = Utilities.computeHmacSha256Signature(message, secret);
  const calc = _toHexLower(digest);
  return calc === String(providedHmac || '').toLowerCase();
}

// Token storage by shop
function _setShopToken(shop, token, scope) {
  PropertiesService.getScriptProperties().setProperty('SHOPIFY_TOKEN__' + shop, token);
  if (scope) PropertiesService.getScriptProperties().setProperty('SHOPIFY_SCOPE__' + shop, scope);
}
function _getShopToken(shop) {
  return PropertiesService.getScriptProperties().getProperty('SHOPIFY_TOKEN__' + shop) || '';
}

// Top-level page (no iframe) to click-through to Shopify
function _renderTopLevelAuthorize(authUrl) {
  var html = HtmlService.createHtmlOutput(
    '<!doctype html><meta charset="utf-8">' +
    '<div style="font:14px/1.4 system-ui,sans-serif;padding:24px">' +
      '<h3>Redirecting to Shopify…</h3>' +
      '<p>Click below to continue the installation on Shopify:</p>' +
      '<p><a href="' + authUrl + '" target="_top" rel="noopener" ' +
      'style="display:inline-block;padding:10px 14px;border:1px solid #ccc;border-radius:6px;text-decoration:none">Continue to Shopify</a></p>' +
      '<p style="margin-top:12px;color:#666">If the button is blocked, copy/paste this URL:</p>' +
      '<code style="display:block;word-break:break-all;background:#f6f6f6;padding:8px;border-radius:6px">' + authUrl + '</code>' +
    '</div>'
  );
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

/**
 * OAuth entrypoint (install + callback)
 * - First hit:  ?shop=...&hmac=...&host=...&timestamp=... → HMAC check → render link to /admin/oauth/authorize
 * - Callback:   ?shop=...&code=...&state=... → exchange code→access_token → store in ScriptProperties → success page
 */
function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const shop = p.shop;
  const code = p.code;
  const hmac = p.hmac;
  const state = p.state;

  // Callback with code
  if (shop && code) {
    const url = `https://${shop}/admin/oauth/access_token`;
    const body = {
      client_id: SHOPIFY_APP.CLIENT_ID,
      client_secret: SHOPIFY_APP.CLIENT_SECRET,
      code
    };
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const txt = res.getContentText();
    let json = {};
    try { json = JSON.parse(txt); } catch(_) {}
    if (!json.access_token) {
      return HtmlService.createHtmlOutput(
        `<pre>Code→token exchange failed (${res.getResponseCode()}):\n${txt}</pre>`
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    _setShopToken(shop, json.access_token, json.scope);

    return HtmlService.createHtmlOutput(
      `<html><body style="font:14px/1.4 system-ui, -apple-system, Segoe UI, Roboto">
        <h3>✅ App installed for <code>${shop}</code></h3>
        <p>Access token saved in ScriptProperties.</p>
        <p>Granted scopes: <code>${json.scope || '(not returned)'}</code></p>
        <p>You can close this tab and rerun your pipeline.</p>
      </body></html>`
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // First hit (install)
  if (shop && hmac) {
    if (!_shopifyHmacValid(p, hmac)) {
      return HtmlService.createHtmlOutput('<pre>Invalid HMAC</pre>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    const stateVal = Utilities.getUuid();
    const installAuthorizeUrl =
      `https://${shop}/admin/oauth/authorize` +
      `?client_id=${encodeURIComponent(SHOPIFY_APP.CLIENT_ID)}` +
      `&scope=${encodeURIComponent(SHOPIFY_APP.SCOPES)}` +
      `&redirect_uri=${encodeURIComponent(SHOPIFY_APP.REDIRECT_URI)}` +
      `&state=${encodeURIComponent(stateVal)}`;
    return _renderTopLevelAuthorize(installAuthorizeUrl);
  }

  // Home/debug
  return HtmlService.createHtmlOutput(
    `<html><body style="font:14px/1.4 system-ui, -apple-system, Segoe UI, Roboto">
      <h3>Shopify OAuth</h3>
      <p>Missing parameters. Use your Partner install link to choose a shop.</p>
    </body></html>`
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/***** ========== IMPLEMENTATION NOTES (what this script does) ========== *****
 * 1) Address matching:
 *    - Extracts <td class="EDM-SHIP-TO-address">…</td> from AliExpress email HTML (FR template by default).
 *    - Heavy normalization (lowercase, strip diacritics/punctuation, collapse spaces).
 *    - Primary test: aliNorm.startsWith(shopNorm) OR shopNorm.startsWith(aliNorm) (robust to variants).
 *    - Fallbacks: Levenshtein distance ≤ 6 (on ~≥20 chars) OR Jaccard ≥ 0.85.
 *
 * 2) Dedup & separators:
 *    - Columns “AE ID(s)” and “Tracking ID(s)” can store multiple values separated by CFG.CELL_SEP.
 *    - Before adding, split + Set() to avoid duplicates.
 *
 * 3) Gmail cleanup:
 *    - Confirmation: if the Ali order ID is already written to the sheet → delete the confirmation email.
 *    - Shipping: if at least one new tracking was added (and pushed to Shopify) → delete the shipping email.
 *
 * 4) Shopify:
 *    - Fetches last N days of **fulfilled** orders per store (GraphQL search query).
 *    - Pushes tracking via GraphQL `fulfillmentTrackingInfoUpdate` (multi numbers, dedup).
 *
 * 5) AliExpress:
 *    - Calls `aliexpress.ds.order.tracking.get` (TOP). Signature = HMAC-SHA256 with strict key sorting.
 *    - Returns an array of tracking numbers (unique).
 *
 * 6) 72h Alert:
 *    - Memorizes first confirmation timestamp per Ali order (ScriptProperties “aliPlacedAt:<id>”).
 *    - If >72h and no tracking → send alert email to CFG.ALERT_RECIPIENTS.
 *
 * 7) Safety / Testing:
 *    - Set CFG.DRY_RUN=true to validate end-to-end without writing Sheet / pushing Shopify / deleting Gmail.
 ********************************************************/

/***** ========== GENERAL UTILITIES ========== *****/
function _now() { return Date.now(); }

function _log(obj, label = '') {
  if (label) console.info(label + ': ' + (typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2)));
  else console.info(typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2));
}

function _hexUpper(bytes) {
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').toUpperCase();
}

function _removeDiacritics(s) {
  return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function _normalizeAddr(s) {
  if (!s) return '';
  s = s.replace(/<br\s*\/?>/gi, ' ').replace(/&nbsp;/gi, ' ').replace(/<[^>]+>/g, ' ');
  s = _removeDiacritics(s.toLowerCase());
  s = s.replace(/[^\p{L}\p{N}\s]/gu, ' ').replace(/\s+/g, ' ').trim();
  return s;
}

function _levenshtein(a, b) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  const dp = Array.from({length: m+1}, (_,i)=>Array(n+1).fill(0));
  for (let i=0;i<=m;i++) dp[i][0]=i;
  for (let j=0;j<=n;j++) dp[0][j]=j;
  for (let i=1;i<=m;i++) {
    for (let j=1;j<=n;j++) {
      const cost = a[i-1]===b[j-1]?0:1;
      dp[i][j]=Math.min(dp[i-1][j]+1, dp[i][j-1]+1, dp[i-1][j-1]+cost);
    }
  }
  return dp[m][n];
}

function _jaccard(a, b) {
  const A = new Set(a.split(' ')); const B = new Set(b.split(' '));
  const inter = new Set([...A].filter(x=>B.has(x))).size;
  const uni = new Set([...A, ...B]).size || 1;
  return inter/uni;
}

function _matchWithMetrics(aliNorm, shopNorm) {
  aliNorm = aliNorm || '';
  shopNorm = shopNorm || '';

  const startsAliShop = aliNorm && shopNorm && aliNorm.startsWith(shopNorm);
  const startsShopAli = aliNorm && shopNorm && shopNorm.startsWith(aliNorm);

  const d = _levenshtein(aliNorm, shopNorm);
  const j = _jaccard(aliNorm, shopNorm);
  const lenOK = Math.max(aliNorm.length, shopNorm.length) >= 20;

  const ok = (startsAliShop || startsShopAli) || (lenOK && d <= 6) || (j >= 0.85);

  return {
    ok,
    startsAliShop,
    startsShopAli,
    d,
    j: Number(j.toFixed(3)),
    lenOK
  };
}

function _afterNDays(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d;
}

function _ts() { return String(Date.now()); }

function _ppToken(email) {
  // ParcelPanel token: reverse email, replace '@' → '_-_', then URL-encode
  return encodeURIComponent(String(email || '')
    .split('').reverse().join('')
    .replace('@', '_-_'));
}

function _parcelPanelUrl(brandKey, orderName, email) {
  const brand = CFG.BRANDS[brandKey] || {};
  const domain = String(brand.publicDomain || '').replace(/\/+$/,''); // no trailing slash
  const orderNo = String(orderName || '').replace(/^#/, '');          // no '#'
  const token = _ppToken(email || '');
  return `https://${domain}/apps/suivi-commande?order=${orderNo}&token=${token}`;
}

function _postForm(url, paramsObj) {
  const payload = Object.keys(paramsObj).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(paramsObj[k])}`).join('&');
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded;charset=UTF-8',
    payload,
    muteHttpExceptions: true
  });
  return res;
}

function _parseMaybeGop(jsonText) {
  try {
    const outer = JSON.parse(jsonText);
    if (outer && typeof outer === 'object' && 'gopResponseBody' in outer) {
      try { return JSON.parse(outer.gopResponseBody); } catch(e) { return outer; }
    }
    return outer;
  } catch(e) {
    throw new Error('Non-JSON response:\n' + jsonText.slice(0, 400));
  }
}

// Helper to read money objects (object OR array)
function _moneyList(m) {
  if (!m) return [];
  if (Array.isArray(m)) return m;
  if (typeof m === 'object' && ('amount' in m || 'currency_code' in m)) return [m];
  return [];
}


/***** ========== ALIEXPRESS SIGNATURES (GOP/TOP) ========== *****/
// — SYSTEM APIs (e.g., /auth/token/create, /auth/token/refresh)
function _gopSignSystem(apiNameForSign, params, appSecret) {
  const keys = Object.keys(params)
    .filter(k => k !== 'sign' && params[k] !== undefined && params[k] !== '')
    .sort();
  let base = apiNameForSign;
  for (const k of keys) base += k + String(params[k]);
  const raw = Utilities.computeHmacSha256Signature(base, appSecret);
  return raw.map(b => ((b & 0xff).toString(16)).padStart(2,'0')).join('').toUpperCase();
}

// — BUSINESS APIs (e.g., aliexpress.ds.order.tracking.get) via /rest
function _gopSignBusiness(params, appSecret) {
  const keys = Object.keys(params).filter(k=>k!=='sign').sort();
  let base = '';
  for (const k of keys) {
    const v = params[k];
    if (v !== undefined && v !== null && v !== '') base += (k + String(v));
  }
  const raw = Utilities.computeHmacSha256Signature(base, appSecret, Utilities.Charset.UTF_8);
  return _hexUpper(raw);
}

function _gopPostSystem(apiPathCanonical, extraParams = {}) {
  const paramKeys = ['api', 'api_name', 'method'];
  const apiVariants = [apiPathCanonical, apiPathCanonical.replace(/^\//, '')];
  let lastText = '';
  for (const key of paramKeys) {
    for (const variant of apiVariants) {
      const p = {
        app_key: CFG.AE.APP_KEY,
        sign_method: 'sha256',
        timestamp: _ts(),
        ...extraParams
      };
      p[key] = variant;
      p.sign = _gopSignSystem(variant, p, CFG.AE.APP_SECRET);
      const res = _postForm(CFG.AE.HOST_REST, p);
      const text = res.getContentText();
      lastText = text;
      const parsed = _parseMaybeGop(text);
      if (!(parsed && parsed.code === 'InvalidApiPath')) {
        return parsed || text;
      }
    }
  }
  throw new Error('Token exchange failed (InvalidApiPath on all variants). Last response: ' + lastText);
}

// TOP (Business) via /sync
function _topSign(params, appSecret) {
  const keys = Object.keys(params)
    .filter(k => k !== 'sign' && params[k] !== undefined && params[k] !== '')
    .sort();
  let base = '';
  for (const k of keys) base += k + String(params[k]) ;
  const raw = Utilities.computeHmacSha256Signature(base, appSecret, Utilities.Charset.UTF_8);
  return _hexUpper(raw);
}

function _topPostBusiness(method, bizParams = {}) {
  const ts = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd HH:mm:ss');
  const p = {
    method,
    app_key: CFG.AE.APP_KEY,
    sign_method: 'sha256',
    timestamp: ts,         // TOP requires Asia/Shanghai timestamp
    v: '2.0',
    simplify: 'true',
    session: _Ali_getValidAccessToken(), // TOP uses "session" = access_token
    ...bizParams
  };
  p.sign = _topSign(p, CFG.AE.APP_SECRET);

  const payload = Object.keys(p).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(p[k])).join('&');
  const res = UrlFetchApp.fetch(CFG.AE.HOST_TOP_SYNC, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded;charset=UTF-8',
    payload,
    muteHttpExceptions: true
  });
  const text = res.getContentText();
  return _parseMaybeGop(text);
}


/***** ========== ALIEXPRESS AUTH (CREATE / REFRESH) ========== *****/
function AliAuth_ExchangeCodeOnce() {
  if (!CFG.AE.AUTH_CODE) throw new Error('CFG.AE.AUTH_CODE is empty. Paste your authorization code then rerun.');
  const apiPath = '/auth/token/create';
  const out = _gopPostSystem(apiPath, { code: CFG.AE.AUTH_CODE });
  _log(out, 'Create token (parsed)');
  if (!out || String(out.code) !== '0' || !out.access_token) {
    throw new Error('Code→token exchange failed. Response:\n' + JSON.stringify(out, null, 2));
  }

  const expiresAt = Date.now() + Number(out.expires_in) * 1000;
  console.info('=== COPY INTO CFG.AE (or keep in ScriptProperties) ===');
  console.info('ACCESS_TOKEN: ' + out.access_token);
  console.info('REFRESH_TOKEN: ' + out.refresh_token);
  console.info('EXPIRES_IN (s): ' + out.expires_in + '  -> EXPIRES_AT_MS: ' + expiresAt);

  const p = PropertiesService.getScriptProperties();
  p.setProperty('AE_ACCESS_TOKEN', out.access_token);
  p.setProperty('AE_REFRESH_TOKEN', out.refresh_token);
  p.setProperty('AE_EXPIRES_AT_MS', String(expiresAt));

  return out;
}

function AliAuth_Refresh() {
  const rt = CFG.AE.REFRESH_TOKEN || PropertiesService.getScriptProperties().getProperty('AE_REFRESH_TOKEN');
  if (!rt) throw new Error('No REFRESH_TOKEN (CFG or ScriptProperties).');

  const apiPath = '/auth/token/refresh';
  const out = _gopPostSystem(apiPath, { refresh_token: rt });
  _log(out, 'Refresh token (parsed)');
  if (!out || String(out.code) !== '0' || !out.access_token) {
    throw new Error('Refresh token failed. Response:\n' + JSON.stringify(out, null, 2));
  }

  const expiresAt = Date.now() + Number(out.expires_in) * 1000;
  console.info('=== COPY INTO CFG.AE ===');
  console.info('ACCESS_TOKEN: ' + out.access_token);
  console.info('REFRESH_TOKEN: ' + out.refresh_token);
  console.info('EXPIRES_IN (s): ' + out.expires_in + '  -> EXPIRES_AT_MS: ' + expiresAt);

  const p = PropertiesService.getScriptProperties();
  p.setProperty('AE_ACCESS_TOKEN', out.access_token);
  p.setProperty('AE_REFRESH_TOKEN', out.refresh_token);
  p.setProperty('AE_EXPIRES_AT_MS', String(expiresAt));

  return out;
}

function _Ali_getValidAccessToken() {
  if (CFG.AE.ACCESS_TOKEN && Number(CFG.AE.EXPIRES_AT_MS) > Date.now() + 5*60*1000) {
    return CFG.AE.ACCESS_TOKEN;
  }
  const p = PropertiesService.getScriptProperties();
  const acc = p.getProperty('AE_ACCESS_TOKEN');
  const exp = Number(p.getProperty('AE_EXPIRES_AT_MS') || 0);
  if (acc && exp > Date.now() + 5*60*1000) return acc;

  const out = AliAuth_Refresh();
  return out.access_token;
}


/***** ========== BUSINESS: TRACKING GET (AliExpress) ========== *****/
function Ali_TrackingGet(ae_order_id, language = 'en_US') {
  // Retry on TOP rate limit
  for (let attempt = 1; attempt <= CFG.AE.ALI_RATE.MAX_RETRIES; attempt++) {
    const out = _topPostBusiness('aliexpress.ds.order.tracking.get', {
      ae_order_id,
      language
    });

    const err = out && out.error_response;
    if (err) {
      const code = String(err.code || '');
      const sub  = String(err.sub_code || '');
      const msg  = String(err.msg || '');
      const isRate =
        code === 'ApiCallLimit' ||
        /api.*limit/i.test(sub) ||
        /Api access frequency exceeds the limit/i.test(msg);

      if (isRate) {
        const waitSec = _aliParseBanSeconds_(msg); // often 1
        const sleepMs = Math.max(1000, waitSec * 1000) + Math.floor(Math.random() * CFG.AE.ALI_RATE.JITTER_MS);
        console.warn(`[RATE_LIMIT] ${ae_order_id} — "${msg}". Waiting ~${sleepMs}ms then retry ${attempt}/${CFG.AE.ALI_RATE.MAX_RETRIES}`);
        Utilities.sleep(sleepMs);
        continue;
      }

      // Other TOP error → return []
      console.warn(`Ali TOP error (non-rate) for ${ae_order_id}: ${JSON.stringify(err)}`);
      return [];
    }

    // Success
    _log(out, 'Tracking (raw parsed)');

    const wrap   = out && out.aliexpress_ds_order_tracking_get_response;
    const result = wrap && wrap.result ? wrap.result : out && out.result ? out.result : null;
    if (!result || result.ret === false) return [];

    const data = result.data || {};
    const list = data.tracking_detail_line_list;

    let lines = [];
    if (list && Array.isArray(list.tracking_detail)) {
      for (const d of list.tracking_detail) if (d && d.mail_no) lines.push(String(d.mail_no));
    }
    if (Array.isArray(list)) {
      for (const d of list) if (d && d.mail_no) lines.push(String(d.mail_no));
    }

    lines = [...new Set(lines.filter(Boolean))];
    return lines;
  }

  console.error(`[RATE_LIMIT] Abandoning after ${CFG.AE.ALI_RATE.MAX_RETRIES} attempts for ${ae_order_id}`);
  return [];
}

// Extract "ban will last N seconds"
function _aliParseBanSeconds_(msg) {
  const m = /ban will last\s+(\d+)\s*second/i.exec(String(msg || ''));
  return m ? Math.max(1, parseInt(m[1], 10)) : 1;
}


/***** ========== BUSINESS: ORDER GET (AliExpress) — amount annotation ========== *****/
function Ali_OrderAmountsGet(ae_order_id) {
  for (let attempt = 1; attempt <= CFG.AE.ALI_RATE.MAX_RETRIES; attempt++) {
    const out = _topPostBusiness('aliexpress.trade.ds.order.get', {
      // TOP expects a JSON string
      single_order_query: JSON.stringify({ order_id: String(ae_order_id) })
    });

    const err = out && out.error_response;
    if (err) {
      const code = String(err.code || '');
      const sub  = String(err.sub_code || '');
      const msg  = String(err.msg || '');
      const isRate = code === 'ApiCallLimit' || /api.*limit/i.test(sub) || /Api access frequency exceeds the limit/i.test(msg);
      if (isRate) {
        const waitSec = _aliParseBanSeconds_(msg);
        const sleepMs = Math.max(1000, waitSec * 1000) + Math.floor(Math.random() * CFG.AE.ALI_RATE.JITTER_MS);
        Utilities.sleep(sleepMs);
        continue;
      }
      // Missing permission / order not found → return null (we'll write raw AE ID)
      console.warn(`Ali GET order error for ${ae_order_id}: ${JSON.stringify(err)}`);
      return null;
    }

    const wrap = out && out.aliexpress_trade_ds_order_get_response;
    const result = wrap && wrap.result ? wrap.result : (out && out.result ? out.result : null);
    if (!result) return null;

    // 2.1 Amount paid by user
    const userArr = _moneyList(result.user_order_amount);
    const user = userArr[0] || {};
    const userPaidAmount = user.amount != null ? String(user.amount) : null;
    const userPaidCurrency = user.currency_code || user.currency || null;

    // 2.2 Sum of product_price * product_count
    let sum = 0;
    let sumCurrency = null;

    const childWrap = result.child_order_list || {};
    const child = childWrap.aeop_child_order_info || childWrap || [];
    const childArr = Array.isArray(child) ? child : [];

    for (const c of childArr) {
      const qty = Number(c.product_count || 1);
      const prices = _moneyList(c.product_price);

      if (prices.length) {
        const p = prices[0];
        const a = parseFloat(p.amount || 0);
        sum += (isFinite(a) ? a : 0) * (isFinite(qty) ? qty : 1);
        if (!sumCurrency) sumCurrency = p.currency_code || p.currency || null;
      } else if (c.actual_fee) {
        const af = _moneyList(c.actual_fee)[0];
        if (af && af.amount) {
          const a = parseFloat(af.amount);
          sum += isFinite(a) ? a : 0;
          if (!sumCurrency) sumCurrency = af.currency_code || af.currency || null;
        }
      }
    }

    if (!sumCurrency) sumCurrency = userPaidCurrency;

    const sumStr = isFinite(sum) ? sum.toFixed(2) : null;

    return {
      product_sum_amount: sumStr,
      product_sum_currency: sumCurrency || '',
      user_paid_amount: userPaidAmount || '',
      user_paid_currency: userPaidCurrency || ''
    };
  }
  return null;
}


/***** ========== GMAIL (scan & FR subject parsing by default) ========== *****/
// These patterns match French AliExpress transactional subjects.
// Adjust to your locale if needed.
const AE_SUBJECTS = {
  CONFIRM: /Commande\s+(\d+)\s*:\s*commande confirmée/i,
  SHIP: /Commande\s+(\d+)\s*:\s*(?:commande expédiée|partiellement expédiée)/i
};
const AE_FROM = 'transaction@notice.aliexpress.com';

function _gmailSearchThreads() {
  const q = [
    `from:${AE_FROM}`,
    `newer_than:${CFG.LOOKBACK_DAYS_EMAILS}d`,
    `(subject:commande)`
  ].join(' ');
  return GmailApp.search(q);
}

function _extractOrderIdFromSubject(subj) {
  if (!subj) return null;
  let m = subj.match(AE_SUBJECTS.CONFIRM); if (m) return m[1];
  m = subj.match(AE_SUBJECTS.SHIP); if (m) return m[1];
  return null;
}

function _extractShipToAddressFromHtml(html) {
  const m = html.match(/<td[^>]*class=["']EDM-SHIP-TO-address["'][^>]*>([\s\S]*?)<\/td>/i);
  return m ? m[1] : '';
}


/***** ========== SHOPIFY (via tokens stored by OAuth) ========== *****/
function _shopifyTokenFor(storeDomain) {
  const tok = _getShopToken(storeDomain);
  if (!tok) throw new Error('No Shopify token for ' + storeDomain + ' (install the app on this shop).');
  return tok;
}
function _shopifyHeadersByDomain(storeDomain) {
  const tok = _shopifyTokenFor(storeDomain);
  return { 'X-Shopify-Access-Token': tok, 'Content-Type': 'application/json' };
}

function _shopifyGetRecentOrders(brandKey) {
  const { storeUrl } = CFG.BRANDS[brandKey];
  const since = new Date(Date.now() - CFG.LOOKBACK_DAYS_EMAILS*24*60*60*1000).toISOString();
  const search = `fulfillment_status:fulfilled AND created_at:>=${since}`;

  const q = `
    query($first: Int!, $query: String!) {
      orders(first: $first, query: $query, sortKey: PROCESSED_AT, reverse: true) {
        nodes {
          id
          name
          email
          shippingAddress { address1 address2 }
        }
      }
    }`;

  const data = _shopifyGraphQL(storeUrl, q, { first: 250, query: search });
  const nodes = (data && data.orders && data.orders.nodes) || [];

  // Return an object compatible with the rest of the script (numeric id + shipping_address)
  return nodes.map(n => ({
    id: Number(String(n.id).split('/').pop()),
    name: n.name,
    email: n.email || '',
    shipping_address: {
      address1: n.shippingAddress && n.shippingAddress.address1 || '',
      address2: n.shippingAddress && n.shippingAddress.address2 || ''
    }
  }));
}

function _normalizeShopifyAddress(order) {
  const a = order?.shipping_address || {};
  // ONLY address1 + address2 for matching
  const parts = [a.address1, a.address2].filter(Boolean).join(' ');
  return _normalizeAddr(parts);
}

/***** ========== Shopify GraphQL ========== *****/
function _shopifyGraphQL(storeDomain, query, variables) {
  const url = `https://${storeDomain}/admin/api/2025-07/graphql.json`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: _shopifyHeadersByDomain(storeDomain),
    contentType: 'application/json',
    payload: JSON.stringify({ query, variables }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const txt = res.getContentText();
  if (code < 200 || code >= 300) throw new Error(`GraphQL ${code}: ${txt}`);
  const out = JSON.parse(txt);
  if (out.errors) throw new Error('GraphQL errors: ' + JSON.stringify(out.errors));
  return out.data;
}

function _shopifyGetFulfillments_GQL(brandKey, orderId) {
  const { storeUrl } = CFG.BRANDS[brandKey];
  const gidOrder = `gid://shopify/Order/${orderId}`;
  const q = `
    query($id: ID!) {
      order(id: $id) {
        fulfillments {
          id
          createdAt
          status
          name
          trackingInfo { number url company }
        }
      }
    }`;
  const d = _shopifyGraphQL(storeUrl, q, { id: gidOrder });
  const arr = d?.order?.fulfillments || [];
  return arr; // list of Fulfillment
}

// Push fulfillment tracking (GraphQL) multi-numbers + dedup
function _shopifyPushTracking(brandKey, orderId, trackingNumbers = [], orderName, orderEmail) {
  if (!trackingNumbers.length) return { ok:false, msg:'No tracking' };

  const { storeUrl } = CFG.BRANDS[brandKey];

  // 1) Target the most recent fulfillment (by createdAt DESC locally)
  const fulf = _shopifyGetFulfillments_GQL(brandKey, orderId);
  if (!fulf.length) return { ok:false, code:0, body:'No existing fulfillment to update' };
  const f = fulf.sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt))[0];
  const fulfillmentId = f.id;

  // 2) Dedup: union of existing + incoming numbers
  const existing = (f.trackingInfo || []).map(t => String(t.number)).filter(Boolean);
  const incoming = [...new Set(trackingNumbers.map(String).filter(Boolean))];
  const union = [...new Set([...existing, ...incoming])];
  const urls = union.map(() => _parcelPanelUrl(brandKey, orderName, orderEmail));

  const same = existing.length === union.length && existing.every(n => union.includes(n));
  if (same) {
    return { ok:true, updated:false, noop:true, msg:'Already up-to-date', fulfillmentId, existing };
  }

  const m = `
    mutation UpdateTracking($id: ID!, $numbers: [String!]!, $urls: [URL!]!, $notify: Boolean!) {
      fulfillmentTrackingInfoUpdate(
        fulfillmentId: $id,
        trackingInfoInput: { numbers: $numbers, urls: $urls, company: "Other" },
        notifyCustomer: $notify
      ) {
        fulfillment { id trackingInfo { number url company } }
        userErrors { field message }
      }
    }`;

  if (CFG.DRY_RUN) {
    _log({id: fulfillmentId, numbers: union, urls, company: "Other", notify: false},
        `[DRY_RUN] fulfillmentTrackingInfoUpdate ${brandKey} ${orderId}`);
    return { ok:true, dry:true, wouldSet: union, wouldUrls: urls, fulfillmentId };
  }

  const data = _shopifyGraphQL(storeUrl, m, {
    id: fulfillmentId,
    numbers: union,
    urls,
    notify: false
  });
  const errs = data?.fulfillmentTrackingInfoUpdate?.userErrors || [];
  if (errs.length) return { ok:false, code:422, body: JSON.stringify(errs) };

  const after = (data.fulfillmentTrackingInfoUpdate.fulfillment?.trackingInfo || []).map(t=>t.number);
  return { ok:true, updated:true, fulfillmentId, before: existing, after };
}


/***** ========== 72h ALERT ========== *****/
function _setPlacedAt(aeOrderId) {
  if (CFG.DRY_RUN) return;
  PropertiesService.getScriptProperties().setProperty('aliPlacedAt:'+aeOrderId, String(Date.now()));
}
function _getPlacedAt(aeOrderId) {
  const v = PropertiesService.getScriptProperties().getProperty('aliPlacedAt:'+aeOrderId);
  return v ? Number(v) : 0;
}
function _maybeAlert72h(aeOrderId) {
  const placedAt = _getPlacedAt(aeOrderId);
  if (!placedAt) return;
  const hours = (Date.now() - placedAt) / 3600000;
  if (hours >= 72) {
    const subj = `72h alert: no tracking for AliExpress order ${aeOrderId}`;
    const body = `Hello,\n\nAliExpress order ${aeOrderId} has no tracking number after ~${Math.round(hours)}h.\nPlease check with the seller.\n\n— Script`;
    if (!CFG.DRY_RUN) MailApp.sendEmail(CFG.ALERT_RECIPIENTS.join(','), subj, body);
    console.warn(subj);
  }
}


/***** ========== GOOGLE SHEET HELPERS (Consolidated - Balance Sheet) ========== *****/
function _sheetOpen_() {
  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + CFG.SHEET_NAME);
  return sh;
}

const SHEET_HEADER_SYNONYMS = {
  ORDER_ID: ['Order #ID'],
  STORE:    ['From store?'],
  AE:       ['AE ID(s)'],
  TR:       ['Tracking ID(s)']
};

// Find header row & columns (search first 5 rows)
function _sheetFindHeaderRowAndCols_(sh) {
  for (let r = 1; r <= 5; r++) {
    const headers = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
    const find = names => {
      const i = headers.findIndex(h => names.some(n => h.toLowerCase() === n.toLowerCase()));
      return i >= 0 ? (i + 1) : 0;
    };
    const colOrder = find(SHEET_HEADER_SYNONYMS.ORDER_ID);
    const colStore = find(SHEET_HEADER_SYNONYMS.STORE);
    let   colAe    = find(SHEET_HEADER_SYNONYMS.AE);
    let   colTr    = find(SHEET_HEADER_SYNONYMS.TR);
    if (colOrder && colStore) {
      // Create AE/Tracking columns if missing (append to header row)
      let last = sh.getLastColumn();
      if (!colAe) { sh.insertColumnAfter(last); last++; _safeSetHeaderIfNew_(sh, r, last, 'AE ID(s)'); colAe = last; }
      if (!colTr) { sh.insertColumnAfter(last); last++; _safeSetHeaderIfNew_(sh, r, last, 'Tracking ID(s)'); colTr = last; }
      return { headerRow: r, colOrder, colStore, colAe, colTr };
    }
  }
  throw new Error('"Order #ID" / "From store?" headers not found (rows 1–5).');
}

// Union write (dedup) — with formula protection
function _sheetUnionWrite_(sh, row, col, values, sep = CFG.CELL_SEP) {
  if (!values || !values.length) return;
  if (_isProtectedCell_(sh, row, col)) {
    const a1 = sh.getRange(row, col).getA1Notation();
    const msg = `${SHEET_SAFETY.LOG_PREFIX} UNION WRITE BLOCKED at ${a1} — formula detected in cell/column`;
    if (SHEET_SAFETY.ABORT_ON_PROTECTED_WRITE) throw new Error(msg);
    console.warn(msg + ' Skipped.');
    return;
  }
  const rng = sh.getRange(row, col, 1, 1);
  const existing = String(rng.getDisplayValue() || '')
    .split(sep).map(s => s.trim()).filter(Boolean);
  const union = [...new Set([...existing, ...values.map(v => String(v).trim()).filter(Boolean)])];
  rng.setValue(union.join(sep));
}

// Find row by numeric Order ID + full store label
function _sheetFindRowByOrderAndStore_(orderNumberNumeric, storeLabelFull) {
  const sh = _sheetOpen_();
  const { headerRow, colOrder, colStore, colAe, colTr } = _sheetFindHeaderRowAndCols_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) return { sh, row: 0, headerRow, colAe, colTr };

  const n = String(orderNumberNumeric).replace(/^#/, '').trim();
  const orderCol = sh.getRange(headerRow + 1, colOrder, lastRow - headerRow, 1).getDisplayValues().map(r => String(r[0]).replace(/^#/, '').trim());
  const storeCol = sh.getRange(headerRow + 1, colStore, lastRow - headerRow, 1).getDisplayValues().map(r => String(r[0]).trim());

  for (let i = 0; i < orderCol.length; i++) {
    if (orderCol[i] === n && storeCol[i] === storeLabelFull) {
      return { sh, row: headerRow + i + 1, headerRow, colAe, colTr };
    }
  }
  return { sh, row: 0, headerRow, colAe, colTr };
}

// Public API to upsert AE ID + Tracking on the correct row — protected
function _sheetUpsertAliAndTracking_ByOrderAndBrand(orderNumberNumeric, brandKey, aeIds = [], trackings = []) {
  const storeLabel = CFG.STORE_NAMES[brandKey]; // full store name, e.g., "Example Store A"
  const ctx = _sheetFindRowByOrderAndStore_(orderNumberNumeric, storeLabel);
  if (!ctx.row) {
    throw new Error(`Row not found for Order #${orderNumberNumeric} @ "${storeLabel}"`);
  }
  // AE IDs
  if (aeIds.length && _isProtectedCell_(ctx.sh, ctx.row, ctx.colAe)) {
    const a1 = ctx.sh.getRange(ctx.row, ctx.colAe).getA1Notation();
    const msg = `${SHEET_SAFETY.LOG_PREFIX} AE WRITE BLOCKED at ${a1} — formula detected in cell/column`;
    if (SHEET_SAFETY.ABORT_ON_PROTECTED_WRITE) throw new Error(msg); else { console.warn(msg + ' Skipped.'); }
  } else {
    _sheetUnionWrite_(ctx.sh, ctx.row, ctx.colAe, aeIds);
  }
  // Tracking numbers
  if (trackings.length && _isProtectedCell_(ctx.sh, ctx.row, ctx.colTr)) {
    const a1 = ctx.sh.getRange(ctx.row, ctx.colTr).getA1Notation();
    const msg = `${SHEET_SAFETY.LOG_PREFIX} TR WRITE BLOCKED at ${a1} — formula detected in cell/column`;
    if (SHEET_SAFETY.ABORT_ON_PROTECTED_WRITE) throw new Error(msg); else { console.warn(msg + ' Skipped.'); }
  } else {
    _sheetUnionWrite_(ctx.sh, ctx.row, ctx.colTr, trackings);
  }
}

// Format "annotated" AE ID for the sheet
function _formatAeAnnotation(aeId, amounts) {
  if (!amounts) return String(aeId);
  const aSum = amounts.product_sum_amount || '';
  const aCur = amounts.product_sum_currency || '';
  const uAmt = amounts.user_paid_amount || '';
  const uCur = amounts.user_paid_currency || '';
  // AEID (SUM CURRENCY • USER CURRENCY)
  return `${aeId} (${aSum} ${aCur} • ${uAmt} ${uCur})`.trim();
}

// Upsert annotated AE ID (replaces any previous entry starting with this AE ID)
function _sheetUpsertAeAnnotated_ByOrderAndBrand(orderNumberNumeric, brandKey, aeId, annotatedText) {
  const storeLabel = CFG.STORE_NAMES[brandKey];
  const ctx = _sheetFindRowByOrderAndStore_(orderNumberNumeric, storeLabel);
  if (!ctx.row) throw new Error(`Row not found for Order #${orderNumberNumeric} @ "${storeLabel}"`);

  if (_isProtectedCell_(ctx.sh, ctx.row, ctx.colAe)) {
    const a1 = ctx.sh.getRange(ctx.row, ctx.colAe).getA1Notation();
    const msg = `${SHEET_SAFETY.LOG_PREFIX} AE ANNOTATION BLOCKED at ${a1} — formula detected in cell/column`;
    if (SHEET_SAFETY.ABORT_ON_PROTECTED_WRITE) throw new Error(msg);
    console.warn(msg + ' Skipped.');
    return;
  }

  const rng = ctx.sh.getRange(ctx.row, ctx.colAe, 1, 1);
  const sep = CFG.CELL_SEP;
  const existing = String(rng.getDisplayValue() || '').split(sep).map(s => s.trim()).filter(Boolean);

  // Remove any entry that starts with this AE ID (e.g., "12345" or "12345 (..)")
  const filtered = existing.filter(v => !v.startsWith(String(aeId)));

  const final = [...new Set([...filtered, String(annotatedText)])];
  rng.setValue(final.join(sep));
}


/***** ========== MAIN PIPE (emails → tracking → Shopify) ========== *****/
function Sync_FromEmails_Main() {
  const threads = _gmailSearchThreads();
  console.info(`Threads found: ${threads.length}`);

  // Index recent Shopify orders by store
  const shopIndex = {};
  for (const brandKey of Object.keys(CFG.BRANDS)) {
    const orders = _shopifyGetRecentOrders(brandKey);
    shopIndex[brandKey] = orders.map(o => ({ order: o, addr: _normalizeShopifyAddress(o) }));
  }

  for (const th of threads) {
    const msgs = th.getMessages();
    for (const m of msgs) {
      const subj = m.getSubject() || '';
      const from = m.getFrom() || '';
      if (!from.toLowerCase().includes(AE_FROM)) continue;

      const aeOrderId = _extractOrderIdFromSubject(subj);
      if (!aeOrderId) continue;

      const html = m.getBody();
      const aeAddrHtml = _extractShipToAddressFromHtml(html);
      const aliAddr = _normalizeAddr(aeAddrHtml);

      // ===== CONFIRM =====
      if (AE_SUBJECTS.CONFIRM.test(subj)) {
        if (!CFG.DRY_RUN) _setPlacedAt(aeOrderId);

        let wrote = false;
        for (const brandKey of Object.keys(shopIndex)) {
          for (const {order, addr} of shopIndex[brandKey]) {
            const metrics = _matchWithMetrics(aliAddr, addr);
            if (!metrics.ok) continue;
            try {
              const ordNum = order.order_number || String(order.name).replace(/^#/, '');
              let annot = String(aeOrderId);
              try {
                const amts = Ali_OrderAmountsGet(aeOrderId);
                annot = _formatAeAnnotation(aeOrderId, amts);
              } catch (e) { console.warn('Ali_OrderAmountsGet failed:', e); }

              if (!CFG.DRY_RUN) {
                _sheetUpsertAeAnnotated_ByOrderAndBrand(ordNum, brandKey, aeOrderId, annot);
              }
              wrote = true;
              break;
            } catch(e) { console.warn('Sheet confirm write failed:', e); }
          }
          if (wrote) break;
        }

        if (wrote && !CFG.DRY_RUN) { try { th.moveToTrash(); } catch(_) {} }
        if (CFG.DRY_RUN) console.info(`[DRY_RUN] CONFIRM: ${aeOrderId} — no placedAt, no sheet write`);
        else             console.info(`CONFIRM: ${aeOrderId} — placedAt stored`);
        continue;
      }

      // ===== SHIP =====
      if (AE_SUBJECTS.SHIP.test(subj)) {
        console.info(`SHIP: ${aeOrderId}`);
        let trackings = [];
        try { trackings = Ali_TrackingGet(aeOrderId, 'en_US'); } catch(e) { console.error(e); }

        if (!trackings.length) {
          console.warn(`No tracking returned for ${aeOrderId}`);
          _maybeAlert72h(aeOrderId);
          continue;
        }

        // Match with Shopify by address
        let pushed = false;
        for (const brandKey of Object.keys(shopIndex)) {
          for (const {order, addr} of shopIndex[brandKey]) {
            const metrics = _matchWithMetrics(aliAddr, addr);

            if (CFG.DEBUG_VERBOSE_MATCH) {
              console.info(JSON.stringify({
                debug: 'MATCH_ATTEMPT',
                brand: brandKey,
                order_id: order.id,
                order_name: order.name,
                ali_norm: aliAddr,
                shop_norm: addr,
                startsAliShop: metrics.startsAliShop,
                startsShopAli: metrics.startsShopAli,
                levenshtein: metrics.d,
                jaccard: metrics.j,
                lenOK: metrics.lenOK,
                matched: metrics.ok
              }));
            }

            if (!metrics.ok) continue;

            // Push tracking (GraphQL) — multi-numbers + dedup
            const r = _shopifyPushTracking(brandKey, order.id, trackings, order.name, order.email);
            _log(r, `Push tracking -> ${brandKey} / order ${order.name || order.id}`);

            const successForCleanup = r.ok === true;
            pushed = pushed || successForCleanup;

            if (successForCleanup && !CFG.DRY_RUN) {
              try {
                const ordNum = order.order_number || String(order.name).replace(/^#/, '');
                // 1) AE ID annotated
                let annot = String(aeOrderId);
                try {
                  const amts = Ali_OrderAmountsGet(aeOrderId);
                  annot = _formatAeAnnotation(aeOrderId, amts);
                } catch (e) { console.warn('Ali_OrderAmountsGet failed:', e); }
                _sheetUpsertAeAnnotated_ByOrderAndBrand(ordNum, brandKey, aeOrderId, annot);

                // 2) Trackings column
                _sheetUpsertAliAndTracking_ByOrderAndBrand(ordNum, brandKey, [], trackings);
              } catch(e) { console.warn('Sheet ship write failed:', e); }

              try { th.moveToTrash(); } catch(e) { console.warn('moveToTrash failed:', e); }
              break;
            }
          }
          if (pushed) break;
        }
      }
    }
  }
}


/***** ========== TESTS / DEBUG ========== *****/
function Test_AE_Tracking() {
  const out = Ali_TrackingGet('REPLACE_WITH_AE_ORDER_ID_FOR_TEST', 'en_US');
  _log(out, 'TRACKINGS');
}

function Test_Auth_Refresh() {
  const out = AliAuth_Refresh();
  _log(out, 'REFRESH');
}

function Debug_LogAllShopifyAddresses() {
  for (const brandKey of Object.keys(CFG.BRANDS)) {
    const orders = _shopifyGetRecentOrders(brandKey);
    console.info(`=== ${brandKey} (${CFG.STORE_NAMES[brandKey]}) — ${orders.length} orders in ${CFG.LOOKBACK_DAYS_EMAILS}d ===`);
    orders.forEach(o => {
      const sa = o.shipping_address || {};
      const raw = [sa.name, sa.address1, sa.address2, sa.zip, sa.city, sa.country].filter(Boolean).join(' ');
      const norm = _normalizeShopifyAddress(o);
      console.info(JSON.stringify({
        brand: brandKey,
        order_id: o.id,
        order_name: o.name,
        shipping_address_raw: raw,
        shipping_address_obj: sa,
        shipping_address_normalized: norm
      }));
    });
  }
}

function Debug_LogShopifyAddresses_OneBrand(brandKey) {
  if (!CFG.BRANDS[brandKey]) {
    console.warn(`Unknown brand: ${brandKey}. Use one of: ${Object.keys(CFG.BRANDS).join(', ')}`);
    return;
  }
  const orders = _shopifyGetRecentOrders(brandKey);
  console.info(`=== ${brandKey} (${CFG.STORE_NAMES[brandKey]}) — ${orders.length} orders ===`);
  orders.forEach(o => {
    const sa = o.shipping_address || {};
    const raw = [sa.name, sa.address1, sa.address2, sa.zip, sa.city, sa.country].filter(Boolean).join(' ');
    const norm = _normalizeShopifyAddress(o);
    console.info(JSON.stringify({
      brand: brandKey,
      order_id: o.id,
      order_name: o.name,
      shipping_address_raw: raw,
      shipping_address_obj: sa,
      shipping_address_normalized: norm
    }));
  });
}

function Debug_TestAliAddrVsShopify(aliAddressInput) {
  let aliRaw = aliAddressInput || '';
  if (/EDM-SHIP-TO-address/i.test(aliRaw)) aliRaw = _extractShipToAddressFromHtml(aliRaw);
  const aliNorm = _normalizeAddr(aliRaw);
  console.info('ALI_RAW:', aliRaw);
  console.info('ALI_NORM:', aliNorm);

  const results = [];
  for (const brandKey of Object.keys(CFG.BRANDS)) {
    const orders = _shopifyGetRecentOrders(brandKey);
    for (const o of orders) {
      const shopNorm = _normalizeShopifyAddress(o);
      const d = _levenshtein(aliNorm, shopNorm);
      const j = _jaccard(aliNorm, shopNorm);
      results.push({
        brand: brandKey,
        order_id: o.id,
        order_name: o.name,
        shop_norm: shopNorm,
        lev: d,
        jaccard: Number(j.toFixed(3))
      });
    }
  }
  results.sort((a, b) => (b.jaccard - a.jaccard) || (a.lev - b.lev));
  console.info('=== TOP 20 similarities (Jaccard desc, Levenshtein asc) ===');
  results.slice(0, 20).forEach((r, i) => {
    console.info(`${String(i+1).padStart(2, '0')}) ${r.brand} ${r.order_name} | J=${r.jaccard} | L=${r.lev} | ${r.shop_norm}`);
  });
  console.info('Rule of thumb: match if startsWith() OR L<=6 (on ~>=20 chars) OR Jaccard>=0.85');
}

function Debug_TestFromAliHtml(html) {
  const addr = _extractShipToAddressFromHtml(html || '');
  Debug_TestAliAddrVsShopify(addr);
}
