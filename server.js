const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '1mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36';
const ED = 'https://www.expireddomains.net';

const cleanDomains = (arr, max = 500) =>
  [...new Set(arr.map((d) => String(d).trim().toLowerCase()).filter((d) => /^[a-z0-9.-]+\.[a-z]{2,}$/.test(d)))].slice(0, max);

// ---------- ExpiredDomains (.net) ----------
// Lists require a FREE expireddomains.net account. Send { edUser, edPass } — used only for this request, never stored.
app.post('/api/expired', async (req, res) => {
  const limit = Math.min(parseInt(req.body.limit || 100, 10), 500);
  const tld = String(req.body.tld || 'com').replace(/[^a-z]/gi, '');
  const edUser = req.body.edUser || process.env.ED_USER || '';
  const edPass = req.body.edPass || process.env.ED_PASS || '';
  const url = `${ED}/domains/expired${tld}/?flist=1`;

  try {
    let cookie = '';
    if (edUser && edPass) {
      const login = await axios.post(
        `${ED}/logincheck/`,
        new URLSearchParams({ login: edUser, password: edPass, rememberme: '1' }).toString(),
        {
          headers: { 'User-Agent': UA, 'Content-Type': 'application/x-www-form-urlencoded' },
          maxRedirects: 0,
          validateStatus: (s) => s < 400,
          timeout: 20000,
        }
      );
      cookie = (login.headers['set-cookie'] || []).map((c) => c.split(';')[0]).join('; ');
      if (!cookie) return res.status(401).json({ error: 'ExpiredDomains login failed — wrong username or password.' });
    }

    const { data } = await axios.get(url, {
      headers: { 'User-Agent': UA, 'Accept-Language': 'en-US,en;q=0.9', ...(cookie ? { Cookie: cookie } : {}) },
      timeout: 20000,
    });

    if (/<title>Login<\/title>/.test(data)) {
      return res.status(401).json({
        error: 'Login required — the site shows only a login page to guests.',
        fix: 'Register FREE at expireddomains.net/register, then enter your username + password in the app (or paste/upload domains instead).',
      });
    }

    const $ = cheerio.load(data);
    let domains = [];
    $('td.fieldDomain a, a.domain-link').each((_, el) => {
      const d = $(el).text().trim().toLowerCase().split(/\s/)[0];
      if (/^[a-z0-9.-]+\.[a-z]{2,}$/.test(d) && !domains.includes(d)) domains.push(d);
      if (domains.length >= limit) return false;
    });
    if (!domains.length) {
      // Fallback: any domain-looking link text on the page
      const re = new RegExp(`>([a-z0-9][a-z0-9-]*\\.${tld})<`, 'gi');
      let m;
      while ((m = re.exec(data)) && domains.length < limit) {
        const d = m[1].toLowerCase();
        if (!domains.includes(d)) domains.push(d);
      }
    }
    res.json({
      source: url,
      count: domains.length,
      domains: domains.slice(0, limit),
      warning: 'Preview list only. For full/filtered lists use the site export, then upload the file here.',
    });
  } catch (e) {
    res.status(502).json({ error: 'Failed to fetch ExpiredDomains.', detail: e.message });
  }
});

// ---------- File upload: CSV / TXT / Excel ----------
// Accepts .csv, .txt, .xls, .xlsx — first sheet / whole text scanned for domain-looking values.
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received.' });
    const name = (req.file.originalname || '').toLowerCase();
    let values = [];
    if (/\.xlsx?$/.test(name)) {
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }).flat().forEach((c) => {
        String(c == null ? '' : c).split(/[\s,;]+/).forEach((x) => values.push(x));
      });
    } else {
      values = req.file.buffer.toString('utf8').split(/[\s,;"']+/);
    }
    const domains = cleanDomains(values);
    if (!domains.length) return res.status(400).json({ error: 'No valid domains found in file.' });
    res.json({ count: domains.length, domains });
  } catch (e) {
    res.status(400).json({ error: 'Could not parse file: ' + e.message });
  }
});

// ---------- Ahrefs DR (API v3, paid token required) ----------
// GET https://api.ahrefs.com/v3/site-explorer/domain-rating?target=..&date=.. (per official OpenAPI spec)
async function getDR(domain, token) {
  if (!token) return { dr: null, drError: 'No Ahrefs token (API is paid-only, no free tier for arbitrary domains).' };
  try {
    const today = new Date().toISOString().slice(0, 10); // date is required by Ahrefs v3
    const { data } = await axios.get('https://api.ahrefs.com/v3/site-explorer/domain-rating', {
      params: { target: domain, date: today, output: 'json' },
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
      timeout: 15000,
    });
    const dr = data?.domain_rating?.domain_rating ?? null;
    return { dr, drError: dr === null ? `Unexpected Ahrefs response: ${JSON.stringify(data).slice(0, 200)}` : null };
  } catch (e) {
    const status = e.response?.status;
    const msg = e.response?.data ? JSON.stringify(e.response.data).slice(0, 300) : e.message;
    let hint = msg;
    if (status === 401) hint = 'Invalid token (401). Check Ahrefs API key / plan includes API access.';
    if (status === 402) hint = 'No API credits left (402). Ahrefs API costs extra beyond subscription.';
    if (status === 429) hint = 'Rate limited (429). Lower concurrency.';
    return { dr: null, drError: hint };
  }
}

// ---------- Moz DA/PA (free tier: ~1200 rows/month with free Community account) ----------
async function getMozBatch(domains, accessId, secret) {
  if (!accessId || !secret) {
    return Object.fromEntries(domains.map((d) => [d, { da: null, pa: null, mozError: 'No Moz keys set.' }]));
  }
  const out = {};
  // Moz allows up to 10 targets per request
  for (let i = 0; i < domains.length; i += 10) {
    const chunk = domains.slice(i, i + 10);
    try {
      const { data } = await axios.post(
        'https://lsapi.seomoz.com/v2/url_metrics',
        { targets: chunk },
        { auth: { username: accessId, password: secret }, timeout: 15000 }
      );
      const results = data?.results || data?.results_by_target || [];
      if (Array.isArray(results)) {
        results.forEach((r, idx) => {
          const d = chunk[idx];
          out[d] = { da: r.domain_authority ?? null, pa: r.page_authority ?? null, mozError: null };
        });
      } else {
        chunk.forEach((d) => { out[d] = { da: null, pa: null, mozError: 'Unexpected Moz response shape.' }; });
      }
    } catch (e) {
      const msg = e.response?.data ? JSON.stringify(e.response.data).slice(0, 300) : e.message;
      chunk.forEach((d) => { out[d] = { da: null, pa: null, mozError: msg }; });
    }
    await sleep(1100); // stay under free-tier rate limits
  }
  return out;
}

// ---------- Availability (RDAP + DNS, no paid keys needed) ----------
// status: available = RDAP 404 (not registered). registered = RDAP 200 or live DNS. unknown = inconclusive.
app.post('/api/availability', async (req, res) => {
  const domains = cleanDomains(Array.isArray(req.body.domains) ? req.body.domains : []);
  const results = [];
  for (const d of domains) {
    let status = 'unknown';
    let detail = '';
    try {
      let hasDNS = false, nx = false;
      try {
        const dns = await axios.get('https://dns.google/resolve', { params: { name: d, type: 'A' }, timeout: 10000 });
        hasDNS = dns.data && dns.data.Status === 0 && Array.isArray(dns.data.Answer) && dns.data.Answer.length > 0;
        nx = dns.data && dns.data.Status === 3;
      } catch (_) { /* DNS failed, rely on RDAP */ }
      let rdap = null;
      try {
        const r = await axios.get(`https://rdap.org/domain/${d}`, { timeout: 12000, maxRedirects: 5, validateStatus: (s) => s < 500 });
        rdap = r.status;
      } catch (e) { rdap = e.response ? e.response.status : null; }
      if (rdap === 404) status = 'available';
      else if (rdap === 200 || hasDNS) status = 'registered';
      detail = `rdap=${rdap === null ? '?' : rdap} dns=${hasDNS ? 'live' : (nx ? 'nxdomain' : 'none')}`;
    } catch (e) { detail = String(e.message).slice(0, 120); }
    results.push({ domain: d, status, detail });
    await sleep(500);
  }
  res.json({ count: results.length, results });
});

// ---------- Main check endpoint ----------
app.post('/api/check', async (req, res) => {
  const raw = Array.isArray(req.body.domains) ? req.body.domains : [];
  const domains = cleanDomains(raw);
  if (!domains.length) return res.status(400).json({ error: 'No valid domains. Send { "domains": ["example.com"] }' });

  const ahrefsToken = req.body.ahrefsToken || process.env.AHREFS_API_TOKEN || '';
  const mozId = req.body.mozAccessId || process.env.MOZ_ACCESS_ID || '';
  const mozSecret = req.body.mozSecretKey || process.env.MOZ_SECRET_KEY || '';

  const mozMap = await getMozBatch(domains, mozId, mozSecret);

  const results = [];
  for (const d of domains) {
    const { dr, drError } = await getDR(d, ahrefsToken);
    const m = mozMap[d] || {};
    results.push({ domain: d, dr, da: m.da ?? null, pa: m.pa ?? null, drError, mozError: m.mozError || null });
    await sleep(400); // be gentle on Ahrefs rate limits
  }
  res.json({ count: results.length, results });
});

app.get('/api/health', (_, res) => res.json({ ok: true }));

app.listen(PORT, '0.0.0.0', () => console.log(`Checker running on :${PORT}`));
