// scan_builtin.js - Built In Boston discovery (run with javascript_tool in a tab on https://www.builtinboston.com).
// Define before running:  DAYS = 1 (normal run) or 3 (catch-up) or 7 (deep sweep)
// Returns relevant early-career AI/ML/data leads with the Built In company slug.
// Built In hides the employer apply link behind a login - NEVER log in. Use the lead only to find the
// employer's own posting (Greenhouse/Ashby/Lever feed, or the company careers site) and verify it there.
const TERMS = ["machine learning","AI engineer","artificial intelligence","data scientist","LLM","NLP",
  "generative AI","applied AI","ML engineer","computer vision","data science","deep learning",
  "research engineer","forward deployed","AI developer","associate data","junior AI","new grad"];
const REL = /(machine learning|\bML\b|\bAI\b|data scien|scientist|applied|LLM|NLP|vision|research engineer|deployed engineer|AI.Native|informatics|analytics engineer|automation engineer)/i;
const BAD = /(manager|director|sales|account|marketing|recruit|partner|product|program|head of|\blead\b|principal|staff|senior|sr\.|co-op|intern|extern|medical|founder|counsel)/i;
const sleep = ms => new Promise(r => setTimeout(r, ms));
const seen = new Map(); const log = [];
async function page(term, p) {
  const u = `/jobs?search=${encodeURIComponent(term)}&daysSinceUpdated=${DAYS}${p > 1 ? '&page=' + p : ''}`;
  const r = await fetch(u);
  if (r.status === 429) { log.push([term, p, 429]); await sleep(8000); return -1; }
  if (!r.ok) { log.push([term, p, r.status]); return 0; }
  const d = new DOMParser().parseFromString(await r.text(), 'text/html');
  const cards = [...d.querySelectorAll('[data-id="job-card"]')];
  for (const c of cards) {
    const id = (c.id || '').replace('job-card-', ''); if (!id || seen.has(id)) continue;
    const ja = c.querySelector('a[href*="/job/"]'), ca = c.querySelector('a[href*="/company/"]');
    const t = c.textContent.replace(/\s+/g, ' ').trim();
    seen.set(id, { id, title: ja ? ja.textContent.trim() : '',
      company: ca ? ((ca.textContent.trim() || (ca.querySelector('img') || {}).alt || '').replace(/ Logo$/, '')) : '',
      slug: ca ? ca.getAttribute('href').replace('/company/', '') : '',
      posted: (t.match(/((Reposted|Posted)?\s*(\d+|An|A)\s+(Minutes?|Hours?|Days?|Weeks?)\s+Ago|Yesterday|Today)/i) || [''])[0].trim(),
      level: (t.match(/(Entry level|Junior|Mid level|Senior level|Expert\/Leader|Internship)/) || [''])[0],
      salary: (t.match(/\d+(\.\d+)?K-\d+(\.\d+)?K Annually|\d+-\d+ Hourly/) || [''])[0],
      where: (t.match(/\b(Remote or Hybrid|In-Office or Remote|Hybrid|Remote|In-Office)\b/) || [''])[0],
      ma: (t.match(/[A-Z][a-z]+(?: [A-Z][a-z]+)?, MA/) || [''])[0],
      url: ja ? 'https://www.builtinboston.com' + ja.getAttribute('href') : '' });
  }
  return cards.length;
}
for (const term of TERMS) {
  for (let p = 1; p <= 8; p++) {
    let c = await page(term, p); if (c === -1) c = await page(term, p);
    await sleep(700);
    if (c < 10) break;
  }
}
const all = [...seen.values()];
const leads = all.filter(j => !['Senior level', 'Expert/Leader', 'Internship'].includes(j.level) && REL.test(j.title) && !BAD.test(j.title));
JSON.stringify({ scanned: all.length, leads: leads.length, errors: log.slice(0, 10), items: leads });
