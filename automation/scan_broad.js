// scan_broad.js - broad CS-role scan with automatic eligibility flags (added 2026-11-12, her "any relevant CS role, contracts OK" request)
// Set KIND, TOKENS, SINCE (YYYY-MM-DD) before running in a tab on the matching API origin.
const INC = /(software|developer|programmer|engineer i\b|engineer 1\b|engineer ii\b|associate engineer|junior|entry|new grad|graduate|early career|university|data analyst|data engineer|analytics engineer|data scien|machine learning|\bml\b|\bai\b|artificial intelligence|llm|genai|generative|nlp|full[- ]?stack|back[- ]?end|front[- ]?end|python|java\b|qa engineer|quality engineer|test engineer|sdet|automation engineer|solutions engineer|implementation engineer|integration engineer|application engineer|applications engineer|cloud engineer|platform engineer|forward deployed|research engineer|bi developer|business intelligence)/i;
const EXC = /(senior|\bsr\b|sr\.|staff|principal|\blead\b|manager|director|head of|\bvp\b|intern\b|internship|co-op|architect|counsel|recruit|sales|account exec|marketing|designer|product manager|\biii\b|\biv\b|\bv\b|finance|legal|talent|customer success|mechanical|electrical|civil|chemical|hardware|firmware|fpga|asic|rf |manufacturing|technician|nurse|physician|clinical|biolog|chemist|tax|attorney|paralegal)/i;
const NONUS = /(canada|toronto|vancouver|montreal|ottawa|united kingdom|\buk\b|london|ireland|dublin|germany|berlin|munich|france|paris|netherlands|amsterdam|spain|madrid|barcelona|poland|warsaw|india|bangalore|bengaluru|hyderabad|pune|singapore|japan|tokyo|korea|seoul|australia|sydney|brazil|mexico|argentina|colombia|israel|tel aviv|switzerland|zurich|sweden|denmark|portugal|lisbon|romania|czech|philippines|emea|apac|latam)/i;
const US = /(united states|\busa\b|\bus\b|u\.s\.|remote|anywhere|[a-z]+, (al|ak|az|ar|ca|co|ct|de|fl|ga|hi|id|il|in|ia|ks|ky|la|me|md|ma|mi|mn|ms|mo|mt|ne|nv|nh|nj|nm|ny|nc|nd|oh|ok|or|pa|ri|sc|sd|tn|tx|ut|vt|va|wa|wv|wi|wy|dc)\b|new york|san francisco|boston|seattle|austin|chicago|los angeles|denver|atlanta|washington)/i;
const SK = ['python','sql','java','spring','react','typescript','javascript','next.js','node','aws','docker','fastapi','rest','api','pytorch','scikit','pandas','numpy','tensorflow','hugging face','nlp','llm','machine learning','git','mysql','linux','tableau','power bi','excel'];
const since = new Date(SINCE + 'T00:00:00-05:00');
const strip = h => { const d = new DOMParser().parseFromString(h || '', 'text/html'); let t = d.documentElement.textContent; if (/<[a-z]/i.test(t)) t = new DOMParser().parseFromString(t, 'text/html').documentElement.textContent; return t.replace(/\s+/g, ' '); };
function flags(t) {
  const yrs = [...t.matchAll(/(\d{1,2})\s*\+?\s*(?:-|–|to)?\s*\d{0,2}\s*\+?\s*years?/gi)].map(m => +m[1]).filter(n => n >= 1 && n <= 15);
  const f = [];
  if (/(u\.?s\.? citizen|citizenship (is )?required|security clearance|active (secret|ts)|secret clearance|ts\/sci|green card holder|permanent resident(s)? only)/i.test(t)) f.push('CITIZEN/CLEARANCE');
  if (/(\bopt\b|\bcpt\b)[^.]{0,60}(not (be )?(eligible|accepted|considered)|ineligible)|not[^.]{0,40}(accept|consider)[^.]{0,30}(opt|f-1|student visa)/i.test(t)) f.push('NO-OPT');
  if (/without (any )?(restrictions? or time limit|time limit)/i.test(t)) f.push('NO-TIME-LIMIT');
  if (/(not|unable to|cannot|can't|won't|will not|do not|does not)[^.]{0,25}(sponsor|provide sponsorship|offer sponsorship)/i.test(t)) f.push('no-sponsor');
  if (/(visa sponsorship (is )?available|will sponsor|sponsorship (is )?available|we sponsor)/i.test(t)) f.push('SPONSOR');
  if (/ph\.?d\.?[^.]{0,40}(required|is required)/i.test(t)) f.push('PHD');
  if (/(contract|contractor|temporary|temp-to-hire|c2h|w-?2 contract|1099)/i.test(t)) f.push('contract');
  const sal = (t.match(/\$\s?\d{2,3}[,.]?\d{0,3}(\.\d+)?\s?[kK]?\s*(?:-|–|to)\s*\$\s?\d{2,3}[,.]?\d{0,3}(\.\d+)?\s?[kK]?/) || [''])[0];
  const sk = SK.filter(s => new RegExp('\\b' + s.replace('.', '\\.') + '\\b', 'i').test(t));
  return { minyrs: yrs.length ? Math.min(...yrs) : null, f, sal, sk };
}
const out = [], stat = { boards: 0, postings: 0, titleHits: 0, dropped: 0 };
function keep(h, t) {
  const x = flags(t); h.x = x;
  if (x.f.some(v => ['CITIZEN/CLEARANCE', 'NO-OPT', 'NO-TIME-LIMIT', 'PHD'].includes(v)) || (x.minyrs !== null && x.minyrs >= 3)) { stat.dropped++; h.drop = true; }
  out.push(h);
}
async function gh(b) {
  const r = await fetch(`https://boards-api.greenhouse.io/v1/boards/${b}/jobs?content=true`); if (!r.ok) return; const d = await r.json(); stat.boards++;
  for (const j of d.jobs || []) { stat.postings++; const loc = j.location && j.location.name || ''; const p = new Date(j.first_published || j.updated_at);
    if (p >= since && INC.test(j.title) && !EXC.test(j.title) && US.test(loc) && !(NONUS.test(loc) && !/united states|usa|remote - us/i.test(loc))) { stat.titleHits++; keep({ s: 'gh', b, id: j.id, t: j.title, loc, d: (j.first_published || '').slice(0, 10), req: j.requisition_id, u: j.absolute_url }, strip(j.content)); } } }
async function ab(b) {
  const r = await fetch(`https://api.ashbyhq.com/posting-api/job-board/${encodeURIComponent(b)}?includeCompensation=true`); if (!r.ok) return; const d = await r.json(); stat.boards++;
  for (const j of d.jobs || []) { stat.postings++; const loc = [j.location, ...(j.secondaryLocations || []).map(x => x.location)].join('; '); const p = new Date(j.publishedAt);
    if (p >= since && INC.test(j.title) && !EXC.test(j.title) && (US.test(loc) || j.isRemote) && !(NONUS.test(loc) && !/united states|usa|\bus\b|remote/i.test(loc))) { stat.titleHits++; keep({ s: 'ab', b, id: j.id, t: j.title, loc: loc + (j.workplaceType ? ' [' + j.workplaceType + ']' : ''), d: (j.publishedAt || '').slice(0, 10), comp: j.compensation && j.compensation.compensationTierSummary, emp: j.employmentType, u: j.jobUrl }, j.descriptionPlain || ''); } } }
async function lv(b) {
  const r = await fetch(`https://api.lever.co/v0/postings/${b}?mode=json`); if (!r.ok) return; const d = await r.json(); if (!Array.isArray(d)) return; stat.boards++;
  for (const j of d) { stat.postings++; const loc = (j.categories || {}).location || '';
    if (j.createdAt >= since.getTime() && INC.test(j.text) && !EXC.test(j.text) && (US.test(loc) || j.country === 'US') && !(NONUS.test(loc) && j.country !== 'US')) { stat.titleHits++;
      const t = (j.descriptionPlain || '') + ' ' + (j.lists || []).map(l => l.text + ' ' + strip(l.content)).join(' ') + ' ' + (j.additionalPlain || '');
      keep({ s: 'lv', b, id: j.id, t: j.text, loc: loc + (j.workplaceType ? ' [' + j.workplaceType + ']' : ''), d: new Date(j.createdAt).toISOString().slice(0, 10), comp: j.salaryRange ? `${j.salaryRange.min}-${j.salaryRange.max} ${j.salaryRange.currency}` : '', emp: (j.categories || {}).commitment, u: j.hostedUrl }, t); } } }
const fn = { greenhouse: gh, ashby: ab, lever: lv }[KIND];
let i = 0; async function w() { while (i < TOKENS.length) { const t = TOKENS[i++]; try { await fn(t); } catch (e) {} } }
await Promise.all(Array.from({ length: 8 }, w));
const kept = out.filter(h => !h.drop);
window.__broad = out;
JSON.stringify({ stat, kept: kept.map(h => [h.s, h.b, h.id, h.t.slice(0, 70), h.loc.slice(0, 60), h.d, h.x.minyrs, h.x.f.join(','), (h.comp || h.x.sal || '').slice(0, 40), h.x.sk.slice(0, 8).join('/'), h.emp || '', h.req || '', h.u]) });
