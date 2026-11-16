// Run with the Browser pane javascript_tool, in a tab whose origin matches the feed:
//   Greenhouse -> https://boards-api.greenhouse.io   Ashby -> https://api.ashbyhq.com   Lever -> https://api.lever.co
// Before running, set: KIND ("greenhouse"|"ashby"|"lever"), TOKENS (array from sources.json), SINCE (ISO date, today-14 ET),
// INC / EXC (regex strings from sources.json). Returns {stat, hits:[...]} with the employer's own posting date.
const inc = new RegExp(INC, 'i'), exc = new RegExp(EXC, 'i'), since = new Date(SINCE + 'T00:00:00-05:00');
const hits = [], stat = {boards_ok: 0, boards_missing: 0, postings: 0};
async function gh(t){ const r = await fetch(`https://boards-api.greenhouse.io/v1/boards/${t}/jobs`); if(!r.ok) return stat.boards_missing++; const d = await r.json(); stat.boards_ok++;
  for (const j of d.jobs||[]) { stat.postings++; const p = new Date(j.first_published||j.updated_at); if (p>=since && inc.test(j.title) && !exc.test(j.title)) hits.push({src:'greenhouse', board:t, id:j.id, title:j.title, loc:j.location&&j.location.name, posted:j.first_published, req:j.requisition_id, url:j.absolute_url}); } }
async function ab(t){ const r = await fetch(`https://api.ashbyhq.com/posting-api/job-board/${encodeURIComponent(t)}?includeCompensation=true`); if(!r.ok) return stat.boards_missing++; const d = await r.json(); stat.boards_ok++;
  for (const j of d.jobs||[]) { stat.postings++; const p = new Date(j.publishedAt); if (p>=since && inc.test(j.title) && !exc.test(j.title)) hits.push({src:'ashby', board:t, id:j.id, title:j.title, loc:[j.location,...(j.secondaryLocations||[]).map(x=>x.location)].join('; '), remote:j.isRemote, workplace:j.workplaceType, posted:j.publishedAt, comp:j.compensation&&j.compensation.compensationTierSummary, url:j.jobUrl}); } }
async function lv(t){ const r = await fetch(`https://api.lever.co/v0/postings/${t}?mode=json`); if(!r.ok) return stat.boards_missing++; const d = await r.json(); if(!Array.isArray(d)) return stat.boards_missing++; stat.boards_ok++;
  for (const j of d) { stat.postings++; if (j.createdAt>=since.getTime() && inc.test(j.text) && !exc.test(j.text)) hits.push({src:'lever', board:t, id:j.id, title:j.text, loc:j.categories&&j.categories.location, workplace:j.workplaceType, country:j.country, posted:new Date(j.createdAt).toISOString(), url:j.hostedUrl}); } }
const fn = {greenhouse: gh, ashby: ab, lever: lv}[KIND];
await Promise.all(TOKENS.map(t => fn(t).catch(() => stat.boards_missing++)));
({stat, hits});
