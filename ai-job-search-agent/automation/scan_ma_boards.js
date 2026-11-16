// Massachusetts company-by-company scan. Run with Browser-pane javascript_tool in a tab on the matching origin:
//   greenhouse -> https://boards-api.greenhouse.io   ashby -> https://api.ashbyhq.com   lever -> https://api.lever.co
// Define first: KIND, TOKENS (array of [employerName, token] for that ATS, from ma_employers.json), SINCE ('YYYY-MM-DD').
// Returns per-employer stats (postings, MA-located postings, title-matched early-career openings in MA or remote-US,
// and those posted on/after SINCE) using the EMPLOYER'S OWN posting dates.
const MA=/\bMA\b|Massachusetts|Boston|Somerville|Waltham|Woburn|Billerica|Chelmsford|Framingham|Marlborough|Watertown|Needham|Natick|Devens|Westborough|Hopkinton|Braintree|Tewksbury|Charlestown|Quincy, M|Worcester, M|Lowell, M|Newton, M|Burlington, M|Lexington, M|Bedford, M|Wilmington, M|Andover, M|Medford|Cambridge(?!.*(UK|England|United Kingdom))/i;
const REM=/remote/i, NONUS=/(UK|England|United Kingdom|Canada|Toronto|India|Germany|Ireland|London|Europe|Poland|Warsaw|Spain|Mexico|Brazil|Japan|Singapore|Australia|Israel|Netherlands|France)/i;
const REL=/(machine learning|\bml\b|\bai\b|artificial intelligence|llm|genai|generative|applied scien|data scien|data engineer|data analyst|analytics engineer|\bnlp\b|deep learning|computer vision|research engineer|forward deployed|software engineer|software developer|developer|full.?stack|backend|new grad|early career|entry|junior|associate engineer|associate software|associate data|engineer i\b)/i;
const EX=/(senior|\bsr\b|sr\.|staff|principal|\blead\b|manager|director|head of|\bvp\b|intern\b|internship|co-?op|architect|counsel|recruit|sales|account exec|marketing|designer|product manager|\biii\b|\biv\b|\bii\b)/i;
const since=new Date(SINCE+'T00:00:00-05:00'), out=[];
const ok=(l,t,remote)=> (MA.test(l)||((remote||REM.test(l))&&!NONUS.test(l))) && REL.test(t) && !EX.test(t);
async function one(name,tok){ try{
  if(KIND==='greenhouse'){ const r=await fetch(`https://boards-api.greenhouse.io/v1/boards/${tok}/jobs`); if(!r.ok) return out.push({name,tok,error:r.status}); const js=(await r.json()).jobs||[];
    const rel=js.filter(j=>ok((j.location&&j.location.name)||'',j.title,false)); out.push({name,tok,total:js.length,ma:js.filter(j=>MA.test((j.location&&j.location.name)||'')).length,rel:rel.length,
      fresh:rel.filter(j=>new Date(j.first_published||j.updated_at)>=since).map(j=>({id:j.id,title:j.title,loc:j.location&&j.location.name,posted:j.first_published,req:j.requisition_id,url:j.absolute_url}))}); }
  else if(KIND==='ashby'){ const r=await fetch(`https://api.ashbyhq.com/posting-api/job-board/${encodeURIComponent(tok)}?includeCompensation=true`); if(!r.ok) return out.push({name,tok,error:r.status}); const js=(await r.json()).jobs||[];
    const L=j=>[j.location,...(j.secondaryLocations||[]).map(x=>x.location)].join('; '); const rel=js.filter(j=>ok(L(j),j.title,j.isRemote));
    out.push({name,tok,total:js.length,ma:js.filter(j=>MA.test(L(j))).length,rel:rel.length,fresh:rel.filter(j=>new Date(j.publishedAt)>=since).map(j=>({id:j.id,title:j.title,loc:L(j),posted:j.publishedAt,workplace:j.workplaceType,comp:j.compensation&&j.compensation.compensationTierSummary,url:j.jobUrl}))}); }
  else { const r=await fetch(`https://api.lever.co/v0/postings/${tok}?mode=json`); if(!r.ok) return out.push({name,tok,error:r.status}); const js=await r.json(); if(!Array.isArray(js)) return out.push({name,tok,error:'not a list'});
    const L=j=>((j.categories&&j.categories.allLocations)||[j.categories&&j.categories.location]).join('; '); const rel=js.filter(j=>ok(L(j),j.text,j.workplaceType==='remote'));
    out.push({name,tok,total:js.length,ma:js.filter(j=>MA.test(L(j))).length,rel:rel.length,fresh:rel.filter(j=>j.createdAt>=since.getTime()).map(j=>({id:j.id,title:j.text,loc:L(j),posted:new Date(j.createdAt).toISOString(),workplace:j.workplaceType,url:j.hostedUrl}))}); }
 }catch(e){ out.push({name,tok,error:String(e)}); } }
await Promise.all(TOKENS.map(([n,t])=>one(n,t)));
({checked:out.length, with_relevant:out.filter(o=>o.rel>0).length, fresh:out.flatMap(o=>(o.fresh||[]).map(f=>({employer:o.name,...f}))), stats:out.map(o=>[o.name,o.tok,o.total,o.ma,o.rel,(o.fresh||[]).length,o.error||null])});
