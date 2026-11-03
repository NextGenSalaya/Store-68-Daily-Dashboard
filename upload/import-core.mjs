export function wildcard(pattern) {
  return new RegExp('^'+pattern.split('*').map(s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('.*')+'$','i');
}
export function planFiles(files, groups) {
  const usable=[...files].filter(f=>!f.name.startsWith('~$'));
  const assignments=[], missing=[], duplicates=[];
  for(const pattern of new Set(groups.flatMap(g=>g.patterns))) {
    const matches=usable.filter(f=>wildcard(pattern).test(f.name));
    if(!matches.length) missing.push(pattern);
    else if(matches.length>1) duplicates.push(pattern);
    else assignments.push({name:pattern.replaceAll('*',''),file:matches[0]});
  }
  const recognized=new Set(assignments.map(a=>a.file));
  return {assignments,missing,duplicates,ignored:usable.filter(f=>!recognized.has(f))};
}
export function validateDatasets(datasets, groups, schemas) {
  const expected=new Set(groups.flatMap(g=>g.outputs));
  const found=new Set();
  for(const d of datasets) {
    if(!expected.has(d.name)||found.has(d.name)||! /^[a-z0-9_]+\.json$/.test(d.name)) throw Error('พบชุดข้อมูลที่ไม่อยู่ในรายการที่เลือก: '+d.name);
    const rows=JSON.parse(d.content);
    if(!Array.isArray(rows)) throw Error('รูปแบบข้อมูลไม่ถูกต้อง: '+d.name);
    const keys=schemas[d.name]||[];
    if(!rows.every(r=>r && typeof r==='object' && !Array.isArray(r) && keys.every(k=>Object.hasOwn(r,k)))) throw Error('คอลัมน์ข้อมูลไม่ตรงกับ Dashboard: '+d.name);
    d.rows=rows.length;
    found.add(d.name);
  }
  const missing=[...expected].filter(n=>!found.has(n));
  if(missing.length) throw Error('แปลงข้อมูลไม่ครบ: '+missing.join(', '));
  return datasets;
}
export async function gitBlobSha(content) {
  const bytes=new TextEncoder().encode(content);
  const header=new TextEncoder().encode(`blob ${bytes.length}\0`);
  const input=new Uint8Array(header.length+bytes.length); input.set(header);input.set(bytes,header.length);
  return [...new Uint8Array(await crypto.subtle.digest('SHA-1',input))].map(v=>v.toString(16).padStart(2,'0')).join('');
}
