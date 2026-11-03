import {gitBlobSha} from './import-core.mjs';
export const REPO='NextGenSalaya/Store-68-Daily-Dashboard';
export const SITE='https://nextgensalaya.github.io/Store-68-Daily-Dashboard/';
const BASE='https://api.github.com/repos/'+REPO;
export function githubClient(token,request=fetch) {
  if(!token?.trim()) throw Error('กรุณาใส่รหัสสิทธิ์ GitHub ก่อนเผยแพร่');
  return async (path,method='GET',body)=>{
    if(!path.startsWith('/')||path.includes('..')) throw Error('Invalid API path');
    let response;
    try {response=await request(BASE+path,{method,signal:AbortSignal.timeout(60000),cache:'no-store',redirect:'error',headers:{Accept:'application/vnd.github+json',Authorization:'Bearer '+token.trim(),'X-GitHub-Api-Version':'2022-11-28',...(body?{'Content-Type':'application/json'}:{})},...(body?{body:JSON.stringify(body)}:{})});}
    catch {throw Error('ติดต่อ GitHub ไม่สำเร็จ กรุณาตรวจสอบอินเทอร์เน็ต ยังไม่ควรกดเผยแพร่ซ้ำจนกว่าจะตรวจสอบสถานะ');}
    const result=await response.json();
    if(!response.ok) {
      const hints={401:'รหัสสิทธิ์ไม่ถูกต้องหรือหมดอายุ',403:'บัญชีไม่มีสิทธิ์เขียน หรือใช้งาน GitHub เกินขีดจำกัด',404:'ไม่พบ repository หรือรหัสสิทธิ์เข้าถึง repository นี้ไม่ได้',409:'มีข้อมูลใหม่บน GitHub กรุณาตรวจสอบก่อนเผยแพร่อีกครั้ง',422:'GitHub ไม่รับการอัปเดต อาจมีข้อมูลใหม่จากผู้ใช้อื่น กรุณาโหลดใหม่'};
      throw Error(hints[response.status]||'GitHub ตอบกลับข้อผิดพลาด '+response.status);
    }
    return result;
  };
}
export async function publishDatasets(api,datasets,summary,onProgress=()=>{}) {
  const head=(await api('/git/ref/heads/main')).object.sha;
  const commit=await api('/git/commits/'+head);
  const existing=await api('/git/trees/'+commit.tree.sha);
  if(existing.truncated) throw Error('อ่านรายการไฟล์ GitHub ได้ไม่ครบ');
  const byPath=new Map(existing.tree.map(e=>[e.path,e.sha]));
  const changed=[];
  for(const d of datasets) {
    if(!/^[a-z0-9_]+\.json$/.test(d.name)||d.name==='data-release.json') throw Error('ชื่อชุดข้อมูลไม่ถูกต้อง');
    const sha=await gitBlobSha(d.content);
    if(byPath.get(d.name)!==sha) changed.push(d);
  }
  if(!changed.length) return {unchanged:true,commit:head};
  const revision=crypto.randomUUID();
  const metadata={revision,publishedAt:new Date().toISOString(),groups:summary,files:datasets.map(d=>({name:d.name,rows:d.rows}))};
  changed.push({name:'data-release.json',content:JSON.stringify(metadata,null,2)+'\n'});
  // Upload blobs sequentially: a failed stage never advances the branch.
  const entries=[];
  for(let i=0;i<changed.length;i++) {
    onProgress(`กำลังส่งชุดข้อมูล ${i+1} / ${changed.length}`);
    const d=changed[i]; const blob=await api('/git/blobs','POST',{content:d.content,encoding:'utf-8'});
    entries.push({path:d.name,mode:'100644',type:'blob',sha:blob.sha});
  }
  if((await api('/git/ref/heads/main')).object.sha!==head) throw Error('มีการอัปเดตเว็บระหว่างประมวลผล จึงยังไม่เผยแพร่ กรุณาตรวจสอบแล้วลองใหม่');
  const tree=await api('/git/trees','POST',{base_tree:commit.tree.sha,tree:entries});
  const next=await api('/git/commits','POST',{message:'Update dashboard data from Excel ('+new Date().toISOString().slice(0,10)+')',tree:tree.sha,parents:[head]});
  try {await api('/git/refs/heads/main','PATCH',{sha:next.sha,force:false});}
  catch(error) {
    // A lost response may follow a successful write. Read before declaring failure.
    let actual; try {actual=(await api('/git/ref/heads/main')).object.sha;} catch {}
    if(actual!==next.sha) {const e=Error(error.message);e.commit=next.sha;e.uncertain=!actual;throw e;}
  }
  return {commit:next.sha,revision};
}
export async function waitForPublication(revision,{request=fetch,onProgress=()=>{},sleep=ms=>new Promise(r=>setTimeout(r,ms)),attempts=40}={}) {
  for(let i=0;i<attempts;i++) {
    onProgress('ส่งข้อมูลแล้ว กำลังรอ GitHub Pages อัปเดตเว็บ…');
    try {
      const r=await request(SITE+'data-release.json?check='+Date.now(),{cache:'no-store',signal:AbortSignal.timeout(10000)});
      if(r.ok && (await r.json()).revision===revision) return true;
    } catch {}
    await sleep(15000);
  }
  return false;
}
