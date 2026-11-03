import {planFiles,validateDatasets} from './import-core.mjs';
import {githubClient,publishDatasets,waitForPublication,REPO,SITE} from './publish.mjs';
const $=id=>document.getElementById(id);
let manifest,files=[],worker,result=null,busy=false,publishing=false,watching=false;
const nf=n=>n.toLocaleString('th-TH');
function notice(id,message,type=''){const e=$(id);e.hidden=false;e.textContent=message;e.className='notice '+type;}
function selected(){return manifest.groups.filter(g=>$('group-'+g.id).checked);}
function invalidate(){result=null;$('review').hidden=true;$('publish-section').hidden=true;$('confirm').checked=false;$('token').value='';$('publish-status').hidden=true;$('result-link').hidden=true;$('commit-link').hidden=true;}
function setBusy(value){busy=value;for(const el of document.querySelectorAll('#groups input,#files,#folder,#clear'))el.disabled=value;$('convert').disabled=value;$('cancel').hidden=!value||publishing;}
function showPlan(){
  if(!manifest)return;
  invalidate();const groups=selected(),plan=planFiles(files,groups);
  $('requirements').replaceChildren(...[...new Set(groups.flatMap(g=>g.patterns))].map(n=>{const li=document.createElement('li');li.textContent=n;return li;}));
  $('file-status').textContent=files.length?`เลือก ${nf(files.length)} ไฟล์ · ใช้ได้ ${nf(plan.assignments.length)} ไฟล์ (${(plan.assignments.reduce((s,a)=>s+a.file.size,0)/1048576).toFixed(1)} MB)`:'ยังไม่ได้เลือกไฟล์';
  const issues=[...plan.missing.map(n=>'ยังขาด: '+n),...plan.duplicates.map(n=>'พบไฟล์ซ้ำ ให้เลือกเพียงไฟล์เดียว: '+n)];
  if(!groups.length)issues.unshift('กรุณาเลือกอย่างน้อยหนึ่งกลุ่มข้อมูล');
  if(plan.ignored.length)issues.push(`ไม่นำเข้า ${nf(plan.ignored.length)} ไฟล์ที่ไม่ตรงกับหัวข้อที่เลือก`);
  if(plan.assignments.reduce((s,a)=>s+a.file.size,0)>250*1048576)issues.push('ไฟล์รวมเกิน 250 MB กรุณาแบ่งอัปเดตทีละกลุ่ม');
  $('issues').replaceChildren(...(files.length?issues:[]).map(s=>{const li=document.createElement('li');li.textContent=s;return li;}));
  $('convert').disabled=!groups.length||!files.length||!!plan.missing.length||!!plan.duplicates.length||plan.assignments.reduce((s,a)=>s+a.file.size,0)>250*1048576;
  return {groups,plan};
}
function showReview(datasets,summaries){
  result={datasets,summaries};$('summary').replaceChildren(...summaries.map(s=>{const tr=document.createElement('tr');for(const v of [s.name,nf(s.files),nf(s.rows),s.dates.length?s.dates.join(', '):'ไม่มีวันที่ในรายงาน']){const td=document.createElement('td');td.textContent=v;tr.append(td);}return tr;}));
  $('review-note').textContent=`พร้อมเผยแพร่ ${nf(datasets.length)} ชุดข้อมูล · หัวข้อที่ไม่ได้เลือกจะคงข้อมูลเดิมไว้ · วันที่แต่ละรายงานอาจไม่ตรงกัน กรุณาตรวจสอบก่อนเผยแพร่`;
  const empty=datasets.filter(d=>!d.rows).map(d=>d.name);$('empty-warning').hidden=!empty.length;$('empty-warning').textContent='ชุดข้อมูลไม่มีรายการ (จะเผยแพร่เป็นตารางว่าง): '+empty.join(', ');
  $('review').hidden=false;$('publish-section').hidden=false;$('publish-status').hidden=true;updatePublish();
}
function updatePublish(){$('publish').disabled=!result||!$('confirm').checked||!$('token').value.trim()||publishing||busy||watching;}
async function monitor(saved){
  watching=true;updatePublish();
  $('publish-section').hidden=false;$('commit-link').hidden=false;$('commit-link').href='https://github.com/'+REPO+'/commit/'+saved.commit;
  const live=await waitForPublication(saved.revision,{onProgress:s=>notice('publish-status',s)});
  if(live){notice('publish-status','เผยแพร่สำเร็จแล้ว เว็บสาธารณะยืนยันข้อมูลชุดใหม่เรียบร้อย','success');$('result-link').hidden=false;$('result-link').href=SITE+'?updated='+encodeURIComponent(saved.revision)+'#sales';sessionStorage.removeItem('store68-publication');}
  else notice('publish-status','ส่งข้อมูลเข้า GitHub แล้ว แต่ยังยืนยันว่าเว็บเปลี่ยนเป็นข้อมูลใหม่ไม่ได้ ตรวจสอบสถานะ GitHub หรือเปิดหน้านี้ใหม่เพื่อรอตรวจสอบต่อ','error');
  watching=false;updatePublish();
}
for(const id of ['files','folder'])$(id).addEventListener('change',()=>{if(busy||watching)return;files=[...$(id).files];showPlan();});
$('clear').onclick=()=>{if(busy||watching)return;files=[];$('files').value='';$('folder').value='';$('progress').hidden=true;showPlan();};
$('cancel').onclick=()=>{if(publishing)return;worker?.terminate();worker=null;setBusy(false);showPlan();notice('progress','ยกเลิกการแปลงแล้ว ข้อมูลบนเว็บยังไม่เปลี่ยน');};
$('convert').onclick=()=>{
  if(busy||watching)return;const {groups,plan}=showPlan();if($('convert').disabled)return;
  setBusy(true);notice('progress','กำลังเริ่มแปลงไฟล์…');
  worker=new Worker('import-worker.js');
  const fail=message=>{worker?.terminate();worker=null;setBusy(false);showPlan();notice('progress','ยังไม่ได้เผยแพร่ข้อมูล: '+message,'error');};
  worker.onerror=()=>fail('ไม่สามารถเริ่มเครื่องมืออ่าน Excel ได้ กรุณาตรวจสอบอินเทอร์เน็ตและลองใหม่');
  worker.onmessage=({data})=>{
    if(data.type==='progress')notice('progress',data.message);
    if(data.type==='error')fail(data.message);
    if(data.type==='done'){
      try{validateDatasets(data.datasets,groups,manifest.schemas);worker.terminate();worker=null;setBusy(false);showReview(data.datasets,data.summaries);notice('progress','แปลงไฟล์ครบแล้ว กรุณาตรวจสอบผลในขั้นตอนที่ 2','success');$('convert').disabled=false;}
      catch(e){fail(e.message);}
    }
  };
  worker.postMessage({groups,files:plan.assignments});
};
for(const id of ['confirm','token'])$(id).addEventListener('input',updatePublish);
$('publish').onclick=async()=>{
  if($('publish').disabled||!result)return;
  publishing=true;setBusy(true);updatePublish();$('confirm').disabled=true;$('token').disabled=true;
  let api=githubClient($('token').value);$('token').value='';
  try {
    notice('publish-status','กำลังตรวจสอบสิทธิ์และข้อมูลบน GitHub…');
    const saved=await publishDatasets(api,result.datasets,result.summaries,s=>notice('publish-status',s));
    api=null;
    if(saved.unchanged)notice('publish-status','ข้อมูลตรงกับชุดที่อยู่บน GitHub แล้ว จึงไม่ได้สร้างการอัปเดตซ้ำ','success');
    else {sessionStorage.setItem('store68-publication',JSON.stringify(saved));await monitor(saved);}
  } catch(e){notice('publish-status',e.message,'error');if(e.commit){$('commit-link').hidden=false;$('commit-link').href='https://github.com/'+REPO+'/commit/'+e.commit;}}
  finally{api=null;publishing=false;setBusy(false);$('confirm').disabled=false;$('confirm').checked=false;$('token').disabled=false;updatePublish();}
};
addEventListener('beforeunload',e=>{if(busy||publishing){e.preventDefault();e.returnValue='';}});
try {
  const response=await fetch('manifest.json',{cache:'no-cache'});if(!response.ok)throw Error('โหลดรายการไฟล์ไม่สำเร็จ');manifest=await response.json();
  $('groups').replaceChildren();const legend=document.createElement('legend');legend.textContent='หัวข้อที่ต้องการอัปเดต';$('groups').append(legend);
  for(const g of manifest.groups){const label=document.createElement('label');label.className='group';const input=document.createElement('input');input.type='checkbox';input.id='group-'+g.id;input.checked=true;input.onchange=showPlan;const span=document.createElement('span');span.textContent=g.name;const small=document.createElement('small');small.textContent=`${g.patterns.length} ไฟล์ต้นฉบับ · ${g.outputs.length} ชุดข้อมูล`;span.append(small);label.append(input,span);$('groups').append(label);}
  showPlan();const saved=sessionStorage.getItem('store68-publication');if(saved){const pending=JSON.parse(saved);if(/^[a-f0-9]{40}$/.test(pending.commit)&&typeof pending.revision==='string')await monitor(pending);}
} catch(e){notice('progress',e.message,'error');}
