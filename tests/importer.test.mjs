import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {planFiles,validateDatasets,gitBlobSha} from '../upload/import-core.mjs';
import {githubClient,publishDatasets,waitForPublication} from '../upload/publish.mjs';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const manifest=JSON.parse(fs.readFileSync(path.join(root,'upload/manifest.json')));
test('recognizes all 31 workbook names, including Thai and mixed-case extensions',()=>{
  const names=[...new Set(manifest.groups.flatMap(g=>g.patterns))].map(p=>({name:p.replaceAll('*','')}));
  const plan=planFiles(names,manifest.groups);
  assert.equal(plan.assignments.length,31);assert.equal(plan.missing.length,0);assert.equal(plan.duplicates.length,0);
  assert.equal(planFiles(names.slice(1),manifest.groups).missing.length,1);
  assert.equal(planFiles([...names,names[0]],manifest.groups).duplicates.length,1);
  assert.equal(planFiles([...names,{name:'~$'+names[0].name}],manifest.groups).duplicates.length,0);
});
test('all currently published datasets meet the import schema; reject missing/unexpected/bad payloads',()=>{
  const all=manifest.groups.flatMap(g=>g.outputs).map(name=>({name,content:fs.readFileSync(path.join(root,name),'utf8')}));
  assert.equal(validateDatasets(all,manifest.groups,manifest.schemas).length,96);
  assert.throws(()=>validateDatasets(all.slice(1),manifest.groups,manifest.schemas),/ไม่ครบ/);
  assert.throws(()=>validateDatasets([{name:'index.html',content:'[]'}],manifest.groups,manifest.schemas),/ไม่อยู่/);
  assert.throws(()=>validateDatasets([{name:'sales.json',content:'[{"bad":1}]'}],manifest.groups,manifest.schemas),/คอลัมน์/);
  assert.throws(()=>validateDatasets([{name:'sales.json',content:'[NaN]'}],manifest.groups,manifest.schemas));
});
test('UTF-8 blob SHA matches Git (including byte length)',async()=>{
  assert.equal(await gitBlobSha('hello\n'),'ce013625030ba8dba906f756967f9e9ca394464a');
});
function fakeApi({same=false,failBlob=false,concurrent=false,lostReply=false}={}){
  const calls=[];let reads=0;const api=async(p,m='GET',b)=>{
    calls.push({p,m,b});
    if(p==='/git/ref/heads/main')return {object:{sha:lostReply&&calls.some(c=>c.m==='PATCH')?'next':concurrent&&++reads>1?'other':'head'}};
    if(p==='/git/commits/head')return {tree:{sha:'base'}};
    if(p==='/git/trees/base')return {tree:same?[{path:'sales.json',sha:await gitBlobSha('[1]')}]:[]};
    if(p==='/git/blobs'){if(failBlob)throw Error('upload failed');return {sha:'blob'};}
    if(p==='/git/trees')return {sha:'tree'};
    if(p==='/git/commits')return {sha:'next'};
    if(p==='/git/refs/heads/main'){if(lostReply)throw Error('response lost');return {object:{sha:'next'}};}
    throw Error('Unexpected API call '+p);
  };return {api,calls};
}
const data=()=>[{name:'sales.json',content:'[1]',rows:1}];
test('publish updates one commit without force, retaining base tree',async()=>{
  const {api,calls}=fakeApi();const saved=await publishDatasets(api,data(),[]);
  assert.equal(saved.commit,'next');assert.ok(saved.revision);
  assert.equal(calls.filter(c=>c.m==='PATCH').length,1);
  assert.deepEqual(calls.find(c=>c.m==='PATCH').b,{sha:'next',force:false});
  assert.equal(calls.find(c=>c.p==='/git/trees').b.base_tree,'base');
  assert.deepEqual(calls.find(c=>c.p==='/git/commits').b.parents,['head']);
});
test('unchanged datasets produce no writes',async()=>{
  const {api,calls}=fakeApi({same:true});assert.equal((await publishDatasets(api,data(),[])).unchanged,true);
  assert.ok(calls.every(c=>c.m==='GET'));
});
test('failed upload cannot advance branch',async()=>{
  const {api,calls}=fakeApi({failBlob:true});await assert.rejects(publishDatasets(api,data(),[]));
  assert.ok(!calls.some(c=>c.m==='PATCH'));
});
test('concurrent update cannot overwrite newer head',async()=>{
  const {api,calls}=fakeApi({concurrent:true});await assert.rejects(publishDatasets(api,data(),[]),/ระหว่าง/);
  assert.ok(!calls.some(c=>c.m==='PATCH'));
});
test('lost final response is checked by reading head, never replayed',async()=>{
  const {api,calls}=fakeApi({lostReply:true});assert.equal((await publishDatasets(api,data(),[])).commit,'next');
  assert.equal(calls.filter(c=>c.m==='PATCH').length,1);
});
test('publication success requires matching served revision, not merely HTTP 200',async()=>{
  const stale=async()=>({ok:true,json:async()=>({revision:'old'})});
  assert.equal(await waitForPublication('new',{request:stale,attempts:2,sleep:async()=>{}}),false);
  const current=async()=>({ok:true,json:async()=>({revision:'new'})});
  assert.equal(await waitForPublication('new',{request:current,attempts:1}),true);
});
test('missing and unauthorized credentials never report success',async()=>{
  assert.throws(()=>githubClient(''),/รหัสสิทธิ์/);
  const api=githubClient('test-token',async()=>({ok:false,status:401,json:async()=>({message:'Bad credentials'})}));
  await assert.rejects(api('/git/ref/heads/main'),/หมดอายุ/);
});
