/* Excel files remain in this worker's memory. Only converted datasets leave it. */
const CDN='https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
self.onmessage=async ({data})=>{
  try {
    const progress=(message)=>self.postMessage({type:'progress',message});
    progress('กำลังเตรียมเครื่องมืออ่าน Excel (ครั้งแรกอาจใช้เวลาสักครู่)');
    importScripts(CDN+'pyodide.js');
    const py=await loadPyodide({indexURL:CDN,stdout:line=>{
      if(line.startsWith('reading ')) progress('กำลังประมวลผล '+line.split('/').pop());
      else if(/^[a-z0-9_]+: \d+ rows/.test(line)) progress('แปลงแล้ว '+line.replace(': ',': ').replace(' rows',' รายการ'));
    },stderr:()=>{}});
    await py.loadPackage('micropip');
    await py.runPythonAsync("import micropip\nawait micropip.install(['openpyxl==3.1.5', 'xlrd==2.0.1', 'pyxlsb==1.0.10'])");
    py.FS.mkdirTree('/input'); py.FS.chdir('/input');
    for(const script of [...data.groups.map(g=>g.script),'unmapped.py']) {
      const response=await fetch(new URL('../'+script,self.location.href),{cache:'no-cache'});
      if(!response.ok) throw Error('โหลดเครื่องมือแปลงไฟล์ไม่ได้: '+script);
      const source=(await response.text()).replace('sys.stdout.reconfigure(encoding="utf-8")','if hasattr(sys.stdout, "reconfigure"): sys.stdout.reconfigure(encoding="utf-8")');
      py.FS.writeFile('/input/'+script,source,{encoding:'utf8'});
    }
    for(const entry of data.files) {
      progress('กำลังอ่าน '+entry.file.name);
      py.FS.writeFile('/input/'+entry.name,new Uint8Array(await entry.file.arrayBuffer()));
      // The existing Over Stock extractor uses the source file's modification date.
      py.FS.utime('/input/'+entry.name,entry.file.lastModified,entry.file.lastModified);
    }
    const datasets=[], summaries=[];
    for(const group of data.groups) {
      progress('กำลังแปลง '+group.name+' — ไฟล์ขนาดใหญ่อาจใช้เวลาหลายนาที');
      py.globals.set('script_name',group.script);
      py.globals.set('output_dir','/input/out_'+group.id);
      const date=await py.runPythonAsync(`
import sys, runpy, gc
if '/input' not in sys.path: sys.path.insert(0, '/input')
sys.argv = [script_name, output_dir]
namespace = runpy.run_path('/input/' + script_name, run_name='__main__')
data_date = str(namespace.get('DATA_DATE') or '')
del namespace
gc.collect()
data_date
`);
      let count=0; const dates=new Set(date?[date]:[]);
      for(const name of group.outputs) {
        const file=name==='t5_drop_cust.json'?'/input/data_private/'+name:'/input/out_'+group.id+'/'+name;
        let content;
        try {content=py.FS.readFile(file,{encoding:'utf8'});} catch {throw Error('ไม่พบผลลัพธ์ '+name+' กรุณาตรวจสอบไฟล์ต้นฉบับ');}
        // Parse here as well: NaN/Infinity or truncated JSON cannot be published.
        const rows=JSON.parse(content); count+=rows.length;
        if(name.endsWith('_meta.json')||name==='kpi.json') for(const r of rows) {
          const value=r.value??r.period;
          if(typeof value==='string' && /^\d{4}-\d{2}-\d{2}(?:$|\s)/.test(value)) dates.add(value.slice(0,10));
        }
        datasets.push({name,content});
      }
      if(!count) throw Error('ไม่พบข้อมูลในกลุ่ม '+group.name);
      summaries.push({id:group.id,name:group.name,files:group.outputs.length,rows:count,dates:[...dates].sort()});
      progress('แปลง '+group.name+' เสร็จแล้ว');
    }
    self.postMessage({type:'done',datasets,summaries});
  } catch(error) {self.postMessage({type:'error',message:String(error.message||error)});}
};
