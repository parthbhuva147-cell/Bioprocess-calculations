const state={
  files:[], mappings:{}, processStart:0, processEnd:null, results:null, carbon:null, plotIds:[]
};

const signalDefs=[
  ["dcw","DCW","g/L","Dry cell weight"],
  ["od","OD600","OD","Optical density"],
  ["glucose","Residual glucose","g/L","Substrate concentration"],
  ["product1","Product 1","mg/L or g/L","Main product / resveratrol"],
  ["product2","Product 2","mg/L or g/L","Second product / piceatannol"],
  ["samplevol","Sample volume","mL","Withdrawal per sample"],
  ["feed","Feed","mL or mL/h","Substrate feed"],
  ["do","DO","% saturation","Dissolved oxygen"],
  ["ph","pH","pH","Reactor pH"],
  ["temp","Temperature","°C","Reactor temperature"],
  ["stirrer","Agitation","rpm","Stirrer speed"],
  ["air","Air inlet","ccm","Air flow"],
  ["o2in","O₂ inlet","ccm","Pure oxygen"],
  ["n2in","N₂ inlet","ccm","Nitrogen flow"],
  ["co2in","CO₂ inlet","ccm","CO₂ addition"],
  ["offco2","Off-gas CO₂","vol.%","Exhaust CO₂"],
  ["offo2","Off-gas O₂","vol.%","Exhaust O₂"],
  ["our","OUR","mol/L/h","Oxygen uptake rate"],
  ["cer","CER","mol/L/h","CO₂ evolution rate"],
  ["rq","RQ","—","Respiratory quotient"]
];

const slotDefs=[
  ["reactorFiles","Bioreactor","reactorCount"],
  ["biomassFiles","OD/DCW","biomassCount"],
  ["hplcFiles","HPLC","hplcCount"],
  ["samplingFiles","Sampling","samplingCount"],
  ["additionalFiles","Other","additionalCount"]
];

function esc(s){return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));}
function fmt(x,d=3){return Number.isFinite(x)?Number(x).toLocaleString(undefined,{maximumFractionDigits:d}):"—";}
function num(id){return Number(document.getElementById(id).value);}
function excelCol(i){let n=i+1,s="";while(n>0){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
function cleanHeader(s){return String(s??"").replace(/^\uFEFF/,"").replace(/\s+/g," ").trim();}
function norm(s){return cleanHeader(s).toLowerCase().replace(/[_\-]+/g," ").replace(/[()[\]{}]/g," ").replace(/\s+/g," ").trim();}
function parseNum(v){
  if(typeof v==="number"&&Number.isFinite(v))return v;
  if(v===null||v===undefined||v==="")return NaN;
  let s=String(v).trim().replace(/\s/g,"");
  if(s.includes(",")&&!s.includes("."))s=s.replace(",",".");
  const x=Number(s);return Number.isFinite(x)?x:NaN;
}
function parseClockHours(v){
  const m=String(v??"").trim().match(/^(\d{1,3}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?$/);
  return m?Number(m[1])+Number(m[2])/60+Number(m[3]||0)/3600:NaN;
}
function parseDate(v){
  if(v instanceof Date&&!isNaN(v))return v;
  if(typeof v==="number"&&v>20000&&v<100000)return new Date((v-25569)*86400*1000);
  const s=String(v??"").trim();if(!s)return null;
  let m=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?$/);
  if(m){
    let y=Number(m[3]);if(y<100)y+=2000;
    const d=new Date(y,Number(m[2])-1,Number(m[1]),Number(m[4]),Number(m[5]),Number(m[6]||0));
    if(!isNaN(d))return d;
  }
  const d=new Date(s);return !isNaN(d)?d:null;
}
function extractInterval(h){
  const m=String(h??"").match(/interval\s*=\s*(\d+(?:\.\d+)?)/i);
  return m&&Number(m[1])>0?Number(m[1]):null;
}
function splitCSV(line,d){
  const out=[];let cur="",q=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(ch==='"'){if(q&&line[i+1]==='"'){cur+='"';i++;}else q=!q;}
    else if(ch===d&&!q){out.push(cur);cur="";}else cur+=ch;
  }
  out.push(cur);return out;
}
function chooseDelimiter(lines){
  let best=",",score=-1;
  for(const d of [",",";","\t"]){
    const s=Math.max(...lines.slice(0,15).map(l=>splitCSV(l,d).length));
    if(s>score){score=s;best=d;}
  }
  return best;
}
function headerScore(cells){
  const j=cells.map(norm).join(" | ");let s=cells.filter(x=>String(x).trim()).length;
  ["time","timestamp","dcw","od","glucose","substrate","resveratrol","piceatannol","stirrer","temp","ph","do","bluevary","our","cer","rq"].forEach(k=>{if(j.includes(k))s+=7;});
  return s;
}
function parseCSVText(text){
  const lines=text.replace(/\r/g,"").split("\n").filter(l=>l.trim()!=="");
  if(lines.length<2)throw new Error("File contains fewer than two data rows.");
  const d=chooseDelimiter(lines);let hr=0,best=-1;
  for(let i=0;i<Math.min(50,lines.length);i++){const s=headerScore(splitCSV(lines[i],d));if(s>best){best=s;hr=i;}}
  return {headers:splitCSV(lines[hr],d).map(cleanHeader),rows:lines.slice(hr+1).map(l=>splitCSV(l,d)).filter(r=>r.some(v=>String(v).trim()!==""))};
}
function parseWorkbook(buf){
  if(typeof XLSX==="undefined")throw new Error("Excel library did not load.");
  const wb=XLSX.read(buf,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const m=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});
  let hr=0,best=-1;
  for(let i=0;i<Math.min(50,m.length);i++){const s=headerScore(m[i]);if(s>best){best=s;hr=i;}}
  return {headers:m[hr].map(cleanHeader),rows:m.slice(hr+1).filter(r=>r.some(v=>String(v).trim()!==""))};
}
function numericCount(file,i){
  let n=0;for(let r=0;r<Math.min(file.rows.length,5000);r++)if(Number.isFinite(parseNum(file.rows[r][i])))n++;
  return n;
}
function classifyHeader(h,i,rows){
  const x=norm(h);
  if(x.includes("timestamp")||x==="time"||x==="time h"||x==="time hours"||x.includes("process time")||x.includes("elapsed time")||x==="hours"||x==="hour")return"time";
  if(/\bdcw\b/.test(x)||x.includes("dry cell"))return"dcw";
  if(/\bod600\b/.test(x)||x==="od"||x.includes("optical density"))return"od";
  if(x.includes("glucose")||x==="substrate"||x.includes("residual substrate"))return"glucose";
  if(x.includes("resveratrol"))return"product1";
  if(x.includes("piceatannol"))return"product2";
  if(x.includes("sample")&&x.includes("volume"))return"samplevol";
  if(x.includes("dissolved oxygen")||/^do\b/.test(x))return"do";
  if(x==="ph"||x.startsWith("ph "))return"ph";
  if((x.startsWith("temp")||x.includes("temperature"))&&!x.includes("bluevary"))return"temp";
  if(x.includes("stirrer")||x.includes("agitation")||x.includes("rpm"))return"stirrer";
  if(x.includes("bluevary")&&x.includes("our"))return"our";
  if(x.includes("bluevary")&&x.includes("cer"))return"cer";
  if(x.includes("bluevary")&&x.includes("rq"))return"rq";
  if((x.includes("subs a")||x.includes("feed"))&&!x.includes("feedback"))return"feed";
  if(x.startsWith("air ")||x==="air ccm"||x.includes("air flow"))return"air";
  if(x.startsWith("o2 ")||x==="o2 ccm")return"o2in";
  if(x.startsWith("n2 ")||x==="n2 ccm")return"n2in";
  if(x.startsWith("co2 ")||x==="co2 ccm")return"co2in";
  if(x.includes("offgas")&&x.includes("co2"))return"offco2";
  if(x.includes("offgas")&&x.includes("o2"))return"offo2";
  if(x.includes("bluevary")&&x.includes("vol")){
    const vals=rows.slice(0,Math.min(rows.length,10000)).map(r=>parseNum(r[i])).filter(Number.isFinite).filter(v=>Math.abs(v)>1e-12).sort((a,b)=>a-b);
    if(vals.length){
      const med=vals[Math.floor(vals.length/2)];
      if(med<10)return"offco2";
      if(med>10&&med<25)return"offo2";
    }
  }
  return null;
}
function detectTimeIndex(file){
  for(let i=0;i<file.headers.length;i++)if(classifyHeader(file.headers[i],i,file.rows)==="time")return i;
  return -1;
}
function inferRole(headers,current){
  const j=headers.map(norm).join(" | ");
  if(current&&current!=="Other")return current;
  if(j.includes("stirrer")||j.includes("bluevary")||j.includes("dissolved oxygen")||j.includes("air ccm"))return"Bioreactor";
  if(j.includes("resveratrol")||j.includes("piceatannol")||j.includes("hplc"))return"HPLC";
  if(j.includes("dcw")||j.includes("od600")||j.includes("optical density"))return"OD/DCW";
  if(j.includes("glucose")||j.includes("sample volume"))return"Sampling";
  return current||"Other";
}
function timeLooksNumericHours(header){
  const x=norm(header);return x==="time"||x==="time h"||x==="hours"||x==="hour"||x.includes("process time")||x.includes("elapsed time");
}
function normalizeTime(file){
  const i=file.timeIndex;
  file.elapsed=[];file.duration=NaN;file.timeWarning="";file.absoluteDates=null;file.timeConfidence="low";
  if(i<0){file.timeMethod="No time column";file.timeWarning="Choose the correct time column.";return;}
  const vals=file.rows.map(r=>r[i]);
  const forced=Number(file.forcedInterval);
  const embedded=extractInterval(file.headers[i]);
  const interval=(forced>0)?forced:embedded;

  const dates=vals.map(parseDate),validDates=dates.filter(Boolean);
  if(interval){
    let inconsistent=false;
    if(validDates.length>=3){
      const ms=validDates.map(d=>d.getTime());
      const span=(Math.max(...ms)-Math.min(...ms))/3600000;
      const expected=(file.rows.length-1)*interval/3600;
      if(expected>0&&(span<=0||span/expected>1.25||span/expected<.75))inconsistent=true;
    }
    if(forced>0||inconsistent||validDates.length<Math.min(3,Math.ceil(file.rows.length*.2))){
      file.elapsed=file.rows.map((_,r)=>r*interval/3600+Number(file.offset||0));
      file.duration=(file.rows.length-1)*interval/3600;
      file.timeMethod=(forced>0?"Forced":"Header")+" interval "+interval+" s";
      file.timeConfidence="high";
      if(inconsistent)file.timeWarning="Spreadsheet date conversion conflicts with the logger interval; interval-based time is used.";
      return;
    }
  }

  if(validDates.length>=Math.min(3,Math.ceil(file.rows.length*.2))){
    const first=validDates[0].getTime();
    file.elapsed=dates.map(d=>d?(d.getTime()-first)/3600000+Number(file.offset||0):NaN);
    file.absoluteDates=dates;
    const f=file.elapsed.filter(Number.isFinite);
    file.duration=f.length?Math.max(...f)-Math.min(...f):NaN;
    file.timeMethod="Timestamps";
    file.timeConfidence="high";return;
  }

  const clocks=vals.map(parseClockHours),vc=clocks.filter(Number.isFinite);
  if(vc.length>=Math.min(3,Math.ceil(file.rows.length*.2))){
    const first=vc[0];
    file.elapsed=clocks.map(v=>Number.isFinite(v)?v-first+Number(file.offset||0):NaN);
    const f=file.elapsed.filter(Number.isFinite);
    file.duration=f.length?Math.max(...f)-Math.min(...f):NaN;
    file.timeMethod="Clock time + offset";file.timeConfidence="medium";return;
  }

  const nums=vals.map(parseNum),vn=nums.filter(Number.isFinite);
  if(vn.length>=Math.min(3,Math.ceil(file.rows.length*.2))){
    const factor=file.numericUnit==="seconds"?1/3600:file.numericUnit==="minutes"?1/60:1;
    // Numeric process-time columns are treated as actual process time, not reset to zero.
    file.elapsed=nums.map(v=>Number.isFinite(v)?v*factor+Number(file.offset||0):NaN);
    const f=file.elapsed.filter(Number.isFinite);
    file.duration=f.length?Math.max(...f)-Math.min(...f):NaN;
    file.timeMethod="Numeric process time ("+file.numericUnit+")";
    file.timeConfidence="medium";return;
  }
  file.timeMethod="Unresolved";file.timeWarning="Time values could not be interpreted.";
}
function alignTrustedTimestampFiles(){
  const ref=state.files.find(f=>f.role==="Bioreactor"&&f.timeMethod==="Timestamps"&&f.absoluteDates);
  if(!ref)return;
  const refStart=ref.absoluteDates.find(Boolean)?.getTime();if(!refStart)return;
  state.files.forEach(f=>{
    if(f===ref||f.timeMethod!=="Timestamps"||!f.absoluteDates)return;
    f.elapsed=f.absoluteDates.map(d=>d?(d.getTime()-refStart)/3600000+Number(f.offset||0):NaN);
    const ff=f.elapsed.filter(Number.isFinite);f.duration=ff.length?Math.max(...ff)-Math.min(...ff):NaN;
    f.timeMethod="Timestamps aligned to reactor start";
  });
}
function updateFileCounts(){
  slotDefs.forEach(([id,role,countId])=>{
    const n=document.getElementById(id).files.length;
    document.getElementById(countId).textContent=n?n+" file"+(n===1?"":"s")+" selected":"No files selected";
  });
}
slotDefs.forEach(([id])=>document.addEventListener("DOMContentLoaded",()=>document.getElementById(id).addEventListener("change",updateFileCounts)));

async function loadAllSelectedFiles(){
  const selected=[];
  slotDefs.forEach(([id,role])=>[...document.getElementById(id).files].forEach(f=>selected.push({file:f,role})));
  if(!selected.length){alert("Select at least one file.");return;}
  state.files=[];state.mappings={};state.results=null;state.carbon=null;
  document.getElementById("loadStatus").textContent="Reading "+selected.length+" file(s)…";
  try{
    for(let k=0;k<selected.length;k++){
      const item=selected[k],f=item.file;
      let p=/\.(xlsx|xls)$/i.test(f.name)?parseWorkbook(await f.arrayBuffer()):parseCSVText(await f.text());
      const obj={
        id:"f_"+Date.now()+"_"+k,name:f.name,headers:p.headers,rows:p.rows,
        role:inferRole(p.headers,item.role),timeIndex:-1,offset:0,forcedInterval:"",
        numericUnit:"hours",elapsed:[],duration:NaN,timeMethod:"",timeWarning:""
      };
      obj.timeIndex=detectTimeIndex(obj);
      normalizeTime(obj);
      state.files.push(obj);
    }
    alignTrustedTimestampFiles();
    autoMap();
    renderFiles();
    renderDetected();
    renderAdvancedMapping();
    updateReferenceTime();
    applyProcessTime(true);
    document.getElementById("loadStatus").textContent=state.files.length+" dataset(s) loaded successfully.";
  }catch(e){
    console.error(e);document.getElementById("loadStatus").textContent="Error: "+e.message;
    alert(e.message);
  }
}
function clearWorkspace(){
  state.files=[];state.mappings={};state.results=null;state.carbon=null;
  slotDefs.forEach(([id])=>document.getElementById(id).value="");
  updateFileCounts();
  document.querySelector("#filesTable tbody").innerHTML="";
  document.getElementById("loadedFileBadge").textContent="0 files";
  document.getElementById("detectedSignals").innerHTML='<span class="placeholder">Load files to detect measurements.</span>';
  document.getElementById("advancedMapping").innerHTML="";
  document.getElementById("referenceDuration").textContent="—";
  document.getElementById("referenceDurationLong").textContent="Load a bioreactor file.";
  document.getElementById("referenceMethod").textContent="No time method selected.";
  document.getElementById("kpiGrid").innerHTML='<div class="kpi empty"><strong>—</strong><span>Upload and calculate to generate results.</span></div>';
  document.getElementById("figureGallery").innerHTML='<div class="card empty-figure">No figures generated yet.</div>';
}
function renderFiles(){
  const tb=document.querySelector("#filesTable tbody");tb.innerHTML="";
  state.files.forEach(f=>{
    const timeOpts=['<option value="-1">None</option>'].concat(f.headers.map((h,i)=>`<option value="${i}" ${i===f.timeIndex?"selected":""}>${esc(h||"Unnamed")} [${excelCol(i)}]</option>`)).join("");
    const tr=document.createElement("tr");
    tr.innerHTML=`<td><b>${esc(f.name)}</b>${f.timeWarning?`<div class="tiny-warning">${esc(f.timeWarning)}</div>`:""}</td>
      <td><select onchange="changeRole('${f.id}',this.value)">${["Bioreactor","OD/DCW","HPLC","Sampling","Feed/Additions","Other"].map(r=>`<option ${r===f.role?"selected":""}>${r}</option>`).join("")}</select></td>
      <td>${f.rows.length.toLocaleString()}</td>
      <td><select onchange="changeTimeColumn('${f.id}',this.value)">${timeOpts}</select></td>
      <td>${esc(f.timeMethod)}</td>
      <td>${fmt(f.duration,2)} h</td>
      <td><input type="number" step="any" value="${f.offset}" onchange="changeOffset('${f.id}',this.value)"></td>
      <td><input type="number" step="any" value="${f.forcedInterval}" placeholder="Auto" onchange="changeInterval('${f.id}',this.value)"></td>
      <td><button class="btn secondary small" onclick="removeFile('${f.id}')">Remove</button></td>`;
    tb.appendChild(tr);
  });
  document.getElementById("loadedFileBadge").textContent=state.files.length+" file"+(state.files.length===1?"":"s");
}
function changeRole(id,v){const f=state.files.find(x=>x.id===id);if(f){f.role=v;updateReferenceTime();}}
function changeTimeColumn(id,v){const f=state.files.find(x=>x.id===id);if(f){f.timeIndex=Number(v);normalizeTime(f);alignTrustedTimestampFiles();renderFiles();updateReferenceTime();renderAdvancedMapping();}}
function changeOffset(id,v){const f=state.files.find(x=>x.id===id);if(f){f.offset=Number(v)||0;normalizeTime(f);alignTrustedTimestampFiles();renderFiles();updateReferenceTime();}}
function changeInterval(id,v){const f=state.files.find(x=>x.id===id);if(f){f.forcedInterval=v;normalizeTime(f);renderFiles();updateReferenceTime();}}
function removeFile(id){state.files=state.files.filter(f=>f.id!==id);autoMap();renderFiles();renderDetected();renderAdvancedMapping();updateReferenceTime();}
function updateReferenceTime(){
  const ref=state.files.find(f=>f.role==="Bioreactor"&&Number.isFinite(f.duration))||state.files.find(f=>Number.isFinite(f.duration));
  if(!ref)return;
  const ff=ref.elapsed.filter(Number.isFinite);if(!ff.length)return;
  const start=Math.min(...ff),end=Math.max(...ff),d=end-start;
  document.getElementById("referenceDuration").textContent=fmt(d,2)+" h";
  document.getElementById("referenceDurationLong").textContent=Math.floor(d/24)+" d "+Math.floor(d%24)+" h "+Math.round((d%1)*60)+" min • "+ref.name;
  document.getElementById("referenceMethod").innerHTML="<b>"+esc(ref.timeMethod)+"</b>"+(ref.timeWarning?" — "+esc(ref.timeWarning):"");
  if(!document.getElementById("processEnd").value)document.getElementById("processEnd").placeholder=fmt(end,3);
}
function applyProcessTime(silent=false){
  const ref=state.files.find(f=>f.role==="Bioreactor"&&f.elapsed.some(Number.isFinite))||state.files.find(f=>f.elapsed.some(Number.isFinite));
  if(!ref){if(!silent)alert("Load data with a time column first.");return;}
  const ff=ref.elapsed.filter(Number.isFinite);const autoEnd=Math.max(...ff);
  let s=Number(document.getElementById("processStart").value),e=Number(document.getElementById("processEnd").value);
  if(!Number.isFinite(s))s=0;if(!Number.isFinite(e))e=autoEnd;
  if(e<=s){if(!silent)alert("Process end must be greater than process start.");return;}
  state.processStart=s;state.processEnd=e;
  document.getElementById("processWindowStatus").innerHTML=`Applied: <b>${fmt(s,3)} → ${fmt(e,3)} h</b> (${fmt(e-s,3)} h).`;
}
function autoMap(){
  state.mappings={};
  state.files.forEach(f=>f.headers.forEach((h,i)=>{
    if(i===f.timeIndex||numericCount(f,i)<2)return;
    const k=classifyHeader(h,i,f.rows);
    if(k&&!state.mappings[k])state.mappings[k]=f.id+"::"+i;
  }));
}
function allNumericColumns(){
  const a=[];state.files.forEach(f=>f.headers.forEach((h,i)=>{
    if(i!==f.timeIndex&&numericCount(f,i)>=2)a.push({key:f.id+"::"+i,file:f,index:i,label:f.name+" → "+(h||"Unnamed")+" ["+excelCol(i)+"]"});
  }));return a;
}
function resolve(k){
  const r=state.mappings[k];if(!r)return null;
  const [fid,idx]=r.split("::"),file=state.files.find(f=>f.id===fid);
  return file?{file,index:Number(idx),header:file.headers[Number(idx)]}:null;
}
function autoDetectedList(){
  return signalDefs.filter(([k])=>resolve(k));
}
function renderDetected(){
  const list=autoDetectedList(),wrap=document.getElementById("detectedSignals");
  document.getElementById("detectedBadge").textContent=list.length+" measurement"+(list.length===1?"":"s");
  document.getElementById("detectedBadge").className="badge"+(list.length?"":" neutral");
  wrap.innerHTML=list.length?list.map(([k,n,u])=>`<span class="signal-chip"><b>${esc(n)}</b> · ${esc(u)}</span>`).join(""):'<span class="placeholder">No known measurement columns were detected automatically.</span>';
}
function renderAdvancedMapping(){
  const opts=allNumericColumns(),wrap=document.getElementById("advancedMapping");wrap.innerHTML="";
  signalDefs.forEach(([k,n,u,hint])=>{
    const div=document.createElement("div");div.className="map-item";
    let o='<option value="">Not mapped</option>';
    opts.forEach(x=>o+=`<option value="${x.key}" ${state.mappings[k]===x.key?"selected":""}>${esc(x.label)}</option>`);
    div.innerHTML=`<label>${esc(n)}<small>${esc(u)} · ${esc(hint)}</small><select onchange="setMap('${k}',this.value)">${o}</select></label>`;
    wrap.appendChild(div);
  });
}
function setMap(k,v){state.mappings[k]=v||null;renderDetected();}
function series(k,windowed=true){
  const r=resolve(k);if(!r)return[];
  const out=[];
  for(let i=0;i<r.file.rows.length;i++){
    const t=r.file.elapsed[i],y=parseNum(r.file.rows[i][r.index]);
    if(!Number.isFinite(t)||!Number.isFinite(y))continue;
    if(windowed&&(t<state.processStart||t>state.processEnd))continue;
    out.push({t,y});
  }
  return out.sort((a,b)=>a.t-b.t);
}
function interp(s,t){
  if(!s.length)return NaN;if(t<=s[0].t)return s[0].y;if(t>=s[s.length-1].t)return s[s.length-1].y;
  let lo=0,hi=s.length-1;while(hi-lo>1){const m=(lo+hi)>>1;if(s[m].t<=t)lo=m;else hi=m;}
  const a=s[lo],b=s[hi];return b.t===a.t?a.y:a.y+(b.y-a.y)*(t-a.t)/(b.t-a.t);
}
function stat(s){
  if(!s.length)return null;let mn=Infinity,mx=-Infinity,sum=0;
  s.forEach(p=>{sum+=p.y;mn=Math.min(mn,p.y);mx=Math.max(mx,p.y);});
  return{first:s[0].y,last:s[s.length-1].y,min:mn,max:mx,mean:sum/s.length,n:s.length};
}
function regression(p){
  const n=p.length;if(n<2)return null;let sx=0,sy=0,sxx=0,sxy=0,syy=0;
  p.forEach(q=>{sx+=q.t;sy+=q.y;sxx+=q.t*q.t;sxy+=q.t*q.y;syy+=q.y*q.y;});
  const den=n*sxx-sx*sx;if(!den)return null;
  const m=(n*sxy-sx*sy)/den,b=(sy-m*sx)/n,tot=syy-sy*sy/n;let res=0;
  p.forEach(q=>res+=(q.y-(b+m*q.t))**2);
  return{slope:m,r2:tot>0?1-res/tot:1};
}
function estimateMu(){
  let s=series("dcw");let source="DCW";
  if(s.length<3){s=series("od");source="OD600";}
  const pts=s.filter(p=>p.y>0).map(p=>({t:p.t,y:Math.log(p.y)}));
  if(pts.length<3)return null;
  let best=null;const maxWin=Math.min(6,pts.length);
  for(let size=3;size<=maxWin;size++)for(let i=0;i<=pts.length-size;i++){
    const r=regression(pts.slice(i,i+size));
    if(r&&r.slope>0&&r.r2>=.95&&(!best||r.slope>best.mu))best={mu:r.slope,r2:r.r2,t1:pts[i].t,t2:pts[i+size-1].t,source};
  }
  return best;
}
function feedVolumeMl(){
  const s=series("feed");if(s.length<2)return 0;
  if(document.getElementById("feedMode").value==="cumulative")return Math.max(0,s[s.length-1].y-s[0].y);
  let v=0;for(let i=1;i<s.length;i++){const dt=s[i].t-s[i-1].t;if(dt>0)v+=.5*(s[i].y+s[i-1].y)*dt;}
  return v;
}
function sampleTimes(){
  const sv=series("samplevol");if(sv.length)return[...new Set(sv.map(p=>p.t.toFixed(6)))].map(Number).sort((a,b)=>a-b);
  const set=new Set();["dcw","od","glucose","product1","product2"].forEach(k=>series(k).forEach(p=>set.add(p.t.toFixed(6))));
  return[...set].map(Number).sort((a,b)=>a-b);
}
function sampleDetails(){
  const times=sampleTimes(),sv=series("samplevol"),def=Number(document.getElementById("defaultSampleVolume").value)||0;
  const gl=series("glucose"),dcw=series("dcw"),p1=series("product1"),p2=series("product2");
  const productScale=document.getElementById("productUnit").value==="mgL"?1/1000:1;
  const out=[];times.forEach(t=>{
    const ml=sv.length?interp(sv,t):def;if(!(ml>0))return;
    const L=ml/1000;
    out.push({t,ml,L,
      glucose:Number.isFinite(interp(gl,t))?interp(gl,t)*L:0,
      biomass:Number.isFinite(interp(dcw,t))?interp(dcw,t)*L:0,
      product1:Number.isFinite(interp(p1,t))?interp(p1,t)*productScale*L:0,
      product2:Number.isFinite(interp(p2,t))?interp(p2,t)*productScale*L:0
    });
  });return out;
}
function finalVolume(){
  const override=Number(document.getElementById("finalVolumeOverride").value);
  if(override>0)return override;
  const init=Number(document.getElementById("initialVolume").value)||0;
  const feed=feedVolumeMl()/1000;
  const samp=sampleDetails().reduce((a,s)=>a+s.L,0);
  return Math.max(0,init+feed-samp);
}
function massMetrics(){
  const initV=Number(document.getElementById("initialVolume").value),feedConc=Number(document.getElementById("feedGlucoseConc").value)||0;
  if(!(initV>0))return null;
  const gl=series("glucose"),dcw=series("dcw"),p1=series("product1"),p2=series("product2");
  const samples=sampleDetails(),feedL=feedVolumeMl()/1000,finalV=finalVolume(),manualG=Number(document.getElementById("manualGlucoseG").value)||0;
  const glOverride=Number(document.getElementById("initialGlucoseOverride").value);
  const initialGlc=Number.isFinite(glOverride)&&glOverride>=0?glOverride:(gl.length?gl[0].y:NaN);
  const finalGlc=gl.length?gl[gl.length-1].y:NaN;
  const substrateIn=(Number.isFinite(initialGlc)?initialGlc*initV:0)+feedL*feedConc+manualG;
  const substrateSample=samples.reduce((a,s)=>a+s.glucose,0);
  const substrateFinal=Number.isFinite(finalGlc)?finalGlc*finalV:NaN;
  const consumed=Number.isFinite(substrateFinal)?substrateIn-substrateFinal-substrateSample:NaN;

  const x0=dcw.length?dcw[0].y:NaN,xf=dcw.length?dcw[dcw.length-1].y:NaN;
  const biomassSample=samples.reduce((a,s)=>a+s.biomass,0);
  const biomassFormed=Number.isFinite(x0)&&Number.isFinite(xf)?xf*finalV+biomassSample-x0*initV:NaN;

  const scale=document.getElementById("productUnit").value==="mgL"?1/1000:1;
  const p10=p1.length?p1[0].y*scale:NaN,p1f=p1.length?p1[p1.length-1].y*scale:NaN;
  const p20=p2.length?p2[0].y*scale:NaN,p2f=p2.length?p2[p2.length-1].y*scale:NaN;
  const p1sample=samples.reduce((a,s)=>a+s.product1,0),p2sample=samples.reduce((a,s)=>a+s.product2,0);
  const p1formed=Number.isFinite(p10)&&Number.isFinite(p1f)?p1f*finalV+p1sample-p10*initV:NaN;
  const p2formed=Number.isFinite(p20)&&Number.isFinite(p2f)?p2f*finalV+p2sample-p20*initV:NaN;

  return{initV,feedL,finalV,samples,initialGlc,finalGlc,substrateIn,substrateSample,substrateFinal,consumed,
    biomassFormed,p1formed,p2formed,p1f,p2f};
}
function buildIntegratedResults(){
  applyProcessTime();
  const dur=state.processEnd-state.processStart,mu=estimateMu(),m=massMetrics();
  const s={dur,mu,m,do:stat(series("do")),ph:stat(series("ph")),temp:stat(series("temp")),stirrer:stat(series("stirrer")),
    od:stat(series("od")),dcw:stat(series("dcw")),glucose:stat(series("glucose")),p1:stat(series("product1")),p2:stat(series("product2")),
    feedMl:feedVolumeMl(),cer:stat(series("cer")),rq:stat(series("rq"))};
  if(mu)s.doubling=Math.log(2)/mu.mu;
  if(m&&m.consumed>0){
    s.substrateEfficiency=m.substrateIn>0?m.consumed/m.substrateIn*100:NaN;
    s.yxs=Number.isFinite(m.biomassFormed)?m.biomassFormed/m.consumed:NaN;
    s.yp1=Number.isFinite(m.p1formed)?m.p1formed/m.consumed:NaN;
    s.yp2=Number.isFinite(m.p2formed)?m.p2formed/m.consumed:NaN;
  }
  if(s.dcw&&dur>0)s.biomassProductivity=(s.dcw.last-s.dcw.first)/dur;
  const scale=document.getElementById("productUnit").value==="mgL"?1/1000:1;
  if(s.p1&&dur>0)s.p1Productivity=(s.p1.last-s.p1.first)*scale/dur;
  if(s.p2&&dur>0)s.p2Productivity=(s.p2.last-s.p2.first)*scale/dur;
  state.results=s;renderResults();
}
function renderResults(){
  const s=state.results,items=[],note=[];
  const add=(v,l,n="")=>items.push([v,l,n]);
  add(fmt(s.dur,2)+" h","Process duration");
  if(s.mu){add(fmt(s.mu.mu,4)+" h⁻¹","Specific growth rate μmax",s.mu.source+" fit, R² "+fmt(s.mu.r2,3));add(fmt(s.doubling,3)+" h","Doubling time","ln(2)/μmax");}
  else add("Not available","Growth rate / doubling","Map DCW or OD600 with ≥3 positive points");
  if(s.dcw)add(fmt(s.dcw.max,3)+" g/L","Maximum DCW");
  if(s.od)add(fmt(s.od.max,2),"Maximum OD600");
  if(Number.isFinite(s.substrateEfficiency))add(fmt(s.substrateEfficiency,1)+" %","Substrate utilization","Mass-based: input − final − samples");
  else add("Not available","Substrate utilization","Requires residual glucose + volume/feed assumptions");
  if(Number.isFinite(s.yxs))add(fmt(s.yxs,3)+" g/g","Biomass yield Yx/s","Mass-based, sampling corrected");
  if(Number.isFinite(s.yp1))add(fmt(s.yp1,4)+" g/g","Product 1 yield Yp/s","Mass-based, sampling corrected");
  if(Number.isFinite(s.yp2))add(fmt(s.yp2,4)+" g/g","Product 2 yield Yp/s","Mass-based, sampling corrected");
  if(s.p1)add(fmt(s.p1.last,3)+" "+(document.getElementById("productUnit").value==="mgL"?"mg/L":"g/L"),"Final Product 1 titre");
  if(s.p2)add(fmt(s.p2.last,3)+" "+(document.getElementById("productUnit").value==="mgL"?"mg/L":"g/L"),"Final Product 2 titre");
  if(Number.isFinite(s.p1Productivity))add(fmt(s.p1Productivity,5)+" g/L/h","Product 1 productivity");
  if(Number.isFinite(s.p2Productivity))add(fmt(s.p2Productivity,5)+" g/L/h","Product 2 productivity");
  if(Number.isFinite(s.biomassProductivity))add(fmt(s.biomassProductivity,4)+" g/L/h","Biomass productivity");
  if(s.glucose)add(fmt(s.glucose.last,3)+" g/L","Final residual glucose");
  if(s.feedMl>0)add(fmt(s.feedMl,2)+" mL","Total feed delivered");
  if(s.do)add(fmt(s.do.min,2)+" %","Minimum DO");
  if(s.ph)add(fmt(s.ph.mean,3),"Mean pH");
  if(s.temp)add(fmt(s.temp.mean,2)+" °C","Mean temperature");
  if(s.stirrer)add(fmt(s.stirrer.max,0)+" rpm","Maximum agitation");
  document.getElementById("kpiGrid").innerHTML=items.map(([v,l,n])=>`<div class="kpi"><strong>${v}</strong><span>${l}</span>${n?`<div class="subnote">${esc(n)}</div>`:""}</div>`).join("");
  if(s.m){
    note.push(`Estimated final liquid volume: ${fmt(s.m.finalV,3)} L from initial volume + mapped feed − sample withdrawals, unless overridden.`);
    if(Number.isFinite(s.m.consumed))note.push(`Estimated substrate consumed: ${fmt(s.m.consumed,3)} g.`);
  }
  note.push("Mass-based yields are preferable to simple concentration differences when feed and sampling change reactor volume.");
  note.push("μmax is estimated from the best local log-linear OD/DCW window with R² ≥ 0.95; inspect the biological plausibility of the selected region.");
  document.getElementById("calculationNotes").innerHTML=note.map(x=>`<p>• ${esc(x)}</p>`).join("");
}
function addCarbonEvent(){
  const tr=document.createElement("tr");
  tr.innerHTML='<td><input class="ce-time" type="number" step="any"></td><td><input class="ce-desc" type="text" placeholder="e.g. glucose bolus"></td><td><select class="ce-dir"><option value="plus">+ input</option><option value="minus">− removal</option></select></td><td><input class="ce-c" type="number" step="any"></td><td><button class="btn secondary small" onclick="this.closest(\'tr\').remove()">Remove</button></td>';
  document.querySelector("#carbonEvents tbody").appendChild(tr);
}
function carbonEvents(){
  return[...document.querySelectorAll("#carbonEvents tbody tr")].map(tr=>({time:Number(tr.querySelector(".ce-time").value),desc:tr.querySelector(".ce-desc").value||"Manual event",dir:tr.querySelector(".ce-dir").value,c:Number(tr.querySelector(".ce-c").value)})).filter(e=>e.c>=0&&Number.isFinite(e.c));
}
function reactorVolumeAt(t,m){
  let v=m.initV;
  const fs=series("feed");
  if(fs.length){
    if(document.getElementById("feedMode").value==="cumulative"){
      const a=interp(fs,state.processStart),b=interp(fs,t);if(Number.isFinite(a)&&Number.isFinite(b))v+=Math.max(0,b-a)/1000;
    }else{
      let add=0;for(let i=1;i<fs.length;i++){const a=fs[i-1],b=fs[i];if(a.t>t)break;const end=Math.min(b.t,t),dt=end-a.t;if(dt>0)add+=.5*(a.y+b.y)*dt/1000;if(b.t>=t)break;}v+=add;
    }
  }
  m.samples.forEach(s=>{if(s.t<=t)v-=s.L;});return Math.max(v,0);
}
function integrateCER(m){
  const s=series("cer").filter(p=>Math.abs(p.y)>1e-12);if(s.length<2)return null;
  let mol=0;for(let i=1;i<s.length;i++){const a=s[i-1],b=s[i],dt=b.t-a.t;if(dt<=0)continue;mol+=.5*(a.y*reactorVolumeAt(a.t,m)+b.y*reactorVolumeAt(b.t,m))*dt;}
  return{mol,gC:mol*12.011,method:"CER integration"};
}
function offgasCO2(m){
  const co2=series("offco2");if(co2.length<2)return null;
  const air=series("air"),o2=series("o2in"),n2=series("n2in"),c2=series("co2in"),offo2=series("offo2");
  if(!air.length&&!o2.length&&!n2.length&&!c2.length)return null;
  const vm=Number(document.getElementById("molarVolume").value)||24.465,airCO2=(Number(document.getElementById("airCO2").value)||0)/100,airN2=(Number(document.getElementById("airN2").value)||79)/100;
  let mol=0,corrected=false;
  function rate(p){
    const t=p.t,A=Number.isFinite(interp(air,t))?interp(air,t):0,O=Number.isFinite(interp(o2,t))?interp(o2,t):0,N=Number.isFinite(interp(n2,t))?interp(n2,t):0,C=Number.isFinite(interp(c2,t))?interp(c2,t):0;
    const fin=(A+O+N+C)/1000;if(fin<=0)return NaN;
    const yco2=p.y/100,cin=A/1000*airCO2+C/1000;
    let fout=fin,yo2=interp(offo2,t);
    if(Number.isFinite(yo2)&&yo2>0&&yo2<100){
      const yn2=1-yo2/100-yco2,nin=A/1000*airN2+N/1000;
      if(yn2>.2&&nin>0){fout=nin/yn2;corrected=true;}
    }
    return (yco2*fout-cin)/vm;
  }
  for(let i=1;i<co2.length;i++){const a=co2[i-1],b=co2[i],dt=(b.t-a.t)*60,r1=rate(a),r2=rate(b);if(dt>0&&Number.isFinite(r1)&&Number.isFinite(r2))mol+=.5*(r1+r2)*dt;}
  return Math.abs(mol)>0?{mol,gC:mol*12.011,method:corrected?"Off-gas CO₂ + N₂ flow correction":"Off-gas CO₂ + inlet-flow approximation"}:null;
}
function gasCO2In(){
  const air=series("air"),c2=series("co2in"),timeline=air.length?air:c2;if(timeline.length<2)return 0;
  const vm=Number(document.getElementById("molarVolume").value)||24.465,airCO2=(Number(document.getElementById("airCO2").value)||0)/100;let mol=0;
  const r=t=>{const A=Number.isFinite(interp(air,t))?interp(air,t):0,C=Number.isFinite(interp(c2,t))?interp(c2,t):0;return(A/1000*airCO2+C/1000)/vm;};
  for(let i=1;i<timeline.length;i++){const a=timeline[i-1],b=timeline[i],dt=(b.t-a.t)*60;if(dt>0)mol+=.5*(r(a.t)+r(b.t))*dt;}
  return mol*12.011;
}
function calculateCarbonBalance(){
  applyProcessTime();if(!state.results)buildIntegratedResults();
  const m=massMetrics();if(!m){alert("Enter a valid initial volume.");return;}
  const gcf=Number(document.getElementById("glucoseCF").value)||.40002,bcf=Number(document.getElementById("biomassCF").value)||.488,p1cf=Number(document.getElementById("product1CF").value)||0,p2cf=Number(document.getElementById("product2CF").value)||0;
  const initialGlucoseC=(Number.isFinite(m.initialGlc)?m.initialGlc*m.initV:0)*gcf;
  const feedGlucose=(m.feedL*(Number(document.getElementById("feedGlucoseConc").value)||0)+(Number(document.getElementById("manualGlucoseG").value)||0));
  const feedC=feedGlucose*gcf;
  const sampleC=m.samples.reduce((a,s)=>a+s.glucose*gcf+s.biomass*bcf+s.product1*p1cf+s.product2*p2cf,0);
  const finalGlucoseC=Number.isFinite(m.finalGlc)?m.finalGlc*m.finalV*gcf:0;
  const dcw=series("dcw"),p1=series("product1"),p2=series("product2"),scale=document.getElementById("productUnit").value==="mgL"?1/1000:1;
  const finalBiomassC=dcw.length?dcw[dcw.length-1].y*m.finalV*bcf:0;
  const finalP1C=p1.length?p1[p1.length-1].y*scale*m.finalV*p1cf:0;
  const finalP2C=p2.length?p2[p2.length-1].y*scale*m.finalV*p2cf:0;
  const gasInC=gasCO2In();
  let co2=null,method=document.getElementById("co2Method").value;
  if(method==="auto"||method==="cer")co2=integrateCER(m);
  if(!co2&&(method==="auto"||method==="offgas"))co2=offgasCO2(m);
  const netCO2=co2?Math.max(0,co2.gC):0,grossOut=co2?netCO2+gasInC:0;
  const ev=carbonEvents(),plus=ev.filter(e=>e.dir==="plus").reduce((a,e)=>a+e.c,0),minus=ev.filter(e=>e.dir==="minus").reduce((a,e)=>a+e.c,0);
  const input=initialGlucoseC+feedC+gasInC+plus;
  const output=sampleC+grossOut+minus;
  const finalInventory=finalGlucoseC+finalBiomassC+finalP1C+finalP2C;
  const accounted=output+finalInventory,unaccounted=input-accounted,recovery=input>0?accounted/input*100:NaN;
  const entries=[
    ["Initial glucose carbon","+",initialGlucoseC],["Feed/manual glucose carbon","+",feedC],["Gas-phase CO₂ entering","+",gasInC],
    ...ev.filter(e=>e.dir==="plus").map(e=>[e.desc,"+",e.c]),
    ["Sampling carbon removed","−",sampleC],["CO₂ carbon in exhaust","−",grossOut],
    ...ev.filter(e=>e.dir==="minus").map(e=>[e.desc,"−",e.c])
  ];
  state.carbon={entries,input,output,finalInventory,unaccounted,recovery,co2,gasInC,grossOut,finalGlucoseC,finalBiomassC,finalP1C,finalP2C};
  renderCarbon();
}
function renderCarbon(){
  const c=state.carbon;
  document.getElementById("carbonLedger").innerHTML=`<table class="ledger-table"><thead><tr><th>Term</th><th>Direction</th><th>g C</th></tr></thead><tbody>${c.entries.map(e=>`<tr><td>${esc(e[0])}</td><td class="${e[1]==="+"?"plus":"minus"}">${e[1]}</td><td>${fmt(e[2],4)}</td></tr>`).join("")}</tbody></table>`;
  document.getElementById("carbonSummary").innerHTML=`<div class="balance-grid">
    <div><strong>${fmt(c.input,3)} g C</strong><span>Total carbon input</span></div>
    <div><strong>${fmt(c.output,3)} g C</strong><span>Known carbon leaving</span></div>
    <div><strong>${fmt(c.finalInventory,3)} g C</strong><span>Measured final inventory</span></div>
    <div><strong>${fmt(c.unaccounted,3)} g C</strong><span>Unaccounted carbon</span></div>
    <div><strong>${fmt(c.recovery,1)} %</strong><span>Carbon recovery</span></div>
    <div><strong>${c.co2?fmt(c.co2.gC,3)+" g C":"Not available"}</strong><span>Net biological CO₂ carbon</span></div>
  </div>
  <div class="warning-box"><b>CO₂ method:</b> ${c.co2?esc(c.co2.method):"No usable non-zero CER or off-gas CO₂ + flow signal was available."}<br>
  <b>Final measured carbon:</b> glucose ${fmt(c.finalGlucoseC,3)}, biomass ${fmt(c.finalBiomassC,3)}, Product 1 ${fmt(c.finalP1C,3)}, Product 2 ${fmt(c.finalP2C,3)} g C.</div>`;
}
function phaseDecorations(){
  if(!document.getElementById("showPhases").checked)return{shapes:[],annotations:[]};
  const b1=Number(document.getElementById("phase1End").value),b2=Number(document.getElementById("phase2End").value),s=state.processStart,e=state.processEnd;
  const bounds=[s];if(Number.isFinite(b1)&&b1>s&&b1<e)bounds.push(b1);if(Number.isFinite(b2)&&b2>bounds[bounds.length-1]&&b2<e)bounds.push(b2);bounds.push(e);
  const names=[document.getElementById("phase1Name").value||"Phase 1",document.getElementById("phase2Name").value||"Phase 2",document.getElementById("phase3Name").value||"Phase 3"];
  const shapes=[],annotations=[];
  for(let i=0;i<bounds.length-1;i++){
    annotations.push({x:(bounds[i]+bounds[i+1])/2,y:1.04,xref:"x",yref:"paper",text:names[i]||("Phase "+(i+1)),showarrow:false,font:{size:14,color:"#6b7280"}});
    if(i<bounds.length-2)shapes.push({type:"line",x0:bounds[i+1],x1:bounds[i+1],y0:0,y1:1,yref:"paper",line:{width:1.5,dash:"dash",color:"#9ca3af"}});
  }
  return{shapes,annotations};
}
function plotLayout(title,ytitle){
  const fs=Number(document.getElementById("figFont").value)||24,d=phaseDecorations();
  return{title:{text:title,x:.5,xanchor:"center",font:{size:fs+4}},paper_bgcolor:"#fff",plot_bgcolor:"#fff",
    font:{family:"Arial, Helvetica, sans-serif",size:fs,color:"#111827"},margin:{l:110,r:40,t:100,b:90},
    xaxis:{title:{text:"Process time (h)",standoff:15},range:[state.processStart,state.processEnd],showline:true,mirror:true,linecolor:"#111827",linewidth:2,ticks:"outside",ticklen:7,tickwidth:2,gridcolor:"#e5e7eb",zeroline:false},
    yaxis:{title:{text:ytitle,standoff:18},showline:true,mirror:true,linecolor:"#111827",linewidth:2,ticks:"outside",ticklen:7,tickwidth:2,gridcolor:"#e5e7eb",zeroline:false},
    legend:{orientation:"h",x:.5,xanchor:"center",y:1.12},hovermode:"x unified",shapes:d.shapes,annotations:d.annotations};
}
function trace(k,name,markers=false){
  const s=series(k),lw=Number(document.getElementById("figLine").value)||3;
  return{x:s.map(p=>p.t),y:s.map(p=>p.y),type:"scatter",mode:markers?"lines+markers":"lines",name,line:{width:lw},marker:{size:9},connectgaps:false};
}
function addFigure(id,title,traces,layout){
  const g=document.getElementById("figureGallery"),card=document.createElement("div");card.className="figure-card";
  card.innerHTML=`<div class="figure-head"><h3>${esc(title)}</h3><div class="figure-actions"><button class="btn secondary small" onclick="downloadFigure('${id}','svg')">SVG</button><button class="btn secondary small" onclick="downloadFigure('${id}','png')">High-res PNG</button></div></div><div id="${id}" class="figure-plot"></div>`;
  g.appendChild(card);Plotly.newPlot(id,traces,layout,{responsive:true,displaylogo:false,modeBarButtonsToRemove:["lasso2d","select2d"]});state.plotIds.push(id);
}
function generateFigures(){
  if(typeof Plotly==="undefined"){alert("Plotly did not load.");return;}
  applyProcessTime();const g=document.getElementById("figureGallery");g.innerHTML="";state.plotIds=[];let n=0;
  const add=(k,id,title,y,markers=false)=>{if(series(k).length>1){addFigure(id,title,[trace(k,title,markers)],plotLayout(title,y));n++;}};
  add("do","fig_do","Dissolved oxygen","DO (% saturation)");
  add("ph","fig_ph","pH profile","pH");
  add("temp","fig_temp","Temperature profile","Temperature (°C)");
  add("stirrer","fig_stirrer","Agitation profile","Stirrer speed (rpm)");
  add("offco2","fig_co2","Off-gas CO₂","CO₂ (vol.%)");
  add("offo2","fig_o2","Off-gas O₂","O₂ (vol.%)");
  if(series("our").length>1||series("cer").length>1){
    const tr=[];if(series("our").length>1)tr.push(trace("our","OUR"));if(series("cer").length>1)tr.push(trace("cer","CER"));
    addFigure("fig_resp","Respiration rates",tr,plotLayout("Respiration rates","Rate (mol/L/h)"));n++;
  }
  add("rq","fig_rq","Respiratory quotient","RQ");
  add("dcw","fig_dcw","Dry cell weight","DCW (g/L)",true);
  add("od","fig_od","Optical density","OD600",true);
  add("glucose","fig_glucose","Residual glucose","Glucose (g/L)",true);
  const unit=document.getElementById("productUnit").value==="mgL"?"mg/L":"g/L";
  add("product1","fig_p1","Product 1 formation","Product 1 ("+unit+")",true);
  add("product2","fig_p2","Product 2 formation","Product 2 ("+unit+")",true);
  if(state.carbon){
    const c=state.carbon,names=c.entries.map(e=>e[0]).concat(["Final measured inventory","Unaccounted"]),vals=c.entries.map(e=>e[1]==="+"?e[2]:-e[2]).concat([-c.finalInventory,-c.unaccounted]);
    const l=plotLayout("Carbon balance","Carbon contribution (g C)");l.xaxis={title:{text:"Carbon term"},tickangle:-28,showline:true,mirror:true,linecolor:"#111827",linewidth:2};l.shapes=[];l.annotations=[];
    addFigure("fig_carbon","Carbon balance",[{type:"bar",x:names,y:vals,name:"Carbon"}],l);n++;
  }
  if(!n)g.innerHTML='<div class="card empty-figure">Map at least one variable with two or more data points.</div>';
}
function downloadFigure(id,format){
  const w=Number(document.getElementById("figWidth").value)||1800,h=Number(document.getElementById("figHeight").value)||1100;
  Plotly.downloadImage(id,{format,filename:id+(format==="png"?"_highres":""),width:w,height:h,scale:format==="png"?2:1});
}
function quickGrowth(){const x1=num("qX1"),x2=num("qX2"),t1=num("qT1"),t2=num("qT2"),dt=t2-t1;if(!(x1>0&&x2>0&&dt>0))return document.getElementById("qGrowthResult").innerHTML="Check inputs.";const mu=(Math.log(x2)-Math.log(x1))/dt,td=mu>0?Math.log(2)/mu:NaN;document.getElementById("qGrowthResult").innerHTML=`μ = ${fmt(mu,4)} h⁻¹<br>Doubling time = ${fmt(td,3)} h`;}
function quickYxs(){const ds=num("qYS1")-num("qYS2");document.getElementById("qYxsResult").textContent=ds>0?"Yx/s = "+fmt((num("qYX2")-num("qYX1"))/ds,4)+" g/g":"Substrate consumed must be > 0.";}
function quickYps(){const ds=num("qPS1")-num("qPS2");document.getElementById("qYpsResult").textContent=ds>0?"Yp/s = "+fmt((num("qP2")-num("qP1"))/ds,4)+" g/g":"Substrate consumed must be > 0.";}
function quickSubstrateEfficiency(){const s0=num("qS0"),sf=num("qSf");document.getElementById("qSubstrateResult").textContent=s0>0?"Substrate consumption = "+fmt((s0-sf)/s0*100,2)+" %":"Initial substrate must be > 0.";}
function quickProductivity(){const dt=num("qProdT2")-num("qProdT1");document.getElementById("qProductivityResult").textContent=dt>0?"Qp = "+fmt((num("qProd2")-num("qProd1"))/dt,4)+" g/L/h":"Time interval must be > 0.";}
function quickBiomassProductivity(){const dt=num("qBioT2")-num("qBioT1");document.getElementById("qBiomassResult").textContent=dt>0?"Qx = "+fmt((num("qBio2")-num("qBio1"))/dt,4)+" g/L/h":"Time interval must be > 0.";}
function quickGlucoseCarbon(){document.getElementById("qGlucoseResult").textContent="Carbon = "+fmt(num("qGlucose")*.40002,3)+" g C";}
function quickCO2Carbon(){const v=num("qCO2"),vm=num("qVm");document.getElementById("qCO2Result").textContent=vm>0?"Carbon = "+fmt(v/vm*12.011,3)+" g C":"Vm must be > 0.";}

document.addEventListener("DOMContentLoaded",()=>{updateFileCounts();quickGrowth();quickYxs();quickYps();quickSubstrateEfficiency();quickProductivity();quickBiomassProductivity();quickGlucoseCarbon();quickCO2Carbon();});
