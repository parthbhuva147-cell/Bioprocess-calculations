
const state = {
  files: [],
  mappings: {},
  processStart: 0,
  processEnd: null,
  analysis: null,
  carbon: null,
  plotIds: []
};

const onlineDefs = [
  ["do","Dissolved oxygen","% saturation","dissolved oxygen / DO"],
  ["ph","pH","pH","reactor pH"],
  ["temp","Temperature","°C","reactor temperature"],
  ["stirrer","Agitation","rpm","stirrer / agitation"],
  ["air","Air flow","ccm","air mass-flow controller"],
  ["o2in","O₂ inlet flow","ccm","pure oxygen addition"],
  ["n2in","N₂ inlet flow","ccm","nitrogen addition"],
  ["co2in","CO₂ inlet flow","ccm","CO₂ gas addition"],
  ["offco2","Off-gas CO₂","vol.%","exhaust CO₂ concentration"],
  ["offo2","Off-gas O₂","vol.%","exhaust O₂ concentration"],
  ["our","OUR","mol/L/h","oxygen uptake rate"],
  ["cer","CER","mol/L/h","carbon dioxide evolution rate"],
  ["rq","RQ","—","respiratory quotient"],
  ["feed","Feed","mL or mL/h","cumulative feed or rate"]
];

const offlineDefs = [
  ["od","OD600","OD","optical density"],
  ["dcw","DCW","g/L","dry-cell-weight concentration"],
  ["glucose","Residual glucose","g/L","substrate concentration"],
  ["product1","Product 1","g/L or mg/L","e.g. resveratrol"],
  ["product2","Product 2","g/L or mg/L","e.g. piceatannol"],
  ["samplevol","Sample volume","mL","individual sample withdrawal"]
];

function esc(s){
  return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
}
function fmt(x,d=3){
  return Number.isFinite(x) ? Number(x).toLocaleString(undefined,{maximumFractionDigits:d}) : "—";
}
function excelCol(index){
  let n=index+1,s="";
  while(n>0){ const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
  return s;
}
function norm(s){
  return String(s??"").replace(/^\uFEFF/,"").toLowerCase().replace(/[_\-]+/g," ").replace(/[()[\]{}]/g," ").replace(/\s+/g," ").trim();
}
function cleanHeader(s){ return String(s??"").replace(/^\uFEFF/,"").replace(/\s+/g," ").trim(); }
function parseNum(v){
  if(typeof v==="number"&&Number.isFinite(v)) return v;
  if(v===null||v===undefined||v==="") return NaN;
  let s=String(v).trim().replace(/\s/g,"");
  if(s.includes(",")&&!s.includes(".")) s=s.replace(",",".");
  const x=Number(s); return Number.isFinite(x)?x:NaN;
}
function parseClockToHours(s){
  const m=String(s??"").trim().match(/^(\d{1,3}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?$/);
  if(!m) return NaN;
  return Number(m[1])+Number(m[2])/60+Number(m[3]||0)/3600;
}
function parseDateValue(v){
  if(v instanceof Date&&!isNaN(v)) return v;
  if(typeof v==="number"&&v>20000&&v<100000) return new Date((v-25569)*86400*1000);
  const s=String(v??"").trim(); if(!s) return null;
  let m=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?$/);
  if(m){
    let y=Number(m[3]); if(y<100)y+=2000;
    const d=new Date(y,Number(m[2])-1,Number(m[1]),Number(m[4]),Number(m[5]),Number(m[6]||0));
    if(!isNaN(d)) return d;
  }
  const d=new Date(s); return !isNaN(d)?d:null;
}
function extractIntervalSeconds(header){
  const m=String(header??"").match(/interval\s*=\s*(\d+(?:\.\d+)?)/i);
  return m&&Number(m[1])>0?Number(m[1]):null;
}
function splitCSVLine(line,d){
  const out=[]; let cur="",q=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(ch==='"'){ if(q&&line[i+1]==='"'){cur+='"';i++;} else q=!q; }
    else if(ch===d&&!q){out.push(cur);cur="";} else cur+=ch;
  }
  out.push(cur); return out;
}
function chooseDelimiter(lines){
  let best=",",score=-1;
  for(const d of [",",";","\t"]){
    const counts=lines.slice(0,15).map(l=>splitCSVLine(l,d).length);
    const s=Math.max(...counts);
    if(s>score){score=s;best=d;}
  }
  return best;
}
function scoreHeader(cells){
  const joined=cells.map(norm).join(" | ");
  let score=cells.filter(x=>String(x).trim()).length;
  for(const k of ["time","timestamp","stirrer","temp","ph","do","air","co2","our","cer","rq","dcw","od","glucose","resveratrol","piceatannol"]){
    if(joined.includes(k))score+=7;
  }
  return score;
}
function parseCSVText(text){
  const lines=text.replace(/\r/g,"").split("\n").filter(l=>l.trim()!=="");
  if(lines.length<2) throw new Error("Not enough data rows.");
  const d=chooseDelimiter(lines);
  let headerRow=0,best=-1;
  for(let i=0;i<Math.min(50,lines.length);i++){
    const c=splitCSVLine(lines[i],d),s=scoreHeader(c);
    if(s>best){best=s;headerRow=i;}
  }
  const headers=splitCSVLine(lines[headerRow],d).map(cleanHeader);
  const rows=lines.slice(headerRow+1).map(l=>splitCSVLine(l,d)).filter(r=>r.some(v=>String(v).trim()!==""));
  return {headers,rows};
}
function parseWorkbook(buf){
  if(typeof XLSX==="undefined") throw new Error("Excel parser failed to load.");
  const wb=XLSX.read(buf,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const m=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});
  let headerRow=0,best=-1;
  for(let i=0;i<Math.min(50,m.length);i++){
    const s=scoreHeader(m[i]);
    if(s>best){best=s;headerRow=i;}
  }
  return {headers:m[headerRow].map(cleanHeader),rows:m.slice(headerRow+1).filter(r=>r.some(v=>String(v).trim()!==""))};
}
function classifyHeader(h,index,rows){
  const x=norm(h),col=excelCol(index);
  if(x.includes("timestamp")||x==="time"||x.includes("process time")||x.includes("elapsed time")||x==="hours"||x==="hour"||x==="time h")
    return "time";
  if(/\bdcw\b/.test(x)||x.includes("dry cell")) return "dcw";
  if(/\bod600\b/.test(x)||x==="od"||x.includes("optical density")) return "od";
  if(x.includes("glucose")||x.includes("substrate concentration")||x==="substrate") return "glucose";
  if(x.includes("resveratrol")) return "product1";
  if(x.includes("piceatannol")) return "product2";
  if(x.includes("sample")&&x.includes("volume")) return "samplevol";
  if(x.includes("dissolved oxygen")||/^do\b/.test(x)) return "do";
  if(x==="ph"||x.startsWith("ph ")) return "ph";
  if(x.startsWith("temp")||x.includes("temperature")) return x.includes("bluevary")?null:"temp";
  if(x.includes("stirrer")||x.includes("agitation")||x.includes("rpm")) return "stirrer";
  if(x.includes("bluevary")&&x.includes("our")) return "our";
  if(x.includes("bluevary")&&x.includes("cer")) return "cer";
  if(x.includes("bluevary")&&x.includes("rq")) return "rq";
  if((x.includes("subs a")||x.includes("feed"))&&!x.includes("feedback")) return "feed";
  if(x.startsWith("air ")||x==="air ccm"||x.includes("air flow")) return "air";
  if(x.startsWith("o2 ")||x==="o2 ccm") return "o2in";
  if(x.startsWith("n2 ")||x==="n2 ccm") return "n2in";
  if(x.startsWith("co2 ")||x==="co2 ccm") return "co2in";
  if(x.includes("bluevary")&&x.includes("vol")){
    const vals=rows.slice(0,Math.min(rows.length,10000)).map(r=>parseNum(r[index])).filter(Number.isFinite).filter(v=>v!==0);
    if(vals.length){
      vals.sort((a,b)=>a-b);
      const med=vals[Math.floor(vals.length/2)];
      if(med<10) return "offco2";
      if(med>10&&med<25) return "offo2";
    }
    return null;
  }
  if(x.includes("offgas")&&x.includes("co2")) return "offco2";
  if(x.includes("offgas")&&x.includes("o2")) return "offo2";
  return null;
}
function guessRole(headers){
  const j=headers.map(norm).join(" | ");
  if(j.includes("stirrer")||j.includes("dissolved oxygen")||j.includes("bluevary")||j.includes("air ccm")) return "Bioreactor";
  if(j.includes("resveratrol")||j.includes("piceatannol")||j.includes("hplc")) return "HPLC";
  if(j.includes("dcw")||j.includes("od600")||j.includes("optical density")) return "OD/DCW";
  if(j.includes("glucose")||j.includes("sample volume")) return "Sampling";
  return "Other";
}
function detectTimeIndex(headers){
  for(let i=0;i<headers.length;i++) if(classifyHeader(headers[i],i,[])==="time") return i;
  return -1;
}
function columnNumericCount(file,index){
  let n=0;
  for(let i=0;i<Math.min(file.rows.length,5000);i++) if(Number.isFinite(parseNum(file.rows[i][index]))) n++;
  return n;
}
function normalizeFileTime(file){
  const idx=file.timeIndex;
  if(idx<0){ file.elapsed=[];file.timeBasis="No time";file.duration=NaN;file.timeWarning="Select a time column.";return; }
  const vals=file.rows.map(r=>r[idx]);
  const forced=Number(file.forcedInterval);
  const headerInterval=extractIntervalSeconds(file.headers[idx]);
  const interval=(Number.isFinite(forced)&&forced>0)?forced:headerInterval;

  if(interval){
    const dates=vals.map(parseDateValue).filter(Boolean);
    let inconsistent=false;
    if(dates.length>=3){
      const ms=dates.map(d=>d.getTime()),span=(Math.max(...ms)-Math.min(...ms))/3600000;
      const expected=(file.rows.length-1)*interval/3600;
      if(expected>0 && (!Number.isFinite(span)||span<=0||span/expected>1.25||span/expected<.75)) inconsistent=true;
    }
    if(forced>0 || inconsistent || dates.length<Math.min(3,Math.ceil(file.rows.length*.2))){
      file.elapsed=file.rows.map((_,i)=>i*interval/3600 + Number(file.offset||0));
      file.timeBasis=forced>0?`Forced ${interval}s interval`:`Logger interval ${interval}s`;
      file.duration=(file.rows.length-1)*interval/3600;
      file.timeWarning=inconsistent?"Spreadsheet timestamps conflict with logger interval; interval-derived process time is used.":"";
      file.timeConfidence="high";
      return;
    }
  }

  const dates=vals.map(parseDateValue);
  const vd=dates.filter(Boolean);
  if(vd.length>=Math.min(3,Math.ceil(file.rows.length*.2))){
    const start=vd[0].getTime();
    file.elapsed=dates.map(d=>d?(d.getTime()-start)/3600000+Number(file.offset||0):NaN);
    const finite=file.elapsed.filter(Number.isFinite);
    file.duration=finite.length?Math.max(...finite)-Math.min(...finite):NaN;
    file.timeBasis="Timestamps";
    file.timeWarning="";
    file.timeConfidence="high";
    return;
  }

  const clocks=vals.map(parseClockToHours);
  const vc=clocks.filter(Number.isFinite);
  if(vc.length>=Math.min(3,Math.ceil(file.rows.length*.2))){
    const start=vc[0];
    file.elapsed=clocks.map(v=>Number.isFinite(v)?v-start+Number(file.offset||0):NaN);
    const finite=file.elapsed.filter(Number.isFinite);
    file.duration=finite.length?Math.max(...finite)-Math.min(...finite):NaN;
    file.timeBasis="Clock time";
    file.timeWarning="";
    file.timeConfidence="medium";
    return;
  }

  const nums=vals.map(parseNum),vn=nums.filter(Number.isFinite);
  if(vn.length>=Math.min(3,Math.ceil(file.rows.length*.2))){
    const start=vn[0];
    file.elapsed=nums.map(v=>Number.isFinite(v)?v-start+Number(file.offset||0):NaN);
    const finite=file.elapsed.filter(Number.isFinite);
    file.duration=finite.length?Math.max(...finite)-Math.min(...finite):NaN;
    file.timeBasis="Numeric elapsed time (assumed h)";
    file.timeWarning="Confirm that the numeric time column is expressed in hours.";
    file.timeConfidence="medium";
    return;
  }
  file.elapsed=[];file.duration=NaN;file.timeBasis="Unresolved";file.timeWarning="Time values could not be interpreted.";file.timeConfidence="low";
}
async function loadFiles(){
  const chosen=[...document.getElementById("multiFiles").files];
  if(!chosen.length){alert("Choose one or more CSV/Excel files.");return;}
  state.files=[];
  try{
    for(let k=0;k<chosen.length;k++){
      const f=chosen[k];
      let p;
      if(/\.(xlsx|xls)$/i.test(f.name)) p=parseWorkbook(await f.arrayBuffer());
      else p=parseCSVText(await f.text());
      const file={
        id:`f${Date.now()}_${k}`,name:f.name,headers:p.headers,rows:p.rows,
        role:guessRole(p.headers),timeIndex:detectTimeIndex(p.headers),offset:0,forcedInterval:"",
        elapsed:[],duration:NaN,timeBasis:"",timeWarning:"",timeConfidence:"neutral"
      };
      normalizeFileTime(file);
      state.files.push(file);
    }
    autoMapSignals();
    renderFilesTable();
    renderMappings();
    suggestProductUnit("product1");
    suggestProductUnit("product2");
    updateGlobalTime();
    applyProcessWindow(true);
  }catch(e){ alert(e.message); }
}
function clearWorkspace(){
  state.files=[];state.mappings={};state.analysis=null;state.carbon=null;
  document.querySelector("#filesTable tbody").innerHTML="";
  document.getElementById("onlineMapping").innerHTML="";
  document.getElementById("offlineMapping").innerHTML="";
  document.getElementById("kpiGrid").innerHTML='<div class="kpi muted-kpi"><strong>—</strong><span>Build analysis to calculate KPIs</span></div>';
  document.getElementById("globalDuration").textContent="—";
  document.getElementById("globalDurationText").textContent="Load data to establish the process timeline.";
  document.getElementById("figureGallery").innerHTML='<div class="card empty-figure">Load and map data, then generate the figures.</div>';
}
function renderFilesTable(){
  const tb=document.querySelector("#filesTable tbody"); tb.innerHTML="";
  state.files.forEach(file=>{
    const tr=document.createElement("tr");
    const timeOpts=['<option value="-1">None</option>'].concat(file.headers.map((h,i)=>`<option value="${i}" ${i===file.timeIndex?"selected":""}>${esc(h||"Unnamed")} [${excelCol(i)}]</option>`)).join("");
    tr.innerHTML=`
      <td><strong>${esc(file.name)}</strong>${file.timeWarning?`<div class="tiny-warning">${esc(file.timeWarning)}</div>`:""}</td>
      <td><select onchange="updateFileRole('${file.id}',this.value)">
        ${["Bioreactor","OD/DCW","HPLC","Sampling","Feed/Additions","Other"].map(r=>`<option ${r===file.role?"selected":""}>${r}</option>`).join("")}
      </select></td>
      <td>${file.rows.length.toLocaleString()}</td>
      <td><select onchange="updateFileTime('${file.id}',this.value)">${timeOpts}</select></td>
      <td>${esc(file.timeBasis)}</td>
      <td>${fmt(file.duration,2)} h</td>
      <td><input type="number" step="any" value="${file.offset}" onchange="updateFileOffset('${file.id}',this.value)" /></td>
      <td><input type="number" step="any" placeholder="Auto" value="${file.forcedInterval}" onchange="updateFileInterval('${file.id}',this.value)" /></td>`;
    tb.appendChild(tr);
  });
}
function updateFileRole(id,v){ const f=state.files.find(x=>x.id===id);if(f)f.role=v; }
function updateFileTime(id,v){ const f=state.files.find(x=>x.id===id);if(f){f.timeIndex=Number(v);normalizeFileTime(f);renderFilesTable();updateGlobalTime();renderMappings();} }
function updateFileOffset(id,v){ const f=state.files.find(x=>x.id===id);if(f){f.offset=Number(v)||0;normalizeFileTime(f);renderFilesTable();updateGlobalTime();} }
function updateFileInterval(id,v){ const f=state.files.find(x=>x.id===id);if(f){f.forcedInterval=v;normalizeFileTime(f);renderFilesTable();updateGlobalTime();} }
function updateGlobalTime(){
  const good=state.files.filter(f=>Number.isFinite(f.duration)&&f.elapsed.some(Number.isFinite));
  if(!good.length)return;
  const reactor=good.find(f=>f.role==="Bioreactor")||good[0];
  const finite=reactor.elapsed.filter(Number.isFinite);
  const min=Math.min(...finite),max=Math.max(...finite),dur=max-min;
  document.getElementById("globalDuration").textContent=`${fmt(dur,2)} h`;
  document.getElementById("globalDurationText").textContent=`${Math.floor(dur/24)} d ${Math.floor(dur%24)} h ${Math.round((dur%1)*60)} min • reference: ${reactor.name}`;
  const conf=document.getElementById("timeConfidence");
  conf.className=`confidence ${reactor.timeConfidence||"neutral"}`;
  conf.textContent=reactor.timeConfidence==="high"?"High-confidence process time":reactor.timeConfidence==="medium"?"Check time units":"Time needs review";
  document.getElementById("timeBasisMessage").innerHTML=`<strong>${esc(reactor.timeBasis)}</strong>${reactor.timeWarning?` — ${esc(reactor.timeWarning)}`:""}`;
  if(!document.getElementById("windowEnd").value) document.getElementById("windowEnd").placeholder=fmt(max,3);
}
function allColumnOptions(){
  const arr=[];
  state.files.forEach(f=>{
    f.headers.forEach((h,i)=>{
      if(i===f.timeIndex)return;
      const n=columnNumericCount(f,i);
      if(n<2)return;
      arr.push({key:`${f.id}::${i}`,file:f,index:i,label:`${f.name} → ${h||"Unnamed"} [${excelCol(i)}]`});
    });
  });
  return arr;
}
function autoMapSignals(){
  state.mappings={};
  const used=new Set();
  for(const f of state.files){
    f.headers.forEach((h,i)=>{
      if(i===f.timeIndex)return;
      const cls=classifyHeader(h,i,f.rows);
      if(cls && !state.mappings[cls] && columnNumericCount(f,i)>=2){
        state.mappings[cls]=`${f.id}::${i}`; used.add(`${f.id}::${i}`);
      }
    });
  }
}
function renderMappings(){
  renderMappingGroup("onlineMapping",onlineDefs);
  renderMappingGroup("offlineMapping",offlineDefs);
}
function renderMappingGroup(id,defs){
  const wrap=document.getElementById(id),opts=allColumnOptions(); wrap.innerHTML="";
  defs.forEach(([key,label,unit,hint])=>{
    const div=document.createElement("div"); div.className="mapping-item";
    let options=`<option value="">Not mapped</option>`;
    opts.forEach(o=>options+=`<option value="${o.key}" ${state.mappings[key]===o.key?"selected":""}>${esc(o.label)}</option>`);
    div.innerHTML=`<label>${label} <span class="hint">${unit} • ${hint}</span>
      <select onchange="setMapping('${key}',this.value)">${options}</select></label>`;
    wrap.appendChild(div);
  });
}
function suggestProductUnit(key){
  if(key!=="product1"&&key!=="product2")return;
  const r=resolveMapping(key); if(!r)return;
  const h=norm(r.header);
  const el=document.getElementById(key==="product1"?"product1Unit":"product2Unit");
  if(!el)return;
  if(h.includes("mg l")||h.includes("mg/l")||h.includes("mg per l")) el.value="mgL";
  else if(h.includes("g l")||h.includes("g/l")||h.includes("g per l")) el.value="gL";
}
function setMapping(k,v){ state.mappings[k]=v||null; suggestProductUnit(k); }
function applyProcessWindow(silent=false){
  updateGlobalTime();
  const reference=state.files.find(f=>f.role==="Bioreactor"&&Number.isFinite(f.duration))||state.files.find(f=>Number.isFinite(f.duration));
  if(!reference){if(!silent)alert("Load a file with usable process time first.");return;}
  const finite=reference.elapsed.filter(Number.isFinite);
  const autoEnd=Math.max(...finite);
  let start=Number(document.getElementById("windowStart").value);
  let end=Number(document.getElementById("windowEnd").value);
  if(!Number.isFinite(start))start=0;
  if(!Number.isFinite(end))end=autoEnd;
  if(end<=start){if(!silent)alert("Process end must be greater than process start.");return;}
  state.processStart=start;state.processEnd=end;
  document.getElementById("windowStatus").innerHTML=`Analysis window: <strong>${fmt(start,3)} to ${fmt(end,3)} h</strong> (${fmt(end-start,3)} h). All calculations use this interval.`;
}
function resolveMapping(key){
  const ref=state.mappings[key]; if(!ref)return null;
  const [fid,idx]=ref.split("::"); const file=state.files.find(f=>f.id===fid);
  if(!file)return null; return {file,index:Number(idx),header:file.headers[Number(idx)]};
}
function getSeries(key,within=true){
  const r=resolveMapping(key); if(!r)return [];
  const out=[];
  for(let i=0;i<r.file.rows.length;i++){
    const t=r.file.elapsed[i],y=parseNum(r.file.rows[i][r.index]);
    if(!Number.isFinite(t)||!Number.isFinite(y))continue;
    if(within&&(t<state.processStart||t>state.processEnd))continue;
    out.push({t,y});
  }
  out.sort((a,b)=>a.t-b.t); return out;
}
function interp(series,t){
  if(!series.length)return NaN;
  if(t<=series[0].t)return series[0].y;
  if(t>=series[series.length-1].t)return series[series.length-1].y;
  let lo=0,hi=series.length-1;
  while(hi-lo>1){const m=(lo+hi)>>1;if(series[m].t<=t)lo=m;else hi=m;}
  const a=series[lo],b=series[hi]; if(b.t===a.t)return a.y;
  return a.y+(b.y-a.y)*(t-a.t)/(b.t-a.t);
}
function stat(series){
  if(!series.length)return null;
  let min=Infinity,max=-Infinity,sum=0;
  series.forEach(p=>{sum+=p.y;if(p.y<min)min=p.y;if(p.y>max)max=p.y;});
  return {min,max,mean:sum/series.length,first:series[0].y,last:series[series.length-1].y,n:series.length};
}
function linearRegression(points){
  const n=points.length;if(n<2)return null;
  let sx=0,sy=0,sxx=0,sxy=0,syy=0;
  points.forEach(p=>{sx+=p.t;sy+=p.y;sxx+=p.t*p.t;sxy+=p.t*p.y;syy+=p.y*p.y;});
  const den=n*sxx-sx*sx;if(den===0)return null;
  const slope=(n*sxy-sx*sy)/den,inter=(sy-slope*sx)/n;
  const ssTot=syy-sy*sy/n;
  let ssRes=0;points.forEach(p=>{const pred=inter+slope*p.t;ssRes+=(p.y-pred)**2;});
  return {slope,inter,r2:ssTot>0?1-ssRes/ssTot:1};
}
function estimateMuMax(){
  const s=getSeries("dcw");
  const pts=s.filter(p=>p.y>0).map(p=>({t:p.t,y:Math.log(p.y)}));
  if(pts.length<3)return null;
  let best=null;
  const win=Math.min(5,pts.length);
  for(let size=3;size<=win;size++){
    for(let i=0;i<=pts.length-size;i++){
      const r=linearRegression(pts.slice(i,i+size));
      if(r&&r.slope>0&&r.r2>=0.95&&(!best||r.slope>best.mu)) best={mu:r.slope,r2:r.r2,t1:pts[i].t,t2:pts[i+size-1].t};
    }
  }
  return best;
}
function feedVolumeMl(){
  const s=getSeries("feed"); if(s.length<2)return 0;
  const mode=document.getElementById("feedMode").value;
  if(mode==="cumulative") return Math.max(0,s[s.length-1].y-s[0].y);
  let ml=0;
  for(let i=1;i<s.length;i++){
    const dt=s[i].t-s[i-1].t;
    ml+=0.5*(s[i].y+s[i-1].y)*dt;
  }
  return ml;
}
function buildAnalysis(){
  applyProcessWindow();
  const a={};
  for(const k of ["do","ph","temp","stirrer","our","cer","rq","od","dcw","glucose","product1","product2"]) a[k]=stat(getSeries(k));
  a.feedMl=feedVolumeMl();
  a.mu=estimateMuMax();
  const dur=state.processEnd-state.processStart;
  const p1=getSeries("product1");
  if(p1.length){
    let last=p1[p1.length-1].y;
    if(document.getElementById("product1Unit").value==="mgL")last/=1000;
    a.productivity1=dur>0?last/dur:NaN;
  }
  const p2=getSeries("product2");
  if(p2.length){
    let last=p2[p2.length-1].y;
    if(document.getElementById("product2Unit").value==="mgL")last/=1000;
    a.productivity2=dur>0?last/dur:NaN;
  }
  const dcw=getSeries("dcw"),glc=getSeries("glucose");
  if(dcw.length>=2&&glc.length>=2){
    const consumed=glc[0].y-glc[glc.length-1].y;
    if(consumed>0)a.apparentYxs=(dcw[dcw.length-1].y-dcw[0].y)/consumed;
  }
  if(p1.length>=2&&glc.length>=2){
    let p0=p1[0].y,pn=p1[p1.length-1].y;
    if(document.getElementById("product1Unit").value==="mgL"){p0/=1000;pn/=1000;}
    const consumed=glc[0].y-glc[glc.length-1].y;
    if(consumed>0)a.apparentYps=(pn-p0)/consumed;
  }
  state.analysis=a; renderKPIs();
}
function renderKPIs(){
  const a=state.analysis,items=[];
  items.push([`${fmt(state.processEnd-state.processStart,2)} h`,"Process duration"]);
  if(a.do)items.push([`${fmt(a.do.min,2)} %`,"Minimum DO"]);
  if(a.ph)items.push([fmt(a.ph.mean,3),"Mean pH"]);
  if(a.temp)items.push([`${fmt(a.temp.mean,2)} °C`,"Mean temperature"]);
  if(a.stirrer)items.push([`${fmt(a.stirrer.max,0)} rpm`,"Maximum agitation"]);
  if(a.feedMl>0)items.push([`${fmt(a.feedMl,2)} mL`,"Feed delivered"]);
  if(a.od)items.push([fmt(a.od.max,2),"Maximum OD600"]);
  if(a.dcw)items.push([`${fmt(a.dcw.max,3)} g/L`,"Maximum DCW"]);
  if(a.glucose)items.push([`${fmt(a.glucose.last,3)} g/L`,"Final residual glucose"]);
  if(a.product1)items.push([`${fmt(a.product1.last,3)} ${document.getElementById("product1Unit").value==="mgL"?"mg/L":"g/L"}`,"Final Product 1 titre"]);
  if(a.product2)items.push([`${fmt(a.product2.last,3)} ${document.getElementById("product2Unit").value==="mgL"?"mg/L":"g/L"}`,"Final Product 2 titre"]);
  if(a.mu)items.push([`${fmt(a.mu.mu,4)} h⁻¹`,"Estimated μmax (R² ≥ 0.95)"]);
  if(Number.isFinite(a.productivity1))items.push([`${fmt(a.productivity1,4)} g/L/h`,"Product 1 volumetric productivity"]);
  if(Number.isFinite(a.apparentYxs))items.push([`${fmt(a.apparentYxs,3)} g/g`,"Apparent concentration Yx/s"]);
  if(Number.isFinite(a.apparentYps))items.push([`${fmt(a.apparentYps,3)} g/g`,"Apparent concentration Yp/s"]);
  document.getElementById("kpiGrid").innerHTML=items.map(([v,l])=>`<div class="kpi"><strong>${v}</strong><span>${l}</span></div>`).join("");
}
function getManualLedger(){
  return [...document.querySelectorAll("#manualLedgerTable tbody tr")].map(tr=>({
    time:Number(tr.querySelector(".evt-time").value),
    desc:tr.querySelector(".evt-desc").value||"Manual event",
    dir:tr.querySelector(".evt-dir").value,
    carbon:Number(tr.querySelector(".evt-c").value)
  })).filter(e=>Number.isFinite(e.carbon)&&e.carbon>=0);
}
function addLedgerRow(time="",desc="",dir="plus",carbon=""){
  const tb=document.querySelector("#manualLedgerTable tbody"),tr=document.createElement("tr");
  tr.innerHTML=`<td><input class="evt-time" type="number" step="any" value="${time}" /></td>
    <td><input class="evt-desc" type="text" value="${esc(desc)}" placeholder="e.g. glucose bolus" /></td>
    <td><select class="evt-dir"><option value="plus" ${dir==="plus"?"selected":""}>+ Carbon input</option><option value="minus" ${dir==="minus"?"selected":""}>− Carbon removal</option></select></td>
    <td><input class="evt-c" type="number" step="any" value="${carbon}" /></td>
    <td><button class="button secondary small" onclick="this.closest('tr').remove()">Remove</button></td>`;
  tb.appendChild(tr);
}
function uniqueSampleTimes(){
  const sampleVol=getSeries("samplevol");
  if(sampleVol.length){
    return [...new Set(sampleVol.map(p=>p.t.toFixed(6)))].map(Number)
      .filter(t=>t>=state.processStart&&t<=state.processEnd).sort((a,b)=>a-b);
  }
  const keys=["dcw","glucose","product1","product2"];
  const set=new Set();
  keys.forEach(k=>getSeries(k).forEach(p=>set.add(p.t.toFixed(6))));
  return [...set].map(Number).filter(t=>t>=state.processStart&&t<=state.processEnd).sort((a,b)=>a-b);
}
function concentrationAt(key,t){
  const s=getSeries(key);return interp(s,t);
}
function sampleCarbon(){
  const times=uniqueSampleTimes(); if(!times.length)return {total:0,volumeMl:0,events:[]};
  const sampleS=getSeries("samplevol"),def=Number(document.getElementById("defaultSampleVolume").value)||0;
  const bcf=Number(document.getElementById("biomassCF").value)||0;
  const p1cf=Number(document.getElementById("product1CF").value)||0;
  const p2cf=Number(document.getElementById("product2CF").value)||0;
  const events=[];let total=0,volTotal=0;
  times.forEach(t=>{
    let ml=sampleS.length?interp(sampleS,t):def;
    if(!Number.isFinite(ml)||ml<=0)return;
    const L=ml/1000;
    const gl=concentrationAt("glucose",t);
    const dcw=concentrationAt("dcw",t);
    let p1=concentrationAt("product1",t),p2=concentrationAt("product2",t);
    if(document.getElementById("product1Unit").value==="mgL"&&Number.isFinite(p1))p1/=1000;
    if(document.getElementById("product2Unit").value==="mgL"&&Number.isFinite(p2))p2/=1000;
    let c=0;
    if(Number.isFinite(gl))c+=L*gl*0.40001998;
    if(Number.isFinite(dcw))c+=L*dcw*bcf;
    if(Number.isFinite(p1))c+=L*p1*p1cf;
    if(Number.isFinite(p2))c+=L*p2*p2cf;
    total+=c;volTotal+=ml;events.push({time:t,carbon:c,volumeMl:ml});
  });
  return {total,volumeMl:volTotal,events};
}
function reactorVolumeAt(t,initialL,sampleInfo){
  let v=initialL;
  const feed=getSeries("feed");
  if(feed.length){
    if(document.getElementById("feedMode").value==="cumulative"){
      const f0=interp(feed,state.processStart),ft=interp(feed,t);
      if(Number.isFinite(f0)&&Number.isFinite(ft))v+=Math.max(0,ft-f0)/1000;
    }else{
      let add=0;
      for(let i=1;i<feed.length;i++){
        const a=feed[i-1],b=feed[i];
        if(a.t>t)break;
        const t2=Math.min(b.t,t),dt=t2-a.t;
        if(dt>0)add+=0.5*(a.y+b.y)*dt/1000;
        if(b.t>=t)break;
      }
      v+=add;
    }
  }
  sampleInfo.events.forEach(e=>{if(e.time<=t)v-=e.volumeMl/1000;});
  return Math.max(v,0);
}
function integrateCER(initialL,sampleInfo){
  const s=getSeries("cer");
  const valid=s.filter(p=>Number.isFinite(p.y)&&p.y!==0);
  if(valid.length<2)return null;
  let mol=0;
  for(let i=1;i<valid.length;i++){
    const a=valid[i-1],b=valid[i],dt=b.t-a.t;
    if(dt<=0)continue;
    const va=reactorVolumeAt(a.t,initialL,sampleInfo),vb=reactorVolumeAt(b.t,initialL,sampleInfo);
    const r1=a.y*va,r2=b.y*vb;
    mol+=0.5*(r1+r2)*dt;
  }
  return {gC:mol*12.011,mol,method:"CER integration (mol/L/h × estimated reactor volume)"};
}
function combinedGasSeries(){
  const co2=getSeries("offco2"); if(co2.length<2)return null;
  const air=getSeries("air"),o2=getSeries("o2in"),n2=getSeries("n2in"),co2in=getSeries("co2in"),offo2=getSeries("offo2");
  if(!air.length&&!o2.length&&!n2.length&&!co2in.length)return null;
  const vm=Number(document.getElementById("molarVolume").value)||24.465;
  const airCO2=(Number(document.getElementById("airCO2").value)||0)/100;
  const airN2=(Number(document.getElementById("airN2").value)||79)/100;
  let mol=0,usedN2=false;
  for(let i=1;i<co2.length;i++){
    const a=co2[i-1],b=co2[i],dtMin=(b.t-a.t)*60;if(dtMin<=0)continue;
    const calcAt=(p)=>{
      const t=p.t,airC=Number.isFinite(interp(air,t))?interp(air,t):0;
      const o2C=Number.isFinite(interp(o2,t))?interp(o2,t):0;
      const n2C=Number.isFinite(interp(n2,t))?interp(n2,t):0;
      const c2C=Number.isFinite(interp(co2in,t))?interp(co2in,t):0;
      const fin=(airC+o2C+n2C+c2C)/1000;
      if(fin<=0)return NaN;
      const yco2=p.y/100;
      const co2InLmin=(airC/1000)*airCO2+(c2C/1000);
      let fout=fin;
      const yo2=interp(offo2,t);
      if(Number.isFinite(yo2)&&yo2>0&&yo2<100){
        const yn2out=1-(yo2/100)-yco2;
        const n2inLmin=(airC/1000)*airN2+(n2C/1000);
        if(yn2out>0.2&&n2inLmin>0){fout=n2inLmin/yn2out;usedN2=true;}
      }
      const netLmin=yco2*fout-co2InLmin;
      return netLmin/vm;
    };
    const r1=calcAt(a),r2=calcAt(b);
    if(Number.isFinite(r1)&&Number.isFinite(r2))mol+=0.5*(r1+r2)*dtMin;
  }
  if(!Number.isFinite(mol)||mol===0)return null;
  return {gC:mol*12.011,mol,method:usedN2?"Off-gas CO₂ with N₂ outlet-flow correction":"Off-gas CO₂ with inlet-flow approximation"};
}
function gasCO2InputCarbon(){
  const vm=Number(document.getElementById("molarVolume").value)||24.465;
  const airCO2=(Number(document.getElementById("airCO2").value)||0)/100;
  const air=getSeries("air"),co2in=getSeries("co2in");
  const timeline=(air.length?air:co2in);
  if(timeline.length<2)return 0;
  let mol=0;
  for(let i=1;i<timeline.length;i++){
    const a=timeline[i-1],b=timeline[i],dtMin=(b.t-a.t)*60;if(dtMin<=0)continue;
    const rateAt=t=>{
      const av=Number.isFinite(interp(air,t))?interp(air,t):0;
      const cv=Number.isFinite(interp(co2in,t))?interp(co2in,t):0;
      return ((av/1000)*airCO2+(cv/1000))/vm;
    };
    mol+=0.5*(rateAt(a.t)+rateAt(b.t))*dtMin;
  }
  return mol*12.011;
}
function finalInventoryCarbon(finalVol){
  const bcf=Number(document.getElementById("biomassCF").value)||0;
  const p1cf=Number(document.getElementById("product1CF").value)||0;
  const p2cf=Number(document.getElementById("product2CF").value)||0;
  const getLast=k=>{const s=getSeries(k);return s.length?s[s.length-1].y:NaN;};
  const gl=getLast("glucose"),dcw=getLast("dcw");let p1=getLast("product1"),p2=getLast("product2");
  if(document.getElementById("product1Unit").value==="mgL"&&Number.isFinite(p1))p1/=1000;
  if(document.getElementById("product2Unit").value==="mgL"&&Number.isFinite(p2))p2/=1000;
  const parts={};
  if(Number.isFinite(gl))parts["Residual glucose"]=finalVol*gl*0.40001998;
  if(Number.isFinite(dcw))parts["Biomass"]=finalVol*dcw*bcf;
  if(Number.isFinite(p1))parts["Product 1"]=finalVol*p1*p1cf;
  if(Number.isFinite(p2))parts["Product 2"]=finalVol*p2*p2cf;
  return parts;
}
function calculateCarbonBalance(){
  applyProcessWindow();
  const initV=Number(document.getElementById("initialVolume").value);
  const initG=Number(document.getElementById("initialGlucose").value);
  const feedConc=Number(document.getElementById("feedGlucoseConc").value);
  if(!(initV>0)||!(initG>=0)){alert("Enter valid initial volume and glucose.");return;}
  const sample=sampleCarbon(),feedMl=feedVolumeMl();
  const initialC=initV*initG*0.40001998;
  const feedC=(feedMl/1000)*(Number.isFinite(feedConc)?feedConc:0)*0.40001998;
  const gasInC=gasCO2InputCarbon();
  const manual=getManualLedger();
  const manualPlus=manual.filter(e=>e.dir==="plus").reduce((a,e)=>a+e.carbon,0);
  const manualMinus=manual.filter(e=>e.dir==="minus").reduce((a,e)=>a+e.carbon,0);
  let co2=null;
  const method=document.getElementById("co2Method").value;
  if(method==="cer"||method==="auto")co2=integrateCER(initV,sample);
  if(!co2&&(method==="offgas"||method==="auto"))co2=combinedGasSeries();

  // CER/off-gas routines return NET CO2 evolution (outlet minus inlet).
  // For the ledger we show gas carbon entering as + and gross exhaust CO2 as -,
  // so gross exhaust = net evolution + gas-phase carbon entering.
  const netCo2C=co2?co2.gC:0;
  const grossCo2OutC=co2?Math.max(0,netCo2C+gasInC):0;

  let finalV=Number(document.getElementById("finalVolumeOverride").value);
  if(!(finalV>0)) finalV=Math.max(0,initV+feedMl/1000-sample.volumeMl/1000);
  const inventory=finalInventoryCarbon(finalV);
  const finalMeasured=Object.values(inventory).reduce((a,b)=>a+b,0);

  const totalIn=initialC+feedC+gasInC+manualPlus;
  const knownOut=sample.total+grossCo2OutC+manualMinus;
  const accounted=knownOut+finalMeasured;
  const unaccounted=totalIn-accounted;
  const recovery=totalIn>0?accounted/totalIn*100:NaN;

  const entries=[
    {name:"Initial glucose carbon",sign:"+",c:initialC},
    {name:"Feed glucose carbon",sign:"+",c:feedC},
    {name:"Gas-phase CO₂ carbon entering",sign:"+",c:gasInC},
    ...manual.filter(e=>e.dir==="plus").map(e=>({name:e.desc,sign:"+",c:e.carbon})),
    {name:"Sampling carbon removed",sign:"−",c:sample.total},
    {name:"CO₂ carbon in exhaust",sign:"−",c:grossCo2OutC},
    ...manual.filter(e=>e.dir==="minus").map(e=>({name:e.desc,sign:"−",c:e.carbon}))
  ];
  state.carbon={entries,totalIn,knownOut,inventory,finalMeasured,unaccounted,recovery,finalV,co2,sample,feedMl,gasInC,netCo2C,grossCo2OutC};
  renderCarbon();
}
function renderCarbon(){
  const c=state.carbon;
  document.getElementById("carbonLedger").innerHTML=`<table class="ledger-table"><thead><tr><th>Carbon event</th><th>Direction</th><th>g C</th></tr></thead><tbody>${
    c.entries.map(e=>`<tr><td>${esc(e.name)}</td><td class="${e.sign==="+"?"plus":"minus"}">${e.sign}</td><td>${fmt(e.c,4)}</td></tr>`).join("")
  }</tbody></table>`;
  const inv=Object.entries(c.inventory).map(([k,v])=>`${k}: ${fmt(v,3)} g C`).join(" • ")||"No final carbon-containing species mapped";
  document.getElementById("carbonSummary").innerHTML=`
    <div class="balance-grid">
      <div><strong>${fmt(c.totalIn,3)} g C</strong><span>Total carbon input</span></div>
      <div><strong>${fmt(c.knownOut,3)} g C</strong><span>Known carbon removed/emitted</span></div>
      <div><strong>${fmt(c.finalMeasured,3)} g C</strong><span>Measured final reactor carbon</span></div>
      <div><strong>${fmt(c.unaccounted,3)} g C</strong><span>Unaccounted carbon</span></div>
      <div><strong>${fmt(c.recovery,1)} %</strong><span>Measured carbon recovery</span></div>
      <div><strong>${fmt(c.finalV,3)} L</strong><span>Final volume used</span></div>
    </div>
    <div class="balance-warning">
      <strong>Final inventory:</strong> ${esc(inv)}<br>
      <strong>CO₂ method:</strong> ${c.co2?esc(c.co2.method):"No usable CER/off-gas CO₂ calculation was possible from the mapped data."}
    </div>`;
}
function plotSeries(key){
  const s=getSeries(key); return {x:s.map(p=>p.t),y:s.map(p=>p.y)};
}
function getPlotRange(){
  let s=Number(document.getElementById("plotStart").value),e=Number(document.getElementById("plotEnd").value);
  if(!Number.isFinite(s))s=state.processStart;if(!Number.isFinite(e))e=state.processEnd;
  return [s,e];
}
function baseLayout(title,ytitle){
  const fs=Number(document.getElementById("figFont").value)||24;
  return {
    title:{text:title,x:.5,xanchor:"center",font:{size:fs+4}},
    paper_bgcolor:"#ffffff",plot_bgcolor:"#ffffff",
    font:{family:"Arial, Helvetica, sans-serif",size:fs,color:"#111827"},
    margin:{l:105,r:35,t:90,b:90},
    xaxis:{title:{text:"Process time (h)",standoff:16},range:getPlotRange(),showline:true,linecolor:"#111827",linewidth:2,mirror:true,ticks:"outside",tickwidth:2,ticklen:7,gridcolor:"#e5e7eb",zeroline:false},
    yaxis:{title:{text:ytitle,standoff:18},showline:true,linecolor:"#111827",linewidth:2,mirror:true,ticks:"outside",tickwidth:2,ticklen:7,gridcolor:"#e5e7eb",zeroline:false},
    legend:{orientation:"h",x:.5,xanchor:"center",y:1.12,yanchor:"bottom"},
    hovermode:"x unified"
  };
}
function lineTrace(key,name){
  const s=plotSeries(key),lw=Number(document.getElementById("figLine").value)||3;
  return {x:s.x,y:s.y,type:"scatter",mode:"lines",name,line:{width:lw},connectgaps:false};
}
function markerTrace(key,name){
  const s=plotSeries(key);
  return {x:s.x,y:s.y,type:"scatter",mode:"lines+markers",name,line:{width:2.5},marker:{size:10},connectgaps:false};
}
function makeFigure(id,title,traces,layout){
  const gallery=document.getElementById("figureGallery"),card=document.createElement("div");card.className="figure-card";
  card.innerHTML=`<div class="figure-head"><h3>${esc(title)}</h3><div class="figure-actions">
    <button class="button secondary small" onclick="downloadFigure('${id}','svg')">Export SVG</button>
    <button class="button secondary small" onclick="downloadFigure('${id}','png')">High-res PNG</button>
  </div></div><div id="${id}" class="figure-plot"></div>`;
  gallery.appendChild(card);
  Plotly.newPlot(id,traces,layout,{responsive:true,displaylogo:false,modeBarButtonsToRemove:["lasso2d","select2d"]});
  state.plotIds.push(id);
}
function generatePublicationFigures(){
  if(typeof Plotly==="undefined"){alert("Plotly did not load. Check your internet connection.");return;}
  applyProcessWindow();
  const gal=document.getElementById("figureGallery");gal.innerHTML="";state.plotIds=[];
  const has=k=>getSeries(k).length>1;
  let n=0;
  if(has("do")){
    makeFigure("fig_do","Dissolved oxygen profile",[lineTrace("do","DO")],baseLayout("Dissolved oxygen profile","DO (% saturation)"));n++;
  }
  if(has("ph")){
    makeFigure("fig_ph","pH profile",[lineTrace("ph","pH")],baseLayout("pH profile","pH"));n++;
  }
  if(has("temp")){
    makeFigure("fig_temp","Temperature profile",[lineTrace("temp","Temperature")],baseLayout("Temperature profile","Temperature (°C)"));n++;
  }
  if(has("stirrer")){
    makeFigure("fig_stirrer","Agitation profile",[lineTrace("stirrer","Agitation")],baseLayout("Agitation profile","Stirrer speed (rpm)"));n++;
  }
  if(has("offco2")||has("offo2")){
    const tr=[];if(has("offco2"))tr.push(lineTrace("offco2","CO₂"));if(has("offo2"))tr.push(lineTrace("offo2","O₂"));
    makeFigure("fig_offgas","Off-gas composition",tr,baseLayout("Off-gas composition","Gas concentration (vol.%)"));n++;
  }
  if(has("our")||has("cer")){
    const tr=[];if(has("our"))tr.push(lineTrace("our","OUR"));if(has("cer"))tr.push(lineTrace("cer","CER"));
    makeFigure("fig_rates","Respiration rates",tr,baseLayout("Respiration rates","Rate (mol/L/h)"));n++;
  }
  if(has("rq")){
    makeFigure("fig_rq","Respiratory quotient",[lineTrace("rq","RQ")],baseLayout("Respiratory quotient","RQ"));n++;
  }
  if(has("dcw")){
    makeFigure("fig_dcw","Dry cell weight",[markerTrace("dcw","DCW")],baseLayout("Dry cell weight","DCW (g/L)"));n++;
  }
  if(has("od")){
    makeFigure("fig_od","Optical density",[markerTrace("od","OD600")],baseLayout("Optical density","OD600"));n++;
  }
  if(has("glucose")){
    makeFigure("fig_glucose","Residual glucose",[markerTrace("glucose","Glucose")],baseLayout("Residual glucose","Glucose (g/L)"));n++;
  }
  if(has("product1")){
    const u1=document.getElementById("product1Unit").value==="mgL"?"mg/L":"g/L";
    makeFigure("fig_product1","Product 1 formation",[markerTrace("product1","Product 1")],baseLayout("Product 1 formation",`Product 1 concentration (${u1})`));n++;
  }
  if(has("product2")){
    const u2=document.getElementById("product2Unit").value==="mgL"?"mg/L":"g/L";
    makeFigure("fig_product2","Product 2 formation",[markerTrace("product2","Product 2")],baseLayout("Product 2 formation",`Product 2 concentration (${u2})`));n++;
  }
  if(state.carbon){
    const c=state.carbon;
    const names=c.entries.map(e=>e.name).concat(Object.keys(c.inventory).map(k=>`Final ${k}`),["Unaccounted"]);
    const vals=c.entries.map(e=>e.sign==="+"?e.c:-e.c).concat(Object.values(c.inventory).map(v=>-v),[-c.unaccounted]);
    const trace={type:"bar",x:names,y:vals,name:"Carbon",marker:{line:{width:1}}};
    const lay=baseLayout("Carbon balance ledger","Carbon contribution (g C)");
    lay.xaxis={title:{text:"Carbon term"},tickangle:-30,showline:true,linecolor:"#111827",linewidth:2,mirror:true};
    makeFigure("fig_carbon","Carbon balance", [trace], lay);n++;
  }
  if(!n)gal.innerHTML='<div class="card empty-figure">No mapped variables contain enough data to create figures.</div>';
}
function downloadFigure(id,format){
  const w=Number(document.getElementById("figWidth").value)||1800,h=Number(document.getElementById("figHeight").value)||1100;
  if(format==="svg") Plotly.downloadImage(id,{format:"svg",filename:id,width:w,height:h});
  else Plotly.downloadImage(id,{format:"png",filename:id+"_highres",width:w,height:h,scale:2});
}
function downloadDemoPack(){
  const reactor=`TimeStamp(UTC+02:00 interval=300),STIRRER(rpm),TEMP(C),PH(pH),DO(%sat),AIR(ccm),SUBS_A(ml),Offgas_CO2(vol.%),Offgas_O2(vol.%),CER(mol/L/h),OUR(mol/L/h),RQ
04/09/26 08:00:00,600,30,5.0,95,1000,0,0.05,20.9,0.002,0.002,1.00
04/09/26 08:05:00,650,30,5.0,80,1000,0,0.12,20.5,0.004,0.0042,0.95
04/09/26 08:10:00,750,30,5.0,65,1000,2,0.30,19.7,0.008,0.0085,0.94
04/09/26 08:15:00,900,30,5.0,45,1100,4,0.55,18.9,0.012,0.0125,0.96
04/09/26 08:20:00,1000,30,5.0,35,1200,7,0.80,18.1,0.016,0.0162,0.99`;
  const sampling=`time_h,OD600,DCW_g_L,glucose_g_L,sample_volume_mL
0,0.2,0.08,40,1.5
4,0.8,0.35,35,1.5
8,2.3,1.10,25,1.5
12,5.8,2.80,12,1.5
16,8.5,4.20,4,1.5`;
  const hplc=`time_h,resveratrol_mg_L,piceatannol_mg_L
0,0,0
4,12,2
8,48,8
12,110,20
16,180,35`;
  downloadBlob(reactor,"demo_bioreactor.csv","text/csv");
  setTimeout(()=>downloadBlob(sampling,"demo_sampling_OD_DCW.csv","text/csv"),250);
  setTimeout(()=>downloadBlob(hplc,"demo_HPLC.csv","text/csv"),500);
}
function downloadBlob(content,name,type){
  const b=new Blob([content],{type}),u=URL.createObjectURL(b),a=document.createElement("a");
  a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(u);
}
addLedgerRow();
