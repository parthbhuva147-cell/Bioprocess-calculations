
let processState = null;

function n(id){ return parseFloat(document.getElementById(id).value); }
function fmt(x,digits=4){
  return Number.isFinite(x) ? Number(x).toLocaleString(undefined,{maximumFractionDigits:digits}) : "—";
}
function setResult(id,html,ok=true){
  const el=document.getElementById(id);
  el.innerHTML=html;
  el.style.background=ok?"#eaf1ff":"#fff0f0";
  el.style.color=ok?"#103d9f":"#9c1c1c";
}

function calcGrowth(){
  const x1=n("gx1"),x2=n("gx2"),t1=n("gt1"),t2=n("gt2"),dt=t2-t1;
  if(!(x1>0&&x2>0&&dt>0)) return setResult("growthResult","Check inputs: X₁, X₂ must be > 0 and t₂ > t₁.",false);
  const mu=(Math.log(x2)-Math.log(x1))/dt,td=mu>0?Math.log(2)/mu:NaN;
  setResult("growthResult",`μ = ${fmt(mu)} h⁻¹<br>Doubling time = ${fmt(td)} h`);
}
function calcYxs(){
  const x1=n("yx1"),x2=n("yx2"),s1=n("ys1"),s2=n("ys2"),ds=s1-s2;
  if(!(ds>0)) return setResult("yxsResult","Substrate consumed (S₁ − S₂) must be > 0.",false);
  setResult("yxsResult",`Yx/s = ${fmt((x2-x1)/ds)} g biomass / g substrate`);
}
function calcYps(){
  const p1=n("yp1"),p2=n("yp2"),s1=n("yps1"),s2=n("yps2"),ds=s1-s2;
  if(!(ds>0)) return setResult("ypsResult","Substrate consumed (S₁ − S₂) must be > 0.",false);
  setResult("ypsResult",`Yp/s = ${fmt((p2-p1)/ds)} g product / g substrate`);
}
function calcProductivity(){
  const p1=n("qp1"),p2=n("qp2"),t1=n("qt1"),t2=n("qt2"),dt=t2-t1;
  if(!(dt>0)) return setResult("qpResult","t₂ must be greater than t₁.",false);
  setResult("qpResult",`Qp = ${fmt((p2-p1)/dt)} g/L/h`);
}
function calcCarbonGlucose(){
  const g=n("glucoseMass");
  if(!(g>=0)) return setResult("carbonGlucoseResult","Enter a valid glucose mass.",false);
  setResult("carbonGlucoseResult",`Carbon in = ${fmt(g*(72.066/180.156),3)} g C`);
}
function calcCarbonCO2(){
  const v=n("co2Vol"),vm=n("molarVol");
  if(!(v>=0&&vm>0)) return setResult("carbonCO2Result","Volume must be ≥ 0 and Vm > 0.",false);
  setResult("carbonCO2Result",`Carbon out = ${fmt((v/vm)*12.011,3)} g C`);
}

function excelCol(index){
  let n=index+1,s="";
  while(n>0){ let r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
  return s;
}
function cleanHeader(h){
  return String(h??"").replace(/^\uFEFF/,"").replace(/\s+/g," ").trim();
}
function norm(h){
  return cleanHeader(h).toLowerCase().replace(/[_\-]+/g," ").replace(/[()[\]{}]/g," ").replace(/\s+/g," ").trim();
}
function parseNum(v){
  if(typeof v==="number" && Number.isFinite(v)) return v;
  if(v===null||v===undefined||v==="") return NaN;
  let s=String(v).trim().replace(/\s/g,"");
  if(s.includes(",") && !s.includes(".")) s=s.replace(",",".");
  const x=Number(s);
  return Number.isFinite(x)?x:NaN;
}
function parseDateValue(v){
  if(v instanceof Date && !isNaN(v)) return v;
  if(typeof v==="number" && v>20000 && v<100000){
    return new Date((v-25569)*86400*1000);
  }

  const s=String(v??"").trim();
  if(!s) return null;

  // Prefer day/month/year for explicit slash/dot/dash timestamps.
  // This avoids interpreting 10/8/26 as October 8 in browsers that assume US dates.
  let m=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?$/);
  if(m){
    let year=Number(m[3]);
    if(year<100) year+=2000;
    const d=new Date(
      year,
      Number(m[2])-1,
      Number(m[1]),
      Number(m[4]),
      Number(m[5]),
      Number(m[6]||0)
    );
    if(!isNaN(d)) return d;
  }

  const d=new Date(s);
  return !isNaN(d) ? d : null;
}

function extractIntervalSeconds(header){
  const m=String(header??"").match(/interval\s*=\s*(\d+(?:\.\d+)?)/i);
  if(!m) return null;
  const seconds=Number(m[1]);
  return Number.isFinite(seconds) && seconds>0 ? seconds : null;
}
function splitCSVLine(line,delimiter){
  const out=[]; let cur="",q=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(ch==='"'){
      if(q && line[i+1]==='"'){cur+='"';i++;}
      else q=!q;
    }else if(ch===delimiter&&!q){out.push(cur);cur="";}
    else cur+=ch;
  }
  out.push(cur);
  return out;
}
function chooseDelimiter(lines){
  const candidates=[",",";","\t"];
  let best=",",bestScore=-1;
  for(const d of candidates){
    const scores=lines.slice(0,20).map(l=>splitCSVLine(l,d).length);
    const score=Math.max(...scores);
    if(score>bestScore){best=d;bestScore=score;}
  }
  return best;
}
function scoreHeader(cells){
  const joined=cells.map(norm).join(" | ");
  let score=cells.filter(x=>String(x).trim()).length;
  const keys=["time","timestamp","stirrer","temp","ph","do","air","o2","co2","bluevary","our","cer","rq","subs"];
  for(const k of keys) if(joined.includes(k)) score+=8;
  return score;
}
function parseCSVText(text){
  const lines=text.replace(/\r/g,"").split("\n").filter(l=>l.trim()!=="");
  if(lines.length<2) throw new Error("The file does not contain enough rows.");
  const delimiter=chooseDelimiter(lines);
  let headerRow=0,best=-1;
  for(let i=0;i<Math.min(lines.length,40);i++){
    const cells=splitCSVLine(lines[i],delimiter);
    const sc=scoreHeader(cells);
    if(sc>best){best=sc;headerRow=i;}
  }
  const headers=splitCSVLine(lines[headerRow],delimiter).map(cleanHeader);
  const rows=[];
  for(let i=headerRow+1;i<lines.length;i++){
    const parts=splitCSVLine(lines[i],delimiter);
    if(parts.length<2) continue;
    rows.push(parts);
  }
  return {headers,rows};
}
function parseWorkbook(arrayBuffer){
  if(typeof XLSX==="undefined") throw new Error("Excel parser did not load. Check your internet connection or upload CSV.");
  const wb=XLSX.read(arrayBuffer,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const matrix=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});
  if(matrix.length<2) throw new Error("The Excel sheet does not contain enough rows.");
  let headerRow=0,best=-1;
  for(let i=0;i<Math.min(matrix.length,40);i++){
    const sc=scoreHeader(matrix[i]);
    if(sc>best){best=sc;headerRow=i;}
  }
  return {
    headers:matrix[headerRow].map(cleanHeader),
    rows:matrix.slice(headerRow+1).filter(r=>r.some(v=>String(v).trim()!==""))
  };
}
function classifyHeader(h,index){
  const x=norm(h), col=excelCol(index);
  if(x.includes("timestamp") || x==="time" || x.startsWith("time ") || x.includes("process time") || x.includes("elapsed time"))
    return {role:"Time",priority:100};
  if(/^do\b/.test(x) || x.includes("dissolved oxygen")) return {role:"DO",priority:90};
  if(x==="ph ph" || x==="ph" || x.startsWith("ph ")) return {role:"pH",priority:88};
  if(x.startsWith("temp") || x.includes("temperature")) return {role:"Temperature",priority:85};
  if(x.includes("stirrer") || x.includes("agitation") || x.includes("rpm")) return {role:"Stirrer",priority:82};
  if(x.startsWith("air ") || x==="air ccm" || x.includes("air flow")) return {role:"Air",priority:78};
  if(x.startsWith("o2 ") || x==="o2 ccm") return {role:"O₂ flow",priority:76};
  if(x.startsWith("n2 ") || x==="n2 ccm") return {role:"N₂ flow",priority:74};
  if(x.startsWith("co2 ") || x==="co2 ccm") return {role:"CO₂ flow",priority:72};
  if(x.includes("subs a") || x.includes("feed")) return {role:"Feed / SUBS_A",priority:80};
  if(x.includes("bluevary") && x.includes("our")) return {role:"OUR",priority:92};
  if(x.includes("bluevary") && x.includes("cer")) return {role:"CER",priority:92};
  if(x.includes("bluevary") && x.includes("rq")) return {role:"RQ",priority:92};
  if(x.includes("bluevary") && x.includes("vol")) return {role:`BlueVary vol.% (${col})`,priority:86};
  if(x.includes("bluevary") && x.includes("pres")) return {role:"BlueVary pressure",priority:55};
  if(x.includes("bluevary") && x.includes("humid")) return {role:"BlueVary humidity",priority:55};
  if(x.includes("bluevary") && x.includes("temp")) return {role:"BlueVary temperature",priority:55};
  if(x.includes("acid")) return {role:"Acid",priority:45};
  if(x.includes("base")) return {role:"Base",priority:45};
  if(x.includes("afoam") || x.includes("antifoam")) return {role:"Antifoam",priority:40};
  if(x.includes("light")) return {role:"Light",priority:35};
  return {role:"Other",priority:0};
}
function buildColumns(headers){
  return headers.map((h,i)=>{
    const c=classifyHeader(h,i);
    return {index:i,header:h||`Column ${excelCol(i)}`,label:`${h||"Unnamed"} [${excelCol(i)}]`,role:c.role,priority:c.priority};
  });
}
function findTimeColumn(columns){
  return columns.find(c=>c.role==="Time") || null;
}
function buildElapsed(rows,timeCol){
  if(!timeCol) return {
    elapsed:[],
    startDate:null,
    endDate:null,
    basis:"unknown",
    intervalSec:null,
    warning:null
  };

  const vals=rows.map(r=>r[timeCol.index]);
  const intervalSec=extractIntervalSeconds(timeCol.header);
  const dates=vals.map(parseDateValue);
  const validDates=dates.filter(Boolean);

  if(validDates.length>=Math.min(3,Math.ceil(rows.length*0.2))){
    const times=validDates.map(d=>d.getTime());
    const minTime=Math.min(...times);
    const maxTime=Math.max(...times);
    const calendarDurationH=(maxTime-minTime)/3600000;

    if(intervalSec){
      const intervalDurationH=((rows.length-1)*intervalSec)/3600;

      // Excel often auto-converts ambiguous D/M/Y CSV dates.
      // Example: 10/8/26 -> Oct 8, 11/8/26 -> Nov 8, while 13/8/26 remains text.
      // If the timestamp span is very different from the acquisition interval implied
      // by the row count, use the declared logger interval for elapsed process time.
      const ratio = intervalDurationH>0 ? calendarDurationH/intervalDurationH : 1;
      const inconsistent =
        !Number.isFinite(calendarDurationH) ||
        calendarDurationH<=0 ||
        ratio>1.25 ||
        ratio<0.75;

      if(inconsistent){
        const elapsed=rows.map((_,i)=>(i*intervalSec)/3600);
        return {
          elapsed,
          startDate:null,
          endDate:null,
          basis:"acquisition interval",
          intervalSec,
          warning:`Timestamp dates are inconsistent with the ${intervalSec} s logging interval. Process time was calculated from row count × logging interval instead.`
        };
      }
    }

    const firstValid=dates.find(Boolean);
    const firstMs=firstValid.getTime();
    const elapsed=dates.map(d=>d?(d.getTime()-firstMs)/3600000:NaN);
    return {
      elapsed,
      startDate:firstValid,
      endDate:validDates[validDates.length-1],
      basis:"timestamps",
      intervalSec,
      warning:null
    };
  }

  const nums=vals.map(parseNum);
  const finite=nums.filter(Number.isFinite);
  if(finite.length){
    const start=finite[0];
    return {
      elapsed:nums.map(v=>Number.isFinite(v)?v-start:NaN),
      startDate:null,
      endDate:null,
      basis:"numeric time",
      intervalSec,
      warning:null
    };
  }

  if(intervalSec){
    return {
      elapsed:rows.map((_,i)=>(i*intervalSec)/3600),
      startDate:null,
      endDate:null,
      basis:"acquisition interval",
      intervalSec,
      warning:`Timestamp values could not be parsed. Process time was calculated from the ${intervalSec} s logging interval.`
    };
  }

  return {
    elapsed:[],
    startDate:null,
    endDate:null,
    basis:"unknown",
    intervalSec:null,
    warning:"Time values could not be interpreted."
  };
}
function statsForColumn(rows,col){
  const a=[];
  for(const r of rows){ const v=parseNum(r[col.index]); if(Number.isFinite(v)) a.push(v); }
  if(!a.length) return null;
  let sum=0,min=Infinity,max=-Infinity;
  for(const v of a){sum+=v;if(v<min)min=v;if(v>max)max=v;}
  return {n:a.length,mean:sum/a.length,min,max,last:a[a.length-1],first:a[0]};
}
function safeText(v){
  if(v instanceof Date) return v.toLocaleString();
  return String(v??"");
}

async function analyzeProcessFile(){
  const file=document.getElementById("processFile").files[0];
  if(!file){ document.getElementById("fileSummary").innerHTML='<span style="color:#9c1c1c">Choose a file first.</span>'; return; }
  try{
    document.getElementById("fileSummary").innerHTML="Reading file…";
    let parsed;
    if(/\.(xlsx|xls)$/i.test(file.name)){
      parsed=parseWorkbook(await file.arrayBuffer());
    }else{
      parsed=parseCSVText(await file.text());
    }
    const columns=buildColumns(parsed.headers);
    const timeCol=findTimeColumn(columns);
    if(!timeCol){
      throw new Error("No time/timestamp column was recognized. The detected headers are shown below for manual review.");
    }
    const timing=buildElapsed(parsed.rows,timeCol);
    const validElapsed=timing.elapsed.filter(Number.isFinite);
    if(!validElapsed.length) throw new Error(`The time column "${timeCol.header}" was found, but its values could not be converted to elapsed time.`);
    const duration=Math.max(...validElapsed)-Math.min(...validElapsed);
    const numericCols=columns.filter(c=>c.role!=="Time" && statsForColumn(parsed.rows,c));
    processState={fileName:file.name,headers:parsed.headers,rows:parsed.rows,columns,timeCol,timing,duration,numericCols};
    renderProcessSummary();
    renderDetectedColumns();
    fillMetricSelect();
    renderProcessTable();
    drawSelectedMetric();
  }catch(err){
    document.getElementById("fileSummary").innerHTML=`<span style="color:#9c1c1c">${err.message}</span>`;
  }
}
function renderProcessSummary(){
  const s=processState;
  const important=["DO","pH","Temperature","Stirrer","OUR","CER","RQ","Feed / SUBS_A"];
  const detected=s.columns.filter(c=>important.includes(c.role) || c.role.startsWith("BlueVary vol")).length;
  const durationDays=Math.floor(s.duration/24);
  const remainderH=s.duration-durationDays*24;
  const durationHours=Math.floor(remainderH);
  const durationMinutes=Math.round((remainderH-durationHours)*60);

  let html=`<div class="analysis-status">Time detected: ${escapeHTML(s.timeCol.label)}</div>
    <div class="summary-kpis">
      <div><strong>${s.rows.length.toLocaleString()}</strong><span>Data rows</span></div>
      <div><strong>${fmt(s.duration,2)} h</strong><span>Process duration (${durationDays} d ${durationHours} h ${durationMinutes} min)</span></div>
      <div><strong>${s.columns.length}</strong><span>Total columns</span></div>
      <div><strong>${detected}</strong><span>Key signals detected</span></div>
    </div>
    <div class="time-basis"><strong>Time basis:</strong> ${escapeHTML(s.timing.basis)}${s.timing.intervalSec?` (${s.timing.intervalSec} s interval)`:""}</div>
    ${s.timing.warning?`<div class="timestamp-warning"><strong>Timestamp correction:</strong> ${escapeHTML(s.timing.warning)}</div>`:""}`;
  document.getElementById("fileSummary").innerHTML=html;
}
function renderDetectedColumns(){
  const wrap=document.getElementById("detectedColumns");
  wrap.innerHTML="";
  processState.columns.filter(c=>c.role!=="Other").forEach(c=>{
    const span=document.createElement("span");
    span.className="detected-chip";
    span.innerHTML=`<strong>${escapeHTML(c.role)}</strong> ← ${escapeHTML(c.label)}`;
    wrap.appendChild(span);
  });
}
function fillMetricSelect(){
  const sel=document.getElementById("metricSelect");
  sel.innerHTML="";
  const cols=[...processState.numericCols].sort((a,b)=>b.priority-a.priority);
  cols.forEach(c=>{
    const op=document.createElement("option");
    op.value=String(c.index);
    op.textContent=`${c.role==="Other"?"":c.role+" — "}${c.label}`;
    sel.appendChild(op);
  });
  const preferred=cols.find(c=>c.role==="DO") || cols.find(c=>c.role.startsWith("BlueVary vol")) || cols[0];
  if(preferred) sel.value=String(preferred.index);
}
function drawSelectedMetric(){
  if(!processState) return clearCanvas("Analyze a file to display a process signal.");
  const index=parseInt(document.getElementById("metricSelect").value,10);
  const col=processState.columns.find(c=>c.index===index);
  if(!col) return clearCanvas("Select a process signal.");
  const pts=[];
  for(let i=0;i<processState.rows.length;i++){
    const x=processState.timing.elapsed[i],y=parseNum(processState.rows[i][index]);
    if(Number.isFinite(x)&&Number.isFinite(y)) pts.push({x,y});
  }
  if(pts.length<2) return clearCanvas("Not enough numeric data for this signal.");
  const maxPts=2200, step=Math.max(1,Math.ceil(pts.length/maxPts));
  const sampled=pts.filter((_,i)=>i%step===0 || i===pts.length-1);
  drawLine(sampled,`${col.role==="Other"?col.header:col.role} (${col.header})`);
}
function clearCanvas(msg){
  const canvas=document.getElementById("trendCanvas"),ctx=canvas.getContext("2d");
  ctx.clearRect(0,0,canvas.width,canvas.height);ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);
  ctx.fillStyle="#637083";ctx.font="18px sans-serif";ctx.fillText(msg,40,60);
}
function drawLine(pts,title){
  const canvas=document.getElementById("trendCanvas"),ctx=canvas.getContext("2d");
  const W=canvas.width,H=canvas.height,p={l:80,r:30,t:52,b:58};
  ctx.clearRect(0,0,W,H);ctx.fillStyle="#fff";ctx.fillRect(0,0,W,H);
  let xmin=pts[0].x,xmax=pts[pts.length-1].x,ymin=Infinity,ymax=-Infinity;
  pts.forEach(q=>{if(q.y<ymin)ymin=q.y;if(q.y>ymax)ymax=q.y;});
  if(xmax===xmin)xmax=xmin+1;if(ymax===ymin){ymax+=1;ymin-=1;}
  const margin=(ymax-ymin)*.08;ymin-=margin;ymax+=margin;
  const X=v=>p.l+(v-xmin)/(xmax-xmin)*(W-p.l-p.r);
  const Y=v=>H-p.b-(v-ymin)/(ymax-ymin)*(H-p.t-p.b);
  ctx.strokeStyle="#e1e7ef";ctx.lineWidth=1;
  for(let i=0;i<=5;i++){
    const yy=p.t+i*(H-p.t-p.b)/5;
    ctx.beginPath();ctx.moveTo(p.l,yy);ctx.lineTo(W-p.r,yy);ctx.stroke();
    const val=ymax-i*(ymax-ymin)/5;ctx.fillStyle="#637083";ctx.font="12px sans-serif";ctx.fillText(fmt(val,3),8,yy+4);
  }
  ctx.fillStyle="#142033";ctx.font="bold 15px sans-serif";ctx.fillText(title,p.l,p.t-22);
  ctx.fillStyle="#637083";ctx.font="13px sans-serif";ctx.fillText("Elapsed time (h)",W/2-45,H-16);
  ctx.strokeStyle="#155eef";ctx.lineWidth=2.5;ctx.beginPath();
  pts.forEach((q,i)=>{ const xx=X(q.x),yy=Y(q.y); if(i===0)ctx.moveTo(xx,yy);else ctx.lineTo(xx,yy); });
  ctx.stroke();
}
function renderProcessTable(){
  const table=document.getElementById("processTable"),thead=table.querySelector("thead"),tbody=table.querySelector("tbody");
  const cols=processState.columns;
  thead.innerHTML="<tr>"+cols.map(c=>`<th>${escapeHTML(c.label)}</th>`).join("")+"</tr>";
  tbody.innerHTML="";
  processState.rows.slice(0,100).forEach(r=>{
    const tr=document.createElement("tr");
    tr.innerHTML=cols.map(c=>`<td>${escapeHTML(safeText(r[c.index]))}</td>`).join("");
    tbody.appendChild(tr);
  });
}
function exportProcessSummary(){
  if(!processState){alert("Analyze a file first.");return;}
  const lines=[["source_file",processState.fileName],["time_column",processState.timeCol.header],["data_rows",processState.rows.length],["duration_h",processState.duration],[]];
  lines.push(["role","source_column","n","mean","min","max","first","last"]);
  processState.columns.filter(c=>c.role!=="Other"&&c.role!=="Time").forEach(c=>{
    const st=statsForColumn(processState.rows,c);
    if(st) lines.push([c.role,c.label,st.n,st.mean,st.min,st.max,st.first,st.last]);
  });
  const csv=lines.map(row=>row.map(csvCell).join(",")).join("\n");
  downloadBlob(csv,"bioprocess_process_summary.csv","text/csv");
}
function csvCell(x){
  const s=String(x??"");
  return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s;
}
function escapeHTML(s){
  return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
}
function downloadInstrumentSample(){
  const csv=`TimeStamp(UTC+02:00 interval=5),STIRRER(rpm),TEMP(C),PH(pH),DO(%sat),AIR(ccm),SUBS_A(ml),BLUEVARY_1(vol.%),BLUEVARY_2(vol.%),BLUEVARY_OUR(mol/l.h),BLUEVARY_CER(mol/l.h),BLUEVARY_RQ(RQ)
2026-09-04 08:00:00,600,30,5.0,95,1000,0,0.5,20.9,0.01,0.01,1.0
2026-09-04 08:05:00,620,30,5.0,88,1000,0,0.6,20.8,0.02,0.018,0.9
2026-09-04 08:10:00,650,30,5.0,75,1000,1,0.7,20.5,0.03,0.028,0.93`;
  downloadBlob(csv,"instrument_export_sample.csv","text/csv");
}
function downloadBlob(content,name,type){
  const blob=new Blob([content],{type}),url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
}
calcGrowth();calcYxs();calcYps();calcProductivity();calcCarbonGlucose();calcCarbonCO2();
