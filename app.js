
let lastAnalysis = null;

function n(id){
  return parseFloat(document.getElementById(id).value);
}
function fmt(x, digits=4){
  return Number.isFinite(x) ? Number(x).toLocaleString(undefined,{maximumFractionDigits:digits}) : "—";
}
function setResult(id, html, ok=true){
  const el=document.getElementById(id);
  el.innerHTML=html;
  el.style.background=ok ? "#eaf1ff" : "#fff0f0";
  el.style.color=ok ? "#103d9f" : "#9c1c1c";
}

function calcGrowth(){
  const x1=n("gx1"), x2=n("gx2"), t1=n("gt1"), t2=n("gt2");
  const dt=t2-t1;
  if(!(x1>0 && x2>0 && dt>0)) return setResult("growthResult","Check inputs: X₁, X₂ must be > 0 and t₂ > t₁.",false);
  const mu=(Math.log(x2)-Math.log(x1))/dt;
  const td=mu>0 ? Math.log(2)/mu : NaN;
  setResult("growthResult",`μ = ${fmt(mu)} h⁻¹<br>Doubling time = ${fmt(td)} h`);
}
function calcYxs(){
  const x1=n("yx1"), x2=n("yx2"), s1=n("ys1"), s2=n("ys2");
  const ds=s1-s2;
  if(!(ds>0)) return setResult("yxsResult","Substrate consumed (S₁ − S₂) must be > 0.",false);
  const y=(x2-x1)/ds;
  setResult("yxsResult",`Yx/s = ${fmt(y)} g biomass / g substrate`);
}
function calcYps(){
  const p1=n("yp1"), p2=n("yp2"), s1=n("yps1"), s2=n("yps2");
  const ds=s1-s2;
  if(!(ds>0)) return setResult("ypsResult","Substrate consumed (S₁ − S₂) must be > 0.",false);
  const y=(p2-p1)/ds;
  setResult("ypsResult",`Yp/s = ${fmt(y)} g product / g substrate`);
}
function calcProductivity(){
  const p1=n("qp1"), p2=n("qp2"), t1=n("qt1"), t2=n("qt2");
  const dt=t2-t1;
  if(!(dt>0)) return setResult("qpResult","t₂ must be greater than t₁.",false);
  const q=(p2-p1)/dt;
  setResult("qpResult",`Qp = ${fmt(q)} g/L/h`);
}
function calcCarbonGlucose(){
  const g=n("glucoseMass");
  if(!(g>=0)) return setResult("carbonGlucoseResult","Enter a valid glucose mass.",false);
  const carbon=g*(72.066/180.156);
  setResult("carbonGlucoseResult",`Carbon in = ${fmt(carbon,3)} g C`);
}
function calcCarbonCO2(){
  const v=n("co2Vol"), vm=n("molarVol");
  if(!(v>=0 && vm>0)) return setResult("carbonCO2Result","Volume must be ≥ 0 and Vm > 0.",false);
  const carbon=(v/vm)*12.011;
  setResult("carbonCO2Result",`Carbon out = ${fmt(carbon,3)} g C`);
}

function parseCSV(text){
  const lines=text.replace(/\r/g,"").split("\n").filter(x=>x.trim()!=="");
  if(lines.length<2) throw new Error("CSV needs a header row and at least one data row.");
  const delimiter=(lines[0].split(";").length > lines[0].split(",").length) ? ";" : ",";
  const headers=lines[0].split(delimiter).map(h=>h.trim().toLowerCase());

  const aliases={
    time:["time","time_h","hours","hour","t"],
    biomass:["biomass","dcw","x","biomass_g_l","dcw_g_l"],
    substrate:["substrate","glucose","s","substrate_g_l","glucose_g_l"],
    product:["product","p","product_g_l","resveratrol"]
  };
  const index={};
  for(const key of Object.keys(aliases)){
    index[key]=headers.findIndex(h=>aliases[key].includes(h));
  }
  if(index.time<0) throw new Error("Could not find a time column.");
  const rows=[];
  for(let i=1;i<lines.length;i++){
    const parts=lines[i].split(delimiter).map(v=>v.trim().replace(",","."));
    const row={};
    for(const key of Object.keys(index)){
      row[key]=index[key]>=0 ? parseFloat(parts[index[key]]) : NaN;
    }
    if(Number.isFinite(row.time)) rows.push(row);
  }
  if(rows.length<2) throw new Error("Need at least two valid time rows.");
  rows.sort((a,b)=>a.time-b.time);
  return rows;
}

function analyzeRows(rows){
  const first=rows[0], last=rows[rows.length-1];
  const dt=last.time-first.time;
  let mu=NaN, td=NaN, yxs=NaN, yps=NaN, qp=NaN;

  if(Number.isFinite(first.biomass) && Number.isFinite(last.biomass) && first.biomass>0 && last.biomass>0 && dt>0){
    mu=(Math.log(last.biomass)-Math.log(first.biomass))/dt;
    td=mu>0 ? Math.log(2)/mu : NaN;
  }
  if(Number.isFinite(first.substrate) && Number.isFinite(last.substrate)){
    const ds=first.substrate-last.substrate;
    if(ds>0 && Number.isFinite(first.biomass) && Number.isFinite(last.biomass)) yxs=(last.biomass-first.biomass)/ds;
    if(ds>0 && Number.isFinite(first.product) && Number.isFinite(last.product)) yps=(last.product-first.product)/ds;
  }
  if(Number.isFinite(first.product) && Number.isFinite(last.product) && dt>0) qp=(last.product-first.product)/dt;

  return {rows, first,last,dt,mu,td,yxs,yps,qp};
}

async function analyzeCSV(){
  const file=document.getElementById("csvFile").files[0];
  if(!file){
    document.getElementById("csvSummary").innerHTML='<span style="color:#9c1c1c">Choose a CSV file first.</span>';
    return;
  }
  try{
    const text=await file.text();
    const rows=parseCSV(text);
    const a=analyzeRows(rows);
    lastAnalysis=a;
    renderSummary(a);
    renderTable(rows);
    drawTrend(rows);
  }catch(err){
    document.getElementById("csvSummary").innerHTML=`<span style="color:#9c1c1c">${err.message}</span>`;
  }
}

function renderSummary(a){
  document.getElementById("csvSummary").innerHTML=`
    <div class="summary-grid">
      <div><strong>${fmt(a.dt,3)} h</strong><span>Analyzed duration</span></div>
      <div><strong>${fmt(a.mu)} h⁻¹</strong><span>Specific growth rate μ</span></div>
      <div><strong>${fmt(a.td)} h</strong><span>Doubling time</span></div>
      <div><strong>${fmt(a.yxs)} g/g</strong><span>Biomass yield Yx/s</span></div>
      <div><strong>${fmt(a.yps)} g/g</strong><span>Product yield Yp/s</span></div>
      <div><strong>${fmt(a.qp)} g/L/h</strong><span>Volumetric productivity</span></div>
    </div>`;
}

function renderTable(rows){
  const tbody=document.querySelector("#dataTable tbody");
  tbody.innerHTML="";
  rows.slice(0,250).forEach(r=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`<td>${fmt(r.time,4)}</td><td>${fmt(r.biomass,4)}</td><td>${fmt(r.substrate,4)}</td><td>${fmt(r.product,4)}</td>`;
    tbody.appendChild(tr);
  });
}

function drawTrend(rows){
  const canvas=document.getElementById("trendCanvas");
  const ctx=canvas.getContext("2d");
  const W=canvas.width,H=canvas.height,pad={l:70,r:30,t:32,b:55};
  ctx.clearRect(0,0,W,H);
  ctx.fillStyle="#ffffff"; ctx.fillRect(0,0,W,H);

  const times=rows.map(r=>r.time).filter(Number.isFinite);
  const series=[
    {key:"biomass",name:"Biomass"},
    {key:"substrate",name:"Substrate"},
    {key:"product",name:"Product"}
  ].filter(s=>rows.some(r=>Number.isFinite(r[s.key])));

  if(times.length<2 || series.length===0){
    ctx.fillStyle="#637083"; ctx.font="18px sans-serif"; ctx.fillText("No plottable data.",40,60); return;
  }
  const xmin=Math.min(...times), xmax=Math.max(...times);
  const ys=[];
  series.forEach(s=>rows.forEach(r=>{ if(Number.isFinite(r[s.key])) ys.push(r[s.key]); }));
  let ymin=Math.min(...ys), ymax=Math.max(...ys);
  if(ymin===ymax){ymin-=1;ymax+=1;}
  const x=v=>pad.l+(v-xmin)/(xmax-xmin)*(W-pad.l-pad.r);
  const y=v=>H-pad.b-(v-ymin)/(ymax-ymin)*(H-pad.t-pad.b);

  ctx.strokeStyle="#dfe5ee"; ctx.lineWidth=1;
  for(let i=0;i<=5;i++){
    const yy=pad.t+i*(H-pad.t-pad.b)/5;
    ctx.beginPath();ctx.moveTo(pad.l,yy);ctx.lineTo(W-pad.r,yy);ctx.stroke();
  }
  ctx.fillStyle="#637083";ctx.font="14px sans-serif";
  ctx.fillText("Time (h)",W/2-25,H-14);
  ctx.save();ctx.translate(20,H/2);ctx.rotate(-Math.PI/2);ctx.fillText("Concentration (g/L)",-65,0);ctx.restore();

  const colors=["#155eef","#0b7a53","#c77612"];
  series.forEach((s,si)=>{
    ctx.strokeStyle=colors[si%colors.length];ctx.lineWidth=3;ctx.beginPath();
    let started=false;
    rows.forEach(r=>{
      if(!Number.isFinite(r[s.key])) return;
      const xx=x(r.time),yy=y(r[s.key]);
      if(!started){ctx.moveTo(xx,yy);started=true;}else ctx.lineTo(xx,yy);
    });
    ctx.stroke();
  });

  let lx=pad.l;
  series.forEach((s,si)=>{
    ctx.fillStyle=colors[si%colors.length];ctx.fillRect(lx,pad.t-18,14,4);
    ctx.fillStyle="#142033";ctx.fillText(s.name,lx+20,pad.t-11);lx+=110;
  });
}

function downloadSample(){
  const csv=`time,biomass,substrate,product
0,0.8,20,0
2,1.3,18.2,0.08
4,2.1,15.6,0.22
6,3.4,12.1,0.55
8,5.1,8.4,1.05
10,6.2,5.9,1.65`;
  downloadBlob(csv,"bioprocess_sample.csv","text/csv");
}

function exportSummary(){
  if(!lastAnalysis){alert("Analyze a CSV first.");return;}
  const a=lastAnalysis;
  const csv=`metric,value,unit
duration,${a.dt},h
specific_growth_rate,${a.mu},1/h
doubling_time,${a.td},h
biomass_yield_yxs,${a.yxs},g/g
product_yield_yps,${a.yps},g/g
volumetric_productivity,${a.qp},g/L/h`;
  downloadBlob(csv,"bioprocess_summary.csv","text/csv");
}

function downloadBlob(content,name,type){
  const blob=new Blob([content],{type});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
  URL.revokeObjectURL(url);
}

calcGrowth();
calcYxs();
calcYps();
calcProductivity();
calcCarbonGlucose();
calcCarbonCO2();
