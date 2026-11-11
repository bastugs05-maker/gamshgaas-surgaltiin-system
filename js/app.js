"use strict";
const SUPABASE_URL="https://djzlkhvkpmhgxoykeozt.supabase.co";
const SUPABASE_KEY="sb_publishable_iBVTq-ruWQDXLWTEO4wqAA_PT1AKG3d";
const $=id=>document.getElementById(id);let sb=null,currentUser=null,profile={},reports=[],adminData=null,trainingRows=[],activeFileUrl="",activeFileName="";
const isReviewer=()=>["admin","reviewer"].includes(profile.role);
function message(id,t,c=""){
  const el=$(id);
  if(!el)return;
  el.textContent=t;
  el.className="message "+c;
}
function safeError(e){return e?.message||String(e)}
async function login(){const email=$("email").value.trim(),password=$("password").value;if(!email||!password)return message("loginStatus","И-мэйл, нууц үгээ оруулна уу.","error");$("loginBtn").disabled=true;message("loginStatus","Нэвтэрч байна...");try{if(!sb)throw Error("Supabase сан ачаалагдсангүй.");let x=await sb.auth.signInWithPassword({email,password});if(x.error)throw x.error;currentUser=x.data.user;await showApp()}catch(e){$("loginSection").classList.remove("hidden");$("appSection").classList.add("hidden");message("loginStatus","Алдаа: "+safeError(e),"error")}finally{$("loginBtn").disabled=false}}
async function logout(){if(sb)await sb.auth.signOut();currentUser=null;profile={};$("appSection").classList.add("hidden");$("loginSection").classList.remove("hidden");$("password").value="";message("loginStatus","Системээс гарлаа.","success")}
async function showApp(){let p=await sb.from("user_profiles").select("*").eq("user_id",currentUser.id).maybeSingle();if(p.error)throw p.error;if(!p.data)throw Error("user_profiles-д эрхийн мэдээлэл алга.");profile=p.data;$("userInfo").textContent=currentUser.email;
$("planBox").classList.toggle("hidden", profile.role !== "admin");
$("dashboardBox").classList.toggle("hidden",profile.role!=="admin");
if(profile.role==="admin"){ $("dashYear").innerHTML="";for(let yy=new Date().getFullYear();yy>=2020;yy--)$("dashYear").add(new Option(yy+" он",yy)); }

if (profile.role === "admin") {
  await initPlanForm();
}
$("reviewBox").classList.toggle("hidden",!isReviewer());$("loginSection").classList.add("hidden");$("appSection").classList.remove("hidden");initYears();await loadProvinces();await loadReports();fillTrainingTypes();await loadTrainingLocations();$("statYear").innerHTML="";for(let y=new Date().getFullYear();y>=2020;y--)$("statYear").add(new Option(y,y));let ps=await sb.from("provinces").select("id,name").order("id");if(ps.error)throw ps.error;(ps.data||[]).forEach(p=>$("statProvince").add(new Option(p.name,p.id)));await loadStatistics();await loadTrainingRecords();if(profile.role==="admin")await loadDashboard();if(isReviewer())await loadReviewQueue()}

function dashRange(y,p){if(p==="q1")return [y+"-01-01",y+"-03-31"];if(p==="q2")return [y+"-04-01",y+"-06-30"];if(p==="q3")return [y+"-07-01",y+"-09-30"];if(p==="q4")return [y+"-10-01",y+"-12-31"];if(p==="h1")return [y+"-01-01",y+"-06-30"];if(p==="h2")return [y+"-07-01",y+"-12-31"];return [y+"-01-01",y+"-12-31"]}
let dashboardExportData=null;
function exportDashboardExcel(){
 if(!dashboardExportData){message("dashStatus","Эхлээд Dashboard-ийг шинэчилнэ үү.","error");return}
 if(!window.XLSX){message("dashStatus","Excel сан ачаалагдсангүй. Интернетээ шалгана уу.","error");return}
 const d=dashboardExportData,wb=XLSX.utils.book_new();
 const summary=[["Удирдлагын Dashboard"],["Тайлангийн он",d.year],["Хугацаа",d.period],[],["Үзүүлэлт","Дүн"],["Нийт сургалт",d.summary.totalTrainings],["Нийт хамрагдсан хүн",d.summary.totalPeople],["Баталгаажсан сургалт",d.summary.approved],["Хүлээгдэж буй",d.summary.pending],["Тайлан ирүүлсэн нэгж",`${d.summary.submittedUnits}/${d.summary.totalUnits}`],["Тайлан ирүүлэлтийн хувь",d.summary.submissionRate.toFixed(1)+"%"],["Нийт тайлан",d.summary.totalReports],["Тайлан ирүүлээгүй байгууллага",d.summary.missingUnits],["Жилийн гүйцэтгэл",d.summary.yearPerformance]];
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),"Нэгдсэн дүн");
 XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d.provinceRows.map(x=>({"Аймаг":x.name,"Бодит сургалт":x.actual,"Хамрагдсан хүн":x.people,"Төлөвлөсөн сургалт":x.planned,"Биелэлт (%)":x.pct===null?"":Number(x.pct.toFixed(1))}))),"Сургалтын гүйцэтгэл");
 XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d.comparison.map((x,i)=>({"Эрэмбэ":i+1,"Аймаг/Нийслэл":x.name,"Тайлан":x.reports,"Сургалт":x.trainings,"Тайлан ирүүлсэн":x.hasReport?"Тийм":"Үгүй"}))),"Тайлангийн харьцуулалт");
 XLSX.writeFile(wb,`Dashboard_${d.year}.xlsx`);message("dashStatus","Excel файл татагдлаа.","success");
}
function exportDashboardPdf(){
 if(!dashboardExportData){message("dashStatus","Эхлээд Dashboard-ийг шинэчилнэ үү.","error");return}
 const d=dashboardExportData;
 const esc=v=>String(v??"").replace(/[&<>\"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'\"':"&quot;","'":"&#39;"}[c]));
 const table=(headers,rows)=>`<table><thead><tr>${headers.map(x=>`<th>${esc(x)}</th>`).join("")}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(x=>`<td>${esc(x)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
 const summary=table(["Үзүүлэлт","Дүн"],[
 ["Нийт сургалт",d.summary.totalTrainings],["Нийт хамрагдсан хүн",d.summary.totalPeople],
 ["Баталгаажсан сургалт",d.summary.approved],["Хүлээгдэж буй",d.summary.pending],
 ["Тайлан ирүүлсэн нэгж",`${d.summary.submittedUnits}/${d.summary.totalUnits}`],
 ["Тайлан ирүүлэлтийн хувь",d.summary.submissionRate.toFixed(1)+"%"],["Нийт тайлан",d.summary.totalReports],["Тайлан ирүүлээгүй байгууллага",d.summary.missingUnits],["Жилийн гүйцэтгэл",d.summary.yearPerformance]]);
 const provinces=table(["Аймаг/Нэгж","Сургалт","Хүн","Төлөвлөгөө","Биелэлт %"],d.provinceRows.map(x=>[x.name,x.actual,x.people,x.planned,x.pct===null?"—":x.pct.toFixed(1)+"%"]));
 const comparison=table(["Эрэмбэ","Аймаг/Нэгж","Тайлан","Сургалт","Төлөв"],d.comparison.map((x,i)=>[i+1,x.name,x.reports,x.trainings,x.hasReport?"Ирүүлсэн":"Ирүүлээгүй"]));
 const w=window.open("","_blank");if(!w){message("dashStatus","Popup цонхыг зөвшөөрөөд дахин оролдоно уу.","error");return}
 w.document.write(`<!doctype html><html lang="mn"><head><meta charset="UTF-8"><title>Dashboard ${d.year}</title><style>
 @page{size:A4 landscape;margin:10mm}body{font-family:"Noto Sans", "Arial", sans-serif;color:#172033;font-size:9.5pt;-webkit-print-color-adjust:exact;print-color-adjust:exact}h1{font-size:20pt;color:#164e86;margin:0 0 8px}h2{font-size:14pt;color:#164e86;margin:18px 0 8px}.meta{margin-bottom:16px;color:#444}table{width:100%;border-collapse:collapse;margin:8px 0 18px;page-break-inside:auto}thead{display:table-header-group}tr{page-break-inside:avoid}th{background:#15558a;color:white;text-align:left}th,td{border:1px solid #cbd5e1;padding:5px 7px}tbody tr:nth-child(even){background:#f1f5f9}.pagebreak{break-before:page;page-break-before:always}h1,h2{break-after:avoid}table{font-size:9pt} .note{margin-top:10px}.note{font-size:9pt;color:#666}@media screen{body{max-width:1100px;margin:24px auto}button{padding:10px 18px;margin-bottom:15px}}
 </style></head><body><button onclick="window.print()">PDF болгон хадгалах / хэвлэх</button><h1>Сургалтын нэгдсэн систем — Dashboard</h1><div class="meta">Тайлангийн он: ${esc(d.year)} | Хугацаа: ${esc(d.period)}</div><h2>Нэгдсэн үзүүлэлт</h2>${summary}<div class="pagebreak"></div><h2>Аймаг, нэгжийн сургалтын гүйцэтгэл</h2>${provinces}<div class="pagebreak"></div><h2>Тайлангийн харьцуулалт</h2>${comparison}<p class="note">Сургалтын нэгдсэн системээс автоматаар үүсгэв.</p><script>window.onload=()=>setTimeout(()=>window.print(),500)<\/script></body></html>`);
 w.document.close();message("dashStatus","Монгол кирилл фонттой PDF хэвлэх цонх нээгдлээ. Print цонхноос Save as PDF сонгоно уу.","success");
}

async function loadDashboard(){
 if(profile.role!=="admin")return;
 const y=Number($("dashYear").value||new Date().getFullYear()),[from,to]=dashRange(y,$("dashPeriod").value);
 message("dashStatus","Dashboard мэдээлэл тооцож байна...");
 try{
  const [rec,prov,plans,unitsRes]=await Promise.all([
   sb.from("training_records").select("province_id,unit_id,participant_total,review_status,training_date").gte("training_date",from).lte("training_date",to),
   sb.from("provinces").select("id,name").order("id"),
   sb.from("training_plans").select("province_id,planned_trainings").eq("plan_year",y),
   sb.from("units").select("id,name,province_id").order("name")
  ]);
  if(rec.error)throw rec.error;if(prov.error)throw prov.error;if(plans.error)throw plans.error;if(unitsRes.error)throw unitsRes.error;
  const all=rec.data||[],approved=all.filter(r=>r.review_status==="Баталгаажсан"),pending=all.filter(r=>["Илгээсэн","Хүлээгдэж буй","Хяналтад"].includes(r.review_status));
  if($("dashTrainings"))$("dashTrainings").textContent=all.length.toLocaleString();if($("dashPeople"))$("dashPeople").textContent=all.reduce((s,r)=>s+Number(r.participant_total||0),0).toLocaleString();if($("dashApproved"))$("dashApproved").textContent=approved.length.toLocaleString();if($("dashPending"))$("dashPending").textContent=pending.length.toLocaleString();
  const a={},people={},ap={},pl={};approved.forEach(r=>{a[r.province_id]=(a[r.province_id]||0)+1;ap[r.province_id]=(ap[r.province_id]||0)+1;people[r.province_id]=(people[r.province_id]||0)+Number(r.participant_total||0)});all.forEach(r=>{if(r.review_status!=="Баталгаажсан")a[r.province_id]=(a[r.province_id]||0)+0});(plans.data||[]).forEach(r=>pl[r.province_id]=(pl[r.province_id]||0)+Number(r.planned_trainings||0));
  const rows=(prov.data||[]).map(p=>({name:p.name,actual:ap[p.id]||0,people:people[p.id]||0,planned:pl[p.id]||0,pct:pl[p.id]?((ap[p.id]||0)/pl[p.id]*100):null}));
  if($("dashProvinceTable"))$("dashProvinceTable").innerHTML=rows.map(r=>`<tr><td>${r.name}</td><td>${r.actual}</td><td>${r.people}</td><td>${r.actual}</td><td>${r.pct===null?"—":r.pct.toFixed(1)+"%"}</td></tr>`).join("");
  const max=Math.max(1,...rows.map(r=>r.actual));if($("dashProvinceBars"))$("dashProvinceBars").innerHTML=rows.filter(r=>r.actual||r.planned).map(r=>`<div style="display:grid;grid-template-columns:120px 1fr 50px;gap:8px;align-items:center"><span>${r.name}</span><div style="height:16px;background:#e5eaf0;border-radius:5px"><div style="height:16px;width:${r.actual/max*100}%;background:#16834a;border-radius:5px"></div></div><b>${r.actual}</b></div>`).join("");
  if(arguments.length===0){} // keep this function compatible with existing refresh controls
  const rpt=arguments; // unused; report result is read from the query below
  const reportsResult=await sb.from("reports").select("province_id,unit_id,report_year,report_month,report_quarter,report_half,review_status,submitted_at").eq("report_year",y);
  if(reportsResult.error)throw reportsResult.error;
  const period=$("dashPeriod").value;
  const filteredReports=(reportsResult.data||[]).filter(r=>{if(period==="all")return true;if(period.startsWith("q"))return Number(r.report_quarter)===Number(period.slice(1));if(period.startsWith("h"))return Number(r.report_half)===Number(period.slice(1));return true});
  // Approved report performance table: only admin-confirmed reports count.
  const approvedReports=(reportsResult.data||[]).filter(r=>r.review_status==="Баталгаажсан");
  const perf={};approvedReports.forEach(r=>{let z=perf[r.province_id]||(perf[r.province_id]={h1:0,h2:0,year:0,total:0});if(Number(r.report_half)===1)z.h1++;else if(Number(r.report_half)===2)z.h2++;else if(!r.report_half&&!r.report_month&&!r.report_quarter)z.year++;z.total++;});
  const perfBody=$("approvedReportPerformance");if(perfBody)perfBody.innerHTML=(prov.data||[]).map(p=>{let z=perf[p.id]||{h1:0,h2:0,year:0,total:0};return `<tr><td>${p.name}</td><td>${z.h1}</td><td>${z.h2}</td><td>${z.year}</td><td><b>${z.total}</b></td></tr>`}).join("");
  // Тайлангийн нэгжийг province_id биш unit_id-аар тоолно.
  // Улаанбаатар province_id=22 мөн адил жагсаалтад орно.
  const submitted=filteredReports.filter(r=>r.submitted_at||["Илгээсэн","Хянаж байна","Баталгаажсан","Буцаасан","Засварласан"].includes(r.review_status));
  // Байгууллагын 32 мөрийг units хүснэгтийн ID-аар тооцно; provinces нь 22 аймаг/нийслэл.
  const reportCounts={},trainingCounts={},submittedUnitIds=new Set();
  submitted.forEach(r=>{
    if(r.unit_id!==null&&r.unit_id!==undefined){
      const k=String(r.unit_id);reportCounts[k]=(reportCounts[k]||0)+1;submittedUnitIds.add(k);
    }
  });
  all.forEach(r=>{if(r.unit_id!==null&&r.unit_id!==undefined){const k=String(r.unit_id);trainingCounts[k]=(trainingCounts[k]||0)+1;}});
  // ===== Харьцуулалтын байгууллагууд — Supabase-ийн report_orgs хүснэгтээс =====
  // kind: aimag -> аймгийн нэрээр (сумын бүртгэл аймагтаа тооцогдоно)
  //       special -> байгууллагын нэрээр (НОБГ, 116-р анги гэх мэт)
  //       district -> аймаг биш бүртгэлийн нэгжийн нэрээр (нийслэлийн дүүрэг)
  const FALLBACK_ORGS=[...["Архангай","Баян-Өлгий","Баянхонгор","Булган","Говь-Алтай","Говьсүмбэр","Дархан-Уул","Дорноговь","Дорнод","Дундговь","Завхан","Орхон","Өвөрхангай","Өмнөговь","Сүхбаатар","Сэлэнгэ","Төв","Увс","Ховд","Хөвсгөл","Хэнтий"].map(n=>({kind:"aimag",name:n,aliases:[n]})),
    ...["Багануур","Багахангай","Баянгол","Баянзүрх","Налайх","Сонгинохайрхан","Сүхбаатар","Хан-Уул","Чингэлтэй"].map(n=>({kind:"district",name:n+" дүүрэг",aliases:[n]})),
    {kind:"special",name:"Нийслэлийн Онцгой байдлын газар",aliases:["Нийслэлийн Онцгой байдлын газар"]},
    {kind:"special",name:"Гамшгаас хамгаалах сургалтын 116 дугаар анги",aliases:["Гамшгаас хамгаалах 116-р анги","Гамшгаас хамгаалах сургалтын 116 дугаар анги"]}];
  const orgRes=await sb.from("report_orgs").select("kind,name,aliases,sort_order").eq("active",true).order("sort_order").order("name");
  const ORGS=(!orgRes.error&&orgRes.data&&orgRes.data.length)?orgRes.data:FALLBACK_ORGS;
  if(orgRes.error)console.warn("report_orgs уншиж чадсангүй, кодын жагсаалтыг ашиглав:",orgRes.error.message);
  const lc=s=>String(s||"").toLocaleLowerCase("mn").replace(/[\s\-–—]+/g," ").trim();
  const stripAimag=n=>lc(n).replace(/\s*(аймгийн\s*онцгой\s*байдлын\s*газар|аймаг)$/,"").trim();
  const stripDist=n=>lc(n).replace(/\s*(дүүргийн\s*онцгой\s*байдлын\s*газар|дүүрэг)$/,"").trim();
  const cmpRows=new Map(),aimagIdx=new Map(),distIdx=new Map(),specials=[];
  ORGS.forEach((o,i)=>{const key="O"+i;cmpRows.set(key,{id:key,name:o.name,reports:0,trainings:0});
    const al=(o.aliases&&o.aliases.length?o.aliases:[o.name]);
    if(o.kind==="aimag")al.forEach(a=>aimagIdx.set(stripAimag(a),key));
    else if(o.kind==="district")al.forEach(a=>distIdx.set(stripDist(a),key));
    else specials.push({key,al:al.map(lc)});});
  const specialOf=n=>{const s=lc(n);if(!s)return null;const f=specials.find(x=>x.al.some(a=>s===a||s.includes(a)));return f?f.key:null};
  const provName={},unitName={};
  (prov.data||[]).forEach(p=>provName[String(p.id)]=p.name);
  (unitsRes.data||[]).forEach(u=>unitName[String(u.id)]=u.name);
  const rowKeyOf=r=>{
    const pn=provName[String(r.province_id)]||"",un=r.unit_id!=null?(unitName[String(r.unit_id)]||""):"";
    const a=aimagIdx.get(stripAimag(pn));if(a)return a;           // аймгийн аль ч сум
    const s=specialOf(un)||specialOf(pn);if(s)return s;           // тусгай байгууллага
    const d=distIdx.get(stripDist(un));if(d)return d;              // нийслэлийн дүүрэг
    return null;
  };
  submitted.forEach(r=>{const k=rowKeyOf(r);if(k)cmpRows.get(k).reports++});
  all.forEach(r=>{const k=rowKeyOf(r);if(k)cmpRows.get(k).trainings++});
  const comparison=[...cmpRows.values()].map(x=>({...x,hasReport:x.reports>0})).sort((a,b)=>b.reports-a.reports||b.trainings-a.trainings||a.name.localeCompare(b.name,"mn"));
  const submittedUnits=comparison.filter(x=>x.hasReport).length,totalUnits=comparison.length,rate=totalUnits?submittedUnits/totalUnits*100:0;
  const missingUnits=Math.max(0,totalUnits-submittedUnits);
  const totalPeople=all.reduce((s,r)=>s+Number(r.participant_total||0),0);
  if($("dashSubmittedUnits")) $("dashSubmittedUnits").textContent=`${submittedUnits} / ${totalUnits}`;
  if($("dashSubmissionRate")) $("dashSubmissionRate").textContent=rate.toFixed(1)+"%";
  const yearPerformance=`${all.length} сургалт / ${totalPeople} хүн`;
  if($("dashTotalReports"))$("dashTotalReports").textContent=submitted.length.toLocaleString();
  if($("dashSubmittedUnits"))$("dashSubmittedUnits").textContent=`${submittedUnits} / ${totalUnits}`;
  if($("dashMissingUnits"))$("dashMissingUnits").textContent=missingUnits.toLocaleString()+" байгууллага";
  if($("dashYearPerformance"))$("dashYearPerformance").textContent=yearPerformance;
  dashboardExportData={year:y,period:$("dashPeriod").selectedOptions[0].text,summary:{totalTrainings:all.length,totalPeople,approved:approved.length,pending:pending.length,submittedUnits,totalUnits,submissionRate:rate,totalReports:submitted.length,missingUnits,yearPerformance},provinceRows:rows,comparison};
  if($("cmpSubmitted"))$("cmpSubmitted").textContent=submittedUnits+" / "+totalUnits;if($('cmpRate'))$('cmpRate').textContent=rate.toFixed(1)+"%";if($('cmpReports'))$('cmpReports').textContent=submitted.length.toLocaleString();if($('cmpTrainings'))$('cmpTrainings').textContent=all.length.toLocaleString();
  const maxReports=Math.max(1,...comparison.map(r=>r.reports));if($('cmpBars'))$('cmpBars').innerHTML=comparison.map((r,i)=>`<div style="display:grid;grid-template-columns:30px minmax(90px,130px) 1fr 35px;gap:8px;align-items:center"><b>${i+1}</b><span>${r.name}</span><div style="height:14px;background:#e5eaf0;border-radius:5px;overflow:hidden"><div style="height:14px;width:${r.reports/maxReports*100}%;background:${r.hasReport?'#16834a':'#cbd5e1'}"></div></div><b>${r.reports}</b></div>`).join("");
  if($('cmpTable'))$('cmpTable').innerHTML=comparison.map((r,i)=>`<tr><td>${i+1}</td><td>${r.name}</td><td>${r.reports}</td><td>${r.trainings}</td><td style="color:${r.hasReport?'#16834a':'#c0392b'};font-weight:bold">${r.hasReport?'Ирүүлсэн':'Ирүүлээгүй'}</td></tr>`).join("");
  message("dashStatus",`${y} он ${$("dashPeriod").selectedOptions[0].text}: ${all.length} сургалт, ${all.reduce((s,r)=>s+Number(r.participant_total||0),0)} хүн.`, "success");
 }catch(e){message("dashStatus","Dashboard алдаа: "+safeError(e),"error")}
}

function initYears(){$("reportYear").innerHTML="";for(let y=new Date().getFullYear();y>=2020;y--)$("reportYear").add(new Option(y+" он",y));$("reportMonth").innerHTML="";for(let m=1;m<=12;m++)$("reportMonth").add(new Option(m+"-р сар",m))}


/* =========================================
   СУРГАЛТЫН ТӨЛӨВЛӨГӨӨ — УДИРДЛАГА
========================================= */

async function initPlanForm() {
  const provinceSelect = $("planProvince");
  const yearSelect = $("planYear");

  provinceSelect.innerHTML =
    '<option value="">Байгууллага сонгоно уу</option>';

  const { data, error } = await sb
    .from("provinces")
    .select("id,name")
    .order("id");

  if (error) {
    message("planStatus", error.message, "error");
    return;
  }

  (data || []).forEach(p => {
    provinceSelect.add(new Option(p.name, p.id));
  });

  yearSelect.innerHTML = "";

  for (let y = new Date().getFullYear(); y >= 2020; y--) {
    yearSelect.add(new Option(y + " он", y));
  }

  await loadPlans();
}

async function savePlan() {
  const provinceId = $("planProvince").value;
  const year = $("planYear").value;
  const period = $("planPeriod").value;
  const count = Number($("planCount").value);
  const editId = $("planEditId").value;

  if (!provinceId || !year) {
    return message(
      "planStatus",
      "Байгууллага, он сонгоно уу.",
      "error"
    );
  }

  if (!Number.isInteger(count) || count < 0) {
    return message(
      "planStatus",
      "Сургалтын тоо 0 буюу түүнээс их бүхэл тоо байна.",
      "error"
    );
  }

  const payload = {
    province_id: Number(provinceId),
    plan_year: Number(year),
    plan_period: period,
    planned_trainings: count,
    created_by: currentUser.id,
    updated_at: new Date().toISOString()
  };

  $("planSaveBtn").disabled = true;

  try {
    let result;

    if (editId) {
      result = await sb
        .from("training_plans")
        .update(payload)
        .eq("id", editId);
    } else {
      result = await sb
        .from("training_plans")
        .upsert(payload, {
          onConflict: "province_id,plan_year,plan_period"
        });
    }

    if (result.error) throw result.error;

    message(
      "planStatus",
      "Сургалтын төлөвлөгөө амжилттай хадгалагдлаа.",
      "success"
    );

    cancelPlanEdit();
    await loadPlans();

  } catch (e) {
    message(
      "planStatus",
      "Алдаа: " + safeError(e),
      "error"
    );
  } finally {
    $("planSaveBtn").disabled = false;
  }
}

async function loadPlans() {
  const { data, error } = await sb
    .from("training_plans")
    .select("*")
    .order("plan_year", { ascending: false })
    .order("plan_period");

  if (error) {
    message("planStatus", error.message, "error");
    return;
  }

  const plans = data || [];
  const ids = [...new Set(plans.map(p => p.province_id))];

  let names = {};

  if (ids.length) {
    const p = await sb
      .from("provinces")
      .select("id,name")
      .in("id", ids);

    (p.data || []).forEach(x => {
      names[x.id] = x.name;
    });
  }

  $("planTable").innerHTML = "";

  plans.forEach((p, i) => {
    const tr = document.createElement("tr");

    [
      i + 1,
      names[p.province_id] || "—",
      p.plan_year,
      p.plan_period,
      p.planned_trainings
    ].forEach(value => {
      const td = document.createElement("td");
      td.textContent = value ?? "";
      tr.appendChild(td);
    });

    const td = document.createElement("td");

    const edit = document.createElement("button");
    edit.textContent = "Засах";
    edit.onclick = () => editPlan(p.id);
    td.appendChild(edit);

    const del = document.createElement("button");
    del.textContent = "Устгах";
    del.className = "danger";
    del.onclick = () => deletePlan(p.id);
    td.appendChild(del);

    tr.appendChild(td);
    $("planTable").appendChild(tr);
  });

  message(
    "planStatus",
    "Нийт " + plans.length + " төлөвлөгөө.",
    "success"
  );
}

function editPlan(id) {
  sb.from("training_plans")
    .select("*")
    .eq("id", id)
    .single()
    .then(({ data, error }) => {
      if (error) {
        message("planStatus", error.message, "error");
        return;
      }

      $("planEditId").value = data.id;
      $("planProvince").value = data.province_id;
      $("planYear").value = data.plan_year;
      $("planPeriod").value = data.plan_period;
      $("planCount").value = data.planned_trainings;

      $("planSaveBtn").textContent = "Өөрчлөлт хадгалах";
      $("planCancelBtn").classList.remove("hidden");
    });
}

function cancelPlanEdit() {
  $("planEditId").value = "";
  $("planCount").value = 0;
  $("planSaveBtn").textContent = "Төлөвлөгөө хадгалах";
  $("planCancelBtn").classList.add("hidden");
}

async function deletePlan(id) {
  if (!confirm("Энэ төлөвлөгөөг устгах уу?")) return;

  const { error } = await sb
    .from("training_plans")
    .delete()
    .eq("id", id);

  if (error) {
    message("planStatus", error.message, "error");
    return;
  }

  message(
    "planStatus",
    "Төлөвлөгөө устгагдлаа.",
    "success"
  );

  await loadPlans();
}
async function loadProvinces() {
  const s = $("province");
  s.innerHTML = '<option value="">Аймаг / Нийслэл сонгоно уу</option>';

  let q = sb
    .from("provinces")
    .select("id,name")
    .order("id");

  if (profile.role !== "admin" && profile.province_id) {
    q = q.eq("id", profile.province_id);
  }

  const { data, error } = await q;

  if (error) {
    message("formStatus", "Аймаг унших алдаа: " + error.message, "error");
    return;
  }

  (data || []).forEach(p => {
    s.add(new Option(p.name, p.id));
  });

  if (profile.province_id) {
    s.value = String(profile.province_id);
    await loadUnits();
  }
}
async function loadUnits(){let pid=$("province").value,s=$("unit");s.innerHTML='<option value="">Сум / Дүүрэг / Байгууллага сонгоно уу</option>';if(!pid)return;let q=sb.from("units").select("id,name").eq("province_id",pid).order("name");if(["district","unit"].includes(profile.role))q=q.eq("id",profile.unit_id);let x=await q;if(x.error)throw x.error;
const units=x.data||[];
units.sort((a,b)=>{
  const priority=n=>n==="Нийслэлийн Онцгой байдлын газар"?0:n==="Гамшгаас хамгаалах 116-р анги"?1:2;
  return priority(a.name)-priority(b.name)||a.name.localeCompare(b.name,"mn");
});
units.forEach(u=>s.add(new Option(u.name,u.id)));if(profile.unit_id)s.value=profile.unit_id}
function changePeriod(){let p=$("period").value;["month","quarter","half"].forEach(k=>$(k+"Box").classList.toggle("hidden",p!==k))}
function periodText(r){if(r.report_month)return r.report_year+" / "+r.report_month+"-р сар";if(r.report_quarter)return r.report_year+" / "+r.report_quarter+"-р улирал";if(r.report_half)return r.report_year+" / "+r.report_half+"-р хагас";return r.report_year+" он"}
async function saveReport(){
 let pid=$("province").value,uid=$("unit").value,title=$("title").value.trim(),edit=$("editId").value,p=$("period").value;
 if(!pid||!uid||!title)return message("formStatus","Аймаг, сум, гарчгийг бөглөнө.","error");
 let old=reports.find(r=>String(r.id)===String(edit));
 if(old&&old.review_status==="Баталгаажсан")return message("formStatus","Батлагдсан тайланг засах боломжгүй.","error");
 let status=$("reportSubmitStatus").value,data={created_by:currentUser.id,province_id:+pid,unit_id:+uid,report_type:$("reportType").value,report_year:+$("reportYear").value,report_month:p==="month"?+$("reportMonth").value:null,report_quarter:p==="quarter"?+$("reportQuarter").value:null,report_half:p==="half"?+$("reportHalf").value:null,title,description:$("description").value.trim(),status:"Бүртгэсэн",review_status:status,submitted_at:status==="Илгээсэн"?new Date().toISOString():null};
 $("saveBtn").disabled=true;let newPath=null;
 try{
  let f=$("reportFile").files[0];
  if(f){
   if(f.size>10*1024*1024)throw Error("Файл 10MB-аас их");
   if(!/\.(pdf|doc|docx|xls|xlsx|jpg|jpeg|png)$/i.test(f.name))throw Error("Зөвшөөрөгдөх файлын төрөл биш");
   newPath=currentUser.id+"/reports/"+Date.now()+"_"+f.name.replace(/[^a-zA-Z0-9._-]/g,"_");
   let up=await sb.storage.from("training-files").upload(newPath,f,{upsert:false});
   if(up.error)throw up.error;
   data.report_file=newPath;
  }
  let result=edit?await sb.from("reports").update(data).eq("id",edit).eq("created_by",currentUser.id):await sb.from("reports").insert(data);
  if(result.error)throw result.error;
  // DB шинэчлэлт амжилттай болсон хойно хуучин файл устгана
  if(old&&f&&old.report_file){let rm=await sb.storage.from("training-files").remove([old.report_file]);if(rm.error)message("formStatus","Тайлан хадгалагдсан. Хуучин файл устгахад алдаа: "+rm.error.message,"error")}
  message("formStatus","Тайлан хадгалагдлаа.","success");cancelEdit();await loadReports();
 }catch(e){
  if(newPath)await sb.storage.from("training-files").remove([newPath]);
  message("formStatus","Алдаа: "+safeError(e),"error")
 }finally{$("saveBtn").disabled=false}
}
async function loadReports(){if(!currentUser)return;let q=sb.from("reports").select("*").order("created_at",{ascending:false});if(!isReviewer())q=q.eq("created_by",currentUser.id);let x=await q;if(x.error)return message("reportStatus",x.error.message,"error");reports=x.data||[];let ids=[...new Set(reports.map(r=>r.province_id))],units=[...new Set(reports.map(r=>r.unit_id))],pm={},um={};if(ids.length){let a=await sb.from("provinces").select("id,name").in("id",ids);(a.data||[]).forEach(v=>pm[v.id]=v.name)}if(units.length){let a=await sb.from("units").select("id,name").in("id",units);(a.data||[]).forEach(v=>um[v.id]=v.name)}$("reportTable").innerHTML="";reports.forEach((r,i)=>{let tr=document.createElement("tr");[i+1,(pm[r.province_id]||"")+" / "+(um[r.unit_id]||""),r.report_type,periodText(r),r.title,r.description,(r.review_status||"Ноорог")+(r.report_file?" | файл":"")].forEach(v=>{let td=document.createElement("td");td.textContent=v||"";tr.appendChild(td)});let td=document.createElement("td");if(profile.role==="admin"|| (isReviewer()&&r.review_status!=="Баталгаажсан")){let d=document.createElement("button");d.textContent="Устгах";d.className="danger";d.onclick=()=>deleteReport(r.id);td.appendChild(d)}else if(r.review_status!=="Баталгаажсан"&&r.created_by===currentUser.id){let b=document.createElement("button");b.textContent="Засах";b.onclick=()=>editReport(r.id);td.appendChild(b);let d=document.createElement("button");d.textContent="Устгах";d.className="danger";d.onclick=()=>deleteReport(r.id);td.appendChild(d)}if(r.report_file){let f=document.createElement("button");f.textContent="Файл нээх";f.onclick=()=>openFile(r.report_file);td.appendChild(f)}tr.appendChild(td);$("reportTable").appendChild(tr)});message("reportStatus","Нийт "+reports.length+" тайлан.","success")}
function editReport(id){let r=reports.find(v=>String(v.id)===String(id));if(!r)return;$("editId").value=r.id;$("reportType").value=r.report_type;$("period").value=r.report_month?"month":r.report_quarter?"quarter":r.report_half?"half":"year";changePeriod();$("reportYear").value=r.report_year;if(r.report_month)$("reportMonth").value=r.report_month;if(r.report_quarter)$("reportQuarter").value=r.report_quarter;if(r.report_half)$("reportHalf").value=r.report_half;$("province").value=r.province_id;loadUnits().then(()=>$("unit").value=r.unit_id);$("title").value=r.title;$("description").value=r.description||"";$("formTitle").textContent="Тайлан засах";$("saveBtn").textContent="Өөрчлөлт хадгалах";$("cancelBtn").classList.remove("hidden");if(r.review_status==="Буцаасан")$("reportSubmitStatus").value="Илгээсэн"}
function cancelEdit(){$("editId").value="";$("title").value="";$("description").value="";$("reportFile").value="";$("formTitle").textContent="Тайлан бүртгэх";$("saveBtn").textContent="Тайлан хадгалах";$("cancelBtn").classList.add("hidden");$("province").value="";$("unit").innerHTML='<option value="">Аймаг сонгоно уу</option>';$("period").value="month";changePeriod()}
async function deleteReport(id){
 let r=reports.find(v=>String(v.id)===String(id));if(!r)return;
 const admin=profile.role==="admin";
 const reviewer=isReviewer();
 if(!admin&&!reviewer&&(r.created_by!==currentUser.id||r.review_status==="Баталгаажсан"))return alert("Өөрийн баталгаажаагүй тайланг устгаж болно.");
 if(reviewer&&!admin&&r.review_status==="Баталгаажсан")return alert("Баталгаажсан тайланг зөвхөн админ устгана.");
 if(!confirm("Сонгосон тайлан болон хавсаргасан файлыг бүрмөсөн устгах уу?"))return;
 let q=sb.from("reports").delete().eq("id",id);if(!admin)q=q.eq("created_by",currentUser.id);
 let db=await q;
 if(db.error)return alert("Тайлан устгахад алдаа: "+db.error.message+"\\nSupabase RLS DELETE эрхийг шалгана уу.");
 if(r.report_file){let st=await sb.storage.from("training-files").remove([r.report_file]);if(st.error)alert("Тайлан устсан боловч Storage файл устгахад алдаа: "+st.error.message)}
 await loadReports();if(isReviewer())await loadReviewQueue();if(profile.role==="admin")await loadDashboard();
}
async function downloadFile(path){let x=await sb.storage.from("training-files").createSignedUrl(path,300);if(x.error)alert(x.error.message);else window.open(x.data.signedUrl,"_blank")}
async function loadReviewQueue(){
 try{
  if(!isReviewer())throw Error("Тайлан хянах эрх шаардлагатай.");
  message("reviewStatus","Ирсэн тайланг ачаалж байна...");
  // Илгээсэн болон буцааж засварлаад дахин илгээсэн тайланг бүгдийг харуулна.
  const statuses=["Илгээсэн","Хянаж байна","Засварласан","Хүлээгдэж буй"];
  const x=await sb.from("reports").select("*").in("review_status",statuses).order("submitted_at",{ascending:true});
  if(x.error)throw x.error;
  $("reviewTable").innerHTML="";
  for(const r of (x.data||[])){
   const [p,u]=await Promise.all([
    sb.from("provinces").select("name").eq("id",r.province_id).maybeSingle(),
    sb.from("units").select("name").eq("id",r.unit_id).maybeSingle()
   ]);
   const tr=document.createElement("tr");
   [(p.data?.name||"")+" / "+(u.data?.name||""),r.report_type,periodText(r),r.title,r.report_file?"Хавсаргасан":"—",r.review_status]
    .forEach(v=>{const td=document.createElement("td");td.textContent=v||"";tr.appendChild(td)});
   const td=document.createElement("td");
   if(r.report_file){const f=document.createElement("button");f.textContent="Файл нээх";f.onclick=()=>openFile(r.report_file);td.appendChild(f)}
   [["Баталгаажуулах","Баталгаажсан","green"],["Буцаах","Буцаасан","danger"],["Хяналтад авах","Хянаж байна","secondary"]]
    .forEach(([label,status,cls])=>{const b=document.createElement("button");b.textContent=label;b.className=cls;b.onclick=()=>reviewReport(r.id,status);td.appendChild(b)});
   if(profile.role==="admin"){const d=document.createElement("button");d.textContent="Устгах";d.className="danger";d.onclick=()=>deleteReport(r.id);td.appendChild(d)}
   tr.appendChild(td);$("reviewTable").appendChild(tr);
  }
  message("reviewStatus",(x.data||[]).length+" тайлан хяналтад байна.","success");
 }catch(e){message("reviewStatus","Тайлан хяналтын алдаа: "+safeError(e),"error")}
}
async function reviewReport(id,status){
 try{
  if(!isReviewer())throw Error("Баталгаажуулах эрх шаардлагатай.");
  const c=$("reviewComment").value.trim();
  if(status==="Буцаасан"&&!c){alert("Буцаах шалтгаан бичнэ үү");return}
  if(!confirm("Тайланг '"+status+"' төлөвт шилжүүлэх үү?"))return;
  message("reviewStatus","Төлөв шинэчилж байна...");
  const payload={review_status:status,review_comment:status==="Баталгаажсан"?"":c,reviewed_by:currentUser.id,reviewed_at:new Date().toISOString()};
  // UPDATE-г SELECT-тэй холбохгүй: UPDATE policy-оос гадна SELECT policy-оос шалтгаалж
  // баталгаажуулалт амжилттай болсон ч хоосон data буцах эрсдэлийг арилгана.
  const upd=await sb.from("reports").update(payload).eq("id",id);
  if(upd.error)throw upd.error;
  // Шинэ төлөвийг дахин уншиж бодитоор хадгалагдсан эсэхийг шалгана.
  const chk=await sb.from("reports").select("id,review_status").eq("id",id).maybeSingle();
  if(chk.error)throw chk.error;
  if(!chk.data||chk.data.review_status!==status)
    throw Error("Төлөв хадгалагдсангүй. reports хүснэгтийн UPDATE/SELECT RLS policy болон role (admin/reviewer)-ийг шалгана уу.");
  $("reviewComment").value="";
  message("reviewStatus","Амжилттай: "+chk.data.review_status,"success");
  await loadReviewQueue();await loadReports();
  if(profile.role==="admin")await loadDashboard();
 }catch(e){message("reviewStatus","Баталгаажуулах алдаа: "+safeError(e),"error")}
}

const trainingTypes=["Гамшгаас хамгаалах иж бүрэн сургууль","Гамшгаас хамгаалах Команд штабын сургууль","Удирдах албан тушаалтны","Сургуулийн ахлах ангийн сурагчдын","Сургуулийн дунд ангийн сурагчдын","Сургуулийн бага ангийн сурагчдын","Сургуулийн өмнөх боловсролын байгууллага","Гамшгаас хамгаалах алба, мэргэжлийн анги","Их, дээд сургууль, МСҮТ, коллежийн","Төрийн захиргаа болон хуулийн этгээдийн","Цэргийн дүйцүүлэх албаны","Сайн дурын иргэн, бүлгийн","Иргэдийн","Хөгжлийн бэрхшээлтэй иргэдийн"];
function fillTrainingTypes(){if($("trType").options.length===0)trainingTypes.forEach(t=>$("trType").add(new Option(t,t)))}

async function loadTrainingLocations() {
  const s = $("trProvince");

  s.innerHTML = '<option value="">Аймаг / Нийслэл сонгоно уу</option>';

  const { data, error } = await sb
    .from("provinces")
    .select("id,name")
    .order("id");

  if (error) {
    message("trStatus", "Аймаг унших алдаа: " + error.message, "error");
    return;
  }

  (data || []).forEach(p => {
    s.add(new Option(p.name, p.id));
  });

  if (profile.province_id) {
    s.value = String(profile.province_id);
  }

  await loadTrainingUnits();
}
async function loadTrainingUnits(){let pid=$("trProvince").value;$("trUnit").innerHTML='<option value="">Сум / Дүүрэг / Байгууллага сонгоно уу</option>';if(!pid)return;let q=sb.from("units").select("id,name").eq("province_id",pid).order("name");if(profile.unit_id)q=q.eq("id",profile.unit_id);let x=await q;if(x.error)throw x.error;
const units=x.data||[];
units.sort((a,b)=>{
  const priority=n=>n==="Нийслэлийн Онцгой байдлын газар"?0:n==="Гамшгаас хамгаалах 116-р анги"?1:2;
  return priority(a.name)-priority(b.name)||a.name.localeCompare(b.name,"mn");
});
units.forEach(u=>$("trUnit").add(new Option(u.name,u.id)));if(profile.unit_id)$("trUnit").value=profile.unit_id}
async function saveTraining(){let pid=$("trProvince").value,uid=$("trUnit").value,date=$("trDate").value,total=+$("trTotal").value,m=+$("trMale").value,f=+$("trFemale").value,topic=$("trTopic").value==="Бусад"?$("trTopicOther").value.trim():$("trTopic").value;if(!pid||!uid||!date||!$("trType").value||!$("trOrgName").value.trim()||!topic)return message("trStatus","Шаардлагатай талбаруудыг бөглөнө.","error");if(m+f>total)return message("trStatus","Эрэгтэй + эмэгтэй нийтээс их байна.","error");try{let path=null,file=$("trFile").files[0];if(file){if(file.size>10*1024*1024)throw Error("Файл 10MB-аас их");path=currentUser.id+"/training/"+Date.now()+"_"+file.name.replace(/[^a-zA-Z0-9._-]/g,"_");let up=await sb.storage.from("training-files").upload(path,file);if(up.error)throw up.error}let x=await sb.from("training_records").insert({created_by:currentUser.id,province_id:+pid,unit_id:+uid,training_date:date,training_type:$("trType").value,organization_type:$("trOrgType").value,organization_name:$("trOrgName").value.trim(),participant_total:total,male_count:m,female_count:f,training_format:$("trFormat").value,training_topic:topic,certificate_file:path,review_status:"Илгээсэн",submitted_at:new Date().toISOString()});if(x.error)throw x.error;message("trStatus","Бүртгэл илгээгдлээ.","success");$("trOrgName").value="";$("trFile").value="";await loadTrainingRecords();await loadStatistics()}catch(e){message("trStatus","Алдаа: "+safeError(e),"error")}}
async function loadTrainingRecords(){if(!currentUser)return;message("trListStatus","Сургалтын бүртгэл ачаалж байна...");let q=sb.from("training_records").select("*").order("created_at",{ascending:false});if(!isReviewer())q=q.eq("created_by",currentUser.id);let x=await q;if(x.error)return message("trListStatus","Алдаа: "+x.error.message,"error");trainingRows=x.data||[];let pids=[...new Set(trainingRows.map(r=>r.province_id))],uids=[...new Set(trainingRows.map(r=>r.unit_id))],pm={},um={};if(pids.length){let a=await sb.from("provinces").select("id,name").in("id",pids);(a.data||[]).forEach(v=>pm[v.id]=v.name)}if(uids.length){let a=await sb.from("units").select("id,name").in("id",uids);(a.data||[]).forEach(v=>um[v.id]=v.name)}$("trListTable").innerHTML="";trainingRows.forEach((r,i)=>{let tr=document.createElement("tr");[i+1,r.training_date,(pm[r.province_id]||"")+" / "+(um[r.unit_id]||""),r.training_type,r.organization_name,r.participant_total,r.review_status||"—"].forEach(v=>{let td=document.createElement("td");td.textContent=v??"";tr.appendChild(td)});let td=document.createElement("td");if(r.certificate_file){let b=document.createElement("button");b.textContent="Файл нээх";b.onclick=()=>openFile(r.certificate_file);td.appendChild(b)}else td.textContent="Хавсаргаагүй";tr.appendChild(td);
let act=document.createElement("td");
if(isReviewer()){
 if(r.review_status!=="Баталгаажсан"){
  let ok=document.createElement("button");ok.textContent="Баталгаажуулах";ok.className="green";ok.onclick=()=>reviewTrainingRecord(r.id,"Баталгаажсан");act.appendChild(ok);
  let back=document.createElement("button");back.textContent="Буцаах";back.className="secondary";back.onclick=()=>reviewTrainingRecord(r.id,"Буцаасан");act.appendChild(back);
 }
 let del=document.createElement("button");del.textContent="Устгах";del.className="danger";del.title="Энэ бүртгэлийг бүрмөсөн устгах";del.onclick=()=>deleteTrainingRecord(r.id,r.certificate_file);act.appendChild(del)
}else{act.textContent="—"}
tr.appendChild(act);$("trListTable").appendChild(tr)});message("trListStatus","Нийт "+trainingRows.length+" бүртгэл.","success")}
async function reviewTrainingRecord(id,status){
 if(!isReviewer())return message("trListStatus","Баталгаажуулах эрх шаардлагатай.","error");
 if(status==="Буцаасан"){const reason=prompt("Буцаах шалтгаан бичнэ үү");if(!reason)return}
 if(!confirm("Сургалтын бүртгэлийг "+status+" төлөвт шилжүүлэх үү?"))return;
 try{
  const payload={review_status:status,reviewed_by:currentUser.id,reviewed_at:new Date().toISOString()};
  const upd=await sb.from("training_records").update(payload).eq("id",id);if(upd.error)throw upd.error;
  const chk=await sb.from("training_records").select("id,review_status").eq("id",id).maybeSingle();if(chk.error)throw chk.error;
  if(!chk.data||chk.data.review_status!==status)throw Error("Төлөв DB-д хадгалагдсангүй. UPDATE/SELECT RLS policy шалгана уу.");
  message("trListStatus","Сургалтын бүртгэл "+status+" төлөвтэй боллоо.","success");
  await loadTrainingRecords();await loadStatistics();if(profile.role==="admin")await loadDashboard();
 }catch(e){message("trListStatus","Баталгаажуулах алдаа: "+safeError(e),"error")}
}
async function deleteTrainingRecord(id,filePath){
 if(!isReviewer())return message("trListStatus","Устгах эрх зөвхөн админ/хянагчид байна.","error");
 if(!confirm("Энэ сургалтын бүртгэл болон хавсаргасан файлыг бүрмөсөн устгах уу? Энэ үйлдлийг буцаах боломжгүй."))return;
 try{
  message("trListStatus","Устгаж байна...");
  // Storage file first; if storage deletion fails, retain the database row.
  if(filePath){let f=await sb.storage.from("training-files").remove([filePath]);if(f.error)throw f.error}
  let d=await sb.from("training_records").delete().eq("id",id);if(d.error)throw d.error;
  message("trListStatus","Бүртгэл устгагдлаа.","success");
  await loadTrainingRecords();await loadStatistics()
 }catch(e){message("trListStatus","Устгах алдаа: "+safeError(e)+" (Supabase DELETE/RLS эрхийг шалгана уу)","error")}
}
async function openFile(path){try{if(!path)throw Error("Файлын зам хоосон байна.");let x=await sb.storage.from("training-files").createSignedUrl(path,600);if(x.error)throw x.error;activeFileUrl=x.data.signedUrl;activeFileName=path.split("/").pop()||"file";let ext=activeFileName.split(".").pop().toLowerCase();$("fileModalTitle").textContent=activeFileName;$("filePreview").innerHTML="";$("fileModalStatus").textContent="";
if(["pdf"].includes(ext)){$("filePreview").innerHTML='<iframe class="preview" title="PDF preview"></iframe>';$("filePreview").firstChild.src=activeFileUrl}
else if(["jpg","jpeg","png","webp"].includes(ext)){$("filePreview").innerHTML='<img style="max-width:100%;max-height:68vh;display:block;margin:auto" alt="Баталгааны зураг">';$("filePreview").firstChild.src=activeFileUrl}
else if(["doc","docx"].includes(ext)){$("filePreview").innerHTML='<div style="padding:20px"><p>Word файлыг браузер дотор шууд үзүүлэхгүй. Доорх <b>Татах</b> товчоор татаж Word/WPS Office-оор нээнэ.</p><p>Татсан файл: '+activeFileName.replace(/[<>&"]/g,"")+'</p></div>'}
else {$("filePreview").textContent="Энэ файлын төрлийг урьдчилан харах боломжгүй. Татаж нээнэ үү."}
$("fileModal").classList.remove("hidden")}catch(e){alert("Файл нээхэд алдаа: "+safeError(e))}}
function downloadActiveFile(){if(!activeFileUrl)return;let a=document.createElement("a");a.href=activeFileUrl;a.download=activeFileName;a.target="_blank";document.body.appendChild(a);a.click();a.remove()}
function closeFileModal(){$("fileModal").classList.add("hidden");$("filePreview").innerHTML="";activeFileUrl=""}
async function loadStatistics(){
 let year=+$("statYear").value,prov=$("statProvince").value;if(!year)return;
 message("statStatus","Статистик тооцож байна...");
 try{
  let all=[],from=0,batch=1000;
  while(true){
   let q=sb.from("training_records").select("training_type,participant_total,male_count,female_count,province_id,training_date")
    .eq("review_status","Баталгаажсан").gte("training_date",year+"-01-01").lte("training_date",year+"-12-31")
    .order("training_date",{ascending:true}).range(from,from+batch-1);
   if(prov)q=q.eq("province_id",+prov);
   let x=await q;if(x.error)throw x.error;
   all=all.concat(x.data||[]);
   if(!x.data||x.data.length<batch)break;
   from+=batch;
  }
  let rows=trainingTypes.map(t=>{let a=all.filter(v=>v.training_type===t);return{t,n:a.length,p:a.reduce((s,v)=>s+(Number(v.participant_total)||0),0),m:a.reduce((s,v)=>s+(Number(v.male_count)||0),0),f:a.reduce((s,v)=>s+(Number(v.female_count)||0),0)}});
  $("statTable").innerHTML=rows.map((v,i)=>`<tr><td>${i+1}</td><td>${v.t}</td><td>${v.n}</td><td>${v.p}</td><td>${v.m}</td><td>${v.f}</td></tr>`).join("");
  let n=rows.reduce((s,v)=>s+v.n,0),p=rows.reduce((s,v)=>s+v.p,0),m=rows.reduce((s,v)=>s+v.m,0),f=rows.reduce((s,v)=>s+v.f,0);
  $("statTotals").innerHTML=`<h3>Нийт ${n} сургалт | ${p} хүн (Э ${m}, Эм ${f})</h3>`;
  message("statStatus",`Баталгаажсан сургалтын дүн. Нийт ${all.length} бүртгэл тооцлоо.`,"success")
 }catch(e){message("statStatus","Статистикийн алдаа: "+safeError(e),"error")}
}
$("dashRefresh").onclick=loadDashboard;$("exportExcel").onclick=exportDashboardExcel;$("exportPdf").onclick=exportDashboardPdf;$("dashYear").onchange=loadDashboard;$("dashPeriod").onchange=loadDashboard;
$("loginBtn").onclick=login;
$("planSaveBtn").onclick = savePlan;
$("planCancelBtn").onclick = cancelPlanEdit;
$("planRefreshBtn").onclick = loadPlans;

$("logoutBtn").onclick=logout;
$("refreshBtn").onclick=loadReports;
$("saveBtn").onclick=saveReport;
$("cancelBtn").onclick=cancelEdit;
$("province").onchange=loadUnits;
$("period").onchange=changePeriod;
$("trProvince").onchange=loadTrainingUnits;
$("trSaveBtn").onclick=saveTraining;
$("trRefreshBtn").onclick=loadTrainingRecords;
$("statRefresh").onclick=loadStatistics;
$("statYear").onchange=loadStatistics;
$("statProvince").onchange=loadStatistics;
$("reviewRefresh").onclick=loadReviewQueue;
$("fileDownloadBtn").onclick=downloadActiveFile;
$("fileCloseBtn").onclick=closeFileModal;
async function init(){try{if(!window.supabase)throw Error("Supabase JS CDN ачаалагдсангүй. Интернетээ шалгана уу.");sb=window.supabase.createClient(SUPABASE_URL,SUPABASE_KEY);let x=await sb.auth.getSession();if(x.error)throw x.error;if(x.data.session){currentUser=x.data.session.user;await showApp()}}catch(e){message("loginStatus","Системийн алдаа: "+safeError(e),"error")}}init();
