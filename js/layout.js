"use strict";
/* ===== Навигаци (edoc маягийн цэс) ===== */
const PAGES={home:["homeBox","Нүүр"],dashboard:["dashboardBox","Удирдлагын Dashboard"],reportForm:["reportFormBox","Тайлан бүртгэх"],reports:["reportListBox","Тайлангууд"],review:["reviewBox","Тайлан хянах"],trainingForm:["trFormBox","Сургалт бүртгэх"],trainingList:["trListBox","Сургалтын бүртгэлүүд"],stats:["statBox","Сургалтын статистик"],plan:["planBox","Сургалтын төлөвлөгөө"]};
const pageAllowed=k=>{const el=$(PAGES[k][0]);return el&&!el.classList.contains("hidden")};
function refreshNav(){document.querySelectorAll("#sideNav [data-go]").forEach(a=>a.style.display=pageAllowed(a.dataset.go)?"":"none")}
function go(k){if(!PAGES[k]||!pageAllowed(k))k="home";
 Object.entries(PAGES).forEach(([key,[id]])=>{const el=$(id);if(el)el.classList.toggle("page-off",key!==k)});
 document.querySelectorAll("#sideNav [data-go]").forEach(a=>a.classList.toggle("active",a.dataset.go===k));
 $("pageTitle").textContent=PAGES[k][1];if(location.hash!=="#"+k)history.replaceState(null,"","#"+k);
 $("appSection").classList.remove("nav-open");window.scrollTo(0,0);if(k==="home")renderHome()}
document.addEventListener("click",e=>{const t=e.target.closest("[data-go]");if(t){e.preventDefault();go(t.dataset.go)}
 const h=e.target.closest(".nav-head");if(h)h.parentElement.classList.toggle("open")});
$("navToggle").onclick=()=>$("appSection").classList.toggle("nav-open");$("navOverlay").onclick=()=>$("appSection").classList.remove("nav-open");
window.addEventListener("hashchange",()=>go(location.hash.slice(1)));
/* ===== Дүрс ===== */
const ICONS={ok:["#2ec4b6","M9 12l2 2 4-4"],work:["#6c5ce7","M12 7v5l3 2"],back:["#e74c3c","M9 9l6 6M15 9l-6 6"],draft:["#f39c12","M8 9h8M8 13h5"],people:["#3b5bdb","M9 11a2.5 2.5 0 1 0 0-5 2.5 2.5 0 0 0 0 5zM4.5 17c.6-2.4 2.3-3.5 4.5-3.5s3.9 1.1 4.5 3.5M15.5 11a2 2 0 1 0 0-4M16 13.6c1.6.3 2.8 1.4 3.3 3.4"]};
document.querySelectorAll("[data-ic]").forEach(s=>{const[c,d]=ICONS[s.dataset.ic];s.innerHTML=`<svg width="40" height="40" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="${c}22" stroke="${c}" stroke-width="1.6"/><path d="${d}" fill="none" stroke="${c}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>`});
/* ===== Нүүр хуудасны тооцоо ===== */
const homeCharts={};
function donutSVG(title,items){const W=640,H=320,cx=150,cy=165,r=105,sw=38,C=2*Math.PI*r,total=items.reduce((s,x)=>s+x.v,0);let off=0;
 const arcs=total?items.filter(x=>x.v>0).map(x=>{const len=x.v/total*C,s=`<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${x.c}" stroke-width="${sw}" stroke-dasharray="${len} ${C-len}" stroke-dashoffset="${-off}" transform="rotate(-90 ${cx} ${cy})"/>`;off+=len;return s}).join(""):`<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#e5e7eb" stroke-width="${sw}"/>`;
 const esc=s=>String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;");
 const leg=items.map((x,i)=>{const p=total?(x.v/total*100).toFixed(2):"0.00";return `<rect x="315" y="${70+i*34}" width="14" height="14" fill="${x.c}"/><text x="337" y="${82+i*34}" font-size="15" fill="#555">${esc(x.l)}: ${x.v} - ${p}%</text>`}).join("");
 return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" font-family="Segoe UI, Arial, sans-serif"><rect width="${W}" height="${H}" fill="#fff"/><title>${esc(title)}</title>${arcs}<text x="${cx}" y="${cy-4}" text-anchor="middle" font-size="30" font-weight="600" fill="#1f2937">${total}</text><text x="${cx}" y="${cy+20}" text-anchor="middle" font-size="13" fill="#64748b">нийт</text>${leg}</svg>`}
function renderHome(){if(!$("homeYear")||!currentUser)return;const y=Number($("homeYear").value||new Date().getFullYear());
 const R=(typeof reports!=="undefined"?reports:[]).filter(r=>Number(r.report_year)===y);
 const T=(typeof trainingRows!=="undefined"?trainingRows:[]).filter(r=>String(r.training_date||"").startsWith(String(y)));
 const rs=s=>R.filter(r=>(r.review_status||"Ноорог")===s).length,ts=s=>T.filter(r=>(r.review_status||"")===s).length;
 const r1=rs("Баталгаажсан"),r3=rs("Буцаасан"),r4=rs("Ноорог"),r2=R.length-r1-r3-r4;
 const t1=ts("Баталгаажсан"),t3=ts("Буцаасан"),t2=T.length-t1-t3,ppl=T.reduce((s,r)=>s+Number(r.participant_total||0),0);
 $("homeRepTotal").textContent="Нийт тайлан "+R.length;$("homeTrTotal").textContent="Нийт сургалт "+T.length;
 [["hR1",r1],["hR2",r2],["hR3",r3],["hR4",r4],["hT1",t1],["hT2",t2],["hT3",t3],["hT4",ppl.toLocaleString()]].forEach(([id,v])=>$(id).textContent=v);
 homeCharts.repChart={title:"Тайлангийн хяналт "+y,svg:donutSVG("Тайлангийн хяналт",[{l:"Баталгаажсан",v:r1,c:"#66cc33"},{l:"Хянагдаж байгаа",v:r2,c:"#f1d230"},{l:"Буцаагдсан",v:r3,c:"#c0254f"},{l:"Ноорог",v:r4,c:"#e8892b"}])};
 homeCharts.trChart={title:"Сургалтын баталгаажилт "+y,svg:donutSVG("Сургалтын баталгаажилт",[{l:"Баталгаажсан",v:t1,c:"#66cc33"},{l:"Хянагдаж байгаа",v:t2,c:"#f1d230"},{l:"Буцаагдсан",v:t3,c:"#c0254f"}])};
 $("repChart").innerHTML=homeCharts.repChart.svg;$("trChart").innerHTML=homeCharts.trChart.svg}
(()=>{const s=$("homeYear");for(let y=new Date().getFullYear();y>=2020;y--)s.add(new Option(y+" он",y));s.onchange=renderHome})();
/* ===== Графикийн экспорт: Print / PNG / JPEG / PDF / SVG ===== */
function closeMenus(){document.querySelectorAll(".menu").forEach(m=>m.remove())}
function svgToCanvas(svg,bg){return new Promise((res,rej)=>{const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:"image/svg+xml;charset=utf-8"}));img.onload=()=>{const c=document.createElement("canvas");c.width=1280;c.height=640;const g=c.getContext("2d");g.fillStyle=bg||"#fff";g.fillRect(0,0,c.width,c.height);g.drawImage(img,0,0,c.width,c.height);URL.revokeObjectURL(url);res(c)};img.onerror=rej;img.src=url})}
function dl(href,name){const a=document.createElement("a");a.href=href;a.download=name;document.body.appendChild(a);a.click();a.remove()}
async function exportChart(id,kind){const ch=homeCharts[id];if(!ch)return;const name=ch.title.replace(/\s+/g,"_");
 try{if(kind==="svg"){const u=URL.createObjectURL(new Blob([ch.svg],{type:"image/svg+xml"}));dl(u,name+".svg");setTimeout(()=>URL.revokeObjectURL(u),1000);return}
 if(kind==="print"){const w=window.open("","_blank");if(!w)return alert("Popup цонхыг зөвшөөрнө үү.");w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${ch.title}</title></head><body style="font-family:Arial;text-align:center"><h2>${ch.title}</h2>${ch.svg}<script>onload=()=>setTimeout(()=>print(),300)<\/script></body></html>`);w.document.close();return}
 const c=await svgToCanvas(ch.svg);
 if(kind==="png")return dl(c.toDataURL("image/png"),name+".png");
 if(kind==="jpeg")return dl(c.toDataURL("image/jpeg",0.95),name+".jpg");
 if(kind==="pdf"){if(!window.jspdf)throw Error("PDF сан ачаалагдсангүй");const pdf=new window.jspdf.jsPDF({orientation:"landscape",unit:"mm",format:"a4"});pdf.addImage(c.toDataURL("image/png"),"PNG",18,30,260,130);pdf.save(name+".pdf")}
 }catch(e){alert("Экспорт алдаа: "+safeError(e))}}
document.addEventListener("click",e=>{const b=e.target.closest(".menu-btn");if(b){e.stopPropagation();const open=b.parentElement.querySelector(".menu");closeMenus();if(open)return;
 const m=document.createElement("div");m.className="menu";m.innerHTML=`<button data-k="print">Print</button><hr><button data-k="png">PNG file</button><button data-k="jpeg">JPEG file</button><button data-k="pdf">PDF file</button><button data-k="svg">SVG file</button>`;
 m.onclick=ev=>{const k=ev.target.dataset.k;if(k){closeMenus();exportChart(b.dataset.chart,k)}};b.parentElement.appendChild(m);return}
 if(!e.target.closest(".menu"))closeMenus()});
/* ===== Хуучин функцуудыг шинэ бүтэцтэй холбох ===== */
{const _show=showApp;showApp=async function(){await _show();refreshNav();go(location.hash.slice(1)||"home")};
 const _lr=loadReports;loadReports=async function(){await _lr();renderHome()};
 const _lt=loadTrainingRecords;loadTrainingRecords=async function(){await _lt();renderHome()};
 const _er=editReport;editReport=function(id){_er(id);go("reportForm")};
 $("refreshBtn").onclick=()=>loadReports();$("trRefreshBtn").onclick=()=>loadTrainingRecords();}
