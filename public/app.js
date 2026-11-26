(function(){
"use strict";
const $ = (id)=>document.getElementById(id);
const state={tool:null,files:[],outputs:[],activities:[],filter:"all",search:"",busy:false,pageItems:[],previewToken:0,previewBusy:false,pageDrag:null,suppressPageClick:false};
const CDN={
 pdf:"https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js",
 pdfjs:"https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
 zip:"https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
 xlsx:"https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
 mammoth:"https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.8.0/mammoth.browser.min.js",
 htmlpdf:"https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js",
 tesseract:"https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js",
 heic:"https://cdn.jsdelivr.net/npm/heic2any@0.0.4/dist/heic2any.min.js"
};
const loaded={};
const tools=[
{id:"merge-pdf",name:"Gabungkan PDF",desc:"Satukan beberapa file PDF menjadi satu.",icon:"▣",color:"purple",cat:"organize",formats:".pdf",min:2,pop:true},
{id:"split-pdf",name:"Pisahkan PDF",desc:"Pisahkan setiap halaman atau rentang halaman.",icon:"⫶",color:"orange",cat:"organize",formats:".pdf",min:1,pop:true,opts:[{id:"mode",label:"Metode pemisahan",type:"select",value:"each",items:[["each","Satu file untuk setiap halaman"],["ranges","Berdasarkan rentang halaman"] ]},{id:"ranges",label:"Rentang halaman",type:"text",value:"1-2,3-5",placeholder:"Contoh: 1-2, 3-5"}]},
{id:"extract-pages",name:"Ekstrak halaman",desc:"Ambil halaman tertentu dari PDF.",icon:"▤",color:"blue",cat:"organize",formats:".pdf",opts:[{id:"pages",label:"Halaman yang ingin diambil",type:"text",value:"1,3-5",placeholder:"Contoh: 1, 3-5"}]},
{id:"remove-pages",name:"Hapus halaman",desc:"Hapus halaman yang tidak diperlukan.",icon:"▧",color:"pink",cat:"organize",formats:".pdf",opts:[{id:"pages",label:"Halaman yang ingin dihapus",type:"text",value:"2",placeholder:"Contoh: 2, 4-6"}]},
{id:"organize-pdf",name:"Atur urutan PDF",desc:"Seret thumbnail untuk mengubah urutan halaman.",icon:"☷",color:"purple",cat:"organize",formats:".pdf",opts:[{id:"orderFallback",label:"Urutan manual (opsional)",type:"text",value:"",placeholder:"Contoh: 3, 1, 2, 4"},{id:"orderHint",label:"",type:"note",text:"Utamakan drag-and-drop thumbnail. Urutan manual hanya digunakan jika pratinjau halaman tidak tersedia."}]},
{id:"rotate-pdf",name:"Putar PDF",desc:"Putar halaman PDF ke orientasi yang benar.",icon:"⟳",color:"blue",cat:"organize",formats:".pdf",opts:[{id:"degrees",label:"Sudut rotasi",type:"select",value:"90",items:[["90","90° searah jarum jam"],["180","180°"],["270","90° berlawanan jarum jam"]]},{id:"pages",label:"Halaman",type:"text",value:"all",placeholder:"all atau 1, 3-5"}]},
{id:"compress-pdf",name:"Kompres PDF",desc:"Optimasi struktur PDF untuk ukuran lebih ringkas.",icon:"⇣",color:"green",cat:"organize",formats:".pdf",opts:[{id:"level",label:"Mode optimasi",type:"select",value:"balanced",items:[["balanced","Seimbang"],["compact","Utamakan ukuran file"] ]},{id:"compressHint",label:"",type:"note",text:"Optimasi struktural tanpa mengurangi resolusi gambar. Penghematan ukuran bergantung pada isi PDF."}]},
{id:"repair-pdf",name:"Perbaiki PDF",desc:"Buka ulang dan tulis ulang struktur dokumen PDF.",icon:"⌁",color:"teal",cat:"organize",formats:".pdf"},
{id:"pdf-to-jpg",name:"PDF ke JPG",desc:"Ubah halaman PDF menjadi gambar JPG.",icon:"▧",color:"orange",cat:"convert",formats:".pdf",opts:[{id:"scale",label:"Kualitas gambar",type:"select",value:"1.6",items:[["1","Standar"],["1.6","Tinggi"],["2.2","Sangat tinggi"]] }]},
{id:"pdf-to-png",name:"PDF ke PNG",desc:"Ekspor setiap halaman PDF sebagai gambar PNG.",icon:"▧",color:"teal",cat:"convert",formats:".pdf",opts:[{id:"scale",label:"Resolusi gambar",type:"select",value:"1.6",items:[["1","Standar"],["1.6","Tinggi"],["2.2","Sangat tinggi"]] }]},
{id:"pdf-to-word",name:"PDF ke Word",desc:"Ekstrak teks PDF ke dokumen DOCX.",icon:"W",color:"blue",cat:"convert",formats:".pdf"},
{id:"pdf-to-excel",name:"PDF ke Excel",desc:"Ekstrak baris teks PDF ke lembar kerja XLSX.",icon:"X",color:"green",cat:"convert",formats:".pdf"},
{id:"pdf-to-text",name:"PDF ke TXT",desc:"Ambil teks dari setiap halaman PDF.",icon:"T",color:"purple",cat:"convert",formats:".pdf"},
{id:"pdf-to-html",name:"PDF ke HTML",desc:"Konversi teks PDF menjadi halaman HTML.",icon:"⌘",color:"teal",cat:"convert",formats:".pdf"},
{id:"word-to-pdf",name:"Word ke PDF",desc:"Konversi DOCX menjadi PDF dari konten dokumen.",icon:"W",color:"blue",cat:"convert",formats:".docx"},
{id:"excel-to-pdf",name:"Excel ke PDF",desc:"Ekspor lembar kerja Excel menjadi PDF.",icon:"X",color:"green",cat:"convert",formats:".xlsx,.xls,.csv"},
{id:"powerpoint-to-pdf",name:"PowerPoint ke PDF",desc:"Ekstrak konten slide PPTX dan susun menjadi PDF.",icon:"P",color:"orange",cat:"convert",formats:".pptx"},
{id:"images-to-pdf",name:"JPG ke PDF",desc:"Gabungkan gambar JPG, PNG, atau WEBP ke PDF.",icon:"▧",color:"pink",cat:"convert",formats:"image/*",min:1,pop:true},
{id:"watermark",name:"Watermark PDF",desc:"Tambahkan teks watermark pada halaman PDF.",icon:"◈",color:"purple",cat:"edit",formats:".pdf",opts:[{id:"text",label:"Teks watermark",type:"text",value:"CONFIDENTIAL",placeholder:"Tulis watermark"},{id:"opacity",label:"Transparansi",type:"select",value:"0.22",items:[["0.12","Sangat tipis"],["0.22","Sedang"],["0.4","Jelas"]]},{id:"size",label:"Ukuran teks",type:"number",value:"42"}]},
{id:"page-numbers",name:"Nomor halaman",desc:"Tambahkan nomor halaman pada PDF.",icon:"№",color:"blue",cat:"edit",formats:".pdf",opts:[{id:"position",label:"Posisi nomor",type:"select",value:"bottom-center",items:[["bottom-center","Bawah tengah"],["bottom-right","Bawah kanan"],["bottom-left","Bawah kiri"],["top-center","Atas tengah"]]},{id:"start",label:"Mulai dari nomor",type:"number",value:"1"},{id:"format",label:"Format",type:"select",value:"plain",items:[["plain","1, 2, 3"],["slash","Page 1 of N"]] }]},
{id:"sign-pdf",name:"Tanda tangani PDF",desc:"Gambar tanda tangan lalu letakkan di halaman pertama.",icon:"✍",color:"green",cat:"edit",formats:".pdf",opts:[{id:"signature",label:"Gambar tanda tangan",type:"signature"}]},
{id:"crop-pdf",name:"Crop PDF",desc:"Potong area halaman PDF dengan ukuran kustom.",icon:"⌗",color:"orange",cat:"edit",formats:".pdf",opts:[{id:"left",label:"Margin kiri (pt)",type:"number",value:"20"},{id:"bottom",label:"Margin bawah (pt)",type:"number",value:"20"},{id:"width",label:"Lebar area (pt)",type:"number",value:"555"},{id:"height",label:"Tinggi area (pt)",type:"number",value:"752"}]},
{id:"add-text",name:"Tambah teks ke PDF",desc:"Letakkan teks tambahan pada halaman pertama.",icon:"T+",color:"pink",cat:"edit",formats:".pdf",opts:[{id:"text",label:"Teks yang ditambahkan",type:"text",value:"Dokumen telah diverifikasi",placeholder:"Masukkan teks"},{id:"x",label:"Posisi X (pt)",type:"number",value:"50"},{id:"y",label:"Posisi Y (pt)",type:"number",value:"50"},{id:"size",label:"Ukuran teks",type:"number",value:"12"}]},
{id:"flatten-pdf",name:"Flatten PDF",desc:"Gabungkan field formulir PDF ke tampilan halaman.",icon:"▱",color:"teal",cat:"edit",formats:".pdf"},
{id:"ocr",name:"OCR PDF & gambar",desc:"Kenali teks dari scan PDF dan gambar.",icon:"⌕",color:"purple",cat:"security",formats:".pdf,image/*",opts:[{id:"lang",label:"Bahasa OCR",type:"select",value:"eng",items:[["eng","English"],["ind","Indonesian"],["eng+ind","English + Indonesian"]]},{id:"ocrHint",label:"",type:"note",text:"OCR menggunakan model bahasa yang dimuat dari CDN. File tetap diproses di browser, namun proses gambar beresolusi tinggi dapat memerlukan memori cukup besar."}]},
{id:"compare-pdf",name:"Bandingkan PDF",desc:"Bandingkan teks dari dua dokumen PDF.",icon:"⇆",color:"blue",cat:"security",formats:".pdf",min:2,max:2},
{id:"image-convert",name:"Konverter gambar",desc:"Ubah PNG, JPG, atau WEBP ke format lain.",icon:"▧",color:"pink",cat:"image",formats:"image/*",opts:[{id:"format",label:"Format hasil",type:"select",value:"image/jpeg",items:[["image/jpeg","JPG"],["image/png","PNG"],["image/webp","WEBP"]]},{id:"quality",label:"Kualitas JPG / WEBP",type:"select",value:"0.9",items:[["0.7","70%"],["0.85","85%"],["0.9","90%"],["1","100%"]]}]},
{id:"heic-to-jpg",name:"HEIC ke JPG",desc:"Konversi foto HEIC menjadi JPG.",icon:"▧",color:"orange",cat:"image",formats:".heic,.heif"},
{id:"resize-image",name:"Resize gambar",desc:"Ubah dimensi gambar berdasarkan lebar dan tinggi.",icon:"⤢",color:"teal",cat:"image",formats:"image/*",opts:[{id:"width",label:"Lebar (px)",type:"number",value:"1200"},{id:"height",label:"Tinggi (px)",type:"number",value:"1200"},{id:"fit",label:"Metode resize",type:"select",value:"contain",items:[["contain","Pertahankan rasio"],["stretch","Paksa ukuran"]]},{id:"format",label:"Format hasil",type:"select",value:"image/jpeg",items:[["image/jpeg","JPG"],["image/png","PNG"],["image/webp","WEBP"]]}]}
];
$("allCount").textContent=tools.length;
$("quickAllCount").textContent=tools.length;
const groupNames={organize:"Organisasi PDF",convert:"Konversi dokumen",edit:"Edit PDF",security:"Keamanan & utilitas",image:"Image tools"};
const groupOrder=["organize","convert","edit","security","image"];
let toastTimer;
function toast(message){const el=$("toast");el.textContent=message;el.classList.add("show");clearTimeout(toastTimer);toastTimer=setTimeout(()=>el.classList.remove("show"),2600)}
function formatSize(n){if(n<1024)return n+" B";if(n<1048576)return (n/1024).toFixed(1)+" KB";return (n/1048576).toFixed(2)+" MB"}
function safeName(n){return (n||"file").replace(/\.[^.]+$/,"").replace(/[^a-z0-9_-]+/gi,"-").replace(/^-|-$/g,"").slice(0,60)||"file"}
function escapeHtml(s){return String(s).replace(/[&<>"']/g,(c)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function loadScript(url){if(loaded[url])return loaded[url];loaded[url]=new Promise((resolve,reject)=>{const s=document.createElement("script");s.src=url;s.async=true;s.onload=()=>resolve();s.onerror=()=>{delete loaded[url];reject(new Error("Tidak dapat memuat pustaka pemrosesan. Periksa koneksi internet dan coba lagi."))};document.head.appendChild(s)});return loaded[url]}
async function needLib(key){await loadScript(CDN[key]);if(key==="pdfjs"&&window.pdfjsLib)window.pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";return true}
function renderTools(){
 const query=state.search.trim().toLowerCase();
 let filtered=tools.filter(t=>{
  const inCategory=state.filter==="all"||(state.filter==="popular"?t.pop:state.filter==="pdf"?t.formats.includes(".pdf"):state.filter==="office"?["word-to-pdf","excel-to-pdf","powerpoint-to-pdf","pdf-to-word","pdf-to-excel"].includes(t.id):state.filter==="image"?t.cat==="image":t.cat===state.filter);
  const matches=!query||(t.name+" "+t.desc+" "+groupNames[t.cat]).toLowerCase().includes(query);
  return inCategory&&matches;
 });
 $("toolGroups").innerHTML=groupOrder.map(cat=>{
  const arr=filtered.filter(t=>t.cat===cat);if(!arr.length)return "";
  return '<section class="tool-group"><h3 class="group-title">'+groupNames[cat]+' <span>'+arr.length+' tools</span></h3><div class="tool-grid">'+arr.map(t=>'<button class="tool-card" data-tool="'+t.id+'"><span class="tool-symbol '+t.color+'">'+t.icon+'</span><span class="tool-text"><strong>'+t.name+'</strong><small>'+t.desc+'</small></span><span class="tool-arrow">↗</span></button>').join("")+'</div></section>';
 }).join("");
 $("emptySearch").classList.toggle("hidden",filtered.length>0);
 document.querySelectorAll("[data-tool]").forEach(b=>b.addEventListener("click",()=>openTool(b.dataset.tool)));
 $("toolHeading").textContent=state.filter==="all"?"Jelajahi semua tools":state.filter==="popular"?"Tools pilihan populer":state.filter==="pdf"?"PDF tools":state.filter==="office"?"Office tools":state.filter==="image"?"Image tools":groupNames[state.filter]||"Jelajahi tools";
}
function setFilter(filter){
 state.filter=filter;document.querySelectorAll(".side-link").forEach(b=>b.classList.toggle("active",b.dataset.category===(filter==="popular"?"all":filter)));
 document.querySelectorAll(".quick-chip").forEach(b=>b.classList.toggle("selected",b.dataset.filter===filter));
 $("crumbCurrent").textContent=filter==="all"?"Semua tools":filter==="popular"?"Tools populer":filter==="pdf"?"PDF tools":filter==="office"?"Office tools":filter==="image"?"Image tools":groupNames[filter]||"Semua tools";
 $("toolSearch").value="";state.search="";renderTools();
}
function showHome(){ $("workspaceView").classList.add("hidden");$("activityView").classList.add("hidden");$("homeView").classList.remove("hidden");$("crumbCurrent").textContent=groupNames[state.filter]||"Semua tools";window.scrollTo({top:0,behavior:"smooth"})}
function openTool(id){
 const tool=tools.find(t=>t.id===id);if(!tool)return;state.tool=tool;state.files=[];state.outputs=[];state.pageItems=[];state.previewToken++;state.previewBusy=false;state.pageDrag=null;state.suppressPageClick=false;
 $("homeView").classList.add("hidden");$("activityView").classList.add("hidden");$("workspaceView").classList.remove("hidden");
 $("workspaceIcon").textContent=tool.icon;$("workspaceTitle").textContent=tool.name;$("workspaceDescription").textContent=tool.desc;$("workspaceCategory").textContent=groupNames[tool.cat].toUpperCase();
 $("crumbCurrent").textContent=tool.name;
 $("fileInput").value="";$("fileInput").accept=tool.formats;$("uploadNote").textContent="Format: "+tool.formats.replaceAll("image/*","gambar umum")+" · Maksimum 100 MB per file";
 $("uploadZone").classList.remove("hidden");$("fileListWrap").classList.add("hidden");$("pagePreviewPanel").classList.add("hidden");$("optionsPanel").classList.add("hidden");$("runRow").classList.add("hidden");$("progressWrap").classList.add("hidden");$("resultPanel").classList.add("hidden");$("errorPanel").classList.add("hidden");
 renderOptions();window.scrollTo({top:0,behavior:"smooth"});
}
function renderOptions(){
 const box=$("toolOptions"),opts=state.tool.opts||[];box.innerHTML=opts.map(o=>{
  if(o.type==="note")return '<div class="field full"><div class="hint">'+escapeHtml(o.text)+'</div></div>';
  if(o.type==="signature")return '<div class="field full"><label>'+o.label+'</label><canvas id="signaturePad" width="700" height="160" style="width:100%;height:118px;border:1px solid #e5e7ef;border-radius:8px;background:#fff;touch-action:none;cursor:crosshair"></canvas><div style="display:flex;justify-content:space-between"><span class="hint">Gambar tanda tangan di area di atas.</span><button type="button" class="text-btn" id="clearSignature">Hapus tanda tangan</button></div></div>';
  const wide=o.type==="textarea"?"full":"";
  let control="";
  if(o.type==="select")control='<select id="opt-'+o.id+'">'+o.items.map(it=>'<option value="'+it[0]+'" '+(o.value===it[0]?"selected":"")+'>'+it[1]+'</option>').join("")+'</select>';
  else if(o.type==="textarea")control='<textarea id="opt-'+o.id+'" placeholder="'+(o.placeholder||"")+'">'+(o.value||"")+'</textarea>';
  else control='<input id="opt-'+o.id+'" type="'+(o.type==="number"?"number":"text")+'" value="'+(o.value||"")+'" placeholder="'+(o.placeholder||"")+'">';
  return '<div class="field '+wide+'"><label for="opt-'+o.id+'">'+o.label+'</label>'+control+(o.hint?'<span class="hint">'+o.hint+'</span>':"")+'</div>';
 }).join("");
 if(opts.some(o=>o.type==="signature"))setupSignature();
}
function setupSignature(){
 const c=$("signaturePad");if(!c)return;const ctx=c.getContext("2d");ctx.strokeStyle="#202942";ctx.lineWidth=3.5;ctx.lineCap="round";ctx.lineJoin="round";let drawing=false,has=false;
 const point=e=>{const r=c.getBoundingClientRect();return {x:(e.clientX-r.left)*c.width/r.width,y:(e.clientY-r.top)*c.height/r.height}};
 c.addEventListener("pointerdown",e=>{drawing=true;has=true;c.setPointerCapture(e.pointerId);const p=point(e);ctx.beginPath();ctx.moveTo(p.x,p.y)});
 c.addEventListener("pointermove",e=>{if(!drawing)return;const p=point(e);ctx.lineTo(p.x,p.y);ctx.stroke()});
 ["pointerup","pointercancel","pointerleave"].forEach(ev=>c.addEventListener(ev,()=>drawing=false));
 $("clearSignature").addEventListener("click",()=>{ctx.clearRect(0,0,c.width,c.height);has=false});
 c.dataset.has="false";c.addEventListener("pointerdown",()=>c.dataset.has="true");
}
function getOptions(){const out={};(state.tool.opts||[]).forEach(o=>{if(o.type==="signature"){const c=$("signaturePad");out.signature=c&&c.dataset.has==="true"?c.toDataURL("image/png"):null}else if(o.type!=="note"){const el=$("opt-"+o.id);out[o.id]=el?el.value:""}});return out}
function accepts(file,formats){if(formats==="image/*")return file.type.startsWith("image/");if(formats.includes("image/*")&&file.type.startsWith("image/"))return true;const allowed=formats.split(",").map(x=>x.trim().toLowerCase());const ext="."+file.name.split(".").pop().toLowerCase();return allowed.includes(ext)}
function addFiles(fileList){if(!state.tool)return;const candidates=Array.from(fileList||[]);const valid=[];candidates.forEach(f=>{if(f.size>100*1024*1024){toast(f.name+" melebihi batas 100 MB.");return}if(!accepts(f,state.tool.formats)){toast("Format "+f.name+" tidak didukung oleh tool ini.");return}valid.push(f)});const multiTools=["merge-pdf","images-to-pdf","ocr","image-convert","resize-image","heic-to-jpg"];const max=state.tool.max||(multiTools.includes(state.tool.id)?20:1);const existing=state.files.length;state.files=state.files.concat(valid).slice(0,max);if(state.files.length>existing)renderFileList();if(valid.length&&state.files.length===existing)toast("Batas file untuk tool ini sudah tercapai.");}
function renderFileList(){
 $("uploadZone").classList.toggle("hidden",state.files.length>0);$("fileListWrap").classList.toggle("hidden",state.files.length===0);
 $("optionsPanel").classList.toggle("hidden",state.files.length===0);$("runRow").classList.toggle("hidden",state.files.length===0);
 $("fileCountLabel").textContent=state.files.length+" file";$("fileList").innerHTML=state.files.map((f,i)=>{
  const ext=f.name.split(".").pop().toUpperCase().slice(0,4);return '<div class="file-row"><div class="file-type">'+escapeHtml(ext)+'</div><div class="file-meta"><strong>'+escapeHtml(f.name)+'</strong><small>'+formatSize(f.size)+'</small></div><span class="file-size">'+formatSize(f.size)+'</span><button class="remove-file" data-remove="'+i+'" title="Hapus file">×</button></div>';
 }).join("");
 document.querySelectorAll("[data-remove]").forEach(b=>b.addEventListener("click",()=>{state.files.splice(Number(b.dataset.remove),1);renderFileList()}));
 refreshPagePreview();
}
function isPdfFile(file){return /\.pdf$/i.test(file.name)||file.type==="application/pdf"}
function previewEligible(){return !!(state.tool&&state.tool.formats.includes(".pdf")&&state.files.some(isPdfFile))}
async function refreshPagePreview(){
 const panel=$("pagePreviewPanel"),token=++state.previewToken;state.pageItems=[];state.previewBusy=false;
 if(!previewEligible()){panel.classList.add("hidden");return}
 panel.classList.remove("hidden");state.previewBusy=true;
 $("pagePreviewHint").textContent="Membaca halaman dan membuat thumbnail...";
 $("pagePreviewGrid").innerHTML='<div class="preview-loading">Menyiapkan pratinjau halaman...</div>';
 $("pagePreviewFoot").textContent="File sedang dibaca langsung di browser.";
 try{
  await needLib("pdfjs");let thumbCount=0;
  for(let fileIndex=0;fileIndex<state.files.length;fileIndex++){
   const file=state.files[fileIndex];if(!isPdfFile(file))continue;
   const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await bytes(file))}).promise;
   for(let pageNo=1;pageNo<=pdf.numPages;pageNo++){
    if(token!==state.previewToken)return;
    const page=await pdf.getPage(pageNo);let thumb=null;
    if(thumbCount<160){
     const base=page.getViewport({scale:1});const scale=Math.min(.34,118/base.width);const viewport=page.getViewport({scale});
     const canvas=document.createElement("canvas");canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
     await page.render({canvasContext:canvas.getContext("2d"),viewport}).promise;
     thumb=canvas.toDataURL("image/jpeg",.64);thumbCount++;canvas.width=0;canvas.height=0;
    }
    state.pageItems.push({id:fileIndex+"-"+pageNo,fileIndex,pageIndex:pageNo-1,pageNumber:pageNo,fileName:file.name,thumb,selected:false});
    if(state.pageItems.length%18===0){$("pagePreviewHint").textContent=state.pageItems.length+" halaman terbaca...";renderPagePreviewGrid()}
   }
   if(pdf.destroy)pdf.destroy();
  }
  if(token!==state.previewToken)return;
  state.previewBusy=false;
  if(!state.pageItems.length)throw new Error("File PDF tidak memiliki halaman yang dapat ditampilkan.");
  renderPagePreviewGrid();
  $("pagePreviewHint").textContent=state.pageItems.length+" halaman dari "+state.files.filter(isPdfFile).length+" dokumen.";
  $("pagePreviewFoot").textContent=thumbCount<state.pageItems.length?"Thumbnail dibatasi untuk menjaga performa. Semua halaman tetap tercantum dan dapat dipilih.":"Thumbnail dibuat di browser. File asli tidak diunggah.";
 }catch(err){
  if(token!==state.previewToken)return;
  state.previewBusy=false;state.pageItems=[];
  $("pagePreviewGrid").innerHTML='<div class="preview-loading preview-error">Pratinjau tidak dapat dibuat: '+escapeHtml(err.message||"PDF tidak dapat dibaca.")+'</div>';
  $("pagePreviewHint").textContent="Gunakan pengaturan halaman di bawah jika diperlukan.";
  $("pagePreviewFoot").textContent="File tetap dapat diproses jika format PDF didukung.";
 }
}
function renderPagePreviewGrid(){
 const grid=$("pagePreviewGrid");if(!grid)return;
 const tool=state.tool;const reorderable=tool&&["merge-pdf","organize-pdf"].includes(tool.id);
 const selectable=tool&&["extract-pages","remove-pages","rotate-pdf","split-pdf"].includes(tool.id);
 $("previewActions").classList.toggle("hidden",!selectable);
 $("pagePreviewHint").textContent=reorderable?"Seret thumbnail untuk mengubah urutan halaman.":selectable?(tool.id==="split-pdf"?"Pilih halaman yang ingin dipisahkan menjadi file tersendiri.":"Klik thumbnail untuk memilih halaman yang diproses."):"Pratinjau isi halaman dokumen.";
 const selected=state.pageItems.filter(p=>p.selected).length;
 $("selectedPageCount").textContent=selected+" dipilih";
 grid.innerHTML=state.pageItems.map((item,index)=>{
  const preview=item.thumb?'<img src="'+item.thumb+'" alt="Pratinjau halaman '+item.pageNumber+'" loading="lazy">':'<div class="page-thumb-placeholder"><span>PDF</span><small>Preview terbatas</small></div>';
  return '<button type="button" class="page-thumb '+(item.selected?"is-selected ":"")+(reorderable?"is-draggable":"")+'" data-page-index="'+index+'" draggable="false" aria-label="Halaman '+item.pageNumber+' dari '+escapeHtml(item.fileName)+'">'+
   '<span class="page-thumb-image">'+preview+'<span class="page-order-badge">'+(index+1)+'</span>'+(reorderable?'<span class="page-drag-handle" aria-hidden="true">⠿</span>':"")+'<span class="page-checkmark">✓</span></span>'+
   '<span class="page-thumb-meta"><strong>Halaman '+item.pageNumber+'</strong><small title="'+escapeHtml(item.fileName)+'">'+escapeHtml(item.fileName)+'</small></span></button>';
 }).join("");
}
function applyPageReorder(from,to){
 if(!Number.isInteger(from)||!Number.isInteger(to)||from<0||to<0||from>=state.pageItems.length||to>=state.pageItems.length||from===to)return false;
 const moved=state.pageItems.splice(from,1)[0];state.pageItems.splice(to,0,moved);return true;
}
function bindPagePreviewInteractions(){
 const grid=$("pagePreviewGrid");
 grid.addEventListener("click",e=>{
  if(state.suppressPageClick){e.preventDefault();e.stopPropagation();return}
  const card=e.target.closest(".page-thumb");if(!card||!state.tool)return;
  if(!["extract-pages","remove-pages","rotate-pdf","split-pdf"].includes(state.tool.id))return;
  const item=state.pageItems[Number(card.dataset.pageIndex)];if(!item)return;
  item.selected=!item.selected;renderPagePreviewGrid();
 });
 $("selectAllPages").addEventListener("click",()=>{state.pageItems.forEach(p=>p.selected=true);renderPagePreviewGrid()});
 $("clearPageSelection").addEventListener("click",()=>{state.pageItems.forEach(p=>p.selected=false);renderPagePreviewGrid()});
 grid.addEventListener("pointerdown",e=>{
  const card=e.target.closest(".page-thumb");
  if(!card||!state.tool||!["merge-pdf","organize-pdf"].includes(state.tool.id))return;
  if(e.pointerType==="mouse"&&e.button!==0)return;
  if(e.pointerType==="touch"&&!e.target.closest(".page-drag-handle"))return;
  const from=Number(card.dataset.pageIndex);if(!Number.isInteger(from)||!state.pageItems[from])return;
  state.pageDrag={pointerId:e.pointerId,currentIndex:from,startX:e.clientX,startY:e.clientY,moved:false,lastTarget:from};
  e.preventDefault();
 });
 document.addEventListener("pointermove",e=>{
  const d=state.pageDrag;if(!d||e.pointerId!==d.pointerId)return;
  if(!d.moved&&Math.hypot(e.clientX-d.startX,e.clientY-d.startY)>5){
   d.moved=true;state.suppressPageClick=true;
   grid.classList.add("is-reordering");
   const initial=grid.querySelector('[data-page-index="'+d.currentIndex+'"]');if(initial)initial.classList.add("dragging");
  }
  if(!d.moved)return;
  e.preventDefault();
  if(e.clientY>window.innerHeight-75)window.scrollBy(0,18);
  else if(e.clientY<75)window.scrollBy(0,-18);
  const under=document.elementFromPoint(e.clientX,e.clientY);
  const targetCard=under&&under.closest?under.closest(".page-thumb"):null;
  if(!targetCard||!grid.contains(targetCard))return;
  const targetIndex=Number(targetCard.dataset.pageIndex);
  if(!Number.isInteger(targetIndex)||targetIndex===d.currentIndex)return;
  const moved=applyPageReorder(d.currentIndex,targetIndex);
  if(moved){
   d.currentIndex=targetIndex;d.lastTarget=targetIndex;
   renderPagePreviewGrid();
   const active=grid.querySelector('[data-page-index="'+d.currentIndex+'"]');
   if(active)active.classList.add("dragging");
  }
 },{passive:false});
 const finishDrag=e=>{
  const d=state.pageDrag;if(!d||e.pointerId!==undefined&&e.pointerId!==d.pointerId)return;
  state.pageDrag=null;
  grid.classList.remove("is-reordering");
  grid.querySelectorAll(".dragging,.drop-target").forEach(el=>el.classList.remove("dragging","drop-target"));
  if(d.moved){
   state.suppressPageClick=true;
   setTimeout(()=>{state.suppressPageClick=false},400);
   renderPagePreviewGrid();
   toast("Urutan halaman diperbarui.");
  }
 };
 document.addEventListener("pointerup",finishDrag);
 document.addEventListener("pointercancel",finishDrag);
}
function setProgress(p,label,detail){$("progressWrap").classList.remove("hidden");$("progressLabel").textContent=label||"Memproses dokumen...";$("progressPercent").textContent=Math.max(0,Math.min(100,Math.round(p)))+"%";$("progressBar").style.width=Math.max(0,Math.min(100,p))+"%";$("progressDetail").textContent=detail||"Memproses file di browser Anda."}
function bytes(file){return file.arrayBuffer()}
function makeBlob(data,type){return new Blob([data],{type:type||"application/octet-stream"})}
function downloadBlob(blob,name){const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000)}
function saveResult(blob,name){state.outputs=[{blob,name}];$("resultPanel").classList.remove("hidden");$("resultDescription").textContent=name+" · "+formatSize(blob.size);$("downloadResult").onclick=()=>downloadBlob(blob,name)}
async function zipOutputs(items,name){
 await needLib("zip");const zip=new JSZip();items.forEach(it=>zip.file(it.name,it.blob));const blob=await zip.generateAsync({type:"blob",compression:"DEFLATE",compressionOptions:{level:6}});saveResult(blob,name||"docuflow-hasil.zip")
}
function pageRange(input,total){
 if(!input||input.trim().toLowerCase()==="all")return Array.from({length:total},(_,i)=>i);
 const set=new Set();input.split(",").map(s=>s.trim()).filter(Boolean).forEach(part=>{
  const m=part.match(/^(\d+)\s*-\s*(\d+)$/);if(m){let a=Number(m[1]),b=Number(m[2]);if(a>b){const t=a;a=b;b=t}for(let p=a;p<=b;p++)if(p>=1&&p<=total)set.add(p-1)}
  else if(/^\d+$/.test(part)){const p=Number(part);if(p>=1&&p<=total)set.add(p-1)}
 });
 return Array.from(set).sort((a,b)=>a-b);
}
async function readPdf(file,opts){await needLib("pdf");return PDFLib.PDFDocument.load(await bytes(file),opts||{})}
async function savePdf(doc){return makeBlob(await doc.save({useObjectStreams:true,addDefaultPage:false}),"application/pdf")}
async function extractPdfLines(file){
 await needLib("pdfjs");const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await bytes(file))}).promise;const all=[];
 for(let n=1;n<=pdf.numPages;n++){const page=await pdf.getPage(n);const content=await page.getTextContent();const rows=[];
  content.items.filter(it=>it.str&&it.str.trim()).forEach(it=>{const y=Math.round(it.transform[5]/3)*3;let row=rows.find(r=>Math.abs(r.y-y)<=3);if(!row){row={y,items:[]};rows.push(row)}row.items.push({x:it.transform[4],text:it.str})});
  rows.sort((a,b)=>b.y-a.y);const text=rows.map(r=>r.items.sort((a,b)=>a.x-b.x).map(x=>x.text).join(" ")).join("\n");all.push(text);
 }
 return all;
}
function textToDocx(paragraphs){
 const zip=new JSZip();const xmlEscape=s=>String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
 const p=paragraphs.map(line=>"<w:p><w:r><w:t xml:space=\"preserve\">"+xmlEscape(line||" ")+"</w:t></w:r></w:p>").join("");
 const documentXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+p+'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
 zip.file("[Content_Types].xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
 zip.folder("_rels").file(".rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
 zip.folder("word").file("document.xml",documentXml);return zip.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
}
function canvasBlob(canvas,mime,quality){return new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error("Browser gagal membuat gambar.")),mime,quality))}
async function pdfPageImages(file,scale,format){
 await needLib("pdfjs");const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await bytes(file))}).promise;const items=[];
 for(let n=1;n<=pdf.numPages;n++){setProgress(10+n/pdf.numPages*70,"Merender halaman PDF...","Halaman "+n+" dari "+pdf.numPages);const page=await pdf.getPage(n);const viewport=page.getViewport({scale:scale});const c=document.createElement("canvas");c.width=Math.ceil(viewport.width);c.height=Math.ceil(viewport.height);await page.render({canvasContext:c.getContext("2d"),viewport}).promise;const blob=await canvasBlob(c,format,0.9);items.push({name:"halaman-"+String(n).padStart(3,"0")+(format==="image/png"?".png":".jpg"),blob});c.width=0;c.height=0}
 return items;
}
async function imageFromFile(file){
 if(file.type==="image/heic"||file.type==="image/heif"||/\.(heic|heif)$/i.test(file.name)){await needLib("heic");const converted=await heic2any({blob:file,toType:"image/png",quality:.95});return await createImageBitmap(Array.isArray(converted)?converted[0]:converted)}
 return await createImageBitmap(file);
}
async function imagesAsPdf(files){
 await needLib("pdf");const doc=await PDFLib.PDFDocument.create();
 for(const file of files){const img=await imageFromFile(file);const c=document.createElement("canvas");c.width=img.width;c.height=img.height;c.getContext("2d").drawImage(img,0,0);const b=await canvasBlob(c,"image/png");const emb=await doc.embedPng(await b.arrayBuffer());const page=doc.addPage([emb.width,emb.height]);page.drawImage(emb,{x:0,y:0,width:emb.width,height:emb.height});img.close();c.width=0;c.height=0}
 return savePdf(doc);
}
async function htmlToPdf(html){
 await needLib("htmlpdf");const box=document.createElement("div");box.style.cssText="position:fixed;left:-10000px;top:0;width:760px;background:white;color:#222;padding:32px;font:14px Arial;";box.innerHTML=html;document.body.appendChild(box);
 try{return await html2pdf().set({margin:12,filename:"hasil.pdf",image:{type:"jpeg",quality:.92},html2canvas:{scale:1.5,useCORS:true},jsPDF:{unit:"mm",format:"a4",orientation:"portrait"}}).from(box).outputPdf("blob")}
 finally{box.remove()}
}
async function xlsxToHtml(file){
 await needLib("xlsx");const wb=XLSX.read(await bytes(file),{type:"array"});let html='<div style="font-family:Arial">';
 wb.SheetNames.forEach((name,i)=>{html+='<h2 style="font-size:18px;border-bottom:1px solid #ccc;padding-bottom:8px">'+escapeHtml(name)+'</h2>'+XLSX.utils.sheet_to_html(wb.Sheets[name])+(i<wb.SheetNames.length-1?'<div style="page-break-after:always"></div>':"")});
 return html+"</div>";
}
function xmlDecode(s){return s.replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&apos;/g,"'")}
async function pptxToHtml(file){
 await needLib("zip");const zip=await JSZip.loadAsync(await bytes(file));const keys=Object.keys(zip.files).filter(k=>/^ppt\/slides\/slide\d+\.xml$/.test(k)).sort((a,b)=>Number(a.match(/slide(\d+)/)[1])-Number(b.match(/slide(\d+)/)[1]));if(!keys.length)throw new Error("Tidak menemukan slide dalam file PPTX.");
 let html='<div style="font-family:Arial">';
 for(let i=0;i<keys.length;i++){const xml=await zip.file(keys[i]).async("text");const matches=[...xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)].map(m=>xmlDecode(m[1]));html+='<section style="page-break-after:always;min-height:900px;padding:28px;border:1px solid #ddd"><div style="font-size:11px;color:#777;margin-bottom:20px">SLIDE '+(i+1)+'</div>'+matches.map((t,j)=>'<p style="font-size:'+(j===0?26:17)+'px;line-height:1.45">'+escapeHtml(t)+'</p>').join("")+'</section>'}
 return html+"</div>";
}
async function processTool(tool,files,opts){
 const f=files[0],base=safeName(f.name);
 if(tool.id==="merge-pdf"){
  await needLib("pdf");const out=await PDFLib.PDFDocument.create();
  if(state.pageItems.length){const docs=[];for(let i=0;i<state.pageItems.length;i++){const item=state.pageItems[i];if(!docs[item.fileIndex])docs[item.fileIndex]=await PDFLib.PDFDocument.load(await bytes(files[item.fileIndex]));setProgress(12+i/state.pageItems.length*72,"Menggabungkan halaman...","Halaman "+(i+1)+" dari "+state.pageItems.length);const cp=await out.copyPages(docs[item.fileIndex],[item.pageIndex]);out.addPage(cp[0])}}
  else{for(let i=0;i<files.length;i++){setProgress(15+i/files.length*65,"Menggabungkan PDF...","Membaca "+files[i].name);const src=await PDFLib.PDFDocument.load(await bytes(files[i]));const pages=await out.copyPages(src,src.getPageIndices());pages.forEach(p=>out.addPage(p))}}
  return [{name:"pdf-gabungan.pdf",blob:await savePdf(out)}];
 }
 if(tool.id==="split-pdf"||tool.id==="extract-pages"||tool.id==="remove-pages"||tool.id==="organize-pdf"){
  await needLib("pdf");const src=await PDFLib.PDFDocument.load(await bytes(f));const total=src.getPageCount();let indexes=[];
  if(tool.id==="split-pdf"){const picked=state.pageItems.filter(p=>p.selected).map(p=>p.pageIndex);if(picked.length){const items=[];for(const i of picked){const out=await PDFLib.PDFDocument.create();const cp=await out.copyPages(src,[i]);out.addPage(cp[0]);items.push({name:"halaman-"+String(i+1).padStart(3,"0")+".pdf",blob:await savePdf(out)})}return items}}
  if(tool.id==="split-pdf"&&opts.mode==="each"){const items=[];for(let i=0;i<total;i++){const out=await PDFLib.PDFDocument.create();const cp=await out.copyPages(src,[i]);out.addPage(cp[0]);items.push({name:"halaman-"+String(i+1).padStart(3,"0")+".pdf",blob:await savePdf(out)})}return items}
  if(tool.id==="split-pdf"){const ranges=(opts.ranges||"").split(",").map(s=>s.trim()).filter(Boolean);if(!ranges.length)throw new Error("Isi rentang halaman yang ingin dipisahkan.");const items=[];for(let i=0;i<ranges.length;i++){const idx=pageRange(ranges[i],total);if(!idx.length)continue;const out=await PDFLib.PDFDocument.create();const cp=await out.copyPages(src,idx);cp.forEach(p=>out.addPage(p));items.push({name:"bagian-"+(i+1)+".pdf",blob:await savePdf(out)})}if(!items.length)throw new Error("Rentang halaman tidak valid.");return items}
  if(tool.id==="extract-pages"){const picked=state.pageItems.filter(p=>p.selected).map(p=>p.pageIndex);indexes=picked.length?picked:pageRange(opts.pages,total);if(!indexes.length)throw new Error("Pilih thumbnail halaman atau isi rentang halaman yang valid.")}
  if(tool.id==="remove-pages"){const picked=state.pageItems.filter(p=>p.selected).map(p=>p.pageIndex);const removed=picked.length?picked:pageRange(opts.pages,total);const removeSet=new Set(removed);indexes=Array.from({length:total},(_,i)=>i).filter(i=>!removeSet.has(i));if(!indexes.length)throw new Error("Semua halaman terpilih untuk dihapus.")}
  if(tool.id==="organize-pdf"){indexes=state.pageItems.length===total?state.pageItems.map(p=>p.pageIndex):(opts.orderFallback||"").trim()?opts.orderFallback.split(",").map(s=>Number(s.trim())-1):Array.from({length:total},(_,i)=>i);if(indexes.some(i=>!Number.isInteger(i)||i<0||i>=total)||indexes.length!==total||new Set(indexes).size!==total)throw new Error("Urutan halaman tidak valid. Setiap halaman harus muncul tepat satu kali.")}
  const out=await PDFLib.PDFDocument.create();const cp=await out.copyPages(src,indexes);cp.forEach(p=>out.addPage(p));return [{name:base+"-"+tool.id+".pdf",blob:await savePdf(out)}];
 }
 if(tool.id==="rotate-pdf"){
  const doc=await readPdf(f);const total=doc.getPageCount();const picked=state.pageItems.filter(p=>p.selected).map(p=>p.pageIndex);const indexes=picked.length?picked:pageRange(opts.pages,total);const angle=Number(opts.degrees);indexes.forEach(i=>{const p=doc.getPage(i);p.setRotation(PDFLib.degrees((p.getRotation().angle+angle)%360))});return [{name:base+"-rotated.pdf",blob:await savePdf(doc)}];
 }
 if(tool.id==="compress-pdf"||tool.id==="repair-pdf"||tool.id==="flatten-pdf"){
  const doc=await readPdf(f,{ignoreEncryption:true});
  if(tool.id==="flatten-pdf"){try{doc.getForm().flatten()}catch(e){throw new Error("Formulir PDF tidak ditemukan atau tidak dapat di-flatten.")}}
  const blob=await savePdf(doc);return [{name:base+(tool.id==="compress-pdf"?"-optimized":tool.id==="repair-pdf"?"-repaired":"-flattened")+".pdf",blob}];
 }
 if(tool.id==="pdf-to-jpg"||tool.id==="pdf-to-png"){return await pdfPageImages(f,Number(opts.scale)||1.6,tool.id==="pdf-to-png"?"image/png":"image/jpeg")}
 if(tool.id==="pdf-to-text"||tool.id==="pdf-to-word"||tool.id==="pdf-to-excel"||tool.id==="pdf-to-html"){
  const pages=await extractPdfLines(f);setProgress(78,"Menyusun dokumen hasil...","Membuat file keluaran.");
  if(tool.id==="pdf-to-text"){return [{name:base+".txt",blob:makeBlob(pages.map((t,i)=>"HALAMAN "+(i+1)+"\n"+t).join("\n\n"),"text/plain;charset=utf-8")}]}
  if(tool.id==="pdf-to-word"){const lines=pages.flatMap((t,i)=>["HALAMAN "+(i+1),...t.split("\n"),""]);const blob=await textToDocx(lines);return [{name:base+".docx",blob}]}
  if(tool.id==="pdf-to-html"){const html="<!doctype html><html><meta charset=\"utf-8\"><title>"+escapeHtml(f.name)+"</title><body style=\"font-family:Arial;max-width:850px;margin:40px auto;line-height:1.6\">"+pages.map((t,i)=>"<section><h2>Halaman "+(i+1)+"</h2><pre style=\"white-space:pre-wrap;font:14px Arial\">"+escapeHtml(t)+"</pre></section>").join("")+"</body></html>";return [{name:base+".html",blob:makeBlob(html,"text/html;charset=utf-8")}]}
  await needLib("xlsx");const rows=[];pages.forEach((t,i)=>{rows.push(["Halaman "+(i+1)]);t.split("\n").forEach(line=>rows.push([line]))});const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Ekstrak PDF");const arr=XLSX.write(wb,{bookType:"xlsx",type:"array"});return [{name:base+".xlsx",blob:makeBlob(arr,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}];
 }
 if(tool.id==="word-to-pdf"){
  await needLib("mammoth");const res=await mammoth.convertToHtml({arrayBuffer:await bytes(f)});if(!res.value.trim())throw new Error("Tidak menemukan konten teks pada DOCX.");const blob=await htmlToPdf(res.value);return [{name:base+".pdf",blob}];
 }
 if(tool.id==="excel-to-pdf"){
  const html=await xlsxToHtml(f);const blob=await htmlToPdf(html);return [{name:base+".pdf",blob}];
 }
 if(tool.id==="powerpoint-to-pdf"){
  const html=await pptxToHtml(f);const blob=await htmlToPdf(html);return [{name:base+".pdf",blob}];
 }
 if(tool.id==="images-to-pdf"){const blob=await imagesAsPdf(files);return [{name:"gambar-ke-pdf.pdf",blob}]}
 if(tool.id==="watermark"||tool.id==="page-numbers"||tool.id==="add-text"||tool.id==="sign-pdf"||tool.id==="crop-pdf"){
  const doc=await readPdf(f);const pages=doc.getPages();const font=await doc.embedFont(PDFLib.StandardFonts.Helvetica);
  if(tool.id==="watermark"){
   const text=(opts.text||"WATERMARK").trim();if(!text)throw new Error("Isi teks watermark.");pages.forEach(p=>{const size=Number(opts.size)||42;const w=p.getWidth(),h=p.getHeight();p.drawText(text,{x:w/2-(font.widthOfTextAtSize(text,size)/2),y:h/2,size,font,color:PDFLib.rgb(.48,.5,.56),opacity:Number(opts.opacity)||.22,rotate:PDFLib.degrees(35)})});
  }
  if(tool.id==="page-numbers"){
   pages.forEach((p,i)=>{const n=Number(opts.start||1)+i;const label=opts.format==="slash"?"Page "+n+" of "+pages.length:String(n);const size=10,w=p.getWidth(),h=p.getHeight(),tw=font.widthOfTextAtSize(label,size);let x=(w-tw)/2,y=22;if(opts.position==="bottom-right")x=w-tw-25;if(opts.position==="bottom-left")x=25;if(opts.position==="top-center"){x=(w-tw)/2;y=h-25}p.drawText(label,{x,y,size,font,color:PDFLib.rgb(.35,.37,.42)})});
  }
  if(tool.id==="add-text"){const text=(opts.text||"").trim();if(!text)throw new Error("Teks tidak boleh kosong.");pages[0].drawText(text,{x:Number(opts.x)||50,y:Number(opts.y)||50,size:Number(opts.size)||12,font,color:PDFLib.rgb(.08,.12,.24)})}
  if(tool.id==="sign-pdf"){
   if(!opts.signature)throw new Error("Gambar tanda tangan terlebih dahulu.");const data=opts.signature.split(",")[1];const emb=await doc.embedPng(Uint8Array.from(atob(data),c=>c.charCodeAt(0)));const p=pages[0],w=Math.min(180,p.getWidth()*.34),h=w*emb.height/emb.width;p.drawImage(emb,{x:p.getWidth()-w-35,y:35,width:w,height:h});
  }
  if(tool.id==="crop-pdf"){pages.forEach(p=>{const w=Number(opts.width),h=Number(opts.height);if(w<=0||h<=0||w>p.getWidth()||h>p.getHeight())throw new Error("Lebar dan tinggi crop harus positif dan tidak melebihi ukuran halaman.");p.setCropBox(Number(opts.left)||0,Number(opts.bottom)||0,w,h)})}
  return [{name:base+"-"+tool.id+".pdf",blob:await savePdf(doc)}];
 }
 if(tool.id==="ocr"){
  await needLib("tesseract");const lang=opts.lang||"eng",texts=[];
  const recognize=async(blob,label)=>{setProgress(20+texts.length*65/Math.max(1,files.length),"Mengenali teks...","Memproses "+label);const r=await Tesseract.recognize(blob,lang,{logger:m=>{if(m.status==="recognizing text")setProgress(Math.min(88,20+m.progress*65),"Mengenali teks...",label)}});texts.push(r.data.text||"")};
  for(const file of files){
   if(file.type==="application/pdf"||file.name.toLowerCase().endsWith(".pdf")){await needLib("pdfjs");const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await bytes(file))}).promise;for(let n=1;n<=pdf.numPages;n++){const page=await pdf.getPage(n);const vp=page.getViewport({scale:1.6});const c=document.createElement("canvas");c.width=vp.width;c.height=vp.height;await page.render({canvasContext:c.getContext("2d"),viewport:vp}).promise;const b=await canvasBlob(c,"image/png");await recognize(b,file.name+" halaman "+n);c.width=0;c.height=0}}
   else await recognize(file,file.name);
  }
  return [{name:"hasil-ocr.txt",blob:makeBlob(texts.join("\n\n---\n\n"),"text/plain;charset=utf-8")}];
 }
 if(tool.id==="compare-pdf"){
  const a=(await extractPdfLines(files[0])).join("\n");const b=(await extractPdfLines(files[1])).join("\n");const al=a.split(/\r?\n/),bl=b.split(/\r?\n/);const max=Math.max(al.length,bl.length),rows=[];
  for(let i=0;i<max;i++){if(al[i]===bl[i])rows.push("  "+(al[i]||""));else{if(al[i]!==undefined)rows.push("- "+al[i]);if(bl[i]!==undefined)rows.push("+ "+bl[i])}}
  const report="PERBANDINGAN DOKUMEN\nDokumen A: "+files[0].name+"\nDokumen B: "+files[1].name+"\n\n"+rows.join("\n");return [{name:"perbandingan-dokumen.txt",blob:makeBlob(report,"text/plain;charset=utf-8")}];
 }
 if(tool.id==="image-convert"||tool.id==="resize-image"||tool.id==="heic-to-jpg"){
  const outputs=[];
  for(let i=0;i<files.length;i++){const file=files[i];setProgress(20+i/files.length*60,"Mengonversi gambar...","Memproses "+file.name);const img=await imageFromFile(file);const c=document.createElement("canvas");
   if(tool.id==="resize-image"){const w=Number(opts.width)||img.width,h=Number(opts.height)||img.height;c.width=w;c.height=h;const ctx=c.getContext("2d");if(opts.fit==="stretch")ctx.drawImage(img,0,0,w,h);else{const scale=Math.min(w/img.width,h/img.height);const dw=img.width*scale,dh=img.height*scale;ctx.fillStyle="#fff";ctx.fillRect(0,0,w,h);ctx.drawImage(img,(w-dw)/2,(h-dh)/2,dw,dh)}}
   else{c.width=img.width;c.height=img.height;c.getContext("2d").drawImage(img,0,0)}
   const mime=tool.id==="heic-to-jpg"?"image/jpeg":tool.id==="image-convert"?opts.format:opts.format;const quality=tool.id==="image-convert"?Number(opts.quality)||.9:.92;const blob=await canvasBlob(c,mime,quality);const ext=mime==="image/png"?"png":mime==="image/webp"?"webp":"jpg";outputs.push({name:safeName(file.name)+"."+ext,blob});img.close();c.width=0;c.height=0;
  }return outputs;
 }
 throw new Error("Tool ini belum dapat dijalankan.");
}
async function runCurrentTool(){
 if(state.busy||!state.tool)return;const tool=state.tool;
 if(state.files.length<(tool.min||1)){showError("Pilih minimal "+(tool.min||1)+" file untuk tool ini.");return}
 if(tool.max&&state.files.length!==tool.max){showError("Tool ini memerlukan tepat "+tool.max+" file.");return}
 if(state.previewBusy&&tool.formats.includes(".pdf")){showError("Pratinjau halaman masih disiapkan. Tunggu sampai thumbnail muncul, lalu jalankan kembali.");return}
 if(state.files.some(f=>f.size===0)){showError("Ada file kosong. Pilih file lain.");return}
 state.busy=true;$("runTool").disabled=true;$("runTool").textContent="Memproses...";$("resultPanel").classList.add("hidden");$("errorPanel").classList.add("hidden");setProgress(5,"Menyiapkan file...","Memuat pustaka pemrosesan yang diperlukan.");
 try{
  const outputs=await processTool(tool,state.files,getOptions());if(!outputs||!outputs.length)throw new Error("Tidak ada hasil yang dibuat.");
  setProgress(95,"Menyiapkan hasil...","Mengemas file untuk diunduh.");
  if(outputs.length===1)saveResult(outputs[0].blob,outputs[0].name);else await zipOutputs(outputs,"docuflow-"+tool.id+".zip");
  setProgress(100,"Selesai","File hasil siap diunduh.");state.activities.unshift({tool:tool.name,files:state.files.length,time:new Date().toLocaleString("id-ID"),output:state.outputs[0].name});
  toast("Berhasil. Hasil siap diunduh.");
 }catch(err){showError(err&&err.message?err.message:"Terjadi kesalahan saat memproses dokumen.");}
 finally{state.busy=false;$("runTool").disabled=false;$("runTool").innerHTML='Jalankan tool <span>→</span>'}
}
function showError(msg){$("errorPanel").classList.remove("hidden");$("errorMessage").textContent=msg;$("progressWrap").classList.add("hidden")}
function showActivity(){ $("homeView").classList.add("hidden");$("workspaceView").classList.add("hidden");$("activityView").classList.remove("hidden");$("activityList").innerHTML=state.activities.length?state.activities.map(a=>'<div class="activity-row"><span class="tool-symbol purple">✓</span><div style="flex:1"><strong>'+escapeHtml(a.tool)+'</strong><small>'+escapeHtml(a.time)+' · '+a.files+' file</small></div><small>'+escapeHtml(a.output)+'</small></div>').join(""):'<div class="activity-empty">Belum ada aktivitas di sesi ini.</div>';window.scrollTo(0,0)}
function bind(){
 bindPagePreviewInteractions();
 $("sidebarHome").addEventListener("click",showHome);$("backHome").addEventListener("click",showHome);$("activityBack").addEventListener("click",showHome);
 $("categoryNav").addEventListener("click",e=>{const b=e.target.closest("[data-category]");if(b){setFilter(b.dataset.category);$("sidebar").classList.remove("open");showHome()}});
 document.querySelectorAll("[data-filter]").forEach(b=>b.addEventListener("click",()=>setFilter(b.dataset.filter)));
 $("toolSearch").addEventListener("input",e=>{state.search=e.target.value;renderTools()});
 $("browseFiles").addEventListener("click",()=>$("fileInput").click());$("addMoreFiles").addEventListener("click",()=>$("fileInput").click());
 $("fileInput").addEventListener("change",e=>{addFiles(e.target.files);e.target.value=""});
 const zone=$("uploadZone");["dragenter","dragover"].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.add("dragover")}));["dragleave","drop"].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.remove("dragover")}));zone.addEventListener("drop",e=>addFiles(e.dataTransfer.files));
 $("runTool").addEventListener("click",runCurrentTool);$("dismissError").addEventListener("click",()=>$("errorPanel").classList.add("hidden"));$("resetTool").addEventListener("click",()=>openTool(state.tool.id));
 $("historyButton").addEventListener("click",showActivity);
 $("aboutLink").addEventListener("click",e=>{e.preventDefault();$("aboutModal").classList.remove("hidden")});$("closeAbout").addEventListener("click",()=>$("aboutModal").classList.add("hidden"));$("understoodAbout").addEventListener("click",()=>$("aboutModal").classList.add("hidden"));
 $("mobileMenu").addEventListener("click",()=>$("sidebar").classList.toggle("open"));
 document.addEventListener("keydown",e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="k"){e.preventDefault();$("toolSearch").focus()}if(e.key==="Escape"){$("sidebar").classList.remove("open");$("aboutModal").classList.add("hidden")}});
}
bind();renderTools();
function initAccount(){
 const area=$("accountArea"),btn=$("accountButton"),A=window.hatchable&&window.hatchable.auth;
 if(!area||!btn||!A){return}
 A.getSession().then(s=>{
  const u=s&&s.user;if(!u)return;
  const display=u.name||u.email||"Akun";
  area.innerHTML='<div class="account-menu"><span class="account-avatar">'+(u.image?'<img src="'+escapeHtml(u.image)+'" alt="">':'<span>'+escapeHtml(display.charAt(0).toUpperCase())+'</span>')+'</span><span class="account-label"><strong>'+escapeHtml(display)+'</strong><small>'+escapeHtml(u.email||"")+'</small></span><button class="account-btn account-out" id="accountSignout">Keluar</button></div>';
  $("accountSignout").addEventListener("click",async()=>{try{await A.signOut();location.reload()}catch(e){toast("Gagal keluar. Coba lagi.")}});
 }).catch(()=>{});
}
initAccount();
})();