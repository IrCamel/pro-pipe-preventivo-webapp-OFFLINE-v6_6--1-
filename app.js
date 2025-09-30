(function(){
  const LOCALITA = ["Bibbona","Campiglia Marittima","Casale Marittimo","Castagneto Carducci","Castellina Marittima","Castelnuovo della Misericordia","Cecina","Chioma","Collemezzano","Collesalvetti","Gabbro","Guardistallo","Livorno","Marina di Bibbona","Montescudaio","Nibbiaia","Orciano Pisano","Pastina","Piombino","Quercianella","Riparbella","Rosignano Marittimo","San Vincenzo","Santa Luce"];
  const eur = (n) => (Number(n)||0).toLocaleString("it-IT",{style:"currency",currency:"EUR"});
  const round2 = (n) => Math.round((Number(n)+Number.EPSILON)*100)/100;
  const ceilToFive = (x) => Math.ceil((Number(x)||0)/5)*5;
  const el = (id)=>document.getElementById(id);
  const qs = (s)=>document.querySelector(s);
  const qsa = (s)=>Array.from(document.querySelectorAll(s));

  // CSV helpers
  function parseCSV(text){
    const delim = text.indexOf(";")>-1 && (text.split(";").length >= text.split(",").length) ? ";" : ",";
    const rows = []; let i=0, field="", row=[], inQuotes=false;
    const pushField=()=>{ row.push(field); field=""; };
    const pushRow=()=>{ rows.push(row); row=[]; };
    while(i<text.length){
      const c = text[i];
      if(c === '"'){ if(inQuotes && text[i+1] === '"'){ field+='"'; i+=2; continue; } inQuotes = !inQuotes; i++; continue; }
      if(!inQuotes && (c === "\n" || c === "\r")){ if(c==="\r" && text[i+1]==="\n") i++; pushField(); pushRow(); i++; continue; }
      if(!inQuotes && c === delim){ pushField(); i++; continue; }
      field += c; i++;
    }
    if(field.length>0 || row.length>0){ pushField(); pushRow(); }
    if(rows.length && rows[rows.length-1].every(c=>c==="")) rows.pop();
    return rows;
  }
  function toCSV(rows, delim=";"){
    return rows.map(r=>r.map(f=>{
      const s = String(f??"");
      if(s.includes(delim) || s.includes('"') || s.includes("\n")) return `"${s.replace(/"/g,'""')}"`;
      return s;
    }).join(delim)).join("\n");
  }

  // Storage keys
  const LS_TARIFFE = "pp_tariffe_v1";
  const LS_URGENZE = "pp_urgenze_v1";
  const LS_CASI = "pp_casi_v2";
  const LS_SERVIZI = "pp_servizi_sel_v1";
  const LS_PRESET_URGENZA = "pp_last_urgenza_v1";
  const LS_ROUND_MODE = "pp_round_mode_v1";
  const LS_DISTANCES = "pp_distanze_v2"; // bump to v2 for km+min

  // Demo data
  const DEMO_TARIFFE = [
    {Codice:"T-INT-130", Voce:"Intervento tecnico base", Unita:"ora", PrezzoIVA:"130", Note:"Tariffa oraria standard"},
    {Codice:"T-VID-COMB", Voce:"Videoispezione in combinata", Unita:"ora", PrezzoIVA:"60", Note:"Applicata in Qtà = ore"},
    {Codice:"T-VID-FIN", Voce:"Videoispezione finale", Unita:"cad", PrezzoIVA:"30", Note:"Importo fisso a fine lavorazione"},
    {Codice:"T-LOC", Voce:"Localizzazione punto", Unita:"cad", PrezzoIVA:"30", Note:"Per punto localizzato"},
    {Codice:"T-TRASF-CL", Voce:"Trasferimento cliente", Unita:"min", PrezzoIVA:"1", Note:"€/min al cliente"},
    {Codice:"T-TRASF-INT", Voce:"Trasferimento interno", Unita:"min", PrezzoIVA:"2", Note:"€/min interno (uso tecnico)"},
    {Codice:"S-LAV-CUC", Voce:"Pulizia scarico cucina", Unita:"cad", PrezzoIVA:"80", Note:"Elemento aggiuntivo"},
    {Codice:"S-COL-SCA", Voce:"Pulizia colonna di scarico", Unita:"cad", PrezzoIVA:"150", Note:"Elemento aggiuntivo"},
    {Codice:"S-TAG-RAD", Voce:"Taglio radici da tubazione", Unita:"cad", PrezzoIVA:"220", Note:"Elemento aggiuntivo"}
  ];
  const DEMO_URGENZE = [
    {Caso:"Intervento entro 4h", Descrizione:"Priorità massima (serale/festivo)", SovrapprezzoIVA:"40", Priorita:"3"},
    {Caso:"Intervento entro 24h", Descrizione:"Alta priorità", SovrapprezzoIVA:"20", Priorita:"2"}
  ];
  const DEMO_CASI = [
    {Codice:"T-INT-130", Tipo:"Intervento tecnico base", Descrizione:"Tariffa oraria standard", TariffaIVA:"130", Note:""},
    {Codice:"T-VID-COMB", Tipo:"Videoispezione", Descrizione:"Ispezione con telecamera", TariffaIVA:"60", Note:""},
    {Codice:"S-LAV-CUC", Tipo:"Pulizia scarico cucina", Descrizione:"Lavello / linea cucina", TariffaIVA:"", Note:""},
    {Codice:"S-COL-SCA", Tipo:"Pulizia colonna di scarico", Descrizione:"Colonna condominiale", TariffaIVA:"", Note:""},
    {Codice:"S-TAG-RAD", Tipo:"Taglio radici da tubazione", Descrizione:"Rimozione radici", TariffaIVA:"", Note:""}
  ];
  const DEMO_DISTANZE = {
    "Bibbona": {km:30, min:55},
    "Campiglia Marittima": {km:70, min:110},
    "Casale Marittimo": {km:28, min:50},
    "Castagneto Carducci": {km:36, min:60},
    "Castellina Marittima": {km:42, min:70},
    "Castelnuovo della Misericordia": {km:24, min:45},
    "Cecina": {km:20, min:50},
    "Chioma": {km:18, min:40},
    "Collemezzano": {km:16, min:35},
    "Collesalvetti": {km:28, min:50},
    "Gabbro": {km:22, min:45},
    "Guardistallo": {km:26, min:50},
    "Livorno": {km:0, min:0},
    "Marina di Bibbona": {km:32, min:58},
    "Montescudaio": {km:30, min:55},
    "Nibbiaia": {km:20, min:45},
    "Orciano Pisano": {km:34, min:60},
    "Pastina": {km:36, min:62},
    "Piombino": {km:100, min:160},
    "Quercianella": {km:14, min:32},
    "Riparbella": {km:32, min:58},
    "Rosignano Marittimo": {km:10, min:25},
    "San Vincenzo": {km:60, min:100},
    "Santa Luce": {km:30, min:55}
  };

  // Safe getters
  function safeParseLS(key, demo){
    try{
      const raw = localStorage.getItem(key);
      if(!raw) return demo.slice ? demo.slice() : JSON.parse(JSON.stringify(demo));
      const val = JSON.parse(raw);
      if(Array.isArray(demo) && (!Array.isArray(val) || val.length===0)) return demo.slice();
      return val;
    }catch(e){
      localStorage.setItem(key, JSON.stringify(demo.slice ? demo.slice() : demo));
      return demo.slice ? demo.slice() : JSON.parse(JSON.stringify(demo));
    }
  }
  const loadTariffe = ()=> safeParseLS(LS_TARIFFE, DEMO_TARIFFE);
  const saveTariffe = (rows)=> localStorage.setItem(LS_TARIFFE, JSON.stringify(rows));
  const loadUrgenze = ()=> safeParseLS(LS_URGENZE, DEMO_URGENZE);
  const saveUrgenze = (rows)=> localStorage.setItem(LS_URGENZE, JSON.stringify(rows));
  const loadCasi = ()=> safeParseLS(LS_CASI, DEMO_CASI);
  const saveCasi = (rows)=> localStorage.setItem(LS_CASI, JSON.stringify(rows));
  const loadServiziSel = ()=> { try{ return JSON.parse(localStorage.getItem(LS_SERVIZI)||"[]"); }catch{return [];} };
  const saveServiziSel = (rows)=> localStorage.setItem(LS_SERVIZI, JSON.stringify(rows));

  // Distances with migration (supports legacy structures)
  const loadDistanze = () => {
    try{
      const raw = localStorage.getItem(LS_DISTANCES);
      if(!raw) return JSON.parse(JSON.stringify(DEMO_DISTANZE));
      const d = JSON.parse(raw);
      const out = {};
      Object.keys(d).forEach(k=>{
        const v = d[k];
        if(typeof v === "number"){
          out[k] = {km: Number(v)||0, min: 0};
        }else{
          out[k] = {km: Number(v.km)||0, min: Number(v.min)||0};
        }
      });
      return out;
    }catch{
      return JSON.parse(JSON.stringify(DEMO_DISTANZE));
    }
  };
  const saveDistanze = (data) => localStorage.setItem(LS_DISTANCES, JSON.stringify(data));
  function ensureDefaultDistanze(){
    try{
      const raw = localStorage.getItem(LS_DISTANCES);
      if(!raw){ localStorage.setItem(LS_DISTANCES, JSON.stringify(DEMO_DISTANZE)); }
    }catch{
      localStorage.setItem(LS_DISTANCES, JSON.stringify(DEMO_DISTANZE));
    }
  }

  // UI refs
  const cliente = el("cliente");
  const localitaSel = el("localita");
  const indirizzo = el("indirizzo");
  const tipoIntervento = el("tipoIntervento");
  const tariffaHint = el("tariffaCaseHint");
  const kmAR = el("kmAR");
  const minutiTrasf = el("minutiTrasf");
  const round5 = el("round5");
  const oreLavoro = el("oreLavoro");
  const tariffaInterventoIvaInc = el("tariffaInterventoIvaInc");
  const videoComb = el("videoComb");
  const videoFinale = el("videoFinale");
  const qtaLoc = el("qtaLoc");
  const trasfClienteIvaInc = el("trasfClienteIvaInc");
  const trasfInternoIvaInc = el("trasfInternoIvaInc");
  const videoCombIvaInc = el("videoCombIvaInc");
  const videoFinaleIvaInc = el("videoFinaleIvaInc");
  const locCadIvaInc = el("locCadIvaInc");
  const lockTariffe = el("lockTariffe");
  const imponibile = el("imponibile");
  const iva = el("iva");
  const totale = el("totale");
  const msgCliente = el("msgCliente");
  const riepilogoTecnico = el("riepilogoTecnico");
  const waBtn = el("waBtn");
  const pdfBtn = el("pdfBtn");
  const copyBtn = el("copyBtn");
  const resetBtn = el("resetBtn");
  const presetUrgenza = el("presetUrgenza");
  const roundMode = el("roundMode");

  // Elementi aggiuntivi
  const servizioSel = el("servizioSel");
  const servizioQta = el("servizioQta");
  const addServizio = el("addServizio");
  const tabServiziBody = el("tabServizi").querySelector("tbody");
  const serviziHint = el("serviziHint");

  // Distanze tab
  const tabDistanzeBody = el("tabDistanze").querySelector("tbody");
  const searchDistanze = el("searchDistanze");
  const addDistanza = el("addDistanza");
  const importDistanze = el("importDistanze");
  const exportDistanze = el("exportDistanze");

  // nav
  function setActiveTab(hash){
    qsa(".tab").forEach(t=>t.classList.remove("active"));
    qsa(".page").forEach(p=>p.classList.add("hidden"));
    if(hash==="#tariffe"){ qs("#tab-tariffe").classList.add("active"); qs("#page-tariffe").classList.remove("hidden"); }
    else if(hash==="#urgenze"){ qs("#tab-urgenze").classList.add("active"); qs("#page-urgenze").classList.remove("hidden"); }
    else if(hash==="#casi"){ qs("#tab-casi").classList.add("active"); qs("#page-casi").classList.remove("hidden"); }
    else if(hash==="#distanze"){ qs("#tab-distanze").classList.add("active"); qs("#page-distanze").classList.remove("hidden"); }
    else { qs("#tab-preventivo").classList.add("active"); qs("#page-preventivo").classList.remove("hidden"); }
  }
  window.addEventListener("hashchange", ()=>setActiveTab(location.hash));
  setActiveTab(location.hash||"#preventivo");

  // init località
  LOCALITA.forEach(c=>{ const o=document.createElement("option"); o.value=c; o.textContent=c; localitaSel.appendChild(o); });

  // helpers
  const ivaAliquota = 0.22;
  const roundToEuro = (n,mode)=> (mode==="up"?Math.ceil(n):mode==="down"?Math.floor(n):Math.round(n));
  const grossToNet = (v) => round2(Number(v)/(1+ivaAliquota));

  // populate elementi dropdown
  const loadTariffeSafe = ()=> loadTariffe().filter(r=> (Number(r.PrezzoIVA)||0) > 0);
  function refreshServiziDropdown(){
    const tar = loadTariffeSafe();
    servizioSel.innerHTML = "";
    const o0=document.createElement("option"); o0.value=""; o0.textContent="— Seleziona elemento —"; servizioSel.appendChild(o0);
    let countValid = 0;
    tar.forEach(r=>{
      const prezzo = Number(r.PrezzoIVA)||0;
      const o=document.createElement("option");
      o.value = r.Codice || r.Voce;
      o.textContent = `${r.Voce} (${r.Codice || "—"}) – ${eur(prezzo)}`;
      o.dataset.codice = r.Codice||"";
      o.dataset.voce = r.Voce||"";
      o.dataset.prezzo = String(prezzo);
      servizioSel.appendChild(o);
      countValid++;
    });
    serviziHint.textContent = countValid ? "" : "Nessun elemento con prezzo valido: aggiungi voci in Tariffe o importa il CSV.";
  }

  function renderServiziSel(){
    const rows = loadServiziSel();
    tabServiziBody.innerHTML = "";
    rows.forEach((r, idx)=>{
      const tr=document.createElement("tr");
      const td = (t)=>{ const d=document.createElement("td"); d.textContent=t; return d; };
      tr.appendChild(td(r.codice||""));
      tr.appendChild(td(r.voce||""));
      tr.appendChild(td(eur(r.prezzo||0)));
      const tdQ=document.createElement("td"); const q=document.createElement("input"); q.type="number"; q.min="1"; q.step="1"; q.value=r.qta||1; q.addEventListener("input", ()=>{ const all=loadServiziSel(); all[idx].qta = Number(q.value)||1; saveServiziSel(all); computePreventivo(); }); tdQ.appendChild(q); tr.appendChild(tdQ);
      tr.appendChild(td(eur((Number(r.prezzo)||0) * (Number(r.qta)||1))));
      const tdA=document.createElement("td"); const del=document.createElement("button"); del.textContent="🗑️"; del.className="apply-btn"; del.addEventListener("click", ()=>{ const all=loadServiziSel(); all.splice(idx,1); saveServiziSel(all); renderServiziSel(); computePreventivo(); }); tdA.appendChild(del); tr.appendChild(tdA);
      tabServiziBody.appendChild(tr);
    });
  }
  addServizio.addEventListener("click", ()=>{
    const opt = servizioSel.selectedOptions[0]; if(!opt || !opt.dataset) return;
    const prezzo = Number(opt.dataset.prezzo)||0; if(!prezzo) { alert("L'elemento selezionato non ha un prezzo valido."); return; }
    const item = { codice: opt.dataset.codice || "", voce: opt.dataset.voce || opt.textContent, prezzo: prezzo, qta: Number(servizioQta.value)||1 };
    const all = loadServiziSel(); all.push(item); saveServiziSel(all); renderServiziSel(); computePreventivo();
  });

  // Tipi intervento (Casi) + mapping
  const tipoInterventoMap = new Map();
  function refreshTipiIntervento(){
    const casi = loadCasi();
    tipoIntervento.innerHTML = "";
    const o0=document.createElement("option"); o0.value=""; o0.textContent="— Seleziona tipologia —"; tipoIntervento.appendChild(o0);
    tipoInterventoMap.clear();
    casi.forEach(r=>{
      const t = r.Tipo || "";
      const cod = r.Codice || "";
      const tIva = r.TariffaIVA || "";
      tipoInterventoMap.set(t, {codice: cod, tariffaIva: tIva});
      const o=document.createElement("option"); o.value=t; o.textContent = t; o.dataset.codice = cod; o.dataset.tariffa = tIva; tipoIntervento.appendChild(o);
    });
  }

  // Urgenze preset
  function refreshPresetUrgenze(){
    const rows = loadUrgenze();
    presetUrgenza.innerHTML = "";
    const o0=document.createElement("option"); o0.value=""; o0.textContent="Nessuna"; presetUrgenza.appendChild(o0);
    rows.sort((a,b)=>Number(b.Priorita||0)-Number(a.Priorita||0)).forEach(r=>{
      const o=document.createElement("option"); o.value=r.SovrapprezzoIVA||""; o.textContent = `${r.Caso} ${r.SovrapprezzoIVA?`(+${eur(r.SovrapprezzoIVA)})`:""}`; presetUrgenza.appendChild(o);
    });
    const last = localStorage.getItem(LS_PRESET_URGENZA)||"";
    presetUrgenza.value = last;
  }

  // Rounding util
  function roundToEuroVal(val){
    const n = Number(val)||0;
    const mode = localStorage.getItem(LS_ROUND_MODE) || "nearest";
    return roundToEuro(n, mode);
  }

  // Compute
  function computePreventivo(){
    const t = tipoIntervento.value;
    let hint = "";
    if(t && !lockTariffe.checked){
      const info = tipoInterventoMap.get(t);
      const tariffe = loadTariffe();
      let applied = null;
      if(info && info.codice){
        const byCode = tariffe.find(x=>String(x.Codice||"")===String(info.codice));
        if(byCode){ tariffaInterventoIvaInc.value = byCode.PrezzoIVA || tariffaInterventoIvaInc.value; applied = {src:"Listino", val: byCode.PrezzoIVA}; }
      }
      if(!applied && info && info.tariffaIva){
        tariffaInterventoIvaInc.value = info.tariffaIva; applied = {src:"Casi", val: info.tariffaIva};
      }
      if(applied){ hint = `Tariffa da ${applied.src}: ${eur(applied.val)}/h`; }
    }
    tariffaHint.textContent = hint;

    const mInput = Number(minutiTrasf.value)||0;
    const minutiEff = round5.checked ? ceilToFive(mInput) : mInput;
    const ore = Number(oreLavoro.value)||0;

    const lavoroNetto = round2((Number(tariffaInterventoIvaInc.value)||0) * ore / (1+0.22));
    const trasferClienteNetto = round2((Number(trasfClienteIvaInc.value)||0)/(1+0.22) * minutiEff);
    const videoCombNetto = videoComb.checked ? round2((Number(videoCombIvaInc.value)||0)/(1+0.22) * ore) : 0;
    const videoFinaleNetto = videoFinale.checked ? round2((Number(videoFinaleIvaInc.value)||0)/(1+0.22)) : 0;
    const locLordo = (Number(locCadIvaInc.value)||0) * (Number(qtaLoc.value)||0);
    const locNetto = round2(locLordo/(1+0.22));

    const servizi = loadServiziSel();
    const serviziLordo = servizi.reduce((s,r)=> s + (Number(r.prezzo)||0) * (Number(r.qta)||1), 0);
    const serviziNetto = round2(serviziLordo/(1+0.22));

    const urgenzaGross = Number(presetUrgenza.value)||0;
    const urgenzaNet = urgenzaGross ? round2(urgenzaGross/(1+0.22)) : 0;

    const impon = round2(lavoroNetto + trasferClienteNetto + videoCombNetto + videoFinaleNetto + locNetto + serviziNetto + urgenzaNet);
    const ivaVal = round2(impon * 0.22);
    const tot = round2(impon + ivaVal);
    const totRounded = roundToEuroVal(tot);

    imponibile.textContent = eur(impon);
    iva.textContent = eur(ivaVal);
    totale.textContent = totRounded.toLocaleString('it-IT',{style:'currency',currency:'EUR'});

    const righe = [];
    righe.push("Buongiorno,");
    if (cliente.value) righe.push(`👤 **Cliente:** ${cliente.value}`);
    righe.push("\n👨‍🔧 **Invio il preventivo per l’intervento richiesto presso:**");
    righe.push(`📍 ${indirizzo.value} – ${localitaSel.value}`);
    if (tipoIntervento.value){ 
      const info = tipoInterventoMap.get(tipoIntervento.value)||{};
      const cod = info.codice ? ` (Cod. ${info.codice})` : "";
      righe.push(`🗂️ **Tipologia intervento:** ${tipoIntervento.value}${cod}`);
    }
    righe.push("\n**Dettaglio costi**");
    righe.push(`📍 Km stimati: ${(Number(kmAR.value)||0)} km A/R`);
    righe.push(`⏱️ Tempo stimato di trasferimento: ${minutiEff} minuti A/R`);
    righe.push(`🚐 Trasferimento tecnico: ${eur(trasferClienteNetto)}`);
    righe.push("🛠️ Intervento tecnico:");
    righe.push(`   • ${ore} h × ${eur((Number(tariffaInterventoIvaInc.value)||0)/(1+0.22))} (netto) = ${eur(lavoroNetto)}`);
    if (videoComb.checked) righe.push(`🎥 Videoispezione in combinata = ${eur(videoCombNetto)}`);
    if (videoFinale.checked) righe.push(`🎥 Videoispezione finale = ${eur(videoFinaleNetto)}`);
    if ((Number(qtaLoc.value)||0)>0) righe.push(`📍 Localizzazione (n. ${qtaLoc.value}) = ${eur(locNetto)}`);
    if (servizi.length){
      righe.push("➕ Elementi aggiuntivi:");
      servizi.forEach(s=> righe.push(`   • ${s.voce} ${s.codice?`(Cod. ${s.codice})`:``} × ${s.qta}`));
    }
    if (urgenzaGross>0) righe.push(`🚨 Urgenza = ${eur(urgenzaNet)} (da ${eur(urgenzaGross)} IVA incl.)`);
    righe.push("\n**Riepilogo**");
    righe.push(`**💰 Totale complessivo stimato (netto): ${eur(impon)}**`);
    righe.push(`**➕ IVA 22%: ${eur(ivaVal)}**`);
    righe.push(`**💰 Totale IVA inclusa: ${totRounded.toLocaleString('it-IT',{style:'currency',currency:'EUR'})}**`);
    righe.push("\n📅 Intervento entro 48h salvo disponibilità. 🚀");
    righe.push("\n---\n\n⚠️ IMPORTANTE:  \nIl costo totale dell’intervento può variare in base a:\n\n• Lunghezza della tubazione  \n• Complessità dell’intervento (curve a 90°, braghe, ecc.)  \n• Tenacia/durezza dell’ostruzione (calcare, radici, grassi calcificati)\n\n#PreventivoTrasparente #NienteSorprese  \n⛑️ Alessandro Cicalini – PRO-PIPE\n\n---");
    msgCliente.value = righe.join("\n");

    const trasferInternoNetto = round2((Number(trasfInternoIvaInc.value)||0)/(1+0.22) * minutiEff);
    const totaleTecnicoNetto = impon;
    const tempoTotMin = ore*60 + minutiEff;
    const valoreOrarioMedioNetto = tempoTotMin>0 ? round2((totaleTecnicoNetto/tempoTotMin)*60) : 0;

    const righeTec = [];
    righeTec.push("RIEPILOGO TECNICO (uso interno)");
    if (tipoIntervento.value){ const info=tipoInterventoMap.get(tipoIntervento.value)||{}; righeTec.push(`- Tipologia: ${tipoIntervento.value}${info.codice?` (Cod. ${info.codice})`:``}`); }
    righeTec.push(`- Cliente: ${cliente.value || "[nome]"}`);
    righeTec.push(`- Tempo e km A/R: ${minutiEff} min, ${(Number(kmAR.value)||0)} km`);
    righeTec.push(`- Trasferimento interno: ${eur(trasfInternoIvaInc.value)} (IVA incl.) → netto ${eur((Number(trasfInternoIvaInc.value)||0)/(1+0.22))}/min → ${eur(trasferInternoNetto)}`);
    righeTec.push(`- Intervento tecnico (netto): ${eur(lavoroNetto)} (da ${eur(tariffaInterventoIvaInc.value)}/h IVA incl.)`);
    if (videoComb.checked) righeTec.push(`- Video combinata: ${eur(videoCombNetto)}`);
    if (videoFinale.checked) righeTec.push(`- Video finale: ${eur(videoFinaleNetto)}`);
    if ((Number(qtaLoc.value)||0)>0) righeTec.push(`- Localizzazioni: ${eur(locNetto)}`);
    if (servizi.length) righeTec.push(`- Elementi aggiuntivi (netto): ${eur(serviziNetto)} (lordo ${eur(serviziLordo)})`);
    righeTec.push(`- Totale netto: ${eur(totaleTecnicoNetto)} | IVA: ${eur(ivaVal)} | Totale arrotondato: ${totRounded.toLocaleString('it-IT',{style:'currency',currency:'EUR'})}`);
    riepilogoTecnico.value = righeTec.join("\n");
  }

  // Listeners
  qsa("input,select").forEach(n=>{ n.addEventListener("input", computePreventivo); n.addEventListener("change", computePreventivo); });
  // Fill km & min on location change from distances table
  const setFromDistances = ()=>{ const d=loadDistanze(); const rec=d[localitaSel.value]; if(rec){ kmAR.value = Number(rec.km)||0; minutiTrasf.value = Number(rec.min)||0; } };
  localitaSel.addEventListener("change", ()=>{ setFromDistances(); computePreventivo(); });
  copyBtn.addEventListener("click", async ()=>{
    try{ await navigator.clipboard.writeText(msgCliente.value); alert("Testo cliente copiato ✅"); }
    catch{ alert("Seleziona e copia manualmente."); }
  });
  pdfBtn.addEventListener("click", ()=>window.print());
  waBtn.addEventListener("click", ()=>{ const text = encodeURIComponent(msgCliente.value); window.open(`https://wa.me/?text=${text}`, "_blank"); });
  presetUrgenza.addEventListener("change", ()=>{ localStorage.setItem(LS_PRESET_URGENZA, presetUrgenza.value); computePreventivo(); });
  roundMode.addEventListener("change", ()=>{ localStorage.setItem(LS_ROUND_MODE, roundMode.value); computePreventivo(); });

  // Reset form (non distruttivo su listini)
  resetBtn.addEventListener("click", ()=>{
    cliente.value=""; indirizzo.value=""; localitaSel.value = LOCALITA[0];
    tipoIntervento.value=""; tariffaHint.textContent="";
    kmAR.value=0; minutiTrasf.value=0; round5.checked=true;
    oreLavoro.value=1; tariffaInterventoIvaInc.value=130;
    videoComb.checked=false; videoFinale.checked=false; qtaLoc.value=0;
    trasfClienteIvaInc.value=1; trasfInternoIvaInc.value=2;
    videoCombIvaInc.value=60; videoFinaleIvaInc.value=30; locCadIvaInc.value=30;
    presetUrgenza.value="";
    localStorage.removeItem(LS_SERVIZI);
    renderServiziSel();
    computePreventivo();
  });

  // Tariffe page
  const tabTariffeBody = el("tabTariffe").querySelector("tbody");
  const searchTariffe = el("searchTariffe");
  const importTariffe = el("importTariffe");
  const exportTariffe = el("exportTariffe");
  const addTariffa = el("addTariffa");
  function renderTariffe(){
    const q = (searchTariffe.value||"").toLowerCase();
    const rows = loadTariffe();
    tabTariffeBody.innerHTML="";
    rows.forEach((r,idx)=>{
      if(q && !(`${r.Codice} ${r.Voce} ${r.Note}`.toLowerCase().includes(q))) return;
      const tr=document.createElement("tr"); const mk=(key,type="text")=>{ const td=document.createElement("td"); const input=document.createElement("input"); input.value=r[key]??""; input.type=type; input.addEventListener("input", ()=>{ const all=loadTariffe(); all[idx][key]=input.value; saveTariffe(all); refreshServiziDropdown(); computePreventivo(); }); td.appendChild(input); return td; };
      tr.appendChild(mk("Codice")); tr.appendChild(mk("Voce")); tr.appendChild(mk("Unita")); tr.appendChild(mk("PrezzoIVA","number")); tr.appendChild(mk("Note"));
      const tdMap=document.createElement("td"); const sel=document.createElement("select"); ["— Seleziona —","Intervento €/h","Video combinata €/h","Video finale fisso","Localizzazione €/cad","Trasferimento cliente €/min","Trasferimento interno €/min"].forEach(v=>{ const o=document.createElement("option"); o.value=v; o.textContent=v; sel.appendChild(o); }); tdMap.appendChild(sel); tr.appendChild(tdMap);
      const tdAct=document.createElement("td"); const apply=document.createElement("button"); apply.textContent="Applica"; apply.className="apply-btn";
      apply.addEventListener("click", ()=>{ const price = Number(r.PrezzoIVA)||0; switch(sel.value){ case "Intervento €/h": tariffaInterventoIvaInc.value = price; break; case "Video combinata €/h": videoCombIvaInc.value = price; break; case "Video finale fisso": videoFinaleIvaInc.value = price; break; case "Localizzazione €/cad": locCadIvaInc.value = price; break; case "Trasferimento cliente €/min": trasfClienteIvaInc.value = price; break; case "Trasferimento interno €/min": trasfInternoIvaInc.value = price; break; default: alert("Seleziona un campo da mappare."); return; } computePreventivo(); location.hash="#preventivo"; });
      const del=document.createElement("button"); del.textContent="🗑️"; del.className="apply-btn"; del.style.marginLeft="6px"; del.addEventListener("click", ()=>{ const all=loadTariffe(); all.splice(idx,1); saveTariffe(all); renderTariffe(); refreshServiziDropdown(); computePreventivo(); });
      tdAct.appendChild(apply); tdAct.appendChild(del); tr.appendChild(tdAct);
      tabTariffeBody.appendChild(tr);
    });
  }
  searchTariffe.addEventListener("input", renderTariffe);
  addTariffa.addEventListener("click", ()=>{ const rows=loadTariffe(); rows.push({Codice:"",Voce:"",Unita:"",PrezzoIVA:"",Note:""}); saveTariffe(rows); renderTariffe(); });
  exportTariffe.addEventListener("click", ()=>{ const rows=loadTariffe(); const headers=["Codice","Voce","Unita","PrezzoIVA","Note"]; const csvRows=[headers].concat(rows.map(r=>headers.map(h=>r[h]??""))); const blob=new Blob([toCSV(csvRows,";")],{type:"text/csv;charset=utf-8"}); const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="Tariffe.csv"; a.click(); });
  importTariffe.addEventListener("change",(ev)=>{ const file=ev.target.files[0]; if(!file) return; const fr=new FileReader(); fr.onload=()=>{ const rows=parseCSV(fr.result); const headers=rows.shift().map(h=>h.trim()); const idx={Codice:headers.indexOf("Codice"), Voce:headers.indexOf("Voce"), Unita:headers.indexOf("Unita"), PrezzoIVA:headers.indexOf("PrezzoIVA"), Note:headers.indexOf("Note")}; const out=rows.map(r=>({Codice:r[idx.Codice]||"", Voce:r[idx.Voce]||"", Unita:r[idx.Unita]||"", PrezzoIVA:r[idx.PrezzoIVA]||"", Note:r[idx.Note]||""})); saveTariffe(out); renderTariffe(); refreshServiziDropdown(); computePreventivo(); }; fr.readAsText(file,"utf-8"); });
  renderTariffe();

  // Urgenze page
  const tabUrgenzeBody = el("tabUrgenze").querySelector("tbody");
  const searchUrgenze = el("searchUrgenze");
  const importUrgenze = el("importUrgenze");
  const exportUrgenze = el("exportUrgenze");
  const addUrgenza = el("addUrgenza");
  function renderUrgenze(){
    const q=(searchUrgenze.value||"").toLowerCase();
    const rows=loadUrgenze();
    tabUrgenzeBody.innerHTML="";
    rows.forEach((r,idx)=>{
      if(q && !(`${r.Caso} ${r.Descrizione}`.toLowerCase().includes(q))) return;
      const tr=document.createElement("tr"); const mk=(key,type="text")=>{ const td=document.createElement("td"); const input=document.createElement("input"); input.value=r[key]||""; input.type=type; input.addEventListener("input", ()=>{ const all=loadUrgenze(); all[idx][key]=input.value; saveUrgenze(all); refreshPresetUrgenze(); }); td.appendChild(input); return td; };
      tr.appendChild(mk("Caso")); tr.appendChild(mk("Descrizione")); tr.appendChild(mk("SovrapprezzoIVA","number")); tr.appendChild(mk("Priorita","number"));
      const tdA=document.createElement("td"); const apply=document.createElement("button"); apply.textContent="Imposta preset"; apply.className="apply-btn"; apply.addEventListener("click", ()=>{ presetUrgenza.value=r.SovrapprezzoIVA||""; localStorage.setItem(LS_PRESET_URGENZA, presetUrgenza.value); computePreventivo(); location.hash="#preventivo"; });
      const del=document.createElement("button"); del.textContent="🗑️"; del.className="apply-btn"; del.style.marginLeft="6px"; del.addEventListener("click", ()=>{ const all=loadUrgenze(); all.splice(idx,1); saveUrgenze(all); renderUrgenze(); refreshPresetUrgenze(); });
      tdA.appendChild(apply); tdA.appendChild(del); tr.appendChild(tdA);
      tabUrgenzeBody.appendChild(tr);
    });
  }
  searchUrgenze.addEventListener("input", renderUrgenze);
  addUrgenza.addEventListener("click", ()=>{ const rows=loadUrgenze(); rows.push({Caso:"",Descrizione:"",SovrapprezzoIVA:"",Priorita:""}); saveUrgenze(rows); renderUrgenze(); });
  exportUrgenze.addEventListener("click", ()=>{ const rows=loadUrgenze(); const headers=["Caso","Descrizione","SovrapprezzoIVA","Priorita"]; const csvRows=[headers].concat(rows.map(r=>headers.map(h=>r[h]??""))); const blob=new Blob([toCSV(csvRows,";")],{type:"text/csv;charset=utf-8"}); const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="Urgenze.csv"; a.click(); });
  importUrgenze.addEventListener("change",(ev)=>{ const file=ev.target.files[0]; if(!file) return; const fr=new FileReader(); fr.onload=()=>{ const rows=parseCSV(fr.result); const headers=rows.shift().map(h=>h.trim()); const idx={Caso:headers.indexOf("Caso"), Descrizione:headers.indexOf("Descrizione"), SovrapprezzoIVA:headers.indexOf("SovrapprezzoIVA"), Priorita:headers.indexOf("Priorita")}; const out=rows.map(r=>({Caso:r[idx.Caso]||"", Descrizione:r[idx.Descrizione]||"", SovrapprezzoIVA:r[idx.SovrapprezzoIVA]||"", Priorita:r[idx.Priorita]||""})); saveUrgenze(out); renderUrgenze(); refreshPresetUrgenze(); computePreventivo(); }; fr.readAsText(file,"utf-8"); });
  renderUrgenze();

  // Casi page
  const tabCasiBody = el("tabCasi").querySelector("tbody");
  const searchCasi = el("searchCasi");
  const importCasi = el("importCasi");
  const exportCasi = el("exportCasi");
  const addCaso = el("addCaso");
  function renderCasi(){
    const q=(searchCasi.value||"").toLowerCase();
    const rows=loadCasi();
    tabCasiBody.innerHTML="";
    rows.forEach((r,idx)=>{
      if(q && !(`${r.Codice} ${r.Tipo} ${r.Descrizione}`.toLowerCase().includes(q))) return;
      const tr=document.createElement("tr"); const mk=(key,type="text")=>{ const td=document.createElement("td"); const input=document.createElement("input"); input.value=r[key]||""; input.type=type; input.addEventListener("input", ()=>{ const all=loadCasi(); all[idx][key]=input.value; saveCasi(all); refreshTipiIntervento(); }); td.appendChild(input); return td; };
      tr.appendChild(mk("Codice")); tr.appendChild(mk("Tipo")); tr.appendChild(mk("Descrizione")); tr.appendChild(mk("TariffaIVA","number")); tr.appendChild(mk("Note"));
      const tdA=document.createElement("td"); const use=document.createElement("button"); use.textContent="Usa nel preventivo"; use.className="apply-btn"; use.addEventListener("click", ()=>{ tipoIntervento.value = r.Tipo || ""; computePreventivo(); location.hash="#preventivo"; });
      const del=document.createElement("button"); del.textContent="🗑️"; del.className="apply-btn"; del.style.marginLeft="6px"; del.addEventListener("click", ()=>{ const all=loadCasi(); all.splice(idx,1); saveCasi(all); renderCasi(); refreshTipiIntervento(); });
      tdA.appendChild(use); tdA.appendChild(del); tr.appendChild(tdA);
      tabCasiBody.appendChild(tr);
    });
  }
  searchCasi.addEventListener("input", renderCasi);
  addCaso.addEventListener("click", ()=>{ const rows=loadCasi(); rows.push({Codice:"",Tipo:"",Descrizione:"",TariffaIVA:"",Note:""}); saveCasi(rows); renderCasi(); });
  exportCasi.addEventListener("click", ()=>{ const rows=loadCasi(); const headers=["Codice","Tipo","Descrizione","TariffaIVA","Note"]; const csvRows=[headers].concat(rows.map(r=>headers.map(h=>r[h]??""))); const blob=new Blob([toCSV(csvRows,";")],{type:"text/csv;charset=utf-8"}); const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="Casi.csv"; a.click(); });
  importCasi.addEventListener("change",(ev)=>{ const file=ev.target.files[0]; if(!file) return; const fr=new FileReader(); fr.onload=()=>{ const rows=parseCSV(fr.result); const headers=rows.shift().map(h=>h.trim()); const idx={Codice:headers.indexOf("Codice"), Tipo:headers.indexOf("Tipo"), Descrizione:headers.indexOf("Descrizione"), TariffaIVA:headers.indexOf("TariffaIVA"), Note:headers.indexOf("Note")}; const out=rows.map(r=>({Codice:r[idx.Codice]||"", Tipo:r[idx.Tipo]||"", Descrizione:r[idx.Descrizione]||"", TariffaIVA:r[idx.TariffaIVA]||"", Note:r[idx.Note]||""})); saveCasi(out); renderCasi(); refreshTipiIntervento(); computePreventivo(); }; fr.readAsText(file,"utf-8"); });
  renderCasi();

  // Distanze page
  function renderDistanze(){
    const q = (searchDistanze.value||"").toLowerCase();
    const rows = loadDistanze();
    tabDistanzeBody.innerHTML = "";
    Object.keys(rows).forEach(loc=>{
      if(q && !loc.toLowerCase().includes(q)) return;
      const tr=document.createElement("tr"); 
      const tdLoc=document.createElement("td"); tdLoc.textContent=loc; tr.appendChild(tdLoc);
      const tdKm=document.createElement("td"); const kmInput=document.createElement("input"); kmInput.type="number"; kmInput.min="0"; kmInput.value=rows[loc].km; kmInput.addEventListener("input", ()=>{ const all=loadDistanze(); all[loc].km=Number(kmInput.value)||0; saveDistanze(all); computePreventivo(); }); tdKm.appendChild(kmInput); tr.appendChild(tdKm);
      const tdMin=document.createElement("td"); const minInput=document.createElement("input"); minInput.type="number"; minInput.min="0"; minInput.value=rows[loc].min; minInput.addEventListener("input", ()=>{ const all=loadDistanze(); all[loc].min=Number(minInput.value)||0; saveDistanze(all); computePreventivo(); }); tdMin.appendChild(minInput); tr.appendChild(tdMin);
      const tdAct=document.createElement("td"); const del=document.createElement("button"); del.textContent="🗑️"; del.className="apply-btn"; del.addEventListener("click", ()=>{ const all=loadDistanze(); delete all[loc]; saveDistanze(all); renderDistanze(); }); tdAct.appendChild(del); tr.appendChild(tdAct);
      tabDistanzeBody.appendChild(tr);
    });
  }
  searchDistanze.addEventListener("input", renderDistanze);
  addDistanza.addEventListener("click", ()=>{
    const rows=loadDistanze();
    const newLoc = prompt("Inserisci il nome della località:");
    if(!newLoc || rows[newLoc]) return;
    rows[newLoc]={km:0,min:0}; saveDistanze(rows); renderDistanze();
  });
  exportDistanze.addEventListener("click", ()=>{
    const rows = loadDistanze();
    const headers = ["Localita","KmAR","TempoMinAR"];
    const csvRows = [headers].concat(Object.keys(rows).map(loc=>[loc, rows[loc].km, rows[loc].min]));
    const blob = new Blob([toCSV(csvRows, ";")], {type:"text/csv;charset=utf-8"});
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "Distanze.csv"; a.click();
  });
  importDistanze.addEventListener("change", (ev)=>{
    const file=ev.target.files[0]; if(!file) return;
    const fr=new FileReader();
    fr.onload=()=>{
      const rows=parseCSV(fr.result);
      const headers=rows.shift().map(h=>h.trim());
      const idxLoc=headers.findIndex(h=>/localit.|localita/i.test(h));
      const idxKm=headers.findIndex(h=>/km|distanza/i.test(h));
      const idxMin=headers.findIndex(h=>/tempo|min/i.test(h));
      const out={};
      rows.forEach(r=>{
        const loc=(r[idxLoc]||"").trim();
        if(!loc) return;
        const km=Number(r[idxKm]||0);
        const min=(idxMin>=0)? Number(r[idxMin]||0) : 0;
        out[loc]={km:km, min:min};
      });
      const current = loadDistanze();
      const merged = Object.assign({}, current, out);
      saveDistanze(merged); renderDistanze();
    };
    fr.readAsText(file, "utf-8");
  });
  renderDistanze();

  // Initial population
  ensureDefaultDistanze();
  roundMode.value = localStorage.getItem(LS_ROUND_MODE) || "nearest";
  if(!localStorage.getItem(LS_ROUND_MODE)) localStorage.setItem(LS_ROUND_MODE, "nearest");

  refreshTipiIntervento();
  refreshPresetUrgenze();
  refreshServiziDropdown();
  renderServiziSel();
  // autofill from distances
  setFromDistances();
  computePreventivo();
})();