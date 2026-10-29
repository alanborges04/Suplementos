/* Fluxo Turbo 2.0 — captura/aplicação de diário Excel Online.
 * Sem IA, sem leitura oculta de arquivos SharePoint, sem requisições externas.
 * Os valores são transferidos por um pacote copiado pelo próprio usuário.
 * Apenas TKS/ANND foi homologada com diário real de 23/10/2026.
 */
(function () {
  "use strict";
  const TYPE = "fluxo-turbo/tks-diario/1";
  const normalize = v => String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
  const money = v => Number.isFinite(v) ? v.toLocaleString("pt-BR", { style:"currency", currency:"BRL" }) : "sem valor";
  const round = v => Math.round((v + Number.EPSILON) * 100) / 100;
  const fmt = d => String(d.getUTCDate()).padStart(2,"0")+"/"+String(d.getUTCMonth()+1).padStart(2,"0")+"/"+d.getUTCFullYear();
  function dateOf(value) {
    if (typeof value === "number" && Number.isInteger(value) && value >= 40000 && value <= 90000) {
      return fmt(new Date(Date.UTC(1899,11,30)+value*86400000));
    }
    const s=String(value ?? "");
    const m=s.match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](20\d\d)/);
    if (!m) return null;
    const d=new Date(Date.UTC(+m[3],+m[2]-1,+m[1]));
    return d.getUTCFullYear()===+m[3] && d.getUTCMonth()===+m[2]-1 && d.getUTCDate()===+m[1] ? fmt(d):null;
  }
  function labelHits(matrix, matcher) {
    const hits=[];
    for(let row=0;row<matrix.length;row++){
      const cells=matrix[row]||[];
      for(let col=0;col<cells.length;col++) if(matcher(normalize(cells[col]))) hits.push({row,col});
    }
    return hits;
  }
  function getField(matrix, label, expected, offset, matcher) {
    const hits=labelHits(matrix,matcher);
    let loc;
    if (hits.length===1) loc=hits[0];
    else if(hits.length>1) {
      const near=hits.filter(p=>Math.abs(p.row-expected[0])<=3 && Math.abs(p.col-expected[1])<=3);
      if(near.length===1)loc=near[0];
    }
    if(!loc) throw new Error("Rótulo não encontrado de forma inequívoca: "+label+" (ocorrências: "+hits.length+").");
    const row=loc.row+offset[0],col=loc.col+offset[1];
    const val=matrix[row]?.[col];
    if(typeof val!=="number" || !Number.isFinite(val)) throw new Error("Valor inválido para "+label+" na posição calculada. Nenhum dado foi capturado.");
    return {value:round(val), address:String.fromCharCode(65+col)+(row+1), labelAddress:String.fromCharCode(65+loc.col)+(loc.row+1)};
  }
  function expectedFileDate(name) {
    const m=String(name||"").match(/(\d{2})\.(\d{2})\.(20\d{2})\.(?:xlsm|xlsx|xls)\b/i);
    return m?dateOf(m[1]+"/"+m[2]+"/"+m[3]):null;
  }
  const state={packet:null,preview:null};
  function element(tag, attrs={}, text) {
    const node=document.createElement(tag);
    for(const [key,value] of Object.entries(attrs)) {
      if(key==="className")node.className=value;
      else if(key==="style")node.style.cssText=value;
      else if(key==="id")node.id=value;
      else node.setAttribute(key,value);
    }
    if(text!==undefined)node.textContent=text;
    return node;
  }
  function setStatus(message,kind="info"){
    const e=document.getElementById("fluxo-capture-status");
    if(e){e.textContent=message;e.style.color=kind==="error"?"#B91C1C":kind==="ok"?"#166534":"#334155";}
  }
  function showPreview(message) {
    const target=document.getElementById("fluxo-preview");
    if(target)target.textContent=message;
  }
  function addUI(){
    const root=document.querySelector(".card");
    if(!root || document.getElementById("fluxo-capture-ui"))return;
    const box=element("div",{id:"fluxo-capture-ui",style:"border-top:1px solid #E5E7EB;margin-top:22px;padding-top:15px"});
    box.appendChild(element("h2",{}, "Fluxo Turbo 2.0 — Diário TKS"));
    box.appendChild(element("p",{}, "1. Abra o diário TKS e capture. 2. Abra o Fluxo Realizado, cole o pacote e visualize. 3. Confirme para preencher C, F e I. Os diários de origem não são alterados."));
    const capture=element("button",{id:"fluxo-capturar",className:"btn",style:"background:#1D4ED8;font-size:12px"},"1. CAPTURAR DIÁRIO ABERTO");
    box.appendChild(capture);
    const copy=element("textarea",{id:"fluxo-pacote",rows:"3",spellcheck:"false",style:"box-sizing:border-box;width:100%;margin-top:10px;padding:8px;border:1px solid #CBD5E1;border-radius:6px;font:11px monospace;resize:vertical",placeholder:"Pacote capturado aparecerá aqui. Na planilha mestre, cole o mesmo conteúdo neste campo."});
    box.appendChild(copy);
    const copyBtn=element("button",{id:"fluxo-copiar",className:"btn",style:"background:#475569;font-size:12px"},"COPIAR PACOTE");
    box.appendChild(copyBtn);
    const preview=element("button",{id:"fluxo-conferir",className:"btn",style:"background:#0F766E;font-size:12px"},"2. CONFERIR DESTINO NA PLANILHA MESTRE");
    box.appendChild(preview);
    box.appendChild(element("pre",{id:"fluxo-preview",style:"font:11px/1.55 Segoe UI, sans-serif;white-space:pre-wrap;word-break:break-word;background:#F1F5F9;border-radius:6px;padding:8px;min-height:30px"},"Nenhuma leitura executada."));
    const confirmLabel=element("label",{style:"display:flex;align-items:start;gap:8px;font:12px Segoe UI,sans-serif;margin-top:9px"});
    const confirm=element("input",{id:"fluxo-substituir",type:"checkbox"});
    confirmLabel.appendChild(confirm);
    confirmLabel.appendChild(element("span",{},"Autorizo substituir valores já preenchidos, se houver."));
    box.appendChild(confirmLabel);
    const apply=element("button",{id:"fluxo-aplicar",className:"btn",style:"background:#15803D;font-size:12px"},"3. APLICAR VALORES CONFERIDOS");
    apply.disabled=true;
    box.appendChild(apply);
    box.appendChild(element("div",{id:"fluxo-capture-status",role:"status",style:"margin-top:10px;font-size:12px;white-space:pre-wrap"},"Leitura por vínculo de arquivo fechado ainda não habilitada: requer autorização Microsoft 365."));
    root.appendChild(box);
    capture.addEventListener("click",()=>captureDaily(capture,copy));
    copyBtn.addEventListener("click",async()=> {
      if(!copy.value.trim())return setStatus("Capture um diário ou cole um pacote antes de copiar.","error");
      copy.select();
      try{await navigator.clipboard.writeText(copy.value);setStatus("Pacote copiado. Abra o suplemento na planilha mestre e cole no campo.");}
      catch{try{document.execCommand("copy");setStatus("Texto selecionado para cópia. Se necessário, use Ctrl+C.");}catch{setStatus("Selecione o texto e use Ctrl+C.");}}
    });
    preview.addEventListener("click",()=>previewTarget(preview,copy.value,apply));
    copy.addEventListener("input",()=>{state.preview=null;apply.disabled=true;showPreview("Pacote alterado. Confira o destino novamente.");});
    confirm.addEventListener("change",()=>{state.preview=null;apply.disabled=true;showPreview("Confirme novamente o destino antes da gravação.");});
    apply.addEventListener("click",()=>applyTarget(apply));
  }
  async function captureDaily(button,field){
    button.disabled=true;state.preview=null;
    try{
      setStatus("Lendo a aba DIÁRIO da pasta de trabalho aberta...");
      const pkt=await Excel.run(async ctx=>{
        const sheets=ctx.workbook.worksheets;
        sheets.load("items/name");await ctx.sync();
        const found=sheets.items.filter(s=>normalize(s.name)==="DIARIO");
        if(found.length!==1)throw new Error("Abra um diário com uma única aba DIÁRIO.");
        const sheet=found[0];
        const r=sheet.getRange("A1:J90");r.load("values");
        await ctx.sync();
        const m=r.values;
        const heading=normalize(m[0]?.[1]);
        if(!heading.includes("TOKAI SALVADOR") || !heading.includes("ANND")) throw new Error("O cabeçalho B1 não identifica TOKAI SALVADOR / ANND. Perfil TKS obrigatório.");
        const dt=dateOf(m[40]?.[8]);
        if(!dt)throw new Error("Data do diário não reconhecida em I41. Nenhum valor foi capturado.");
        const credit=getField(m,"TOTAL CRÉDITOS",[40,7],[0,-2],x=>x==="TOTAL CREDITOS");
        const debit=getField(m,"PAGAMENTOS",[41,7],[0,-2],x=>x==="PAGAMENTOS");
        const balance=getField(m,"SALDO FINAL B/C.I.",[54,3],[0,-2],x=>x.startsWith("SALDO FINAL B C I"));
        if(credit.value<0)throw new Error("Recebimentos negativos: verificar o diário.");
        if(debit.value>0)throw new Error("Pagamentos positivos: conferir regra de sinal antes de continuar.");
        const origin=String(Office.context.document?.url||"");
        const file=decodeURIComponent(origin.split("/").pop()?.split("?")[0]||"");
        const inFileDate=expectedFileDate(file);
        if(inFileDate && inFileDate!==dt)throw new Error("A data do arquivo aberto difere da data I41.");
        return {type:TYPE,store:"TKS",date:dt,credit:credit.value,debit:debit.value,balance:balance.value,
          source:{header:String(m[0]?.[1]||""),file:file||"diário aberto",ranges:{credit:credit.address,debit:debit.address,balance:balance.address}},
          capturedAt:new Date().toISOString()};
      });
      state.packet=pkt;
      field.value=JSON.stringify(pkt);
      showPreview("Captura concluída: "+pkt.store+" | "+pkt.date+"\nCréditos: "+money(pkt.credit)+"\nPagamentos: "+money(pkt.debit)+"\nSaldo final B/C.I.: "+money(pkt.balance)+"\nFonte: "+pkt.source.ranges.credit+", "+pkt.source.ranges.debit+", "+pkt.source.ranges.balance);
      setStatus("Valores capturados da planilha aberta. Copie o pacote e abra o Fluxo Realizado.","ok");
    }catch(e){setStatus("Captura interrompida: "+String(e.message||e),"error");}
    finally{button.disabled=false;}
  }
  function checkPacket(text) {
    let p;
    try{p=JSON.parse(text);}catch{throw new Error("Pacote não é um JSON válido.");}
    if(!p || p.type!==TYPE || p.store!=="TKS")throw new Error("Pacote incompatível. Nesta versão, somente TKS.");
    if(dateOf(p.date)!==p.date)throw new Error("Data do pacote inválida.");
    if(![p.credit,p.debit,p.balance].every(v=>typeof v==="number" && Number.isFinite(v) && Math.abs(v)<1e10))throw new Error("Valores do pacote inválidos.");
    if(p.credit<0 || p.debit>0)throw new Error("Sinais incompatíveis com o diário.");
    if(!p.source || !p.source.ranges || !/ANND/i.test(p.source.header))throw new Error("Origem do pacote não identifica a loja.");
    if(!p.capturedAt || Number.isNaN(Date.parse(p.capturedAt)))throw new Error("Data da captura ausente.");
    return p;
  }
  async function targetSnapshot(pkt) {
    return Excel.run(async ctx=>{
      const list=ctx.workbook.worksheets;
      list.load("items/name");await ctx.sync();
      const matches=list.items.filter(s=>normalize(s.name)==="TKS");
      if(matches.length!==1)throw new Error("Abra a planilha mestre com a aba TKS.");
      const sheet=matches[0];
      const dates=sheet.getRange("A5:A65");dates.load("values,text");
      await ctx.sync();
      const rows=[];
      for(let i=0;i<61;i++){
        const day=dateOf(dates.values[i]?.[0]) || dateOf(dates.text[i]?.[0]);
        if(day===pkt.date)rows.push(i+5);
      }
      if(rows.length!==1)throw new Error("Data "+pkt.date+" não encontrada de forma única na coluna A da aba TKS (ocorrências: "+rows.length+").");
      const row=rows[0];
      const items={};
      for(const c of ["C","F","I"]){const range=sheet.getRange(c+row);range.load("values,formulas,text");items[c]=range;}
      await ctx.sync();
      const before={};
      for(const c of ["C","F","I"]){
        before[c]={value:items[c].values[0][0],formula:items[c].formulas[0][0],text:items[c].text[0][0]};
      }
      return {sheet:"TKS",row,before};
    });
  }
  function snapshotId(s) {return JSON.stringify({sheet:s.sheet,row:s.row,before:s.before});}
  function alreadyNumeric(s) {
    return ["C","F","I"].filter(c=>typeof s.before[c].value==="number" && !(typeof s.before[c].formula==="string" && s.before[c].formula.startsWith("=")));
  }
  async function previewTarget(button,raw,apply){
    button.disabled=true;apply.disabled=true;state.preview=null;
    try{
      const pkt=checkPacket(raw);
      const snap=await targetSnapshot(pkt);
      const nums=alreadyNumeric(snap);
      const needOverride=nums.length>0;
      const allowed=document.getElementById("fluxo-substituir").checked;
      const dest=["C","F","I"];
      const fields=[pkt.credit,pkt.debit,pkt.balance];
      const lines=["Loja TKS | Data "+pkt.date+" | Linha "+snap.row,"",...dest.map((c,i)=>c+snap.row+": "+String(snap.before[c].text||"(vazio)")+"  →  "+money(fields[i])),"","Fonte: "+pkt.source.file];
      if(needOverride)lines.push("ATENÇÃO: já há valores digitados em "+nums.join(", ")+". É exigida autorização de substituição.");
      showPreview(lines.join("\n"));
      if(needOverride && !allowed){setStatus("Marque autorização de substituição e confira novamente.","error");return;}
      state.preview={pkt,snap,id:snapshotId(snap)};
      apply.disabled=false;
      setStatus("Prévia validada. Use APLICAR somente se loja, data e valores estiverem corretos.","ok");
    }catch(e){setStatus("Prévia interrompida: "+String(e.message||e),"error");}
    finally{button.disabled=false;}
  }
  async function applyTarget(button) {
    const p=state.preview;
    if(!p)return setStatus("Confira o destino antes de aplicar.","error");
    button.disabled=true;
    try{
      const current=await targetSnapshot(p.pkt);
      if(snapshotId(current)!==p.id)throw new Error("A planilha foi alterada após a prévia. Refaça a conferência.");
      const numeric=alreadyNumeric(current);
      if(numeric.length && !document.getElementById("fluxo-substituir").checked)throw new Error("Substituição não autorizada.");
      await Excel.run(async ctx=>{
        const sheet=ctx.workbook.worksheets.getItem("TKS");
        const row=p.snap.row;
        sheet.getRange("C"+row).values=[[p.pkt.credit]];
        sheet.getRange("F"+row).values=[[p.pkt.debit]];
        sheet.getRange("I"+row).values=[[p.pkt.balance]];
        await ctx.sync();
      });
      const after=await targetSnapshot(p.pkt);
      if(["C","F","I"].some((c,i)=>Math.abs(Number(after.before[c].value)-[p.pkt.credit,p.pkt.debit,p.pkt.balance][i])>0.0001))
        throw new Error("A gravação retornou, mas a conferência não confirmou todos os valores. Verifique o fluxo.");
      setStatus("Valores gravados e conferidos em TKS!C"+p.snap.row+", F"+p.snap.row+" e I"+p.snap.row+". Diário de origem inalterado.","ok");
      showPreview("CONCLUÍDO: "+p.pkt.date+"\nCréditos: "+money(p.pkt.credit)+"\nDébitos: "+money(p.pkt.debit)+"\nSaldo final: "+money(p.pkt.balance));
      state.preview=null;
    }catch(e){setStatus("Gravação interrompida: "+String(e.message||e),"error");}
    finally{button.disabled=true;}
  }
  Office.onReady(info=>{if(info.host===Office.HostType.Excel)addUI();});
})();
