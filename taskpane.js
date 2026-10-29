Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("run-btn").onclick = runAutomation;
    }
});

function setStatus(msg) { 
    document.getElementById("status").innerText = msg; 
}

function downloadErrorReport(errors) {
    if (errors.length === 0) return;
    let text = "RELATÓRIO DE ERROS - FLUXO REALIZADO\n========================================\n\n";
    errors.forEach(err => { text += `Aba: ${err.sheet}\nCélula: ${err.address}\nMotivo: ${err.reason}\n\n`; });
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url; link.download = "ERROS_ATUALIZACAO_FLUXO.txt";
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
}

async function runAutomation() {
    const btn = document.getElementById("run-btn"); 
    btn.disabled = true; 
    let errorList = [];
    
    try {
        await Excel.run(async (context) => {
            setStatus("Iniciando varredura das abas...");
            const sheets = context.workbook.worksheets; 
            sheets.load("items/name, items/visibility"); 
            await context.sync();
            
            let visibleSheets = sheets.items.filter(s => s.visibility === Excel.SheetVisibility.visible);
            
            for (let i = 0; i < visibleSheets.length; i++) {
                let sheet = visibleSheets[i]; 
                let sheetName = sheet.name;
                setStatus(`Processando ${sheetName} (${i+1}/${visibleSheets.length})...`);
                
                let usedRange = sheet.getUsedRange(); 
                usedRange.load("values, rowCount, columnCount, rowIndex"); 
                await context.sync();
                
                if (!usedRange.values) continue;
                
                let totalRowIndex = -1;
                let obsRowIndex = -1;
                
                // 1. Mapeia onde estão os limites (TOTAIS e OBSERVAÇÕES)
                for (let r = 0; r < usedRange.rowCount; r++) {
                    let rowText = usedRange.values[r].join(" ").toUpperCase();
                    
                    if (totalRowIndex === -1 && (rowText.includes("TOTAIS") || rowText.includes("TOTAL"))) {
                        totalRowIndex = r + usedRange.rowIndex;
                    }
                    if (obsRowIndex === -1 && (rowText.includes("OBSERVAÇÕES") || rowText.includes("OBSERVACOES"))) {
                        obsRowIndex = r + usedRange.rowIndex;
                    }
                }
                
                // ==========================================
                // BLOCO 1: ATUALIZAÇÃO DO FLUXO (PARTE DE CIMA)
                // ==========================================
                let sourceRow = -1;
                let targetRow = -1;
                
                if (totalRowIndex > 5) {
                    let relativeTotalIndex = totalRowIndex - usedRange.rowIndex;
                    for (let r = relativeTotalIndex - 1; r >= 0; r--) {
                        let val = String(usedRange.values[r][0] || "").trim();
                        if (val !== "" && val !== "-") {
                            sourceRow = r + usedRange.rowIndex;
                            targetRow = sourceRow + 1; 
                            break;
                        }
                    }
                }
                
                if (sourceRow !== -1 && targetRow < totalRowIndex) {
                    let colCount = 10; // Trava na coluna J
                    
                    let sourceRange = sheet.getRangeByIndexes(sourceRow, 0, 1, colCount); 
                    let targetRange = sheet.getRangeByIndexes(targetRow, 0, 1, colCount);
                    
                    targetRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
                    await context.sync();
                    
                    let oldDateCell = sheet.getCell(sourceRow, 0);
                    let newDateCell = sheet.getCell(targetRow, 0);
                    oldDateCell.load("text");
                    newDateCell.load("text");
                    await context.sync();
                    
                    let oldDateStr = oldDateCell.text[0][0]; 
                    let newDateStr = newDateCell.text[0][0]; 
                    let oldDateDot = oldDateStr.split('/').join('.');
                    let newDateDot = newDateStr.split('/').join('.');
                    
                    let colsToUpdate = [2, 5, 8]; // Colunas C, F e I
                    for (let c of colsToUpdate) {
                        let cell = sheet.getCell(targetRow, c);
                        cell.load("formulas");
                        await context.sync();
                        
                        let formula = cell.formulas[0][0];
                        if (formula && formula.startsWith("=")) {
                            formula = formula.split(oldDateDot).join(newDateDot);
                            cell.formulas = [[formula]];
                        }
                    }
                    await context.sync();
                }

                // ==========================================
                // BLOCO 2: ATUALIZAÇÃO DAS OBSERVAÇÕES (PARTE DE BAIXO)
                // ==========================================
                if (obsRowIndex !== -1) {
                    let obsSourceRow = -1;
                    let obsTargetRow = -1;
                    
                    let relativeObsIndex = obsRowIndex - usedRange.rowIndex;
                    // Procura de baixo pra cima a última linha preenchida nas observações
                    for (let r = usedRange.rowCount - 1; r > relativeObsIndex; r--) {
                        let val = String(usedRange.values[r][0] || "").trim();
                        if (val !== "" && val !== "-") {
                            obsSourceRow = r + usedRange.rowIndex;
                            obsTargetRow = obsSourceRow + 1; 
                            break;
                        }
                    }

                    if (obsSourceRow !== -1) {
                        let colCount = 10;
                        let obsSourceRange = sheet.getRangeByIndexes(obsSourceRow, 0, 1, colCount); 
                        let obsTargetRange = sheet.getRangeByIndexes(obsTargetRow, 0, 1, colCount);
                        
                        obsTargetRange.copyFrom(obsSourceRange, Excel.RangeCopyType.all);
                        await context.sync();

                        let oldObsDateCell = sheet.getCell(obsSourceRow, 0);
                        let newObsDateCell = sheet.getCell(obsTargetRow, 0);
                        oldObsDateCell.load("text");
                        newObsDateCell.load("text");
                        await context.sync();

                        let oldObsDateStr = oldObsDateCell.text[0][0]; 
                        let newObsDateStr = newObsDateCell.text[0][0]; 
                        let oldObsDateDot = oldObsDateStr.split('/').join('.');
                        let newObsDateDot = newObsDateStr.split('/').join('.');

                        // O PULO DO GATO DA MESCLAGEM: Atualiza apenas a coluna B (índice 1)
                        let obsCell = sheet.getCell(obsTargetRow, 1);
                        obsCell.load("formulas");
                        await context.sync();
                        
                        let formula = obsCell.formulas[0][0];
                        if (formula && formula.startsWith("=")) {
                            formula = formula.split(oldObsDateDot).join(newObsDateDot);
                            formula = formula.split(oldObsDateStr).join(newObsDateStr);
                            obsCell.formulas = [[formula]];
                        }
                        await context.sync();
                    }
                }
            }
            
            setStatus("Calculando planilha na nuvem..."); 
            context.workbook.application.calculate(Excel.CalculationType.full); 
            await context.sync();
            
            setStatus("Processo concluído com sucesso!");
        });
        
        if (errorList.length > 0) { 
            downloadErrorReport(errorList); 
            setStatus(`Concluído com ${errorList.length} aviso(s). Verifique o TXT baixado.`); 
        }
    } catch (error) {
        console.error(error); 
        setStatus("Erro! " + error.message);
    } finally {
        btn.disabled = false; 
        setTimeout(() => { if(btn.disabled === false) setStatus(""); }, 5000);
    }
}


/* Fluxo Turbo 2.0: módulo isolado de diagnóstico somente leitura */
(function fluxoDiagnosticModule() {
  "use strict";
  const targets = { C: "Créditos realizados", F: "Débitos realizados", I: "Saldo realizado" };
  const expectedTKS = { C: "F41", F: "F42", I: "B55" };
  const cols = ["C", "F", "I"];
  function dateOf(v) {
    const m = String(v || "").match(/\b(\d{1,2})\/(\d{1,2})\/(20\d{2})\b/);
    if (!m) return null;
    const d = new Date(Date.UTC(+m[3], +m[2]-1, +m[1]));
    if (d.getUTCDate() !== +m[1] || d.getUTCMonth()+1 !== +m[2]) return null;
    return String(m[1]).padStart(2,"0") + "/" + String(m[2]).padStart(2,"0") + "/" + m[3];
  }
  function external(formula) {
    const f = String(formula || "");
    const m = f.match(/^\s*=\s*'([^']*?)\[([^\]]+\.(?:xlsx|xlsm|xls))\]([^']+)'\s*!\s*(\$?[A-Z]{1,3}\$?\d+)\s*$/i);
    if (!m) return null;
    const dm = m[2].match(/\b(\d{2})\.(\d{2})\.(20\d{2})\b/);
    return { name:m[2], sourceSheet:m[3].trim(), sourceCell:m[4].replace(/\$/g,"").toUpperCase(), fileDate:dm ? dm[1]+"/"+dm[2]+"/"+dm[3]:null, sharepoint:/^https:\/\/[^']*sharepoint\.com/i.test(m[1]) };
  }
  function saveReport(rows,n) {
    const header = [
      "FLUXO TURBO - DIAGNÓSTICO DE VÍNCULOS - SOMENTE LEITURA",
      "Abas analisadas: " + n,
      "Referências detectadas: " + rows.filter(r=>r.link).length,
      "NENHUMA CÉLULA FOI ALTERADA. NENHUM ARQUIVO EXTERNO FOI ABERTO.",
      "O diagnóstico não testa permissões do Microsoft Graph.",
      ""
    ];
    for (const r of rows) {
      header.push(r.sheet+" | "+r.date+" | "+r.dest+" | "+r.field);
      header.push("  "+r.status+(r.error?" (Excel exibe erro de conexão)":""));
      if(r.link)header.push("  Fonte: "+r.link.name+" | "+r.link.sourceSheet+"!"+r.link.sourceCell);
      if(r.expected && r.link && r.expected!==r.link.sourceCell)header.push("  ATENÇÃO: célula de origem validada no diário TKS: "+r.expected);
      header.push("");
    }
    const u=URL.createObjectURL(new Blob([header.join("\n")],{type:"text/plain;charset=utf-8"}));
    const a=document.createElement("a");a.href=u;a.download="DIAGNOSTICO_FLUXO_TURBO.txt";document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(u),2500);
  }
  Office.onReady(info => {
    if(info.host!==Office.HostType.Excel)return;
    const old=document.getElementById("run-btn");
    if(!old || document.getElementById("diag-fluxo-btn"))return;
    const b=document.createElement("button");
    b.id="diag-fluxo-btn";b.textContent="🔎 DIAGNOSTICAR LINKS (SOMENTE LEITURA)";
    b.className="btn";b.style.background="#1D4ED8";b.style.fontSize="12px";
    b.title="Examina vínculos da planilha aberta, sem abrir diários nem modificar valores";
    const status=document.createElement("div");status.id="diag-fluxo-status";status.style.cssText="font-size:12px;white-space:pre-wrap;margin-top:10px;color:#334155";
    old.parentElement.appendChild(b);old.parentElement.appendChild(status);
    b.addEventListener("click",async()=>{
      b.disabled=true;status.textContent="Identificando abas, datas e referências externas...";
      try{
        const result=await Excel.run(async ctx=>{
          const sh=ctx.workbook.worksheets;sh.load("items/name,items/visibility");await ctx.sync();
          const active=sh.items.filter(s=>s.visibility===Excel.SheetVisibility.visible && s.name.trim().toUpperCase()!=="CONSOLIDADO");
          const loaded=[];
          for(const s of active){
            const a=s.getRange("A5:A65");a.load("text");
            const r={};
            for(const col of cols){r[col]=s.getRange(col+"5:"+col+"65");r[col].load("formulas,text");}
            loaded.push({name:s.name,a,r});
          }
          await ctx.sync();
          const rows=[];
          for(const item of loaded){
            for(let j=0;j<61;j++){
              const dt=dateOf(item.a.text[j]&&item.a.text[j][0]);
              if(!dt)continue;
              for(const col of cols){
                const f=item.r[col].formulas[j]&&item.r[col].formulas[j][0];
                const display=item.r[col].text[j]&&item.r[col].text[j][0];
                if(typeof f!=="string"||!f.startsWith("="))continue;
                const link=external(f), tks=item.name.trim().toUpperCase()==="TKS";
                const expected=tks?expectedTKS[col]:null;
                let state=link?"endereco_identificado":"formula_sem_endereco_reconhecido";
                if(link && link.fileDate && link.fileDate!==dt)state="ALERTA: data do nome não corresponde à coluna A";
                if(link && expected && expected!==link.sourceCell)state="ALERTA: célula de origem divergente";
                rows.push({sheet:item.name,date:dt,dest:col+(j+5),field:targets[col],link,status:state,
                  error:/#(?:CONECTAR!|CONNECT!|REF!)/i.test(String(display||"")),expected});
              }
            }
          }
          return {rows,n:active.length};
        });
        const links=result.rows.filter(x=>x.link).length;
        const alerts=result.rows.filter(x=>x.status.startsWith("ALERTA")).length;
        status.textContent=result.n+" aba(s) | "+links+" link(s) | "+alerts+" alerta(s).\nRelatório TXT gerado.\nNão foram feitas alterações.";
        saveReport(result.rows,result.n);
      }catch(e){console.error(e);status.textContent="Erro no diagnóstico: "+(e.message||String(e));}
      finally{b.disabled=false;}
    });
  });
})();
