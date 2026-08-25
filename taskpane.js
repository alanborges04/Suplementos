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
