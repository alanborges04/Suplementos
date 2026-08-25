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
                let sourceRow = -1;
                let targetRow = -1;
                
                // 1. Acha em qual linha está a palavra TOTAIS
                for (let r = 0; r < usedRange.rowCount; r++) {
                    let val = String(usedRange.values[r][0] || "").trim().toUpperCase();
                    if (val === "TOTAL" || val === "TOTAIS" || val.startsWith("TOTAIS ")) { 
                        totalRowIndex = r + usedRange.rowIndex; 
                        break; 
                    }
                }
                
                // 2. Procura de baixo pra cima qual foi a última linha com data
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
                
                // 3. Copia a linha e ajusta a fórmula com a data correta
                if (sourceRow !== -1 && targetRow < totalRowIndex) {
                    let colCount = 10; // Mantém a trava na coluna J
                    
                    let sourceRange = sheet.getRangeByIndexes(sourceRow, 0, 1, colCount); 
                    let targetRange = sheet.getRangeByIndexes(targetRow, 0, 1, colCount);
                    
                    // Copia tudo primeiro
                    targetRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
                    await context.sync();
                    
                    // Lê a data velha e a data nova da coluna A (em formato de texto)
                    let oldDateCell = sheet.getCell(sourceRow, 0);
                    let newDateCell = sheet.getCell(targetRow, 0);
                    oldDateCell.load("text");
                    newDateCell.load("text");
                    await context.sync();
                    
                    // Exemplo: pega "20/08/2026" e "21/08/2026"
                    let oldDateStr = oldDateCell.text[0][0]; 
                    let newDateStr = newDateCell.text[0][0]; 
                    
                    // Converte as barras para pontos: "20.08.2026" e "21.08.2026"
                    let oldDateDot = oldDateStr.split('/').join('.');
                    let newDateDot = newDateStr.split('/').join('.');
                    
                    // Índices das colunas que precisam de alteração: C(2), F(5) e I(8)
                    let colsToUpdate = [2, 5, 8];
                    
                    for (let c of colsToUpdate) {
                        let cell = sheet.getCell(targetRow, c);
                        cell.load("formulas");
                        await context.sync();
                        
                        let formula = cell.formulas[0][0];
                        if (formula && formula.startsWith("=")) {
                            // Substitui a data antiga pela nova dentro do link da fórmula
                            formula = formula.split(oldDateDot).join(newDateDot);
                            cell.formulas = [[formula]];
                        }
                    }
                    await context.sync();
                    
                    // Checagem final de erros
                    targetRange.load("formulas, values"); 
                    await context.sync();
                    
                    for (let c = 0; c < colCount; c++) {
                        let val = String(targetRange.values[0][c] || "");
                        if (val.includes("#REF!") || val.includes("#NOME?")) {
                            errorList.push({ sheet: sheetName, address: `Linha ${targetRow + 1}, Coluna ${c + 1}`, reason: `Erro na fórmula: ${val}` });
                        }
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
