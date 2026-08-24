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
                
                // 2. Se achou os Totais, procura de baixo pra cima qual foi a última linha preenchida com dados (Data)
                if (totalRowIndex > 5) {
                    let relativeTotalIndex = totalRowIndex - usedRange.rowIndex;
                    for (let r = relativeTotalIndex - 1; r >= 0; r--) {
                        let val = String(usedRange.values[r][0] || "").trim();
                        // Se a coluna da data não estiver vazia, achamos a fonte!
                        if (val !== "" && val !== "-") {
                            sourceRow = r + usedRange.rowIndex;
                            targetRow = sourceRow + 1; // O alvo é exatamente a linha em branco debaixo dela
                            break;
                        }
                    }
                }
                
                // 3. Se achou a última linha preenchida E a linha debaixo ainda está antes dos Totais
                if (sourceRow !== -1 && targetRow < totalRowIndex) {
                    let colCount = usedRange.columnCount;
                    
                    let sourceRange = sheet.getRangeByIndexes(sourceRow, 0, 1, colCount); 
                    let targetRange = sheet.getRangeByIndexes(targetRow, 0, 1, colCount);
                    
                    // O PULO DO GATO: Copia direto pra linha debaixo, SEM INSERIR UMA NOVA
                    targetRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
                    await context.sync();
                    
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
