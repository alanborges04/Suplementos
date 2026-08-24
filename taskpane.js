Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("run-btn").onclick = testarBotao;
    }
});

async function testarBotao() {
    document.getElementById("status").innerText = "Testando conexão...";
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            let range = sheet.getRange("A1");
            
            // Escreve um texto de teste na célula A1
            range.values = [["O botão do Fantomas tá vivo!"]];
            
            await context.sync();
            document.getElementById("status").innerText = "Sucesso! Célula A1 alterada.";
        });
    } catch (error) {
        document.getElementById("status").innerText = "Erro: " + error.message;
    }
}
