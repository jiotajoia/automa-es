const fs = require('fs');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');

async function extrairTextoPDF(pdfPath) {
    const pdfData = fs.readFileSync(pdfPath);
    const textoPDF = await pdfParse(pdfData);
    return textoPDF.text;
}

async function compararPlanilhaComPDF(planilhaPath, pdfPath, colunaPlanilha, coluna) {
    
    const planilha = XLSX.readFile(planilhaPath);
    const aba = planilha.Sheets[planilha.SheetNames[0]];
    const dados = XLSX.utils.sheet_to_json(aba);

    const texto = await extrairTextoPDF(pdfPath);
    const linhasTabela = texto.split('\n');

    console.log(linhasTabela);

    dados.forEach((linhaPlanilha, idx) => {
        const valorPlanilha = linhaPlanilha[colunaPlanilha];
        let encontrado = false;
        
        // Verifica se o valor da planilha está em alguma linha da tabela
        linhasTabela.forEach(linhaPDF => {
            if (linhaPDF.includes(valorPlanilha)) {
                // Se encontrado, marca como "Encontrado"
                encontrado = true;
            }
        });

        if (encontrado) {
            linhaPlanilha[coluna] = 'S';
        } else {
            linhaPlanilha[coluna] = 'N';
        }
    });


    const novoAba = XLSX.utils.json_to_sheet(dados);
    planilha.Sheets[planilha.SheetNames[0]] = novoAba;

    XLSX.writeFile(planilha, planilhaPath);
    
}

async function main() {
    const pdfPath = 'teste.pdf';
    const planilhaPath = 'teste.xlsx';
    const colunaPlanilha = 'NOME'; 
    const coluna = 'RETORNO';

    await compararPlanilhaComPDF(planilhaPath, pdfPath, colunaPlanilha, coluna);
}

main();

