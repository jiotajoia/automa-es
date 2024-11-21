const fs = require('fs');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');

// Função para extrair todo o texto do PDF
async function extrairTextoPDF(pdfPath) {
    const pdfData = fs.readFileSync(pdfPath);
    const textoPDF = await pdfParse(pdfData);
    return textoPDF.text;
}

// Função para comparação entre PDF e planilha
async function compararPlanilhaComPDF(planilhaPath, pdfPath, colunaPlanilha, colunaRetorno) {
    
    // Lê o conteúdo da planilha na aba 1 e extrai convertendo os dados para um json
    const planilha = XLSX.readFile(planilhaPath);
    const aba = planilha.Sheets[planilha.SheetNames[0]];
    const dados = XLSX.utils.sheet_to_json(aba);

    // Extrai o texto do PDF, caso tabela, faz o split
    const texto = await extrairTextoPDF(pdfPath);
    const linhasTabela = texto.split('\n');

    // Itera sobre cada linha da planilha
    dados.forEach((linhaPlanilha) => {
        // por coluna especificada
        const valorPlanilha = linhaPlanilha[colunaPlanilha];
      
        // Verifica se o valor da planilha está em alguma linha da tabela
        let encontrado = false;
        linhasTabela.forEach(linhaPDF => {
            if (linhaPDF.includes(valorPlanilha)) {
                // Se encontrado, marca como "true"
                encontrado = true;
            }
        });

        // Retorno especificado na coluna de retorno especificada
        if (encontrado) {
            linhaPlanilha[colunaRetorno] = 'S';
        } else {
            linhaPlanilha[colunaRetorno] = 'N';
        }
    });

    // Registra o novo valor na planilha
    const novoAba = XLSX.utils.json_to_sheet(dados);
    planilha.Sheets[planilha.SheetNames[0]] = novoAba;

    XLSX.writeFile(planilha, planilhaPath);
    
}

async function main() {
    const pdfPath = 'teste.pdf';
    const planilhaPath = 'teste.xlsx';
    const colunaPlanilha = 'NOME'; 
    const colunaRetorno = 'RETORNO';

    await compararPlanilhaComPDF(planilhaPath, pdfPath, colunaPlanilha, colunaRetorno);
}

main();