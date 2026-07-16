// Ler o arquivo
import pkg from 'xlsx';
const XLSX = pkg.default ?? pkg;
const workbook = XLSX.readFile("vendas_pratica.xlsx");

// Transformar o arquivo excel em um array de objetos
const sheet = workbook.Sheets['Vendas']
const vendas = XLSX.utils.sheet_to_json(sheet);
// console.log(vendas);

// Fazer o filtro e retornar os dados 
const maioresValores = (vendas , limite) => {
const  filter = vendas.filter(venda => venda.Total >= limite)
return filter
}

//Transformar em planilha
const resultado = maioresValores(vendas, 1000)
const novaSheet = XLSX.utils.json_to_sheet(resultado)
const novoWorkbook = XLSX.utils.book_new()
XLSX.utils.book_append_sheet(novoWorkbook, novaSheet, "vendas_caras")
XLSX.writeFile(novoWorkbook, "vendas_caras.xlsx")


// Workbook -> objeto que representa o arquivo Excel inteiro,
// e guarda dentro dele todas as abas (sheets) que existem nesse arquivo