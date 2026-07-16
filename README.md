## JS + Excel

Script em Node.js para ler dados de uma planilha Excel, filtrar valores acima de um limite e gerar uma nova planilha com o resultado.

## Requisitos

- Node.js instalado
- Pacote `xlsx` (SheetJS)

## Instalação

```bash
npm init -y
npm install xlsx
```

## Estrutura do projeto

```
.
├── script.js
├── vendas_pratica.xlsx
├── package.json
└── README.md
```

## Como funciona

O script segue quatro etapas:

1. Leitura do arquivo `vendas_pratica.xlsx` com `XLSX.readFile`
2. Conversão da aba `Vendas` para um array de objetos com `XLSX.utils.sheet_to_json`
3. Filtro dos registros com `Total` maior ou igual a um valor mínimo
4. Geração de uma nova planilha (`vendas_caras.xlsx`) com o resultado, usando `json_to_sheet`, `book_new`, `book_append_sheet` e `writeFile`

## Rodar o script 

```bash
node script.js
```

O script imprime no console os registros filtrados e gera o arquivo `vendas_caras.xlsx` na raiz do projeto.

## Funções principais utilizadas

| Função | Finalidade |
|---|---|
| `XLSX.readFile` | Ler um arquivo `.xlsx` do disco |
| `XLSX.utils.sheet_to_json` | Converter uma aba em array de objetos |
| `XLSX.utils.json_to_sheet` | Converter um array de objetos em aba |
| `XLSX.utils.book_new` | Criar um workbook vazio |
| `XLSX.utils.book_append_sheet` | Anexar uma aba a um workbook |
| `XLSX.writeFile` | Salvar um workbook em disco |

## Notas

- O `Total` de cada venda é lido como valor já calculado, não como fórmula, pois o arquivo original foi salvo com os valores em cache.
- O import da biblioteca `xlsx` requer tratamento de interoperabilidade CommonJS/ESM:

```js
import pkg from 'xlsx';
const XLSX = pkg.default ?? pkg;
```
