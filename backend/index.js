const xlsx = require('xlsx')

//Insira o nome e extensão do arquivo que deseja copiar. Obs ele lê a partir da raiz.
const myFile = 'FOR 0219.xlsx' 

//Nome da aba que deseja copiar
const mySheet= 'Ficha de Campo e Calibração' 

//Lê o arquivo xlsx armazenado em 'myFile' e armazena as informações em 'wb'
const workBook = xlsx.readFile(myFile)

//Cria uma nova pasta de trabalho
const newWB = xlsx.utils.book_new('NEW WORKBOOK')
const newWS = workBook.Sheets[mySheet]

//Acrescenta a nova WorkSheet(newWS) com o nome 'Recovery' na WorkBook(newWB)
xlsx.utils.book_append_sheet(newWB, newWS, 'Recovery') 

//Gera um arquivo com a pasta de trabalho 'newWB' com o nome 'RecoveredFile.xlsx'
xlsx.writeFile(newWB, 'RecoveredFile ' + Date.now()+ '.xlsx') 

console.log('---XLSX-JS---');
console.log('A Aba (' + mySheet + ') foi copiada para um novo arquivo.')