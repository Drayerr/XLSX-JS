const xlsx = require('xlsx')

const myFile = 'FOR 609.xlsx'

const wb = xlsx.readFile(myFile)  //Lê o arquivo xlsx armazenado em 'myFile' e armazena as informações em 'wb'

const newWB = xlsx.utils.book_new('NEW WORKBOOK') //Cria uma nova pasta de trabalho
const newWS = wb.Sheets['Carta de Controle Pd']                  //Pega as informações da aba 'Plan1' do arquivo que foi lido e insere em 'newWS'

xlsx.utils.book_append_sheet(newWB, newWS, 'Recovery')  //Acrescenta a nova Worksheet(newWS) com o nome 'Recovery' na Workbook(newWB)

console.log(wb.SheetNames);
xlsx.writeFile(newWB, 'recovery 2.xlsx')      //Gera um novo arquivo com a pasta de trabalho 'newWB' com o nome de 'recovery'