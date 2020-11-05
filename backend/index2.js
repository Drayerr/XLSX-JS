const xlsx = require('xlsx')
const openExplorer = require('open-file-explorer')

const path = 'C:\\users'

const file = openExplorer(path, (err) => {
  if(err){

  } else {
    console.log('file selected');
  }
})

console.log(file);