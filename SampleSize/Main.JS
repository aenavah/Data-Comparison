id_old = "1wvp0hunTlJBSO6MULC42b3x9XpvFmEWvMQ6vIvfBCC0"
var number_batches = 8
id_new = "1IOfFu4VH5PPXIXR7A6IthKJbFc8_s1sksiZxJ5SkhGE"
teamcount = 8
/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/
var sheet_old = "CDA Team - Batch " + number_batches
var sheet_new = "Sample Size " + number_batches
/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/

var data_old = SpreadsheetApp.openById(id_old)   
var data_new = SpreadsheetApp.openById(id_new)   

if(data_new.getSheetByName(sheet_new) == null){
  spreadsheet.insertSheet(sheet_new); /* create sheet "Output" */}

var Batch_New = data_new.getSheetByName(sheet_new); 
var Data = data_old.getSheetByName(sheet_old);

var SampleSize_List = []

function SampleSize() {
  counter = 2
  for (counter; counter <= teamcount+1; counter++)
  var cell = Data.getRange(counter, 5).getValue()
  SampleSize_List.push(cell)
}

console.log(SampleSize_list)
