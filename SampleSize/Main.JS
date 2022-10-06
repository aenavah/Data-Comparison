// @ts-nocheck
id_rawdata = "1lnJG-_u0gwhBzH0wnMUmMaHA5CoTmIjAMo3ar89f42U" /*Fake Data Book id*/
var number_batches = 8
id_new = "1WmFm1FkcUuCi-MCNZHRTaqVFpyRdsr-PlkSromqR-Es" /*Script Output Book id*/
teamcount = 8
teammember = 2 /*insert row team member is on*/


/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/
var sheet_old = "CDA Team - Batch " + number_batches
var sheet_new = "Sample Size " + number_batches
/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/

var raw_data = SpreadsheetApp.openById(id_rawdata)   /*raw_data book*/
var data_new = SpreadsheetApp.openById(id_new)      /*book to write to*/
var Data = raw_data.getSheetByName(sheet_old);       /*raw_data sheet*/
var Batch_New = data_new.getSheetByName(sheet_new); /*sheet to write to*/


if(data_new.getSheetByName(sheet_new) == null){
  spreadsheet.insertSheet(sheet_new); /* create sheet to write to */}

var SampleSize_list = []                           /*list of all team scores   --- NEED1? */
var Sum = 0                                        /* keeps track to get mean*/

function SampleSize() {
  counter = 2
  for (counter; counter <= teamcount+1; counter++){
  var cell = Data.getRange(counter, 5).getValue()
  SampleSize_list.push(cell)                      /*NEED1?*/        
  Sum = Sum + cell 
  }
  Avg = Sum/teamcount
}

var Rater = Data.getRange(teammember, 5).getValue() /*Rater's Sample Size*/
var Diff = Avg - Rater

function Format_Numbers(color, text){

}

