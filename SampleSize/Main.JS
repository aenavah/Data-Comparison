number_batches = 6
teamcount = 8
teammember = 2 /*insert row team member is on*/

id_old = "1lnJG-_u0gwhBzH0wnMUmMaHA5CoTmIjAMo3ar89f42U" /*Old data*/
id_rawdata = "1lnJG-_u0gwhBzH0wnMUmMaHA5CoTmIjAMo3ar89f42U" /*Fake Data Book id*/

/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/

var raw_data = SpreadsheetApp.openById(id_rawdata)   /*raw_data book*/
var data_old = SpreadsheetApp.openById(id_old)    /*old data book*/

var Data = raw_data.getSheetByName("QA - CDA TEAM - Batch " + number_batches);       /*raw_data sheet*/
var Sheet_Old = data_old.getSheetByName("Sample Size");                           /*sheet with old data*/

if(data_old.getSheetByName("Sample Size") == null){
  spreadsheet.insertSheet("Sample Size"); /* create sheet to write to */}


var Sum = 0                                        /* keeps track to get mean*/
function SampleSize() {
  counter = 2
  for (counter; counter <= teamcount+1; counter++){
    console.log(counter)
    var cell = Data.getRange(counter, 5).getValue() 
    Sum = Sum + cell 
  }
  Avg = Sum/teamcount
  var Rater = Data.getRange(teammember, 5).getValue() /*Rater's Sample Size*/
  var Diff = Rater - Avg

  Format(4, number_batches+1, "Batch " + number_batches, "#fff656","") 
  Format(5, number_batches+1, Rater, "", "string", "") 
  Format(7, number_batches+1, "Batch " + number_batches, "#ffd387", "")
  Format(8, number_batches+1, Avg, "", "")
  Format(11, (3*number_batches)-1, "BATCH " + number_batches, "#fff656", "")
  Sheet_Old.getRange(11, (3*number_batches)-1, 1,3).merge()
  Format(12, (3*number_batches)-1, "CDA", "#ffe4bc", "")  
  Format(12, 3*number_batches, "Vs.", "#ffe4bc", "")
  Format(12, (3*number_batches)+1, "Rater", "#ffe4bc", "")
  Format(13, (3*number_batches)-1, Avg, "", "")
  Format(13, 3*number_batches, Diff, "", 1)  
  Format(13, (3*number_batches)+1, Rater, "", "")
}

function Format(row, column, text, color, bin_textcolor){
  if (bin_textcolor==1){
    if (text > 0){
      text_color = "#008000"}
    else{
      text_color = "red"}
  }
  if(bin_textcolor == ""){
    text_color = ""}
  Sheet_Old.getRange(row, column).setValue(text).setBackgroundColor(color).setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center").setFontColor(text_color)
}



