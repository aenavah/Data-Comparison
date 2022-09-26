/*------------------------------------------user inputs-----------------------------------------*/
/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/
id = "1FepeHyFf6Mn_lGANLN3uZdOrV7N-hyO5eP9EtHOgpMU"
var number_batches = 8
var sheet_old = "Data Analysis"
var sheet_new = "Data Analysis " + number_batches
/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/
/*----------------------------------------------------------------------------------------------*/


var spreadsheet = SpreadsheetApp.openById(id)   
if(spreadsheet.getSheetByName(sheet_new) == null){
  spreadsheet.insertSheet(sheet_new); /* create sheet "Output" */}
var Data_Analysis = spreadsheet.getSheetByName(sheet_new); 
var Data_Analysis_Last = spreadsheet.getSheetByName(sheet_old);

function get_Prog_Avg_Comp() {
  row_counter = 1
  for (row_counter; row_counter <= 12; row_counter ++){

    var color = "white"
    if (row_counter == (3||8||13)){
      continue
    }
    if(row_counter == 4){color = "#fff656"}
    if(row_counter == 5){color = "#c7ffc7"}
    if(row_counter == 6){color = "#c7e5ff"}
    if(row_counter == 7){color = "#e7c2ff"}
    if(row_counter == 9){color = "#ffd38f"}
    column_counter = 1
    for (column_counter; column_counter <= number_batches + 1; column_counter++){
      if (row_counter == 1){
        if (column_counter == (1||2))
          {color = "#c7ffc7"}
        if (column_counter == (3||4))
          {color = "#c7e5ff"}
        if (column_counter > 4){
          color = "white"
        }
      }
      var base_data = Data_Analysis_Last.getRange(row_counter, column_counter).getValue()
      if(base_data.toString().includes(".",0)){ /*if it's a number*/
        Data_Analysis.getRange(row_counter, column_counter).setValue((base_data*100) + "%").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment ("center").setBackground(color)
      }
      else { /*not a number*/

        Data_Analysis.getRange(row_counter, column_counter).setValue(base_data).setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center").setBackground(color)

      }
    }
  }

/*------------------------------------------putting data in lists------------------------------------------*/

  batch_counter = 2
  writing_counter = 1
  number_batches++ 

  var tm_greenscores = []
  var tm_bluescores = []
  var tm_violetscores = []

  var cda_greenscores = []
  var cda_bluescores = []
  var cda_violetscores = []

  for (let batch_counter = 2; batch_counter <= number_batches; batch_counter++){
    let tm_green = grab_data(5,batch_counter)   
    tm_greenscores.push(tm_green)
    let tm_blue = grab_data(6,batch_counter)   
    tm_bluescores.push(tm_blue)
    let tm_violet = grab_data(7,batch_counter)   
    tm_violetscores.push(tm_violet)
    let cda_green = grab_data(10,batch_counter)   
    cda_greenscores.push(cda_green)
    let cda_blue = grab_data(11,batch_counter)   
    cda_bluescores.push(cda_blue)
    let cda_violet = grab_data(12,batch_counter)   
    cda_violetscores.push(cda_violet)
  }


/*------------------------------------------TM Progression------------------------------------------*/
/*making row titles*/
  Data_Analysis.getRange(14, 1).setValue("TM Progression").setBackground("#fff656").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
  Data_Analysis.getRange(15, 1).setValue("Green Score").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
  Data_Analysis.getRange(16, 1).setValue("Blue Score").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
  Data_Analysis.getRange(17, 1).setValue("Violet Score").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")

/*copying scores*/
var skipped = 0
var batch_number = 1

  for (batch_number; batch_number < number_batches; batch_number++){
    if (tm_violetscores[batch_number-1] == "No Data"){   /*checks if team member has empty score for batch from arbitrary list*/
      continue
    } 

    /*creates headers for tm progression and highlighting*/
    if (1 < batch_number){
      batch_before = Data_Analysis.getRange(14, (2*writing_counter)-2).getValue().replace("Batch ", "")
      Data_Analysis.getRange(14, 2 * writing_counter - 1).setValue(batch_before + " to " + batch_number).setBackground("#fff656").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")}
    Data_Analysis.getRange(14, 2*writing_counter).setValue("Batch " + batch_number).setBackground("#fff656").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")

    Data_Analysis.getRange(15, 2*writing_counter).setValue(tm_greenscores[batch_number-1]).setBackground("#c7ffc7")
    fix_format_number(15, 2*writing_counter, tm_greenscores[batch_number-1])

    Data_Analysis.getRange(16, 2*writing_counter).setValue(tm_bluescores[batch_number-1]).setBackground("#c7e5ff")
    fix_format_number(16, 2*writing_counter, tm_bluescores[batch_number-1])

    Data_Analysis.getRange(17, 2*writing_counter).setValue(tm_violetscores[batch_number-1]).setBackground("#e7c2ff")
    fix_format_number(17, 2*writing_counter, tm_violetscores[batch_number-1])

    writing_counter++
  }

  /*calculating differences*/
  batch_number = 1
  writing_counter = 1

  for (writing_counter; writing_counter <= number_batches; writing_counter++){
    if (Data_Analysis.getRange(15, 2*writing_counter+2).getValue() == ''){
      continue
    }
    improvement_color(Data_Analysis, 15, writing_counter)
    fix_format_number(15, 2 * writing_counter + 1, Data_Analysis.getRange(15, writing_counter).getValue())

    improvement_color(Data_Analysis, 16, writing_counter)
    fix_format_number(16, 2 * writing_counter + 1, Data_Analysis.getRange(16, writing_counter).getValue())

    improvement_color(Data_Analysis, 17, writing_counter)
    fix_format_number(17, 2 * writing_counter + 1, Data_Analysis.getRange(17, writing_counter).getValue())


  }

/*------------------------------------------TM/Team Average------------------------------------------*/
/*making row titles*/
  Data_Analysis.getRange(20, 1).setValue("TM / Team Average").setBackground("#ffdd79").setFontSize(12).setFontFamily("Arial")
  Data_Analysis.getRange(22, 1).setValue("Green Score").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
  Data_Analysis.getRange(23, 1).setValue("Blue Score").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
  Data_Analysis.getRange(24, 1).setValue("Violet Score").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
 
  batch_counter = 1
  writing_counter = 1

  for (batch_counter; batch_counter < number_batches; batch_counter++){
    if (tm_bluescores[batch_counter - 1] == "No Data"){
      continue;
    }
    var left = 3*writing_counter-1
    var middle = 3*writing_counter 
    var right = 3*writing_counter+1  

    /*formatting rows 20, 21, 22, 23, 24*/
    Data_Analysis.getRange(20, middle).setValue("Batch " + batch_counter).setBackground("#fff656").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(20, left, 1, 3).merge()
    Data_Analysis.getRange(21, left).setValue("CDA").setBackground("#ffe4bc").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(21, middle).setValue("Vs.").setBackground("#ffe4bc").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(21, right).setValue("RATER").setBackground("#ffe4bc").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(22, left).setValue(((cda_greenscores[batch_counter - 1])*100) + "%").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(23, left).setValue(cda_bluescores[batch_counter - 1]).setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(24, left).setValue(cda_violetscores[batch_counter - 1]).setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")    
    Data_Analysis.getRange(22, right).setValue(tm_greenscores[batch_counter - 1]).setBackground("#c7ffc7").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(23, right).setValue(tm_bluescores[batch_counter - 1]).setBackground("#c7e5ff").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
    Data_Analysis.getRange(24, right).setValue(tm_violetscores[batch_counter - 1]).setBackground("#e7c2ff").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")

  team_avg_color(tm_greenscores, cda_greenscores, 22, batch_counter, middle)
  team_avg_color(tm_bluescores, cda_bluescores, 23, batch_counter, middle)
  team_avg_color(tm_violetscores, cda_violetscores, 24, batch_counter, middle)


    writing_counter++
  }}
  /*------------------------------------------Formatting Numbers------------------------------------------*/

function fix_format_number(row_counter, cell_counter, current_value){
    current_value = Data_Analysis.getRange(row_counter, cell_counter).getValue()
    Data_Analysis.getRange(row_counter, cell_counter).setValue((current_value*100)+"%").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
}
function grab_data(row,batch_counter){ /* this function gets the cell value given the row, batch*/ 
  var cell = Data_Analysis_Last.getRange(row,batch_counter).getValue()
    if (cell == ""){
      cell = "No Data"}
    return cell
}
function improvement_color(Data_Analysis, row, writing_counter){
  var batch_before = Data_Analysis.getRange(row, 2*writing_counter).getValue()
  var batch_after = Data_Analysis.getRange(row, 2*writing_counter+2).getValue()

  var difference = batch_after - batch_before
  if (difference > 0){
     Data_Analysis.getRange(row, 2*writing_counter+1).setValue(batch_after - batch_before).setFontColor("green")
  }
  if (difference < 0){
     Data_Analysis.getRange(row, 2*writing_counter+1).setValue(batch_after - batch_before).setFontColor("red")
  }
}
function team_avg_color(tm_color, cda_color, row, batch_counter, middle){
   if((tm_color[batch_counter-1] - cda_color[batch_counter-1]) > 0){
     var new_number = ((tm_color[batch_counter-1]-cda_color[batch_counter-1])*100)+"%"
     Data_Analysis.getRange(row, middle).setValue(new_number).setFontColor("green").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")
   }
   if((tm_color[batch_counter-1] - cda_color[batch_counter-1]) < 0){
     var new_number = ((tm_color[batch_counter-1]-cda_color[batch_counter-1])*100)+"%"
     Data_Analysis.getRange(row, middle).setValue(new_number).setFontColor("red").setFontSize(12).setFontFamily("Arial").setHorizontalAlignment("center")}
 }
