function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Capacity Functions')
      .addItem('Update Projects', 'updateProjects')
      .addItem('Update People', 'updatePeople')
      .addItem('Add Person', 'addPerson')
      .addItem('Add Project', 'addProject')
      .addToUi();
}

function addPerson(){
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Let\'s get to know each other!',
      'Please enter your name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(text);
    sheet.deleteRows(100, sheet.getMaxRows() - 100);
    sheet.getRange('A1').setValue("Name");
    sheet.getRange('C1').setValue("Type");
    sheet.getRange('A2').setValue(text);
    sheet.getRange('C2').setValue("person");
    sheet.getRange('B4').setValue("Booked");
    sheet.getRange('B6').setValue("Availability");
    sheet.getRange('B7').setValue("Base");

    sheet.getRange('C4:BB4').setValue("=SUM(C7:INDEX(C7:C65536,MATCH(TRUE,INDEX(ISBLANK(C7:C65536),0,0),0)-1,0))");
    sheet.getRange('C4:BB4').setNumberFormat("0%");
    sheet.getRange('C7:BB100').setNumberFormat("0%");
    sheet.getRange('C5').setValue(this.firstMonday(0,2018));
    sheet.getRange('D5:BB5').setValue("=C5+7");
    sheet.getRange('C6:BB6').setValue(5);
                                      
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get your name.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}

function firstMonday (month, year){

 var d = new Date(year, month, 1, 0, 0, 0, 0)

 var day = 0

// check if first of the month is a Sunday, if so set date to the second

 if (d.getDay() == 0) {

 day = 2

 d = d.setDate(day)

 d = new Date(d)
 
 

 }
 return d 
}
  


function updateProjects(){
  var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Project Overview");
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();  
  var count = 0;
    overviewSheet.clearConditionalFormatRules()
  for (i = 0; i < sheets.length; i++) { 
    var value = sheets[i].getRange('C2').getValues()[0][0];
    if (value == "project"){
      var row = 4 + count;
      count +=3 ;

      overviewSheet.getRange('B'+ row).setValue("=HYPERLINK(\"#gid="+sheets[i].getSheetId()+"\",\""+sheets[i].getName()+"\")");
      overviewSheet.getRange('C'+ row).setValue("Plan");
      overviewSheet.getRange('D'+row+':N'+row).setFormula("=indirect(CONCATENATE($B"+row+",\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"6\"))");
      overviewSheet.getRange('C'+ (row+1)).setValue("Actual");
      overviewSheet.getRange('D'+(row+1)+':N'+(row+1)).setFormula("=indirect(CONCATENATE($B"+(row)+",\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"7\"))");
      
      var rules = overviewSheet.getConditionalFormatRules();

      var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=eq(D"+row+",D"+(row+1)+")=FALSE")
      .setBackground("#EDC9C4")
      .setRanges([overviewSheet.getRange('D'+(row+1)+':N'+(row+1))])
      .build();
      
      rules.push(rule1);
      
       var rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=eq(D"+row+",D"+(row+1)+") = TRUE")
      .setBackground("#55c170")
      .setRanges([overviewSheet.getRange('D'+(row+1)+':N'+(row+1))])
      .build();

      rules.push(rule2);

      
      overviewSheet.setConditionalFormatRules(rules);
      
    }
    
  }
}

function updatePeople(){
  var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("People Overview");
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();  
  var count = 0;
  overviewSheet.clearConditionalFormatRules()
  for (i = 0; i < sheets.length; i++) { 
    var value = sheets[i].getRange('C2').getValues()[0][0];
    if (value == "person"){
      var row = 6 + count;
      count ++;
      overviewSheet.getRange('B'+ row).setValue("=HYPERLINK(\"#gid="+sheets[i].getSheetId()+"\",\""+sheets[i].getName()+"\")");
      overviewSheet.getRange('D'+row+':N'+row).setFormula("=indirect(CONCATENATE($B"+row+",\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"4\"))");
      
      var rules = overviewSheet.getConditionalFormatRules();

      var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(.5, .8)
      .setBackground("#55c170")
      .setRanges([overviewSheet.getRange('D'+row+':N'+row)])
      .build();

      rules.push(rule1);

      
      var rule2 = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(.8)
      .setBackground("#EDC9C4")
      .setRanges([overviewSheet.getRange('D'+row+':N'+row)])
      .build();
      rules.push(rule2);
      
      overviewSheet.setConditionalFormatRules(rules);
      
      //overviewSheet.getRange('C'+ (row+1)).setValue("Actual");
      //overviewSheet.getRange('D'+(row+1)+':N'+(row+1)).setFormula("=indirect(CONCATENATE($B"+(row)+",\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"5\"))");
    }
    
  }
}

function addFormula(){
  var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview");
  
    overviewSheet.getRange('C16:G16').setFormula("=indirect(CONCATENATE($B4,\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"4\"))");

  
}
