var firstMonth = 6;

function onOpen() {
  
  var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Project Overview");
  if (overviewSheet == null){
    createProjectOverview();
    
  }
  
  var personSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("People Overview");
  if (personSheet == null){
    createPersonOverview();
    
  }
  
  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('Capacity Functions')
  .addItem('Update Projects', 'updateProjects')
  .addItem('Update People on this Project', 'updateProjectAssignments')
  .addItem('Update People', 'updatePeople')
  .addItem('Add Person', 'addPerson')
  .addItem('Add Project', 'addProject')
  .addItem('Hide Projects', 'hideProjects')
  .addItem('Hide People', 'hidePeople')
      .addToUi();
}

function hidePeople(){

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();  

                   
  for (i = 0; i < sheets.length; i++) { 
      var value = sheets[i].getRange('C2').getValues()[0][0];
  
    if (value == "person"){
      sheets[i].hideSheet();
    }  
  }
    
}


function hideProjects(){
  
   var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();  

                   
  for (i = 0; i < sheets.length; i++) { 
      var value = sheets[i].getRange('C2').getValues()[0][0];
  
    if (value == "project"){
      sheets[i].hideSheet();
    }  
  }
}


function createProjectOverview(){
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Project Overview");
  sheet.getRange('A1').setValue("Name");

  sheet.getRange('C1').setValue("Type");

  sheet.getRange('A2').setValue("Project Overview");

  sheet.getRange('C2').setValue("Project Overview");
    
  sheet.getRange('D3').setValue(this.firstMonday(this.firstMonth,2018));
  sheet.getRange('E3:BB3').setValue("=D3+7");
}

function createPersonOverview(){
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("People Overview");
  sheet.getRange('A1').setValue("Name");
  sheet.getRange('B1').setValue("Total Points");
  sheet.getRange('C1').setValue("Type");

  sheet.getRange('A2').setValue("People Overview");

  sheet.getRange('C2').setValue("People Overview");
    
  sheet.getRange('C5').setValue(this.firstMonday(this.firstMonth,2018));
  sheet.getRange('D5:BB5').setValue("=C5+7");
}


function updateProjectAssignments(){
 
  var sheet = SpreadsheetApp.getActiveSheet();
  var value = sheet.getRange('C2').getValues()[0][0];
    if (value == "project"){
      var projectName = sheet.getRange('A2').getValues()[0][0];
      this.updatePeopleOnProject(projectName);
    }
}

function addProject(){
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please enter your project name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var projectName = result.getResponseText();
  if (button == ui.Button.OK) {
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(projectName);
    sheet.deleteRows(100, sheet.getMaxRows() - 100);
    sheet.getRange('A1').setValue("Name");
    sheet.getRange('B1').setValue("Total Points");
    sheet.getRange('C1').setValue("Type");
    sheet.getRange('D1').setValue("Status");

    sheet.getRange('A2').setValue(projectName);
    sheet.getRange('B2').setValue(0);
    sheet.getRange('C2').setValue("project");
    sheet.getRange('D2').setValue("active");

    
    
    //set static values
    sheet.getRange('C4:C10').setValues(
      [
        ["Plan"],["Actual"],["PlanRemainig"],["Actual Remaining"],["Added Points"],["Delta"],["Total Drift"]
      ]
    );
  
        //set plan remaining row
    sheet.getRange('D6:BB6').setValue("=$B$2-sum($D$4:D4)");
    
      
    //set plan row
    sheet.getRange('D4:BB4').setValue("=sum(D13:D24)");
    
    
        //set actual remaingin
    sheet.getRange('D7:BB7').setValue("=if(D5<>\"\",$B$2+sum($D$8:D8)-sum($D$5:D5),D6)");
    
    
        //set delta row
    sheet.getRange('D9:BB9').setValue("=if(D5<>\"\",D5-D4,\"\")");
    
    //set dates
    sheet.getRange('D12').setValue(this.firstMonday(this.firstMonth,2018));
    sheet.getRange('E12:BB12').setValue("=D12+7");
    
    updatePeopleOnProject(projectName);
    
                                      
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get your name.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
  
  updatePeople();
  updateProjects();
}

function updatePeopleOnProject(projectNameVal){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(projectNameVal);
  var value = sheet.getRange('C2').getValues()[0][0];
  
  if (value == "project"){
    var projectName = sheet.getRange('A2').getValues()[0][0];
    if (projectName != projectNameVal){
      Logger.log("bad project sheet");
      return;
    }
  }

  var data = getProjectsFromPeople();

  Logger.log(data);
  
  var people = data[projectName];
  
  Logger.log(people);
  
  
  var startRow = 13;
  
  for (i = 0; i < people.length; i++) { 
    var row = startRow + i;
    var personFormula = "=iferror(vlookup($A$2, INDIRECT (CONCATENATE($C"+row+",\"!B\"&match($A$2,indirect(concatenate($C"+row+"&\"!$B:$B\")),0)&\":L\"& match($A$2,indirect(concatenate($C"+row+"&\"!$B:$B\")),0))),column()-2,0) * Indirect($C"+row+"&\"!\"&SUBSTITUTE(ADDRESS(1,column()-1,4),1,\"\")&\"6\") ,\"\")";
    
    sheet.getRange('C'+row).setValue(people[i]);
    sheet.getRange('D'+row+":BB"+row).setValue(personFormula);
  }
  
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
    sheet.getRange('C5').setValue(this.firstMonday(this.firstMonth,2018));
    sheet.getRange('D5:BB5').setValue("=C5+7");
    sheet.getRange('C6:BB6').setValue(5);
                                      
    updatePeople();
    
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
    overviewSheet.clearConditionalFormatRules()
    
  var projects = getProjectsFromPeople()
  for (i=0;i< projects.length; i++){
   
    getDataFromProjectSheet(projects[i],i);
    
  }
}

function getProjectsFromPeople(){
  var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("People Overview");
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();  
  var count = 0;
  overviewSheet.clearConditionalFormatRules()
  projectList = [];
                   
  for (i = 0; i < sheets.length; i++) { 
    var value = sheets[i].getRange('C2').getValues()[0][0];
    if (value == "person"){
      
      var projects = sheets[i].getRange('B8:B100').getValues()
      var person = sheets[i].getRange('A2').getValue();
      
      for (j=0; j<projects.length;j++){ 
        var projectName = projects[j][0];

        if (projectName == ""){
          continue;
        }
        
        //if not already there, add this project to the list
        if(this.projectList.indexOf(projectName) === -1) {
           this.projectList.push(projectName);
          
          //make it an array so we can add people to it
           this.projectList[projectName]=[];
        }

        
        if(this.projectList[projectName].indexOf(person) === -1) {
           this.projectList[projectName].push(person);
        }
      }
      
    }
    
  }
  Logger.log(projectList);
  return projectList;
}

function getDataFromProjectSheet(projectName,count){
  
  Logger.log("loading:"+projectName);
  var overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Project Overview");
  
  if (count ==0){
    overviewSheet.clearConditionalFormatRules()
    overviewSheet.getRange('B4:BB100').clearContent();
    overviewSheet.getRange('B4:BB100').clearFormat();
  }
    
  var projectSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(projectName);
  

  
  var row = 4 + (count * 3);

  
  if (projectSheet == null){
    Logger.log("not found:"+projectName);
    overviewSheet.getRange('B'+ row).setValue(projectName);
    overviewSheet.getRange('B'+ row).setBackground("red");
    return;
  }
  

  
    var value = projectSheet.getRange('C2').getValues()[0][0];
    
    if (value == "project"){
      

      overviewSheet.getRange('A'+ row).setFormula("=indirect(CONCATENATE($B"+row+",\"!D2\"))");
      overviewSheet.getRange('A'+ (row+1)).setFormula("=indirect(CONCATENATE($B"+row+",\"!D2\"))");
      overviewSheet.getRange('B'+ row).setValue("=HYPERLINK(\"#gid="+projectSheet.getSheetId()+"\",\""+projectSheet.getName()+"\")");
      overviewSheet.getRange('C'+ row).setValue("Plan");
      overviewSheet.getRange('D'+row+':BB'+row).setFormula("=indirect(CONCATENATE($B"+row+",\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"6\"))");
      overviewSheet.getRange('C'+ (row+1)).setValue("Actual");
      overviewSheet.getRange('D'+(row+1)+':BB'+(row+1)).setFormula("=indirect(CONCATENATE($B"+(row)+",\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"7\"))");
      
      var rules = overviewSheet.getConditionalFormatRules();

      var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=eq(D"+row+",D"+(row+1)+")=FALSE")
      .setBackground("#EDC9C4")
      .setRanges([overviewSheet.getRange('D'+(row+1)+':BB'+(row+1))])
      .build();
      
      rules.push(rule1);
      
       var rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=eq(D"+row+",D"+(row+1)+") = TRUE")
      .setBackground("#55c170")
      .setRanges([overviewSheet.getRange('D'+(row+1)+':BB'+(row+1))])
      .build();

      rules.push(rule2);

      
      overviewSheet.setConditionalFormatRules(rules);
      
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
      overviewSheet.getRange('D'+row+':BB'+row).setFormula("=indirect(CONCATENATE($B"+row+",\"!\",SUBSTITUTE(ADDRESS(1,COLUMN(),4), \"1\", \"\"),\"4\"))");
      
      var rules = overviewSheet.getConditionalFormatRules();

      var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(.5, .8)
      .setBackground("#55c170")
      .setRanges([overviewSheet.getRange('D'+row+':BB'+row)])
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
