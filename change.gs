function doGet() {
  var sheetName = getSheet('name');
  var sheetMatter = getSheet('案件')
  var name = sheetName.getDataRange().getValues();
  var matter = sheetMatter.getDataRange().getValues();
  var t = HtmlService.createTemplateFromFile("salesChange");
  t.name = name;
  t.matter = matter;
  return t.evaluate().setTitle("");
}
 
function doPost(e) {
  var sheetName = getSheet('name');
  var sheetMatter = getSheet('案件');
  var name = sheetName.getDataRange().getValues();
  var matter = sheetMatter.getDataRange().getValues();
  var t = HtmlService.createTemplateFromFile("salesChange");
  var sheet = getSheet('売上');
  var values = sheet.getDataRange().getValues(); 
  values.shift();
  values.shift();

  var startDate = new Date(e.parameter.start);
  var startYear = startDate.getFullYear(); //入力された開始データから年を取得
  var startMonth = startDate.getMonth();　//入力された開始データから月を取得 
  var lastRow = sheet.getLastRow();
  t.name = name;
  t.matter = matter;

  
  for(var i=2; i<=lastRow; i++){
    if((e.parameter.name == sheet.getRange(i,9).getValue()) && (e.parameter.matter == sheet.getRange(i,10).getValue())){
      var value = sheet.getRange(i,7).getValue();
      if(startDate.getTime() <= value.getTime()){
        if(e.parameter.manager){
          sheet.getRange(i,2).clearContent();
          sheet.getRange(i,2).setValue(e.parameter.manager);
        }
        if(e.parameter.division){
          sheet.getRange(i,3).clearContent();
          sheet.getRange(i,3).setValue(e.parameter.division);
        }
        if(e.parameter.commodity){
          sheet.getRange(i,4).clearContent();
          sheet.getRange(i,4).setValue(e.parameter.commodity);
        }
        if(e.parameter.name){
          sheet.getRange(i,9).clearContent();
          sheet.getRange(i,9).setValue(e.parameter.name);
        }
        if(e.parameter.matter){
          sheet.getRange(i,10).clearContent();
          sheet.getRange(i,10).setValue(e.parameter.matter);
        }
        if(e.parameter.achievement){
          sheet.getRange(i,12).clearContent();
          sheet.getRange(i,12).setValue(e.parameter.achievement);
        }
        if(e.parameter.hourlyWage){
          sheet.getRange(i,13).clearContent();
          sheet.getRange(i,15).clearContent();
          sheet.getRange(i,13).setValue(e.parameter.hourlyWage);
        }
        if(e.parameter.productionCosts){
          sheet.getRange(i,14).clearContent();
          sheet.getRange(i,14).setValue(e.parameter.productionCosts);
        }
        if(e.parameter.budget){
          sheet.getRange(i,15).clearContent();
          sheet.getRange(i,13).clearContent();
          sheet.getRange(i,14).clearContent();
          sheet.getRange(i,15).setValue(e.parameter.budget);
        }
        if(e.parameter.hourlyWage || e.parameter.productionCosts || e.parameter.budget){
          sheet.getRange(i,11).clearContent();
          if(e.parameter.budget){
            sheet.getRange(i,11).setValue(e.parameter.budget);
          }else if(e.parameter.hourlyWage || e.parameter.productionCosts){
            var budget = sheet.getRange(i,13).getValues() * sheet.getRange(i,14).getValues();
            sheet.getRange(i,11).setValue(budget);
          }
        }

      }
    }
  }

  return t.evaluate();
}

 

function myFunc(){
  var sheet = getSheet('売上');
  var lastColumn = sheet.getLastRow();
  console.log(lastColumn);
}


function getSheet(name){
  var ssId = '1m9wjay5-PfFnUUDCrFV0Roxp5_BfezvRjfA2JYVJbMI';
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(name);
  return sheet;
}