function doGet() {
  return HtmlService.createTemplateFromFile("labor").evaluate();
}
 
function doPost(e) {
  var sheetName = getSheet('name');
  var sheet = getSheet('人件費');
  
  var values = sheet.getDataRange().getValues(); 
  values.shift();
  values.shift();

  var valuesName = sheetName.getDataRange().getValues();
  valuesName.shift();
  
  var enterDate = new Date(e.parameter.start);
  var startDate = new Date(enterDate.getFullYear(), enterDate.getMonth(),1);
  var startYear = startDate.getFullYear(); //入力された開始データから年を取得
  var startMonth = startDate.getMonth()+1;　//入力された開始データから月を取得
  var nameLastRow = sheetName.getLastRow();
  var flag = false;

  //期間を計算
  if(startMonth<=3){
    var term = 3 - startMonth;
  }else{
    var term = 12 - startMonth + 3;
  }

  for(var i=2; i<=nameLastRow; i++){
    if(sheetName.getRange(i,2).getValue() == e.parameter.name){
      flag = true;
      break;
   }
  }

  if(flag==false){
    valuesName.push([
      ,
      e.parameter.name,
      e.parameter.division,
      e.parameter.employment,
      startDate,
    ]);
  }
  valuesName.sort(((a, b) => {
  if (a[2] < b[2]) {
    return 1;
  } else {
    return -1;
  }
}));
  
  var nameColumn = sheetName.getDataRange().getLastColumn();
  var nameRow = valuesName.length;
  sheetName.getRange(2,1,nameRow,nameColumn).setValues(valuesName);
  sheetName.getRange(2,1,nameRow,nameColumn).setBorder(true, true, true, true, true, true);

  for(var i=0; i<term+1; i++){ 
    startDate.setMonth(startDate.getMonth()+i);
    var startDate1 = Utilities.formatDate(startDate,'JST','yyyy/MM/dd');
    startDate.setMonth(startDate.getMonth()+1);
    var provision = Utilities.formatDate(startDate,'JST','yyyy/MM/dd');
    values.push([
      Utilities.formatDate(enterDate,'JST','yyyy/MM/dd'),
      e.parameter.period,
      startDate.getFullYear() ,
      startDate1,
      provision,
      e.parameter.division,
      e.parameter.name,
      e.parameter.employment,
      e.parameter.budget,
      e.parameter.achievement,
      e.parameter.bonus,
    ]);
    startDate.setMonth(startDate.getMonth()-1);
    startDate.setMonth(startDate.getMonth()-i);
  }
  values.sort(function(a,b){return new Date(a[3]) - new Date(b[3]);});
  
  var column = sheet.getDataRange().getLastColumn();
  var row = values.length;
  sheet.getRange(3,1,row,column).setValues(values);
  sheet.getRange(3,1,row,column).setBorder(true, true, true, true, true, true);
  
  return HtmlService.createTemplateFromFile("labor").evaluate();
}
 

function getSheet(name){
  var ssId = '1m9wjay5-PfFnUUDCrFV0Roxp5_BfezvRjfA2JYVJbMI';
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(name);
  return sheet;
}