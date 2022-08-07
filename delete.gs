function doGet() {
  var ss = getSheet('name');
  var msg = ss.getDataRange().getValues();
  var t = HtmlService.createTemplateFromFile("deleteMember");
  t.msg = msg;
  return t.evaluate().setTitle("");
}
 
function doPost(e) {
  var ss = getSheet('name');
  var msg = ss.getDataRange().getValues();
  var t = HtmlService.createTemplateFromFile("deleteMember");
  var sheet = getSheet('人件費');
  var values = sheet.getDataRange().getValues(); 
  values.shift();
  values.shift();
  
  var startDate = new Date(e.parameter.start);
  var startYear = startDate.getFullYear(); //入力された開始データから年を取得
  var startMonth = startDate.getMonth();　//入力された開始データから月を取得 
  var lastRow = sheet.getLastRow();
  t.msg = msg;

  for(var i=2; i<=lastRow+1; i++){
    if(e.parameter.name == sheet.getRange(i,7).getValue()){
      var value = sheet.getRange(i,4).getValue();
      if(startDate.getTime() <= value.getTime()){
        sheet.getRange(i,1,1,30).clearContent();
      }
    }
  }

  var range = sheet.getRange(2,1,sheet.getLastRow(), sheet.getLastColumn());
  var data = range.getValues();
  var result = [];

  for(let i = 0; i < data.length; i++){
    if(data[i].join('').length > 0){ 
      result.push(data[i]); //１文字以上あれば、結果の配列に追加する
    }
  }
  range.clearContent();
  range = sheet.getRange(2,1,result.length, result[0].length);
  range.setValues(result);
  return t.evaluate();
}

 

function myFunc(){
  var sheet = getSheet('人件費');
  var lastColumn = sheet.getLastRow();
  console.log(lastColumn);
}


function getSheet(name){
  var ssId = '1m9wjay5-PfFnUUDCrFV0Roxp5_BfezvRjfA2JYVJbMI';
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(name);
  return sheet;
}