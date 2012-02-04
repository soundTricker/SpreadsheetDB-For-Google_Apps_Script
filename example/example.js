/*
load spreadsheetService.js
and Ready Spreadsheet.
sheet as table on Spreadsheet.
Row1 is Header.
*/


function testInsert() {
  var ss = SpreadsheetApp.getActive();
  var spreadsheetService = new SpreadsheetService(ss.getId());

  //When you use SpreadsheetService,You should call init method.
  spreadsheetService.init(); //When your first call,Need oauth authentification.
  
  var entry = {
    "id":1,
    "name":"soundTricker"
  };
  
  var entry2 = {
    "id":2,
    "name": "soundTricker2"
  };
  
  //insert
  //If your need insert row, you user insert method
  //insert method args
  //arg0 : sheetName
  //arg1 : insert object, {column : value}
  spreadsheetService.insert(ss.getSheets()[0].getName(), entry);

  //insert
  spreadsheetService.insert(ss.getSheets()[0].getName(), entry2);

  
  //select
  //If you need search rows,you use query method.
  //query method args
  //arg0 : sheetName
  //arg1 : query. query reference is http://code.google.com/intl/ja/apis/spreadsheets/data/3.0/reference.html#ListParameters
  //arg2 : advanceObject
  //if you need sorted result or revesed result or full-text query,you set advanceObject.
  //advanceObject field is 
  // { orderby : "column:columnName" ,//Specifies what column to use in ordering the entries in the feed.
  //   reverse:true/false , // Specifies whether to sort in descending or ascending order.
  //   q: full-text-query for rows
  // };
  
  //if you need all rows,arg1 set empty string.
  var rows = spreadsheetService.query(ss.getSheets()[0].getName(), "");
  
  //maybe rows length is 2, and rows[0]'is 1, rows[1]'id is 2
  //query result always return as a array. if result length is 1.
  var row = spreadsheetService.query(ss.getSheets()[0].getName() , "id=1");
  
  row[0].name = "soundTricer3";
  
  //update
  //If you need update row,you use update method.
  //update method args
  //arg0 : sheetName
  //arg1 : update object. it shoud be selected object.
  spreadsheetService.update(ss.getSheets()[0].getName() , row[0]);
  
  
  //delete
  //If you need delete row,you use delete method
  //delete method args
  //arg0 ; sheetName
  //arg1 : delete object. it shoud be selected object.
  spreadsheetService.delete(ss.getSheets()[0].getName() , row[0]);
  
  
  //if you update sheetName or add sheet.
  //you should call refleshListKey method
  spreadsheetService.refleshListKey();
}


