# SpreadsheetDB
## What's SpreadsheetDB?
- This Script can be Spreadsheet as a DB for Google Apps Script.
- It can do select/insert/update/delete Row(s) on Spreadsheet  by query.
- It can full text query to Spreadsheet rows.

## How it work?
-  It use List Base Feed on Google Spreadsheet API.

## How use it?
- You can use it as a [Library](https://developers.google.com/apps-script/guide_libraries)
 - Library ID: `MlMEyCNWeCV97wefWrBRcEQV62A3znfOO`
- Do you need code assist? You can use mock library and mini hack for code assist.
 - Mock Library ID: `MORnDiFLK6dWUCyoTNsS1HQV62A3znfOO`
 - Use adobe library id and below code.


    var spreadsheetService = GSSDB.createService(spreadsheetId , consumerKey , consumerSecret);
    //var spreadsheetService = SpreadsheetService; it's hack for gas editor. you can do code assist after code.
    

## Example
### Initializer
  
    var spreadsheetService = GSSDB.createService(ss.getId());
    //if you need set ConsumerKey and ConsumerSecret
    var spreadsheetService = GSSDB.createService(ss.getId() , "consumerKey" , "consumerSecret");

  - args1 spreadsheet id 
  - args2 consumerKey
  - args3 consumerSecret

### insert row.

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
    spreadsheetService.insert(ss.getSheets()[0].getName(), entry);
    
  - arg0 : sheetName
  - arg1 : insert object, {column : value}

### select row(s).

    var rows = spreadsheetService.query(ss.getSheets()[0].getName(), ""); //if you need all rows,arg1 set empty string.
  
    //maybe rows length is 2, and rows[0]'is 1, rows[1]'id is 2
    //query result always return as a array. if result length is 1.
    var row = spreadsheetService.query(ss.getSheets()[0].getName() , "id=1");


  - arg0 : sheetName
  - arg1 : query. query reference is http://code.google.com/intl/ja/apis/spreadsheets/data/3.0/reference.html#ListParameters
  - arg2 : advanceObject
  - if you need sorted result or revesed result or full-text query,you set advanceObject.
  - advanceObject field is 
    
    { orderby : "column:columnName" ,//Specifies what column to use in ordering the entries in the feed.
     reverse:true/false , // Specifies whether to sort in descending or ascending order.
       q: full-text-query for rows
     };
  

### update row

    spreadsheetService.update(ss.getSheets()[0].getName() , row[0]);
   
   - arg0 : sheetName
   - arg1 : update object. it shoud be selected object by query method.
   
### delete row

    spreadsheetService.deleteEntry(ss.getSheets()[0].getName() , row[0]);

   - arg0 : sheetName
   - arg1 : delete object. it shoud be selected object by query method.

### if you update sheetName or add sheet. you should call refleshListKey method

     spreadsheetService.refleshListKey();
