function fix_links() { 
  
  var DocCol = 'C'; 
  //get spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  //get selected rows
  var selectedRange = String(ss.getActiveRange().getA1Notation());
  
  Logger.log('Range Selected: ' + selectedRange);

  // gets starting and ending row nums
  var splitRange = selectedRange.split(":");
  var vertStart = parseInt(String(splitRange[0].slice(1)));
  var vertEnd = parseInt(String(splitRange[1]).slice(1));
  
  Logger.log('Vertical Start: ' + vertStart);
  Logger.log('Vertical End: ' + vertEnd); 

  for (var i = vertStart; i < vertEnd + 1; i++) {
    
    // get current cell range
    var currentRange = sheet.getRange(DocCol + i)
    
    // get current row value
    var thisRowValue = currentRange.getValue();
    var thisRowFormula = currentRange.getFormula();
    
    // convert values to string
    var thisRowValueAsString = String(thisRowValue);
    var thisRowFormulaAsString = String(thisRowFormula); 
   
    //Logger.log("This row value: " + thisRowValue);   
    
    // PUT CHECKS IN LOOP, getOldDocName() should after checks
  
   // Logger.log(thisRowValueAsString);
   // Logger.log(thisRowFormulaAsString);
    
    
   
    if (thisRowFormulaAsString.substring(20, 31) == 'docs.google') {
      // link already fixed
      Logger.log('already fixed');
    } 
    else if (thisRowValueAsString.split('\\').length > 1) {
      // link is in old local format, get new document by name and insert
      Logger.log(thisRowValueAsString.split('\\')); 
      var local_name = thisRowValueAsString.split('\\')[1];
      Logger.log(local_name);
      insertNewDocLinkAsFormula(currentRange, local_name);  
    }
     else if (thisRowFormulaAsString == '') {
       // row formula is blank
      Logger.log('formula is blank, turn red'); 
      currentRange.setBackground('red'); 
    }
    else if (thisRowValueAsString = '') {
      // row value is blank 
      Logger.log('Value is blank, turn red');
      currentRange.setBackground('red')
    }
    else {
      // bad, turn red
      Logger.log('bad, turn red'); 
    }   
      
  }
 
  
  function getOldDocName (rowString) {
    
    var old_file_name = 'nullFile'; 
   // demo - returuns 'ull'
    //var old_file_name = old_file_name.substring(1, 4); 
    Logger.log(rowString.substring(12, 20)); 
    if(rowString.split('\\').length == 2) {
      Logger.log('YEP');
    }   
    
    return old_file_name; 
  }

  function insertNewDocLinkAsFormula(thisRow, fileName) {     
    // search for file by name
    var file_iterator = DriveApp.getFilesByName(fileName);

    while(file_iterator.hasNext()) {
     var file = file_iterator.next();
     var fileUrl = file.getUrl();
    }

    // add on the '=hyperlink()'
    var linkAsFormula = '=hyperlink(\"' + fileUrl + '\",\"Document\")';
    
    
    thisRow.setValue(linkAsFormula); 
  }
}

// drive link: 
// =hyperlink("https://docs.google.com/a/edmva.com/document/d/1ShbqF8NL2GEuUrfDY-sxeZXkts52RjFTi5NMD7lbxQ0/edit?usp=drivesdk","Document")
