function copyNewTimestampWithUniqueID() {
    var ss = SpreadsheetApp.openById("1048SyHN-e4XujUkJOXj1j5rMQdZ4R4ZwIEYcLTK0OJk"); 
    var sourceSheet = ss.getSheetById(1447833263); 
    var destSheet = ss.getSheetById(684280098); 
    
    var lastRow = sourceSheet.getLastRow(); 
    if (lastRow < 2) return;

    var timestamp = sourceSheet.getRange(lastRow, 1).getValue(); 
    destSheet.appendRow([timestamp, 'Chair']); 
}