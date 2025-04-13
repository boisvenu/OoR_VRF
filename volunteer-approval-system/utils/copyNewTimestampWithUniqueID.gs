function copyNewTimestampWithUniqueID() {
    var ss = SpreadsheetApp.openById("1048SyHN-e4XujUkJOXj1j5rMQdZ4R4ZwIEYcLTK0OJk"); // Google Sheet ID
    
    var sourceSheet = ss.getSheetById(1447833263); // Source Tab ID
    var destSheet = ss.getSheetById(684280098); // Destination Tab ID
    
    var lastRow = sourceSheet.getLastRow(); // Get the last row from the source tab
    
    if (lastRow < 2) return; // Prevent copying if there's no new data

    var timestamp = sourceSheet.getRange(lastRow, 1).getValue(); // Get the last row's timestamp from Column A
        
    destSheet.appendRow([timestamp, 'Chair']); // Append Unique ID + Timestamp to the destination sheet
}
