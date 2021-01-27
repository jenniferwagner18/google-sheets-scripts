/* 

Designed for instructors who want to share a Sheet with their students, and 
assign one protected column to each student. Protects 20 cells in a column below 
the email address entered into first row. Only the owner and the person specified 
by the email address can edit those cells when Sharing permissions are set to group 
(university or school) and Editor. (It's also a good idea to protect the first row 
so that only the owner can edit it.)

This function will be called when a cell in the first row is edited (except A1). 
No need to run script from the Script Editor.

Created by Dr. Jennifer Wagner - jenniferlynnwagner.com - January 2021

 */  

function onEdit(event){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0]; // 0 == first sheet
  var cellRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();

  // if a cell in first row (except A1) is not empty
  if (cellRange.getRow() == 1 && cellRange.getColumn() > 1 && event.value != null){ 

      // get the email address of the student from first row
      var email = sh.getRange(cellRange.getRow(), cellRange.getColumn()).getValue(); 

      // select the 20 rows underneath the email address (change green 20 below if you need a different number)
      var protectedRows = sh.getRange(cellRange.getRow() + 1, cellRange.getColumn(), 20); 

      // protect this range and set description for protected cells
      var protection = protectedRows.protect().setDescription(email + " & instructor can edit"); 

      // remove all other group editors (except instructor who is Owner of spreadsheet)
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) { protection.setDomainEdit(false); }

      // add student whose email address is in first row as editor for that column
      protection.addEditor(email);
  } 
  else
  {
    // if instructor deletes email address, remove protection for that column
    var protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      if (protections[i].getRange().getColumn() == cellRange.getColumn()) {
        protections[i].remove();
      }
    }  
  }
}
