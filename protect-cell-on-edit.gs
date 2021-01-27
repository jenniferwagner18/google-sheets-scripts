/* 

   This function automatically protects a cell in a certain column after it is edited. Only the owner of the 
   spreadsheet and the original editor can then change or unprotect that cell. This prevents students from 
   accidentally (or not) deleting other students' names and taking their timeslots for presentation sign-ups.
   If a student needs to change their timeslot, the protection will automatically be removed when they delete
   their name so that another student can then edit that cell.
   
   This can be done manually by selecting the cell to protect, open Data and Protect sheets and ranges, click 
   + Add a sheet or range, Set Permissions, Restrict who can edit this range and remove everyone else. (You cannot
   remove yourself or the owner of the spreadsheet.) This script just automates that process so that students
   do not have to manually protect or unprotect the cell they put their name in.

   Make sure to protect the entire sheet and set an exception for the range where you want students to add 
   their names. This way the students cannot change the timeslots either or add anything else to the spreadsheet.

   Created by Dr. Jennifer Wagner - jenniferlynnwagner.com - January 2021
   
*/

function onEdit(event)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var cellRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();

  if (cellRange.getColumn() == 2 && event.value != null) // number refers to column in spreadsheet that students can edit
  { 
    // if edited cell is not empty (in column 2 only as specified above), then protect it
    setProtected(cellRange)
  } 
  else
  { 
    // if student deletes their own name, remove the protection
    var protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      if (protections[i].getRange().getRow() == cellRange.getRow() && protections[i].getRange().getColumn() == cellRange.getColumn()) {
        protections[i].remove(); 
      }
    }
  }
} 

function setProtected(cellRange)
{
  // sets the desired range to be protected
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var range = sh.getRange(cellRange.getRow(), cellRange.getColumn());
  var protection = range.protect().setDescription("Protected cell");
  
  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script throws an exception upon removing the group.
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
