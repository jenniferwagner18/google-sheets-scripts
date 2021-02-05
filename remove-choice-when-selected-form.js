/*
This function is used to update multiple choice options on a Form. It can be used by 
students for signing up for presentations. 

The following Sheet includes this code, but you will need to add a trigger:
https://docs.google.com/spreadsheets/d/1dRU4dnsazVqqNbWVr14kUb1QfuaXpLJUjVY4txELzdQ/copy?usp=sharing

To add a trigger in Sheets, go to Tools -> Script Editor -> Trigger (left sidebar): 
+ Add Trigger with On form submit under Select event type (leave the rest as they are)

On the spreadsheet, go to Tools and Create a Form to link the Sheet and Form. The Options 
listed on the Sheet should be entered into your Form as the multiple choice question 
options. Copy and paste the full URL of the form on the Config sheet. You may also need to 
update the formula in the C column on Options, depending on what the name of the Form responses 
sheet will be. (It's usually Form Responses and a number.)

On the Form, change the settings to "Collect email adressses" as well as "Limit to 
1 response" and "Edit after submit" if you want your students to be able to change 
their responses. Their original choice will then be placed back into the Form for 
another student to select.
*/

function openSlots() {
  
  // get spreadsheet data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName('Config');
  var configData = configSheet.getDataRange().getValues();
  var optionsSheet = ss.getSheetByName('Options');
  var optionsData = optionsSheet.getDataRange().getValues();
  
  // get Form Url from spreadsheet on Config sheet
  var formUrl = configData[0][1];
  Logger.log('formUrl is: ' + formUrl);
  
  // open Form
  var form = FormApp.openByUrl(formUrl);
  
  // create empty array to push available options
  var options = [];
  
    // loop through available slots and push into array
  var optionsDataLength = optionsData.length;
  for (var i=1; i<optionsDataLength; i++) {
    
    var choice = optionsData[i][0];
    var left = optionsData[i][2];
    
    // check if slot not blank and the option available (not been used)
    if ((choice != '') && (left > 0)) {
      options.push(choice); // add to array
    }
    
  }
  // end of loop through available slots and push into array
  
  // set item type to Multiple Choice for accessing Form Options
  var formList = FormApp.ItemType.MULTIPLE_CHOICE;
  
  // access Form Options list
  var formItems = form.getItems(formList);
  
  // access first list on Form and rewrite all options to ones that are available
  formItems[0].asMultipleChoiceItem().setChoiceValues(options);
}
