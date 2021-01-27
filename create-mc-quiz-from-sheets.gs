/*
  Create multiple choice quiz in Google Forms from a Sheet. Put question in Column A, 
  correct answer in column B, and three wrong answers in Columns C through E. First row
  should have headers because they will be excluded.
  
  Dr. Jennifer Wagner - jenniferlynnwagner.com - January 2021
*/

// creates a Create Quiz menu in GUI so no need to open Script Editor to use
function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('Create Quiz');
  menu.addItem('Create MC Quiz', 'createQuiz').addToUi();
}

function createQuiz() {
  let file = SpreadsheetApp.getActive();
  let sheet = file.getSheetByName("Sheet1");
  let range = sheet.getDataRange();
  let values = range.getValues();

  // create the form as a quiz
  var form = FormApp.create('Multiple Choice Quiz');
  form.setIsQuiz(true);

  // remove first row of headers
  values.shift(); 

  // create multiple choice questions
  values.forEach(q => {
    let choices = [q[1], q[2], q[3], q[4]];
    let title = q[0];

    createShuffledChoices(form, title, choices)
  });
}

function createShuffledChoices(form, title, choices){

  let item = form.addMultipleChoiceItem();
  item.setTitle(title)
  .setPoints(1)

  let shuffledChoices = [];
  let correctAnswerChosen = false;

  for (let i = choices.length; i != 0; i--) {
    let rand = Math.floor(Math.random() * (i - 1));
    if (rand == 0 && correctAnswerChosen == false) {
      shuffledChoices.push(item.createChoice(choices.splice(rand, 1)[0], true));
      correctAnswerChosen = true;
    } else {
      shuffledChoices.push(item.createChoice(choices.splice(rand, 1)[0]));
    }  
  }
  
  item.setChoices(shuffledChoices);
}
