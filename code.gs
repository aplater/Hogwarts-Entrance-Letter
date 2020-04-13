function setting() {

var form = FormApp.openById('1FunkM2Wipxv-VppCsQeK0H79QUm_pdu05KbSxnjV8tU'); // defining the form through its id
ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
}

function onFormSubmit(e) {

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses'); // the spreadsheet that collects the form responses

for(var i=1; i < 10000; i++) {


if(sheet.getRange(i,4).getValue() != 0 && sheet.getRange(i,5).getValue() == 0) {

reportFill(i); 
/* 
This is a loop through the sheet. 
If the column 4 (email, that is optional) isn't empty and the column 5 (verification for emails sent) is empty - what means that:
for the indicated email, the script didn't send the letter - then the reportFill function will run for that row (i) content
*/

}

else { continue };
  
}
}

function reportFill(row) {

//getting information from the Google Spreadsheet 

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
var name = sheet.getRange(row,2).getValue();
var email = sheet.getRange(row,4).getValue();

//getting information from the Google Slides (used for the template)
  
var template = DriveApp.getFileById('1WmsVVqtQ-QjNNuGdv_s-Jl0tk8dYHYvtlPc4U2RB44E');
var letter = template.makeCopy('Welcome to Hogwarts '+name);
var editt = SlidesApp.openById(letter.getId());
var url = letter.getUrl();

// Replace the text {name} with the indicated Name on the form.

var presentationId = letter.getId();
var resource = {
  "requests": [
    {
      "replaceAllText": {
        "containsText": {
          "text": "{name}"
        },
        "replaceText": name // If this is not defined, the text is searched from all slides.
      }
    }
  ]
};
Slides.Presentations.batchUpdate(resource, presentationId);
  
// Create a PDF with the letter filled with Forms Name
  
var pdf = DriveApp.getFileById(editt.getId()).getAs('application/pdf').getBytes();
var attach = {fileName:'Welcome to Hogwarts '+name+'.pdf',content:pdf, mimeType:'application/pdf'};

// Send and e-mail, set a confirmation and the letter url (as a google slides) on the Spreadsheet 
  
mail(email,name,attach);
sheet.getRange(row,5).setValue('Email enviado');
sheet.getRange(row,6).setValue(url);

}

function mail(email,name,attach) {

  MailApp.sendEmail(email,'Welcome to Hogwarts, '+name+'!','Prezado, '+name+'\nGostaríamos de dar às boas-vindas à Escola de Magia e Bruxaria de Hogwarts!\nA sua coruja já chegou com a sua carta!\n\nP.S.: Esperamos que você tenha um ótimo RPG!\nAtt,\nOrganização Hogwarts Tamroc 2020',{attachments:[attach]});

}
