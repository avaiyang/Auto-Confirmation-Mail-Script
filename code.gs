/*
#############################################################################################################################################################
README:

Auto Confirmation mail Script.

Note:
Select function Initialize() and run it.

function GenerateID(e): 
Generates auto incrimented unique application ID. 
The newly generated id will be written in column BL.

function SendConfirmationMail(e) :
Sends confirmation email to Applicant.


function InformRecommenders(e):
Sends Request for filling out recommendation forms to the two recommenders.

Author: Avaiyang Garg

##############################################################################################################################################################
*/


var applicantID = 0;
var applicant_email_column = "Your Email Address";
var reco1_email_column = "Recommender # 1 email address";
var reco2_email_column = "Recommender # 2 email address";
/* Send Confirmation Email with Google Forms */
 
function Initialize() {
  
 
  
  var triggers = ScriptApp.getProjectTriggers();
 
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  ScriptApp.newTrigger("GenerateID")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
 
  ScriptApp.newTrigger("SendConfirmationMail")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  
  ScriptApp.newTrigger("InformRecommenders")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  

}
 
function SendConfirmationMail(e) {
 
  try {
 
    var ss, cc, sendername, subject, columns;
    var message, value, textbody, sender;
    
    var uniqID = GenerateID(e);
    
    // This is your email address and you will be in the CC
 
    // This will show up as the sender's name
    sendername = "NYU Mechatronics";
 
    // Optional but change the following variable
    // to have a custom subject for Google Docs emails
    subject = "RET 2018 Application Received";
     
 
    // This is the submitter's email address
    // Make sure you have a field called Email Address in the Google Form
    sender = e.namedValues["Your Email Address"].toString();
    
    var applicant_name = e.namedValues["First Name"].toString() + " " + e.namedValues["Last Name"].toString();
    Logger.log(uniqID);
    
   
    var html = HtmlService.createHtmlOutputFromFile('Index').getContent().replace("xyz123", applicant_name).replace("123xyz", uniqID); // where template is the name of our html file!;
    
    var message = "Dear "+ applicant_name + "\n\nThank you for the expressed interest in our application. The unique id for your application is: " + uniqID+
    "\n\nWe will contact you regarding further process once your recommenders have provided with the recommendation letter.";
    
    
    MailApp.sendEmail({name: sendername, to: sender, subject: subject, htmlBody: html});
    
    

  } catch (e) {
    Logger.log(e.toString());
  }
 
}




function InformRecommenders(e){

  try {
    
    var uniqID = GenerateID(e);
    var ss, cc, sendername, subject, columns;
    var message, value, textbody, sender, message2;
    var applicant_name = e.namedValues["First Name"].toString() + " " + e.namedValues["Last Name"].toString();
    var email_id = e.namedValues["Your Email Address"];
    
    var reco_1_fullname = e.namedValues["Recommender # 1 First Name (Principal)"].toString() + " " + e.namedValues["Recommender # 1 Last Name"].toString();
    var reco_2_fullname = e.namedValues["Recommender # 2 First Name"].toString() + " " + e.namedValues["Recommender # 2 Last Name"].toString();
    var reco_1_email = e.namedValues["Recommender # 1 email address"].toString();
    var reco_2_email = e.namedValues["Recommender # 2 email address"].toString();
   
    
    // This is your email address and you will be in the CC
 
    // This will show up as the sender's name
    sendername = "NYU Mechatronics";
 
  
    // to have a custom subject for Google Docs emails
    subject = "RET 2018 Recommendation Request";
 
    // This is the body of the auto-reply
    message =  "Dear "  + reco_1_fullname + "," + "\n\nApplicant "+applicant_name+" has listed you as a recommender for his/her application to a teacher professional development program. " + "The unique id of the application is: " + uniqID+ 
      "\n\nKindly submit your evaluation about the suitability of applicant for this program. "+
    "In your letter, please explicitly indicate how you and your school will support the applicant to integrate his/her professional development experiences in classroom teaching and learning. "+
      "\n\nPlease click on the link to upload your recommendation letter: https://goo.gl/forms/X8sKtuPorewEign42 " +
        "\n\nNOTE: Please upload the recommendation letter in PDF file format."+
                "\n\nBest Regards,\nThe RET Team\nNew York Unversity";
 

    // This is the submitter's email address
    // Make sure you havea  field called Email Address in the Google Form
    sender = reco_1_email;
    
    //for 1st recommender
    var html = HtmlService.createHtmlOutputFromFile('Index1').getContent().replace("recommender123", reco_1_fullname).replace("student123", applicant_name).replace("email123", email_id).replace("uniq123", uniqID); // where template is the name of our html file!;
    
    MailApp.sendEmail({name: sendername, to: reco_1_email, subject: subject, htmlBody: html});
 
    
    
    
    sender = reco_2_email;
    //for 2nd recommender
    var html = HtmlService.createHtmlOutputFromFile('Index1').getContent().replace("recommender123", reco_2_fullname).replace("student123", applicant_name).replace("email123", email_id).replace("uniq123", uniqID); // where template is the name of our html file!;
    
    message2 =  "Dear "  + reco_2_fullname + "," + "\n\nApplicant "+applicant_name+" has listed you as a recommender for his/her application to a teacher professional development program. " + "The unique id of the application is: " + uniqID+ 
      "\n\nKindly submit your evaluation about the suitability of applicant for this program. "+
    "In your letter, please explicitly indicate how you and your school will support the applicant to integrate his/her professional development experiences in classroom teaching and learning. "+
      "\n\nPlease click on the link to upload your recommendation letter: https://goo.gl/forms/X8sKtuPorewEign42 " +
        "\n\nNOTE: Please upload the recommendation letter in PDF file format."+
                "\n\nBest Regards,\nThe RET Team\nNew York Unversity";
 
    //var html = HtmlService.createHtmlOutputFromFile('Index1').getContent().replace("recommender123", reco_2_fullname).replace("student123", applicant_name).replace("email123", email_id).replace("uniq123", uniqID); // where template is the name of our html file!;
    MailApp.sendEmail({name: sendername, to: reco_2_email, subject: subject, htmlBody: html});    
 
  } catch (e) {
    Logger.log(e.toString());
  }
  
}





function GenerateID(e) {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRowIndex() - 1;
  if(sheet.getRange(row, 35).getValue() == 0 || sheet.getRange(row, 35).getValue() == "uniqID")
  {
    applicantID = 20180001;
    sheet.getRange(row+1, 35).setValue(applicantID);
    return (applicantID.toString()); 
  }
  
  applicantID = sheet.getRange(row, 35).getValue() + 1;
  sheet.getRange(row+1, 35).setValue(applicantID);
  
  return (applicantID.toString());
  
}
