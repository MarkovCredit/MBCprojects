///MLevine 5/3/2021 \\\\

//The purpose of this script is to email our evangelists as a means 
//to follow up with community members and further the gospel and possible
//membership at MBC. 


//check emails remaining (I believe its 100 per day)
var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
console.log(emailQuotaRemaining);

//lets create some variables we may need globally
var sheet = SpreadsheetApp.getActive().getSheetByName("evangelist_database");
var customMessage_ = SpreadsheetApp.getActive().getRangeByName("input_custom_email_message!customMessage").getValue();

console.log(customMessage_);
//check the formURL from the linked google Doc
try {
  const formURL = sheet.getFormUrl();
  // console.log(formURL);
}
catch(e) {
  // console.log(e)
};

var startRow = 2;
var numRows = sheet.getLastRow();
//headers
var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
//check if the email sent column is present; if it is do nothing, else add it in there!
if(headers.indexOf('EmailSent') > -1){

}
else{
  
  sheet.getRange(1,sheet.getLastColumn()+1).setValue('EmailSent')
  
}

//specifying the index of the columns we will ultimately send out via email
var emailIndex = headers.indexOf('Email');
var emailSentIndex = headers.indexOf('EmailSent')
//set our dataRange and grab the values
//we need to find the last row and column similar to Excel where you go to the 
//bottom and come up (or the right and come back left for columns)


var dataRange = sheet.getRange(startRow,1,numRows,sheet.getLastColumn());
var data_ = dataRange.getValues();


class prepData{
  constructor(preppedData,distinctEmails){
    console.log('the constructor is being created!')
    this.preppedData = new Array;
    this.distinctEmails = new Array;
  }
    prepareSheetsData() {

      
      //grab the emails column
      var emails =  sheet.getRange(startRow,emailIndex+1,numRows-1).getValues();  
      //flatten the two dimensional array 
      var flattened_emails = [].concat(...emails);  
      //use filter with index of to only take distinct emails
      this.distinctEmails = flattened_emails.filter((e,i)=>
      flattened_emails.indexOf(e)===i);
      // this.preppedData = new Array;
      
      for (var e = 0; e < this.distinctEmails.length; ++e){
        var key = this.distinctEmails[e]
        this.preppedData[key] = []
        for (var i = 0; i < data_.length; ++i){
          var row = data_[i];
          //  console.log(row)
          if(row[3]=== key){
            this.preppedData[key].push(row)
          }
          
          }
        }
        return this.preppedData
      }
    
  };

    
//function to grab the data from sheets and prep it for use in HTML

 //we need to itearte over distinct emails; then with the array inside of each;
 //run a script in the HTML that only looks at the array inside of it! 
// console.log(distinct_emails)
 function getEmailHtml(emailData) {
   
   var htmlTemplate = HtmlService.createTemplateFromFile("Template.html");
   htmlTemplate.customMessage = customMessage_
   htmlTemplate.messages = emailData;
   var htmlBody = htmlTemplate.evaluate().getContent();
   return htmlBody;

};


function sendEmail() {
  var emailData_ = new prepData().prepareSheetsData();
  
  for (let k in emailData_){
    // console.log(`Printing each row from the email messages ${emailData_[k]}`);
    var htmlBody_ = getEmailHtml(emailData_[k])
    console.log(htmlBody_);
    
    MailApp.sendEmail({
    to: k,
    subject: "Evangelism Update and Follow Up",
    
    htmlBody: htmlBody_
  });
    emailData_[k].forEach((row,index)=>{
      console.log(`Row${row[1]}`)
      sheet.getRange(row[1],emailSentIndex+1).setValue('1')


    }

  );
};
};