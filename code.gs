//Date utility function
Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

//takes shot1+shot2 and aptDate and filters based on date 
function filtering(columnnumber) {
  var sss = SpreadsheetApp.openById('**********************'); //blocked out sheetID
  var ss = sss.getSheetByName('Master');   //replace with source Sheet tab name
  var range = ss.getRange('A:Z');        //assign the range you want to copy
  var rawData = range.getValues()        // get value from spreadsheet 1

  var filterdData = []                       // Filtered Data will be stored in this array
  var now = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  for (var i = 0; i < rawData.length ; i++){
    if(rawData[i][columnnumber]){
      if(Utilities.formatDate(new Date(rawData[i][columnnumber]), "GMT", "yyyy-MM-dd") === now)            // Check to see if shot date 1 is today
        {
          filterdData.push(rawData[i])
        }
    }
  }
  return filterdData
};

function createTable(data){ 
  var cells = [];
  //This would be the header of the table
  var table = "<br><table border=1><tr><th>DMS #</th><th>Shelter</th><th>OPH Name</th><th>DMS Link</th><th>Review Notes</tr></br>";

  //the body of the table is build in 2D (two foor loops)
  for (var i = 0; i < data.length; i++){
    
      cells = data[i]; //puts each cell in an array position
      table = table + "<tr>";

      for (var u = 0; u < cells.length; u++){
        if(u==0||u==2 ||u==4 ||u==20 ||u==10){
          table = table + "<td>"+ cells[u] +"</td>";
        }
      }
  table = table + "</tr></table>"
  }

  return table;
}



//prepares report to send
function Shelter_Vaccination_Report(shotinput,aptinput) {
  var Emailtest=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails")
  var now = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  var subject = "OPH Vaccination Follow-up Reminder -";
  var subject=subject.concat(now);
  var emailList =Emailtest.getRange(2, 1, 100).getValues();

  var shotreport=createTable(shotinput)
  var aptreport=createTable(aptinput)

    for (let index = 0; index < emailList.length; ++index) {
      const email = emailList[index][0];


      if(email){
        var message="<html><body>"
        var name = email.split("@")[0]; // Second column
        var message = '<p>Dear '+name+", </p>"

        if((shotinput === undefined || shotinput.length == 0)==false){
          var message=message+"<p>"+'The following pet(s) are scheduled to receive shot(s) today (' + now+'):'; // Second column
          message+=shotreport
        }
        message+="</p>"

        if ((aptinput === undefined || aptinput.length == 0)==false){
          var message = message+'<p>The following pet(s) have an appointment today (' + now+'):'; // Second column
          message+=aptreport
        }
        message+="</p>"

        var message=message+ "<p> If you have any questions/concerns about this report contact:</p> <p>Hirsh  or Jon  </p>"

        var message=message+"</body></html>"
        console.log("Email Sent to:", email)
        MailApp.sendEmail(email, subject, "",{ htmlBody: message });
      };
    
    //link to spreadsheet
    }
}


//runs script
function runscript(){
  var shot1=(filtering(19))
  var shot2=(filtering(20))

  var aptDate=(filtering(22))

  var combined=shot1.concat(shot2)

  let set  = new Set(combined.map(JSON.stringify));
  let arr2 = Array.from(set).map(JSON.parse);

  if (((arr2 === undefined || arr2.length == 0)==false) || ((aptDate === undefined || aptDate.length == 0)==false)){
    Shelter_Vaccination_Report(arr2,aptDate)
  }
  else{
    console.log("No appointments today")
  }
}




