/*********************************************
Voyage 2011-2013 copyright
This file is part of the Voya.ge rental manager.
Voya.ge rental manager is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation.
Voya.ge rental manager is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with Voya.ge rental manager.  If not, see <http://www.gnu.org/licenses/>.
For more information and to contribute visit forum.voya.ge 
***************************************************/

/**
 * Loads up Rentals Manager menu on Open 
 */
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Rentals Manager", [{name: "Preferences", functionName: "showManagerPreferences"},{name: "Send Emails", functionName: "addNewBookings"}, 
                                 {name: "Initialize Templates", functionName: "initializeTemplates"}, 
                                 {name: "Refresh Charts", functionName: "updateAllCharts"}, 
                                 {name: "Refresh Feed", functionName: "runScripts"}]);
}


/**
 * Runs major scripts to update feed
 */
function runScripts(){
  importData();
  dailyStats();
  setAvailability();
}


/**
 * Shows the Preferences for the software
 */
function showManagerPreferences() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Manager preferences');

  var grid = app.createGrid(7, 3);
  
  var percentageBox = app.createTextBox().setName('percentageBox').setText(ScriptProperties.getProperty('Deposit percentage')*100);
  grid.setWidget(0, 0, app.createLabel('To confirm booking'));
  grid.setWidget(0, 1, percentageBox);
  grid.setWidget(0, 2, app.createLabel('% must be paid'));
  
  var balanceDueBox = app.createTextBox().setName('balanceDueBox').setText(ScriptProperties.getProperty('Balance due date'));
  grid.setWidget(1, 0, app.createLabel('Balance due'));
  grid.setWidget(1, 1, balanceDueBox)
  grid.setWidget(1, 2, app.createLabel('days before arrival'));
  
  var reminderBox = app.createTextBox().setName('reminderBox').setText(ScriptProperties.getProperty('Days before reminder sent'));
  grid.setWidget(2, 0, app.createLabel('A payment reminder is sent'));
  grid.setWidget(2, 1, reminderBox);
  grid.setWidget(2, 2, app.createLabel('days before balance is due '));
  
  var replyToBox = app.createTextBox().setName('replyToBox').setText(ScriptProperties.getProperty('ReplyTo Email'));
  grid.setWidget(3, 0, app.createLabel('Clients send emails to'));
  grid.setWidget(3, 1, replyToBox);
  
  grid.setWidget(4, 0, app.createLabel('Voya.ge API Key'));
  grid.setWidget(4, 1, app.createTextBox().setName('apiBox').setText(ScriptProperties.getProperty('API Key')));
  
  grid.setWidget(5, 0, app.createLabel('Language codes for Emails'));
  grid.setWidget(5, 1, app.createTextBox().setName('languagesBox').setText(ScriptProperties.getProperty('Confirmation Languages')));
  grid.setWidget(5, 2, app.createLabel('comma seperated EX: FR,EN'));


  var button = app.createButton('Save').setStyleAttribute("fontSize", "12pt" );
  var panel = app.createVerticalPanel().setSize("100%", "100%")
      .setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
  
  var handler = app.createServerHandler('savePreferences')
      .validateRange(percentageBox, 0, 100).validateInteger(balanceDueBox).validateInteger(reminderBox).validateEmail(replyToBox)
      .addCallbackElement(grid);
  button.addClickHandler(handler);
  
  panel.add(grid);
  panel.add(button);
  app.add(panel);
  ss.show(app);
  
  return app
}

/**
 * Saves the preferences into ScriptProperties
 */
function savePreferences(e){
  var app = UiApp.getActiveApplication();
  
  ScriptProperties.setProperty('Deposit percentage', e.parameter.percentageBox/100);
  ScriptProperties.setProperty('Balance due date', e.parameter.balanceDueBox);
  ScriptProperties.setProperty('Days before reminder sent', e.parameter.reminderBox);
  ScriptProperties.setProperty('ReplyTo Email', e.parameter.replyToBox);
  ScriptProperties.setProperty('Confirmation Languages', e.parameter.languagesBox);
  
  var apiKey = e.parameter.apiBox;
  
  if(apiKey){ 
    ScriptProperties.setProperty('API Key', apiKey);
    importData();
  }
  
  app.close();
  return app;
}


/**
 * Creates the templates for all the correspondances
 */
function initializeTemplates(){  
  var templateTypes = ["Deposit Request", "Rental Confirmation", "Balance Reminder", "Balance Received", "Rental Agreement", "House instructions"];
  var languages = ScriptProperties.getProperty('Confirmation Languages').replace(/ /g, "").toUpperCase().split(",");
  var folder;
  
  //Initialize statuses if there are none
  initializeStatuses();
  
  //find out if the main folder exists, if it doesn't make it
  try{folder = DocsList.getFolderById(ScriptProperties.getProperty("Voyage Templates Folder"));}
  catch(err){
    folder = DocsList.createFolder("Voyage Rentals Templates");
    ScriptProperties.setProperty("Voyage Templates Folder", folder.getId());
  }
    
  var lang;
  //loop that goes through languages
  while(languages.length>0){
    lang = languages.pop();
    //loop that goes through template types
    for(var i=0; i<templateTypes.length; i++){
      //does it exist?
      try{DocsList.getFileById(ScriptProperties.getProperty(lang + " " + templateTypes[i]));}
      //no it doesn't, make it
      catch(err){
        ScriptProperties.setProperty(lang + " " + templateTypes[i], addTemplate(folder, templateTypes[i], lang));
      }
    }
  }
}

/**
 * Adds a template from the originals to a folder
 * @param {Folder} folder The docsList folder to which the templates must be added
 * @param {string} templateType The title of the template to be added
 * @param {string} lang The language of the added template
 * @return {string} the template's ID
*/
function addTemplate(folder, templateType, lang){
  var originalTemplateID;

  //depending on the template, get the right ID
  switch(templateType){
    case "Deposit Request":
      originalTemplateID = "1KqvVaJpT3O_k1NigFjLp4a3htfPq0AJJoI0Vur8R5JE";
      break;    
    case "Rental Confirmation":
      originalTemplateID = "1ohUqk10wLkTPgkz4UFzYgmeL7eT_x3alyJGOudtx9bc";
      break;
    case "Balance Reminder":
      originalTemplateID = "13GkA4WE6rz1XSLBybiMihMAgZLr_BPW5upodaDHmgtE";
      break;
    case "Balance Received":
      originalTemplateID = "1whVqq5rDS4Ml3RVDvKLCEy63LDIoC26_hHH1MIcEQvc";
      break;
    case "Rental Agreement":
      originalTemplateID = "1Q8jY-rdgkYQQshPkZ-aq1ChLafkaRH7MxYm5VqFgcTw";
      break;
    case "House instructions":
      originalTemplateID = "1fblfemXq0isxl9-MeCX8fDpTWHDDH1LsptMc0vC61YE";
      break;
    default:
      throw("bad templateType");
  }
  //put that template in the folder
  var newDoc = DocsList.getFileById(originalTemplateID).makeCopy(lang + " " + templateType);
  newDoc.addToFolder(folder);
  return newDoc.getId();
}


/**
 * If there is nothing in the statuses, this adds the booking #s and PAID for all bookings
 */
function initializeStatuses(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var emailStatesSheet = ss.getSheetByName("Statuses");
  
  //if there are things in the array stop to not override
  if(emailStatesSheet.getRange('A2').getValue()!=''){
    throw("There are things in the Statuses, I wont override");
    return;
  }
  
  var bookingsSheet = ss.getSheetByName('Data import');
  var bookings = bookingsSheet.getDataRange().getValues();
  bookings.shift();
  bookings.sort(function(a,b){return a[0]-b[0]});
  
  //remove all bookings before today
  var today= new Date();
  while(bookings.length>0&&today>bookings[0][0]) {bookings.shift();}
  
  //get the booking refNum column
  var newBookings = getColumn(bookings, 7);
  
  //make an array with bookings marked as paid
  var paidArray = new Array();
  while(newBookings.length>0){
    paidArray.push([newBookings.shift(), 'PAID']);
  }
  //put in titles and push to sheet
  paidArray.unshift(["Booking#", "Status"]);
  emailStatesSheet.clearContents();
  emailStatesSheet.getRange(1, 1, paidArray.length, 2).setValues(paidArray);  
}

/**
 * Grabs data from feed and puts it on the spreadsheet
 */
function importData() {
  var url = "https://voya.ge/account/csv_export/feed?api_key="+ScriptProperties.getProperty('API Key');
  //get the feed
  try{var doc = UrlFetchApp.fetch(url).getContentText();}
  catch(err){throw("Can't fetch feed. "+url+" Verify your api key and try again"); return;}
 
  //for big feeds, split in half
  if(doc.length>150000){
    var placeToSplit = (doc.length/2).toFixed();
    while(doc[placeToSplit+1]!=='2' || doc[placeToSplit+5]!='/' || doc[placeToSplit+11]!=',' || doc[placeToSplit+12]!='2' || doc[placeToSplit+16]!='/' || doc[placeToSplit+22]!=',' || doc[placeToSplit+23]!='2' || doc[placeToSplit+27]!='/' || doc[placeToSplit+33]!=','){placeToSplit++;}
    var doc1 = doc.slice(0, placeToSplit);
    var doc2 = doc.slice(placeToSplit+1, doc.length-1);
    var bookings1 = CSVToArray(doc1);
    var bookings2 = CSVToArray(doc2);
    var bookings = bookings1.concat(bookings2);
  }else{
    var bookings = CSVToArray(doc);
  }
  if(!bookings[bookings.length-1][0]){bookings.pop();}

  bookings.sort(function(a,b){return new Date(b[0])-new Date(a[0])});
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taxesSheet = ss.getSheetByName('Tax settings')
  var taxes = taxesSheet.getRange(2, 1, taxesSheet.getLastRow()-1, 5).getValues();
  
  bookings = calculateAllTaxes(bookings, taxes);
      
  var dataSheet = ss.getSheetByName("Data import");
  dataSheet.clearContents();
  dataSheet.getRange(1, 1, bookings.length, bookings[0].length).setValues(bookings);  
}


/** 
 * FROM http://www.bennadel.com/blog/1504-Ask-Ben-Parsing-CSV-Strings-With-Javascript-Exec-Regular-Expression-Command.htm
 * This will parse a delimited string into an array of arrays. The default delimiter is the comma.
 * @param {string} strData The string to be parsed
 * @param {string} strDelimiter A delimiter to override the default ,
 * @return An array built of the CSV
 */
function CSVToArray( strData, strDelimiter ){
  // Check to see if the delimiter is defined. If not,
  // then default to comma.
  strDelimiter = (strDelimiter || ",");

  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +
      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){

    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){
      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );
    }

    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){
      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );
    } else {
      // We found a non-quoted value.      
      var strMatchedValue = (arrMatches[ 3 ] || null);
    }
    
    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }

  // Return the parsed data.
  return( arrData );
}


/** 
 * Returns all bookings with tax details
 * @param {Array} bookings An array with all bookings
 * @param {Array} taxes: all historical tax information
 * @return All bookings with taxes*/
function calculateAllTaxes(bookings, taxes){
  //add tax names
  bookings[0].push('Before Tax');
  for(var i = 0; i<taxes.length; i++){if(taxes[i][0] !== ''){bookings[0].push(taxes[i][0]);}}
  
  for(var i=1; i<bookings.length; i++){
    bookings[i] = bookings[i].concat(getTaxes(bookings[i], taxes));
  }
  return bookings;
}


/** 
 * Returns the value of the booking's tax amounts
 * @param {Array} booking A booking
 * @param {Array} taxes  All historical tax information
 * @return The booking's tax amounts
 */
function getTaxes(booking, taxes){
  var price = removeTaxes(booking, taxes);
  var taxesTotal = 0;
  var taxesArray = new Array();
  //first thing to be returned
  taxesArray.push(price);
  
  var arrivalDate = new Date(booking[0]);
  var departureDate = new Date(booking[1]);
  var stayLength = (departureDate-arrivalDate)/(1000*60*60*24);

  //loop that calculates taxes 
  for(var i = 0; i<taxes.length; i++){
    
    //Taxes went up after booking date, charge old level
    if(booking[2] < taxes[i][4] && taxes[i][1] > taxes[i+1][1]){ continue; }
    
    //find our tax dates
    if(arrivalDate >= taxes[i][4] && taxes[i][4]){
      var tax = 0;
      
      //tax is a Percentage
      if(taxes[i][2]==='%'){
        
        //compounded, previous taxes to tax
        if(taxes[i][3].toLowerCase()==='y'){tax = (price+taxesTotal) * (taxes[i][1]/100);}
        
        //not compounded, just calculate tax
        else{tax =  price * (taxes[i][1]/100);}
      }
      //per night tax
      else{tax = stayLength*taxes[i][1];}
      
      //add this tax to the others
      taxesTotal = taxesTotal + tax;
      //then go to the next tax
      while(i+1<taxes.length && taxes[i+1][0] === ''){i++;}
      
      //push the tax
      taxesArray.push(tax);
    }
  }
  return taxesArray;
}


/**
 * Returns the value of the booking without taxes
 * @param {Array} booking A booking
 * @param {Array} taxes All historical tax information
 * @return The booking amount before taxes
 */
function removeTaxes(booking, taxes){
  var price = booking[4];
  if(price == 0){return 0;}
  var arrivalDate = new Date(booking[0]);
  var departureDate = new Date(booking[1]);
  var stayLength = (departureDate-arrivalDate)/(1000*60*60*24);
  var taxPercent = 1.0;
  var taxPerNight = 0.0;
  
  //loop that calculates taxes level
  for(var i = 0; i<taxes.length; i++){
    //find tax for date
    if(booking[2] < taxes[i][4] && taxes[i][1]>taxes[i+1][1]){ continue; }
    if(arrivalDate >= taxes[i][4]){
      if(taxes[i][2]==='%'){
        //compounded, multiply
        if(taxes[i][3].toLowerCase()==='y'){
          taxPercent = taxPercent * (1+(taxes[i][1]/100));
          taxPerNight = taxPerNight * (1+(taxes[i][1]/100));
        }
        //not compounded, add
        else{taxPercent = taxPercent + (taxes[i][1]/100);}
      }
      else{taxPerNight = taxPerNight + taxes[i][1];}
      
      //then go to the next tax
      while(i+1<taxes.length && taxes[i+1][0] === ''){i++;}
    }
  }

  return ((price/taxPercent)-(taxPerNight*stayLength));
}


/**
 * Gives us a total of taxes and revenues for that period
 * @param {Date} startDate The date when we want to start the report
 * @param {Date} endDate The date we want to end the report
 * @param {Array} bookings All the bookings with taxes
 * @return An array of the total of taxes for that period
 */
function taxReport(startDate, endDate, bookings){
  //remove emptycells
  while(!bookings[bookings.length-1][0]){bookings.pop();}
  //add block
  bookings.push(makeArray(bookings.length,null,endDate))
  
  bookings.sort(function(a,b){return a[0]-b[0]});
  var taxes = new Array();
  //make the array
  taxes.push(bookings.shift().slice(22));
  taxes.push(new Array(taxes[0].length));
  for(var i=0; i<taxes[0].length; i++){taxes[1][i] = 0;}
  
  //loop that finds the beginning day
  while(bookings[0][0]<startDate){bookings.shift();}
  
  //loop that looks at what we need
  while(bookings[0][0]<endDate){
    
    //loop that makes the total for each piece
    for(var i=0; i<taxes[0].length; i++){taxes[1][i] = taxes[1][i] + bookings[0][i+22];}
    
    bookings.shift();
  }
  return taxes; 
}


/**
 * Returns a 1d array of the right length with the locations filles
 * @param {number} length The length of the array
 * @param fill What to fill the array with
 * @return An array filled of the length
 */
function makeArray(length, fill, firstField){
  var filledArray = [(firstField || fill)];
  while(length>1){
    filledArray.push(fill);
    length--;
  }
  return filledArray;
}


/**
 * Grabs bookings from spreadsheet, sends emails, updates statuses
 */
function addNewBookings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookingsSheet = ss.getSheetByName('Data import');
  var bookings = bookingsSheet.getDataRange().getValues();
  bookings.sort(function(a,b){return a[0]-b[0]});
  
  //iferror in dataimport, do nothing so that the information in the list isn't lost
  if(bookings.shift()[0]!="Start Date"){
    throw "Sorry I couldn't import range at this time. Try again later";
    return;
  }

  //remove all bookings before today
  var today= new Date();
  while(bookings.length>0&&today>bookings[0][0]) {bookings.shift();}
  //get the booking refNum column
  var newBookings = getColumn(bookings, 7);
  
  //get the old bookings statuses
  var emailStatesSheet = ss.getSheetByName("Statuses");
  var emailStatesRange = emailStatesSheet.getDataRange();
  var oldBookings = emailStatesRange.getValues();
  oldBookings.shift();

  //merge old and new bookings
  var statesMerged = mergeBookings(oldBookings, newBookings);
  
  //Send emails, change statuses of bookings that require it
  statesMerged = emailCrawler(statesMerged, bookings);

  //push the bookings merged to the emailStatesSheet
  emailStatesSheet.clearContents();
  statesMerged.unshift(["Booking#", "Status"]);
  emailStatesSheet.getRange(1, 1, statesMerged.length, 2).setValues(statesMerged);
}


/**
 * Updates old bookings adds new bookings
 * @param {Array} oldBookings Array of old booking ref# and states
 * @param {Array} newBookings Array of new bookings ref# 
 * @return New clean array with only new bookings and their state
 */
function mergeBookings(oldBookings, newBookings){
  //merge both arrays and orders them
  var bookingsMerging = oldBookings.concat(newBookings);
  bookingsMerging.sort();
  var cleanBookings = new Array();
  //Goes through the array, add state if empty, delete if not in newBookings
  for(var i=0; i<bookingsMerging.length; i++){
    //the booking is there twice, keep it and its state
    if(i+1<bookingsMerging.length && bookingsMerging[i]==bookingsMerging[i+1][0]){
      i++;
      cleanBookings.push(bookingsMerging[i]);
      
      //the booking is in the new booking only, add it as new
    }else if(!(bookingsMerging[i].constructor ==Array)){
      cleanBookings.push([bookingsMerging[i],"new"]);
    }
  } 
  return cleanBookings;
}


/**
 * Goes down the statuses array and sends emails when needed then updates the status
 * @param {Array} statuses Array of booking ref#s and statuses
 * @param {Array} bookings Array of all bookings in the future
 * @return Array of updated booking ref#s and statuses
 */
function emailCrawler(statuses, bookings) {
  //loop that crawls bookings and takes action on each
  for(var j=0; j<statuses.length; j++){
    statuses[j][1]=bookingAction(bookings[findBooking(bookings, statuses[j][0])], statuses[j][1]);
  }
  return statuses;
}


/** 
 * If a booking needs attention, sends email and returns new status
 * @param {Array} booking The booking entry
 * @param {string} state  The booking's state
 * @return state of booking after check
 */
function bookingAction(booking, state){
  
  //if no email address get out
  if(!booking[9]){return "NO EMAIL";}
  
  var daysToBooking = (booking[0].getTime()-new Date().getTime())/(1000*60*60*24);
  var cost = Number(booking[4]);
  var paid = Number(booking[5]);
  state = state.toUpperCase();

  //*******Lots of conditions to know what to do with booking!**********
  //free booking do nothing
  if(cost==0){state="FREE";}
  else{
    
    //all paid and emailed do nothing
    if(state=="PAID"){}
    else{
    
      //paid send final email
      if(paid>=cost){sendMail(booking, "Balance Received"); state="PAID";}
      else{
        
        //balance reminded, not paid, do nothing (if your would like to send yourself an email this would be a good place)
        if(state=="REMINDED"){}
        else{
          
          //Reminder of balance due
          if(state=="CONFIRMED" && daysToBooking<Number(ScriptProperties.getProperty('Balance due date'))+Number(ScriptProperties.getProperty('Days before reminder sent'))){sendMail(booking, "Balance Reminder"); state="REMINDED";}
          else{
            
            //Deposit received send confirmation of booking
            if(paid>0){
              if(state!="CONFIRMED"){sendMail(booking, "Rental Confirmation");state="CONFIRMED";}}
            else{
              
              //booking received send deposit request
              if(!(paid>0)){
                if(state!="DEPOSIT REQUESTED"){sendMail(booking, "Deposit Request"); state="DEPOSIT REQUESTED";}
              }
            }
          }
        }
      }
    }
  }
  return state; 
}


/**
 * Finds the booking's index with its ref#
 * @param {Array} bookings Array of bookings in which to look
 * @param {number} refNum Booking reference number to look for
 * @return the index at which that booking is
 */
function findBooking(bookings, refNum){
  //loop that goes through the bookings to find the right one
  for(var i=0; i<bookings.length; i++){
    if(bookings[i][7]==refNum) return i;
  }
}


/**
 * Send the client the right email with booking information
 * @param {Array} booking The booking that needs an email sent
 * @param {string} emailSubject The type of message to be sent
 * @return true if an email was sent, false otherwise
 */
function sendMail(booking, emailSubject){
  var total = booking[4].toFixed(2);
  var deposit = booking[5].toFixed(2);
  var email = booking[9];
  var lang = booking[10].toUpperCase();
  var replyToEmail = ScriptProperties.getProperty('ReplyTo Email');
  
  var id = ScriptProperties.getProperty(lang + " "+ emailSubject);
  var url = 'https://docs.google.com/feeds/';
  try{var doc = UrlFetchApp.fetch(url+'download/documents/Export?exportFormat=html&format=html&id='+id,googleOAuth_('docs',url)).getContentText();}
  catch(err){throw("Can't fetch template. Make sure they were generated in this email's language and are present in ScriptProperties. If they were never generated, select Initialize Templates from menu.");}
  
  
  //if there is no deposit paid calculate it
  if(!(deposit>0)){deposit = (total*ScriptProperties.getProperty('Deposit percentage')).toFixed(2);}
  var balance = total-deposit;
  
  //add all the booking information in the email
  doc = doc.
  replace(/keyarrival/gi, getFormattedDate(booking[0])).
  replace(/keydeparture/gi, getFormattedDate(booking[1])).
  replace(/keyproperty/gi, booking[3]).
  replace(/keyguests/gi, booking[6]).
  replace(/keyname/gi, booking[8]).
  replace(/keyref#/gi, booking[7]).
  replace(/keytotal/gi, total).
  replace(/keydeposit/gi, deposit).
  replace(/keybalance/gi, balance).
  replace(/keydays/gi, ScriptProperties.getProperty('Balance due date')).
  replace(/keytaxes/gi, getTaxString(booking.slice(23, booking.length)));
  

  //email with PDF
  if(emailSubject == "Balance Received" || emailSubject == "Deposit Request"){
    
    var pdfTitle = " Rental Agreement";
    if(emailSubject == "Balance Received"){pdfTitle = " House instructions";}
    var pdfID = ScriptProperties.getProperty(lang + pdfTitle);
    var pdf = getPDF(pdfID, booking, balance);
    
    try{MailApp.sendEmail(email, DocsList.getFileById(id).getName(), 'html only', 
                          {htmlBody:doc, attachments: pdf, replyTo:replyToEmail, bcc:replyToEmail});}
    catch(err){sendError_(booking, replyToEmail); return false;}
    
    return true;  
  }
  
  //email without pdf
  else{
    try{ MailApp.sendEmail(email, DocsList.getFileById(id).getName(), 'html only', 
                           {htmlBody:doc, replyTo: replyToEmail, bcc:replyToEmail});}
    catch(err){sendError_(booking, replyToEmail); return false;}
    
    return true;
  }
  
  return false;
}


/**
 * Sends out an error email
 * @param {string} booking The booking that could be get its email sent
 * @param {string} replyToEmail The email to which the error should be sent
 */
function sendError_(booking, replyToEmail){
  MailApp.sendEmail(replyToEmail, "Voya.ge manager tried to send an email and failed", 'Sorry, this is booking' + booking);
}


/**
 * gets a formatted date string in YYYY/MM/DD format
 * @param {Date} rawDate The date to be formatted
 * @return the date as string in YYYY/MM/DD format
 */
function getFormattedDate(rawDate){
  var endDate = rawDate.getFullYear().toString()+"/"+(rawDate.getMonth()+101).toString().substr(1)+"/"+(rawDate.getDate()+100).toString().substr(1);  
  return endDate;
}


/**
 * Generates a string with tax amounts and names
 * @param {Array} bookingTaxes the booking
 * @return A strinc containing tax names followed by their amount
 */
function getTaxString(bookingTaxes){
  var taxString = " ";
  
  //get tax names from spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookingsSheet = ss.getSheetByName('Data import');
  var taxNames = bookingsSheet.getRange(1, 24, 1, bookingsSheet.getLastColumn()).getValues()[0];
  
  //merge to booking
  for(var i=0; i<bookingTaxes.length; i++){
    taxString = taxString + taxNames[i] +": "+ bookingTaxes[i].toFixed(2).toString() + " ";
  }
  return taxString;
}

/**
 * Gets a PDF formatted for a particular booking
 * @param {string} id The PDF document template's id
 * @param {Array} booking The booking from which to change information
 * @param {number} balance The balance left to be paid on the booking
 * @return The PDF document
 */
function getPDF(id, booking, balance){
  var ogiginalDoc = DocsList.getFileById(id);
  var docName = ogiginalDoc.getName();
  var copyId   = ogiginalDoc.makeCopy(docName+" for "+booking[8]).getId();
  var copyDoc  = DocumentApp.openById(copyId);
  var copyBody = copyDoc.getActiveSection();
  
  //replace booking info
  copyDoc.replaceText("keyarrival", getFormattedDate(booking[0]));
  copyDoc.replaceText("keydeparture", getFormattedDate(booking[1]));
  copyDoc.replaceText("keyproperty", booking[3]);
  copyDoc.replaceText("keytotal", booking[4].toFixed(2));
  copyDoc.replaceText("keydeposit", booking[4].toFixed(2)-balance);
  copyDoc.replaceText("keyguests",booking[6]);
  copyDoc.replaceText("keyref#", booking[7]);
  copyDoc.replaceText("keyname", booking[8]);
  copyDoc.replaceText("keybalance", balance);
  copyDoc.replaceText("keydays", ScriptProperties.getProperty('Balance due date'));
  var temp = copyDoc.getText();
  
  copyDoc.saveAndClose();
  var pdf = DocsList.getFileById(copyId).getAs("application/pdf"); 
  
  //put in garbage, could be archived if desired
  DocsList.getFileById(copyId).setTrashed(true);
  
  return pdf;
}


function googleOAuth_(name,scope) {
  var oAuthConfig = UrlFetchApp.addOAuthService(name);
  oAuthConfig.setRequestTokenUrl("https://www.google.com/accounts/OAuthGetRequestToken?scope="+scope);
  oAuthConfig.setAuthorizationUrl("https://www.google.com/accounts/OAuthAuthorizeToken");
  oAuthConfig.setAccessTokenUrl("https://www.google.com/accounts/OAuthGetAccessToken");
  oAuthConfig.setConsumerKey('anonymous');
  oAuthConfig.setConsumerSecret('anonymous');
  return {oAuthServiceName:name, oAuthUseToken:"always"};
}

/**
 * Uses the daily stats to go black out rented dates in the availabilities
 */
function setAvailability() {
 var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  //get the days booked
  var dailyStatsSheet = ss.getSheetByName('Daily stats');
  var dailyStats = dailyStatsSheet.getDataRange().getValues();
  dailyStats.shift();
  dailyStats.shift();
  
  //get the rates page
  var ratesSheet = ss.getSheetByName('Rates');
  var ratesDates = ratesSheet.getRange(2, 2, ratesSheet.getLastRow()-1,1).getValues();
  var colorTable = getColorTable(getStatsTable(dailyStats, ratesDates));
  
  //set background **not sure about ratesSheet.getLastColumn()-1 vs. -2
  ratesSheet.getRange(2,3,ratesSheet.getLastRow()-1, ratesSheet.getLastColumn()-2).setBackgroundColors(colorTable);
}


/**
 * Returns cell format array for the rates sheet
 * @param {Array} statsTable the dailyStats with proper fitting dates for rates
 * @return a table of ratesTable size with blacked out cells for booked days
 */
function getColorTable(statsTable){
  var colorTable = new Array();
  
  //go through stats, pushing new info to colorTable
  while(statsTable.length>0){
    var daysStats = statsTable.shift();
    var daysColors = new Array();
    //go through the day and get colors
    for(i=2; i<daysStats.length; i=i+4){
      if(daysStats[i]===0){daysColors.push("grey");}
      else{
        if(daysStats[i]>0){daysColors.push("black");}
        else{daysColors.push("white")}
      }
    }
    colorTable.push(daysColors);
  }
  return colorTable;
}


/** 
 * Returns all dailystats that fit within the rates table dates
 * @param {Array} dailyStats All daily Stats
 * @param {Array} ratesDates All dates for which there are rates
 * @return the daily statsTable extended or tructated to fit with ratesDates
 */
function getStatsTable(dailyStats, ratesDates){
  
//---set the first day in the dailyStats
  //if the dailyStats first day is smaller than the rate's iterate on dailyStats
  if(dailyStats[0][0]<ratesDates[0][0]){    
    while(dailyStats[0][0]­<ratesDates[0][0]){dailyStats.shift();}
  }//otherwise iterate to reach ratesDates
  else{
    while(dailyStats[0][0]­>ratesDates[0][0]){
      dailyStats.unshift(new Array(dailyStats[0].length));
      dailyStats[0][0] = new Date(dailyStats[1][0]);  
      dailyStats[0][0].setDate(dailyStats[0][0].getDate()-1);
    }
  }

//---set the last day of dailyStats  
  //if the dailyStats last day is larger than the rate's iterate on dailyStats
  if(dailyStats[dailyStats.length-1][0]>ratesDates[ratesDates.length-1][0]){    
    while(dailyStats[dailyStats.length-1][0]­>ratesDates[ratesDates.length-1][0]){dailyStats.pop();}
    
  }//otherwise iterate to reach ratesDates
  else{
    var dayToAdd = new Date(dailyStats[dailyStats.length-1]);
    while(dayToAdd<ratesDates[ratesDates.length-1][0]){
      dayToAdd.setDate(dayToAdd.getDate()+1);
      dailyStats.push(makeArray(dailyStats[0].length, null, new Date(dayToAdd)));
    }
  }
  return dailyStats;
}

/**
 * Returns stats of each night in an array
 * @param {Array} allBookings A 2 dimentional array with all bookings
 * @return {Array} Table of basic daily statistics
 */
function dailyStats(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allBookings = ss.getSheetByName('Data import').getDataRange().getValues();
  allBookings.shift();
  for(;!allBookings[allBookings.length-1][0];allBookings.pop()){}
  allBookings.sort(function(a,b){return a[0]-b[0]});
  
  //get properties from bookings and add empty field
  var properties = removeDoubles(getColumn(allBookings, 3));
  var statTypes = new Array("Name", "R/N", "Res.", "Size");
  //generate the empty array with dates and properties
  var firstArrival = allBookings[0][0];
  var lastDeparture = allBookings[allBookings.length-1][1];
  var allDays = getDays(firstArrival, lastDeparture, properties.length*statTypes.length+1);
  allDays = getTitles(properties, statTypes).concat(allDays);
  var valuePerNight = 0;
  var timeStamp = new Date();
  var clientName = '';
  var groupSize = 1;
  var arrivalDay = new Date();
  var departureDay = new Date();
  var oneDay = 1000*60*60*24;

  //loop goes through bookings
  for(var i=0; i<allBookings.length; i++){
 
    //grab information for booking and calculate per night value
    clientName = allBookings[i][8];
    arrivalDay = allBookings[i][0].setHours(0,0,0,0);
    departureDay = allBookings[i][1].setHours(0,0,0,0);
    valuePerNight = allBookings[i][22]/((departureDay-arrivalDay)/oneDay);
    timeStamp = allBookings[i][2];
    groupSize = allBookings[i][6];
    
    //find the right property
    for(var j=1; j<allDays[0].length;j++){
      if(allBookings[i][3]==allDays[0][j]){
        
        //add valuePerNight for each date
        for(var k=1; k<allDays.length-1 && allDays[k+1][0]<=departureDay ;k++){
          if(allDays[k][0]>=arrivalDay){
            allDays[k][j] = clientName;
            allDays[k][j+1] = valuePerNight;
            allDays[k][j+2] = timeStamp;
            allDays[k][j+3] = groupSize;
          }
        }
        //you have found the right property break
        break;        
      }
    } 
  }
  //push to sheet
  ss.getSheetByName("Daily stats").getRange(1, 1, allDays.length, allDays[0].length).setValues(allDays);

}


/**
 * Returns a column from a 2d array
 * @param {Array} twodArray A 2 dimentional array
 * @param {number} column The column to be returned starts at 0
 * @return {Array} a one dimentional array containing the values in the column
 */
function getColumn(twodArray, column){
  var onedArray = new Array();
  for(var i=0; i<twodArray.length; i++){
    onedArray.push(twodArray[i][column]);
  }
  return onedArray;
}


/**
 * Makes an array with all date in first column and empty fields 
 * @param {Date} firstArrival Is the beginning day of our first booking
 * @param {Date} lastDeparture Is the last day of our last booking
 * @param {number} columns The number of empty columns after the date
 * @return {Array} Array of days between first arrival and last departure
 */
function getDays(firstArrival, lastDeparture, columns){
  var allDays = new Array();
  var firstOfMonth = new Date(firstArrival.getFullYear(), firstArrival.getMonth(), 1);
  var lastOfMonth = new Date(lastDeparture.getFullYear(), lastDeparture.getMonth()+1, 0);
  
  //empty cells to be added to date to create full table
  var emptyCells = new Array();
  for(var i=1; i<columns; i++){
    emptyCells.push(null); 
  }
  
  //loop that adds each day
  for(var dayToAdd = new Date(firstOfMonth); dayToAdd<=lastOfMonth ; dayToAdd.setDate(dayToAdd.getDate()+1)){
    var temp = new Array(new Date(dayToAdd));
    temp = temp.concat(emptyCells);
    allDays.push(temp);
  }
  return allDays;
}


/**
 * Generates the 2 title rows for  stats
 * @param {Array.<string>} properties An an array with all properties
 * @param  statTypes {Array.<string>} the different stats types for each property
 * @return Array of tiles on 2 rows
 */
function getTitles(properties, statsTypes){
  var titlesArray = [[''],['']];
  //loop that adds stat types to each property
  for(var i=0; i<properties.length; i++){
    titlesArray[1] = titlesArray[1].concat(statsTypes);
    for(var j=0; j<statsTypes.length; j++){
      titlesArray[0].push(properties[i]);
    }
  }
  return titlesArray;
}


/**
 * Builds basic monthly stats
 * @param {Array} dayStats: a 2 dimentional array containing nightly statistics
 * @return Array of basic monthly statistics including nights booked and revenue
 */
function monthltyStats(dayStats){
  var statTitles = ["Rev.", "Nights", "Occ.", "N/P"];
  //get properties
  var properties = new Array();
  for(var i=1; i<dayStats[0].length; i=i+statTitles.length){properties.push(dayStats[0][i]);}
  properties.push("Total");
  
  //get header for monthStats
  var allMonths = getTitles(properties, statTitles);
  dayStats.shift();  dayStats.shift();
  var zerosArray = [];
  for(var i=1;i<allMonths[0].length;i++){zerosArray.push(0);}
  
  //remove emptycells
  while(!dayStats[dayStats.length-1][0]){dayStats.pop();}
      
  //loop months
  while(dayStats.length>0){
    var month = dayStats[0][0];
    allMonths.push([month].concat(zerosArray));
    
  
    //loop days
    while(dayStats.length>0 && month.getMonth()==dayStats[0][0].getMonth()){
      var monthLength=32-new Date(month.getFullYear(),month.getMonth(),32).getDate();
      
      //loop properties
      for(var i=1; i<dayStats[0].length; i=i+4){
        if(dayStats[0][i+1]>0){
          allMonths[allMonths.length-1][i] += dayStats[0][i+1];
          allMonths[allMonths.length-1][i+1]++;
          allMonths[allMonths.length-1][i+3] += Number(dayStats[0][i+3]);
        }
      }
      dayStats.shift() 
    }
    //put occ and totals
    for(var i=1; i<allMonths[0].length-4; i=i+4){
      allMonths[allMonths.length-1][allMonths[0].length-4] += allMonths[allMonths.length-1][i];
      allMonths[allMonths.length-1][allMonths[0].length-3] += allMonths[allMonths.length-1][i+1];
      allMonths[allMonths.length-1][allMonths[0].length-1] += allMonths[allMonths.length-1][i+3];
      allMonths[allMonths.length-1][i+2] = allMonths[allMonths.length-1][i+1]/monthLength;
    }
    allMonths[allMonths.length-1][allMonths[0].length-2] = allMonths[allMonths.length-1][allMonths[0].length-3]/(monthLength*(properties.length-1));
  }
  //super total
  var total = ['total'].concat(zerosArray) ;
  for(var i=2; i<allMonths.length;i++){
    for(var j=1; j<total.length-1; j+=4){
      total[j] = total[j] + allMonths[i][j];
      total[j+1] = total[j+1] + allMonths[i][j+1];
      total[j+2] = total[j+2] + allMonths[i][j+2]/(allMonths.length-2);
      total[j+3] = total[j+3] + allMonths[i][j+3];
    }
  }
  allMonths.push(total);
  
  return allMonths;
}


/**
 * returns the booking progress for a given period
 * @param {Date} periodStart the beginning date of the period to analyze
 * @param {Date} periodEnd the end date of the period to analyze
 * @param {Array} dayStats the Daystats Array
 * @return an array containing the progress of bookings for a given period through time
 */
function bookingForecast(periodStart, periodEnd, dayStats){

  //if the dates have no stats exit
  if(periodStart<dayStats[0][0]||periodEnd>dayStats[dayStats.length-1][0]){return false;}
  
  //make array to put stats into
  var forecast = new Array(49);
  forecast[0]=periodStart.getFullYear();
  var thirtyDays = 1000*60*60*24*30;
  for(var i=1; i<forecast.length; i++){
    forecast[i]=0;
  }

  //loop goes through the period (this should be with a while and not have the push after. FIX
  var currentDay=dayStats[dayStats.length-1];
  for(;currentDay[0]>=periodStart; currentDay=dayStats.pop()){
    //if we haven't reached the right day yet pop again
    if(currentDay[0]>=periodEnd){continue;}
    
    //if the dates fit look through houses
    for(var i=2; i<currentDay.length ; i=i+4){
      if(currentDay[i]>0){
      
        //calculate how long in advance we are
        var advance = (periodEnd-currentDay[i+1])/thirtyDays;
        
        if(advance>=11){
          advance=11;
        }else{
          advance=Math.floor(advance);
        }
        var daysCell = (11-advance)*4 + 1;
        forecast[daysCell]=forecast[daysCell]+1;
        forecast[daysCell+2]=forecast[daysCell+2]+currentDay[i];
      } 
    }
  }

  //push the last day back on because it wasn't used
  dayStats.push(currentDay);
  
  
  //loop adds cumulative stats
  forecast[2]=forecast[1];
  forecast[4]=forecast[3];
  for(var i=5; i<forecast.length; i=i+4){
    forecast[i+1]=forecast[i-3]+forecast[i];
    forecast[i+3]=forecast[i-1]+forecast[i+2];
    }
  //get last years forecast (recursive) and add it  
  var forecasts = bookingForecast(new Date(periodStart.setFullYear(periodStart.getFullYear()-1)), new Date(periodEnd.setFullYear(periodEnd.getFullYear()-1)), dayStats);
  if(forecasts!=false){
    forecasts.push(forecast);
  }else{
    forecasts = new Array();
    forecasts.push(forecast);
  }

  return forecasts;
}

/**
 * merges all bookings on a specific fields keeping most recent information, addind revenues
 * @param {Array} sortedList All the bookings
 * @param {number} column the column on which to merge identical information
 * return an array of bookings merged on the identical information of column
 */
function clientList(sortedList, column){
  //take out empties
  while(!sortedList[sortedList.length-1][0]){sortedList.pop();}
  
  var clientList = new Array();
  var currentClient = sortedList.pop();
  var tempClient;

//goes through bookings and merges same clients bookings
  while(sortedList.length > 0){
    tempClient = sortedList.pop();
    
    //same client merge info
    if(currentClient[column-1].toLowerCase() == tempClient[column-1].toLowerCase()){
      currentClient[4] = Number(currentClient[4])+Number(tempClient[4]);
      currentClient[6] = Number(currentClient[6])+Number(tempClient[6]);
      for(var i = 8;i<currentClient.length;i++){
        if(tempClient[i]){currentClient[i]=tempClient[i];}
      } 
    }else{
      //different clients, add the last one and move to next
      clientList.push(currentClient);
      currentClient = tempClient;
    }
  }
  
  clientList.push(currentClient);
  return clientList;
}


/**
 * returns unique values from an array
 * @param {Array} doublesArray An array that might have doubles
 * @return array sorted and without doubles
 */
function removeDoubles(doublesArray){
  var uniquesArray = new Array();
  doublesArray.sort();
  uniquesArray.push(doublesArray[0]);

  for(var i=1; i<doublesArray.length; i++){
    if(doublesArray[i]!=doublesArray[i-1]){
     uniquesArray.push(doublesArray[i]); 
    }
  }
  if(uniquesArray[0]==""){uniquesArray.shift();}
  return uniquesArray;
}


/*function: Grabs bookings from spreadsheet, sends emails containing upcoming bookings*/
function sendEmployeeReport(bookings, startDate, endDate) {
  //get all bookings that leave after startDate and that arrive before endDate
  bookings.sort(function(a,b){return a[0]-b[0]});
  while(bookings[bookings.length-1][0]>endDate){bookings.pop();}
  bookings.sort(function(a,b){return a[1]-b[1]});
  while(bookings[0][1]<startDate){bookings.shift();}

  //build the report for all 14 days
  var reportArray = new Array();
  while(startDate<=endDate){
    reportArray.push([new Date(startDate), '', '', '', '', '']);
    
    var daysEvents = new Array();
    for(var i=0; i<bookings.length;i++){
      if(bookings[i][1].getTime()==startDate.getTime()){
          daysEvents.push(['',"Departure", bookings[i][3], bookings[i][8], bookings[i][6], bookings[i][4]-bookings[i][5]]);
      }else{
        if(bookings[i][0].getTime()==startDate.getTime()){
        daysEvents.push(['',"Arrival", bookings[i][3], bookings[i][8], bookings[i][6], bookings[i][4]-bookings[i][5]]);
        }
      }
    }
    daysEvents.sort(function(a,b){return (a[2] < b[2] ? -1 : (a[2] > b[2] ? 1 : 0));});
    while(daysEvents.length>0){reportArray.push(daysEvents.shift());}
    startDate.setDate(startDate.getDate()+1);
  }
  return reportArray;
}

/**
 * Updates or adds charts to the montly stats sheet
 */
function updateAllCharts(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monthlyStatsSheet = ss.getSheetByName('Monthly stats');
  var lastColumn = monthlyStatsSheet.getLastColumn();
  var lastRow = monthlyStatsSheet.getLastRow() -1;
  var charts = monthlyStatsSheet.getCharts();
  var revChart;
  var occChart;

  //find if the charts exist
  while(charts.length>0){
    if(charts[0].getOptions().get('title')=='Monthly Revenue'){
      revChart = charts.shift().modify();
    }else{ 
      if(charts[0].getOptions().get('title')=='Monthly Occupancy Rate'){
        occChart = charts.shift().modify();
      }else{charts.shift()}//just in case some other charts were added
    }
  }
  
  //if revenue doesn't exist make it
  if(!revChart){
    revChart = monthlyStatsSheet.newChart();
    revChart.setChartType(Charts.ChartType.AREA).setOption('title', 'Monthly Revenue')
    .setOption('isStacked',true)
    .setOption('width', 1000)
    .setOption('height', 600)
    .setPosition(5, 3, 0, 0);
  }
  
  //if occupancy doesn't exist make it
  if(!occChart){
    occChart = monthlyStatsSheet.newChart();
    occChart.setChartType(Charts.ChartType.LINE).setOption('title', 'Monthly Occupancy Rate')
    .setOption('isStacked',true)
    .setOption('width', 1000)
    .setOption('height', 600)
    .setPosition(5, 5, 0, 0);
  }
  
  //remove ranges
  var ranges = revChart.getRanges();
  while(ranges.length>0){revChart=revChart.removeRange(ranges.pop());}
  ranges = occChart.getRanges();
  while(ranges.length>0){occChart=occChart.removeRange(ranges.pop());}
  
  //add ranges
  revChart.addRange(monthlyStatsSheet.getRange(1, 1, lastRow, 1));
  occChart.addRange(monthlyStatsSheet.getRange(1, 1, lastRow, 1));
  for(var i=2; i<lastColumn-4; i+=4){
    revChart = revChart.addRange(monthlyStatsSheet.getRange(1, i, lastRow, 1));
    occChart = occChart.addRange(monthlyStatsSheet.getRange(1, i+2, lastRow, 1));
  }
  
  try{monthlyStatsSheet.insertChart(revChart.build());
      monthlyStatsSheet.insertChart(occChart.build());}
  catch(err){
    throw('Google is picky when it feels charts are to close to the side. Move or delete them or make sheet larger and run again.')
  }
}
