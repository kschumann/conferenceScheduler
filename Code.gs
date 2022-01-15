function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  try{
    var ui = SpreadsheetApp.getUi();
    ui.createAddonMenu()
    .addItem('Setup Conference Host Schedule', 'openSettings')  
    .addItem('Create Conference Attendee Schedule', 'translateSchedule')
    .addToUi();  
  } catch(e){
  }
}

function openSettings(){
  try{
  var html = HtmlService.createHtmlOutputFromFile('Settings')
      .setTitle('Set Up Conference Host Schedule')
      .setWidth(500).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Set Up Conference Schedule');// Or DocumentApp or SlidesApp or FormApp.
      //.showSidebar(html);
  } catch(e){
    console.error("Open Settings: Could not open settings - " + e);
  }
}

function translateSchedule() {
   try{ 
     //Prep Variables     
     var ss = SpreadsheetApp.getActive();
     var ssId = ss.getId();
     var hostSheet = ss.getSheetByName('HostSchedule');
     var attendeeSheet = ss.getSheetByName('AttendeeSchedule');
     var hostArray = hostSheet.getRange(1, 1, hostSheet.getLastRow(), hostSheet.getLastColumn()).getValues();
     var attendees = [];
     var attendeeSchedules = [[]];
     
     //Create list of attendees and meetings that each family has
     for (var i = 1; i<hostArray.length; i++){
       for (var j = 1; j<hostArray[i].length; j++){
         if(!(hostArray[i][j] == "" || hostArray[i][j] == null || hostArray[i][j] == " ")){
           var index = attendees.indexOf(hostArray[i][j]);
           var hours = hostArray[i][0].getHours()<13 ? hostArray[i][0].getHours() : hostArray[i][0].getHours()-12;
           var minutes = hostArray[i][0].getMinutes() == 0 ? "00" : hostArray[i][0].getMinutes();
           var period = hostArray[i][0].getHours()<12 ? " AM" : " PM";        
           var meetingTime = hours + ":" + minutes + period;
           if(index == -1){
             attendees.push(hostArray[i][j]);
             attendeeSchedules.push([hostArray[i][j],meetingTime + ", " + hostArray[0][j]]);
           } 
           else{
             attendeeSchedules[index+1].push(meetingTime + ", " + hostArray[0][j]);    
           }
         }      
       }
     }
     
     //If schedules are not set up, throw alert and end script  
     if(attendeeSchedules.length == 0){
       SpreadsheetApp.getUi().alert("It looks like the Conference Host schedule has not been set up.  Please complete Host Setup and then try again.");
       return;
     }   
     
     //Add Attendee Schedule list to Attendee Sheet, replacing whatever was there before.
     attendeeSheet.clear();
     var attendeeHeader = [""];
     for(var i=0; i<hostSheet.getLastColumn()-1;i++){//Create new list of meetings for header
       var mtgNumb = i+1;
       attendeeHeader.push("Session " + mtgNumb);
     }
     
     attendeeSheet.getRange(1, 1, 1, hostSheet.getLastColumn()).setValues([attendeeHeader]);
     for(var k = 1; k<attendeeSchedules.length;k++){
       attendeeSheet.getRange(k+1,1,1,attendeeSchedules[k].length ).setValues([attendeeSchedules[k]]);   
       if(k % 2 == 0){
         attendeeSheet.getRange(k+1,1,1,attendeeHeader.length).setBackground('#efefef');
       } 
     }
     //Apply Formatting to Attendee Sheet
     attendeeSheet.getRange(1, 1, 1, attendeeHeader.length+1).setFontWeight('900').setBackground('#000000').setFontColor('#ffffff');
     attendeeSheet.getRange(2,1,attendeeSchedules.length,1).setFontWeight('900').setBackground('#d9d9d9');
     attendeeSheet.getRange(1,1).setBackground('#ffffff');
     attendeeSheet.setRowHeights(2, attendees.length, 30);
     attendeeSheet.setColumnWidths(2, attendeeHeader.length-1, 150);
     
     //Trim excess Rows and Columns     
     var rowsInSheet = attendeeSheet.getMaxRows();
     var firstRowDeleted = attendees.length+2;
     var rowsDeleted = rowsInSheet - firstRowDeleted+1;
     if(rowsDeleted>0){
       attendeeSheet.deleteRows(firstRowDeleted, rowsDeleted);   
     }
     var columnsInSheet = attendeeSheet.getMaxColumns();
     var firstColumnDeleted = attendeeHeader.length+1;
     var columnsDeleted = columnsInSheet - firstColumnDeleted +1;  
     if(columnsDeleted >0){
       attendeeSheet.deleteColumns(firstColumnDeleted, columnsDeleted);   
     } 
     
     //Set the Active Sheet to Attendee
    ss.setActiveSheet(attendeeSheet);  
     console.info("Schedule Translated. Id: " + ssId);
  } catch(e){
    console.error("translateSchedule(). Host Schedule Not Set Up. Error: " + e);
     SpreadsheetApp.getUi().alert("It looks like the host schedule has not been set up.  Please complete host setup and then try again.");
    }
}


function setupAttendeeSchedule(numTeachers,start,end,interval){
  try{
    console.info("Initiating Setup");
    //Setup variables
    numTeachers = parseInt(numTeachers);
    start = parseInt(start);
    end = parseInt(end);
    interval = parseInt(interval);
    var ss = SpreadsheetApp.getActive();
    var ssId = ss.getId();
    console.info("Getting Sheets");    
    //Create Required Sheets and switch to Host sheet to be active
    if(ss.getSheetByName('HostSchedule')){
      var teacher =  ss.getSheetByName('HostSchedule').clear();
    } else {
      var teacher = ss.insertSheet('HostSchedule'); 
    }
    if(ss.getSheetByName('AttendeeSchedule')){
      ss.getSheetByName('AttendeeSchedule').clear();
    } else {
      var parent = ss.insertSheet('AttendeeSchedule');
    }  
    ss.setActiveSheet(teacher);  
    console.info("Creating headers and Times");    
    //Create and format headers and times
    var header = generateHeaders(numTeachers);
    teacher.getRange(1, 1, 1, numTeachers+1).setValues(header).setFontWeight('900').setBackground('#000000').setFontColor('#ffffff');
    var times = generateTimes(start,end,interval);
    teacher.getRange(2,1,times.length,1).setValues(times).setFontWeight('900').setBackground('#d9d9d9');
    teacher.getRange(1,1).setBackground('#ffffff');
    teacher.setRowHeights(2, times.length, 60);
    teacher.setColumnWidths(2, header[0].length-1, 150);
    
    //Trim excess rows and columns, if any
    var rowsInSheet = teacher.getMaxRows();
    var firstRowDeleted = times.length+2;
    var rowsDeleted = rowsInSheet - firstRowDeleted+1;
    if(rowsDeleted>0){
      teacher.deleteRows(firstRowDeleted, rowsDeleted);   
    }
    var columnsInSheet = teacher.getMaxColumns();
    var firstColumnDeleted = header[0].length+1;
    var columnsDeleted = columnsInSheet - firstColumnDeleted +1;  
    if(columnsDeleted >0){
      teacher.deleteColumns(firstColumnDeleted, columnsDeleted);   
    }
    console.info("Attendee Schedule Setup. ID: " + ssId);
  } catch(e) {
        console.error("setupAttendeeSchedule(). Attendee Schedule Not Set Up. Error: " + e);
        SpreadsheetApp.getUi().alert("It looks like something went wrong while setting up the Conference Host Schedule template.  Try removing any existing sheets and running setup agian.");
  }
}


function checkForExisting(){
  var ss = SpreadsheetApp.getActive();
  var hostSheet = ss.getSheetByName('HostSchedule');
  if(hostSheet){
    var hostValue = hostSheet.getRange(1,2).getValue();
    var existing = hostValue ? true : false;
  } else {
    var existing = false;
  } 
  Logger.log(existing);
  return existing;
}


function generateHeaders(numTeachers){
  var headerArray = [['']];
  for(var i=0; i<numTeachers; i++){
    headerArray[0].push('Host'+(i+1));
  }
 return headerArray;
}


function generateTimes(start, end, interval){
  var timeArray = [];
  var intervalSize = 60/interval;
  for (var i=start; i<end; i++){
    for(var j=0; j<intervalSize; j++){
      var hours = i<13 ? i : i-12;
      var minutes = j*interval == 0 ? "00" : j*interval;
      var period = i<12 ? "AM" : "PM";
      timeArray.push([hours + ":" + minutes + " " + period])
    }
  }
  return timeArray;
}


