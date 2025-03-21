const dataCol = 1 ;                     // Column containing unique user IDs - 1 (A) using default layout 
const startRow = 7;                     // Row number containing the first ID - 7 using default layout
const listSheetName = "Create-Share";   // Name of worksheet tab containing user data
//const templateID = "1OfOsEODmX1S5EAoBY26DATy0WZKNRXXLyi3-2bQ4kms" 
var functionNumber = 11;
//var ui = SpreadsheetApp.getUi();

function getNumRows(listSheet) {

  var IDlist = listSheet.getRange(startRow,dataCol).getDataRegion(SpreadsheetApp.Dimension.ROWS).getA1Notation();
  var numRows = listSheet.getRange(IDlist).getLastRow() - (startRow - 1);
  return(numRows)

}



function createTemplates(functionNumber) {
  // comment out the next line to run the script!
  //return; // do not execute... files have been created for current semester
  
  

  var list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheetName);
  var numRows = getNumRows(list);
  var templateURL=list.getRange("E1").getValue();
  Logger.log(templateURL);
  var ss = SpreadsheetApp.openByUrl(templateURL); 

  // Fetch the range of cells
  var dataRange = list.getRange(startRow, 1, numRows, 4)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var netId = row[0];  // First column
    var email = row[1];
    var processFlag = row[2];
    var newFile = row[3];
    Logger.log(row);

    if(processFlag) {
      switch(functionNumber) {
        case 1: // ===============  Create Files  =============================
          // create template copy with name newFile as ss2   
          var ss2 = ss.copy(newFile);
          // create [link] and url in column G 
          list.getRange(startRow + i, 7).setValue('=HYPERLINK("'+ss2.getUrl()+'","[Link]")');

          // add user as editor.  Does this automatically generate an email?
          ss2.addEditor(email);
  

          // access the first sheet in the new file  
          var sheet = ss2.getSheets()[0];

          // write data to specified locations
          var cell = sheet.getRange("B2");
          cell.setValue(email);
          Logger.log(newFile, email);
          break;


        case 2:
          Logger.log("Boo!")
          break;
        case 3: // remove sharing permission from editors
          var ss2 = DriveApp.getFilesByName(newFile) ;
          while (ss2.hasNext()) {
            var temp=ss2.next();
            //temp.removeEditor(email);
            //temp.addEditor(email);
            temp.setShareableByEditors(false);
            Logger.log(temp.getEditors());
          }
          Logger.log("Boo!")
          break;
        case 4: // add editor access
          var ss2 = DriveApp.getFilesByName(newFile) ;
          while (ss2.hasNext()) {
            var temp=ss2.next();
            temp.removeEditor(email);
            //temp.addEditor(email);
            Logger.log(temp.getEditors());
          }
          break;

        case 5: // remove editor access
          var ss2 = DriveApp.getFilesByName(newFile) ;
          while (ss2.hasNext()) {
            var temp=ss2.next();
            temp.removeEditor(email);
            //temp.addEditor(email);
            Logger.log(temp.getEditors());
          }
          break;

        case 6: //insert sheet
          const sheetName = "Risk Analysis"; //should be in the template file identified on the worksheet
          var source = ss.getSheetByName(sheetName);       
          var ss2 = DriveApp.getFilesByName(newFile) ;
          while (ss2.hasNext()) {
            var target = SpreadsheetApp.open(ss2.next());
            source.copyTo(target).setName(sheetName);  
          }
          break;

        case 7: // add viewer access
          var ss2 = DriveApp.getFilesByName(newFile) ;
          while (ss2.hasNext()) {
            var temp=ss2.next();
            temp.removeEditor(email);
            //temp.addViewer(email);
            Logger.log(temp.getEditors());
            Logger.log(email)
          }
          break;

        case 10:  // retrieve grade
          var ss2 = DriveApp.getFilesByName(newFile) ;
          while (ss2.hasNext()) {
          var spreadsheet = SpreadsheetApp.open(ss2.next());
          var sheet = spreadsheet.getSheetByName("Overview");
          var grade1 = sheet.getRange("F2").getValue();
          Logger.log(grade1); 
          list.getRange(i+startRow, 8).setValue(grade1);
          }
          break;

        case 11:  // get last update and share status
          var ss2 = DriveApp.getFilesByName(newFile) ;
          while (ss2.hasNext()) {
          var item=ss2.next(); 
          list.getRange(startRow + i, 5).setValue(item.getLastUpdated());
          list.getRange(startRow + i, 6).setValue(item.getAccess(netId+'@plattsburgh.edu')); 
          //list.getRange(startRow + i, 7).setValue('=HYPERLINK("'+item.getUrl()+'","[LINK]")');
          list.getRange(startRow + i, 7).setValue('=HYPERLINK("'+item.getUrl()+'#gid=1369031038'+'","[LINK]")');
          //list.getRange(startRow + i, 7).setValue('=HYPERLINK("'+item.getUrl()+'#gid=1796068497'+'","[LINK]")');
          }
          break;


        default:
        Logger.log("Invalid function number "+functionNumber);

      }
    }


  }


  // Make sure the cell is updated right away in case the script is interrupted
  SpreadsheetApp.flush();
  return;  
};


function getGrades() {
createTemplates(10);
return;
}

function getLastEdit() {
  createTemplates(11);
  return;
}

function createFiles() {
  createTemplates(1);
  return;
}

function stopSharing() {
  createTemplates(3);
  return;
}

function insertSheet() {
  createTemplates(6);
  return;
}

function addViewAccess() {
  createTemplates(7);
  return;
}
function addEditAccess() {
  createTemplates(4);
  return;
}
