/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Reports', 'showSidebar')
      //.addItem('Analyze Class', 'analyzeClass')
      .addToUi();
}
/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('ToM Reports');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function scanData() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  if (sheets.length == 1) {
    var dataSheet = sheets[0];
    dataSheet.setName('studentResults');
    var numStudents = dataSheet.getLastRow()-4;
    var colD = dataSheet.getRange(4,4).getValue();
    dataSheet.getRange("P5:Q").setNumberFormat("@");
    var reportType;
    if (colD == "Class") {
      var obj = {
        type : "Teacher",
        students : numStudents
      };
    } else if (colD == "Teacher") {
      var obj = {
        type : "Building",
        students : numStudents
      };
    } else if (colD == "School") {
      var obj = {
        type : "District",
        students : numStudents
      };
    }
    return obj;
  } else {
    var obj = false;
    return obj;
  }
}

function teacherAdjust() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("studentResults");
  sheet.insertColumnsAfter(3, 2);
  sheet.getRange(4, 4).setValue("School");
  sheet.getRange(4, 5).setValue("Teacher");
  sheet.getRange("N5:O").setNumberFormat('@');
}
function buildingAdjust() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("studentResults");
  sheet.insertColumnsAfter(3, 1);
  sheet.getRange(4, 4).setValue("School");
  sheet.getRange("N5:O").setNumberFormat('@');
}

function districtAdjust() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("studentResults");
  sheet.getRange("N5:O").setNumberFormat('@');
}

function makeSheets(teacher) {
  //generate sheets for every student
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var dataSheet = sheets[0];
  var numStudents = dataSheet.getLastRow()-4;
  if (teacher) {
    dataSheet.getRange(5,5,numStudents).setValue(teacher);
  }
  var studentIDs = [];
  var nameDataRange = dataSheet.getRange(5, 1, numStudents, 3);
  var nameData = nameDataRange.getValues();
  for (i = 0; i < nameData.length; i++) {
    ss.insertSheet(nameData[i][0] + ', ' + nameData[i][1], i+1);
    var currentSheet = ss.getSheetByName(nameData[i][0] + ', ' + nameData[i][1]).setColumnWidths(1, 10, 87);
    var page2rows = [54,58,62,67,71,75,79,83];
    for (j in page2rows) {
        var currentHeight = currentSheet.getRowHeight(page2rows[j]);
        currentSheet.setRowHeight(page2rows[j], currentHeight-5);
      }
    currentSheet.setRowHeight(50, 78).deleteColumns(11, 16);
    studentIDs.push(nameData[i][2]);
    currentSheet.getRange("B4").setValue(nameData[i][2]).setFontColor("#ffffff");
  }
  Logger.log("function makeSheets completed");
}

function buildEmptyReport(count,grade) {
  Logger.log("function buildEmptyReport started");
  var ss = SpreadsheetApp.getActive();
  var template = SpreadsheetApp.openById("ENTER YOUR TEMPLATE FILE ID HERE");
  var rangeSheet = template.getSheetByName("Ranges");
  var rangesToMerge = rangeSheet.getRange(2,1,116).getValues();
  var rangesToBorder = rangeSheet.getRange(2,2,60).getValues();
  var rangesToHead = rangeSheet.getRange(2,3,24).getValues();
  var grade3meeting = rangeSheet.getRange(2,4,20).getValues();
  var grade4meeting = rangeSheet.getRange(2,5,20).getValues();
  var sheets = ss.getSheets();
  var dataSheet = sheets[0];
  var sheet  = sheets[count+1];
  if (sheet.getSheetName() != "studentResults" && sheet.getSheetName() != "DemoReport") {
    var currentID = sheet.getRange("B4").getValue();
    var currentRow = count + 5;
    var requestedLanguage = dataSheet.getRange(currentRow, 29).getValue();
    var labelSheet = template.getSheetByName(requestedLanguage);
    var labels = labelSheet.getRange(2,1,116).getValues();
    var ELAlabels = labelSheet.getRange(118,1,172).getValues();
    var headings = labelSheet.getRange(2,2,24).getValues();
    for (j in rangesToMerge) {
      if (j != 0) { // Sets labels for merged cells other than the title #see below#
        var currentRange = rangesToMerge[j][0].split(",");
        sheet.getRange(currentRange[0],currentRange[1],currentRange[2],currentRange[3]).mergeAcross().setBorder(true, true, true, true, true, true).setValue(labels[j]).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
      } else { // Sets the Title
          var currentRange = rangesToMerge[j][0].split(",");
          sheet.getRange(currentRange[0],currentRange[1],currentRange[2],currentRange[3]).mergeAcross().setValue(labels[j]).setFontWeight("bold").setHorizontalAlignment('center');
      }
    }//end rangesToMerge Loop 
    for (j in rangesToBorder) {
      var currentRange = rangesToBorder[j][0].split(",");
      sheet.getRange(currentRange[0],currentRange[1]).setBorder(true, true, true, true, true, true).setValue(ELAlabels[j]).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
    for (j in rangesToHead) {
       var currentRange = rangesToHead[j][0].split(",");
      sheet.getRange(currentRange[0],currentRange[1]).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(false).setValue(headings[j]);
    }
    //Grade based highlighting
    if (grade == 3) {
      for (j in grade3meeting) {
        var currentRange = grade3meeting[j][0].split(",");
        sheet.getRange(currentRange[0],currentRange[1]).setBackground("#90EE90");
      }
    } else if (grade == 4) {
      for (j in grade4meeting) {
        var currentRange = grade4meeting[j][0].split(",");
        sheet.getRange(currentRange[0],currentRange[1]).setBackground("#90EE90");
      }  
    }
  }
  Logger.log("function buildEmptyReport completed");
  return count+1;
}

function finishReport(count) {
  var ss = SpreadsheetApp.getActive();
  var template = SpreadsheetApp.openById("1uf8wSzkYpPbh-Cx0ouFeJDPsKMb_5IJoMPVzH35fn4A");
  var linkit = template.getSheetByName('LinkIt');
  var sheets = ss.getSheets();
  var data = sheets[0];
  var sheet = sheets[count];
  if (sheet.getSheetName() != "studentResults" && sheet.getSheetName() != "DemoReport") {
    var currentRow = count+4;
    var name = sheet.getRange("A3").setValue(sheet.getSheetName());
    var teacher = data.getRange(currentRow,5,1,1).getValue();
    var testDate = sheet.getRange("G3").setValue(data.getRange(currentRow,7,1,1).getValue());
    var daysAbsent = sheet.getRange("B2").setHorizontalAlignment("center").setValue(data.getRange(currentRow, 28,1,1).getValue());
    sheet.getRange("D3").setValue(teacher);
    var socialEmotional = data.getRange(currentRow,8,1,6).getValues();
    for (i=0; i<socialEmotional[0].length; i++) {
      var seRow = [8,12,16,20,24,28];
      var column;
      switch(socialEmotional[0][i]) {
        case 0:
          column = 1;
          break;
        case 1:
          column = 4;
          break;
        case 2:
          column = 7;
          break;
      }
      sheet.getRange(seRow[i], column).setFormula("=char(10004)");
    }
    var writing = data.getRange(currentRow,14,1,1).getValue();
    var linkitWriting = linkit.getRange(113,1,8).getValues();
    var writingCol;
    switch(writing) {
      case linkitWriting[0][0]:
        writingCol = 1;
        break;
      case linkitWriting[1][0]:
        writingCol = 2;
        break;
      case linkitWriting[2][0]:
        writingCol = 3;
        break;
      case linkitWriting[3][0]:
        writingCol = 4;
        break;
      case linkitWriting[4][0]:
        writingCol = 5;
        break;
      case linkitWriting[5][0]:
        writingCol = 6;
        break;
      case linkitWriting[6][0]:
        writingCol = 7;
        break;
      case linkitWriting[7][0]:
        writingCol = 8;
        break;
      }
    
    sheet.getRange(33, writingCol).setFormula("=char(10004)");
    var verbal = data.getRange(currentRow,15,1,1).getValue();
    var linkitVerbal = linkit.getRange(129, 1, 7).getValues();
    var verbalCol;
    switch(verbal) {
      case linkitVerbal[0][0]:
        verbalCol = 2;
        break;
      case linkitVerbal[1][0]:
        verbalCol = 3;
        break;
      case linkitVerbal[2][0]:
        verbalCol = 4;
        break;
      case linkitVerbal[3][0]:
        verbalCol = 5;
        break;
      case linkitVerbal[4][0]:
        verbalCol = 6;
        break;
      case linkitVerbal[5][0]:
        verbalCol = 7;
        break;
      case linkitVerbal[6][0]:
        verbalCol = 8;
        break;
      }
    sheet.getRange(37, verbalCol).setFormula("=char(10004)");    
    var readLang = data.getRange(currentRow,16,1,3).getValues();
    var linkitRecog = linkit.getRange(143, 1,5).getValues();
    var linkitAcq = linkit.getRange(153, 1, 5).getValues();
    for (i=0; i<readLang[0].length; i++) {
      var rlRow = [41,45,49];
      var rlColumn;
      if(readLang[0][i] == linkitRecog[0][0] || readLang[0][i] == linkitAcq[0][0]) {
        rlColumn = 3;
        Logger.log(linkitAcq[0][0]);
      } else if (readLang[0][i] == linkitRecog[1][0] || readLang[0][i] == linkitAcq[1][0]){
        rlColumn = 4;
        Logger.log(linkitAcq[1][0]);
      } else if (readLang[0][i] == linkitRecog[2][0] || readLang[0][i] == linkitAcq[2][0]){
        rlColumn = 5;
        Logger.log(linkitAcq[2][0]);
      } else if (readLang[0][i] == linkitRecog[3][0] || readLang[0][i] == linkitAcq[3][0]){
        rlColumn = 6;
        Logger.log(linkitAcq[3][0]);
      } else if (readLang[0][i] == linkitRecog[4][0] || readLang[0][i] == linkitAcq[4][0]){
        rlColumn = 7;
        Logger.log(linkitAcq[4][0]);
      }
      sheet.getRange(rlRow[i], rlColumn).setFormula("=char(10004)");
    }
    var vListen = data.getRange(currentRow, 19, 1, 2).getValues();
    var linkitVoc = linkit.getRange(37, 1, 5).getValues();
    var linkitListen = linkit.getRange(47, 1, 5).getValues();
    for (i=0; i<vListen[0].length; i++) {
      var vlRow = [53,57];
      var vlColumn;
      if(vListen[0][i] == linkitVoc[0][0] || vListen[0][i] == linkitListen[0][0]) {
        vlColumn = 1;
      } else if (vListen[0][i] == linkitVoc[1][0] || vListen[0][i] == linkitListen[1][0]){
        vlColumn = 3;
      } else if (vListen[0][i] == linkitVoc[2][0] || vListen[0][i] == linkitListen[2][0]){
        vlColumn = 5;
      } else if (vListen[0][i] == linkitVoc[3][0] || vListen[0][i] == linkitListen[3][0]){
        vlColumn = 7;
      } else if (vListen[0][i] == linkitVoc[4][0] || vListen[0][i] == linkitListen[4][0]){
        vlColumn = 9;
      }
      sheet.getRange(vlRow[i], vlColumn).setFormula("=char(10004)");
    }
      var rhyming = data.getRange(currentRow,21,1,1).getValue();
      var linkitRhy = linkit.getRange(57, 1, 4).getValues();
      var rhyCol;
      if (rhyming == linkitRhy[0][0]) {
        rhyCol = 2;
      } else if (rhyming == linkitRhy[1][0]) {
        rhyCol = 4;
      } else if (rhyming == linkitRhy[2][0]) {
        rhyCol = 6;
      } else if (rhyming == linkitRhy[3][0]) {
        rhyCol = 8;
      }
      sheet.getRange(61,rhyCol).setFormula("=char(10004)");
    
    var mathResults = data.getRange(currentRow,22,1,6).getValues();
    var linkitCount = linkit.getRange(65, 1, 4).getValues();
    var linkitObj = linkit.getRange(73, 1, 4).getValues();
    var linkitNumer = linkit.getRange(81, 1, 4).getValues();
    var linkitShapes = linkit.getRange(89, 1, 4).getValues();
    var linkitClass = linkit.getRange(97, 1, 4).getValues();
    var linkitMeas = linkit.getRange(105, 1, 4).getValues();
    for (i=0; i<mathResults[0].length; i++) {
      var mathRows = [66,70,74,78,82,86];
      var mathCol;
      if(mathResults[0][i] == linkitCount[0][0] || mathResults[0][i] == linkitObj[0][0] || mathResults[0][i] == linkitNumer[0][0] || mathResults[0][i] == linkitShapes[0][0] || mathResults[0][i] == linkitClass[0][0] || mathResults[0][i] == linkitMeas[0][0]) {
        mathCol = 2;
      } else if(mathResults[0][i] == linkitCount[1][0] || mathResults[0][i] == linkitObj[1][0] || mathResults[0][i] == linkitNumer[1][0] || mathResults[0][i] == linkitShapes[1][0] || mathResults[0][i] == linkitClass[1][0] || mathResults[0][i] == linkitMeas[1][0]) {
        mathCol = 4;
      } else if(mathResults[0][i] == linkitCount[2][0] || mathResults[0][i] == linkitObj[2][0] || mathResults[0][i] == linkitNumer[2][0] || mathResults[0][i] == linkitShapes[2][0] || mathResults[0][i] == linkitClass[2][0] || mathResults[0][i] == linkitMeas[2][0]) {
        mathCol = 6;
      } else if(mathResults[0][i] == linkitCount[3][0] || mathResults[0][i] == linkitObj[3][0] || mathResults[0][i] == linkitNumer[3][0] || mathResults[0][i] == linkitShapes[3][0] || mathResults[0][i] == linkitClass[3][0] || mathResults[0][i] == linkitMeas[3][0]) {
        mathCol = 8;
      }
      sheet.getRange(mathRows[i], mathCol).setFormula("=char(10004)");
    }
  }
  return count;
}

function delSheets() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  for (i=1; i < sheets.length; i++) {
    ss.deleteSheet(sheets[i]);
  }
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('ToM Reports');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function buildAnalysis() {
  var ss = SpreadsheetApp.getActive();
  ss.insertSheet("Data Analysis", 1);
  var sheet = ss.getSheetByName("Data Analysis");
  sheet.getRange("A1:D23").setBorder(true, true, true, true, false, false)
  sheet.setColumnWidth(1, 352);
  sheet.getRange("A1:D1").merge().setBackgroundRGB(47, 84, 150);
  sheet.getRange("A2:D2").merge().setBackgroundRGB(142, 170, 219);
  sheet.getRange("A3:D3").setWrap(true).setHorizontalAlignment("center").setFontWeight("bold");
  var row3 = ["Domain", "Below expectations", "Meeting expectations", "Exceeding expectations"];
  for (var i = 0; i<row3.length; i++) {
     sheet.getRange(3, i+1).setValue(row3[i]);
  }
  var indicators = ["Uses language to regulate behavior", "Can social problem-solve", "Uses rules during learning activities", "Able to focus attention until task is finished", "Engages in positive interactions with peers", "Has task persistence. Keeps trying...", "Writing: SW Level", "Verbal Planning", "Letter Recognition", "Letter Sound Recognition", "Language Acquisition", "Vocabulary", "Listening Comprehension", "Phonological Awareness", "Rote Counting", "Counting Objects", "Recognizing Numbers", "Shapes", "Sorting by Attribute", "Measurement"];
  for (var i = 0; i < indicators.length; i++) {
    sheet.getRange(i+4, 1).setValue(indicators[i]);
  }
  var cols = ["H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA"];
  var ratings = ["0", "1", "2"];
  var expectations = sheet.getRange("B4:D23");
  var formulas = [];
  for (col in cols) {
    var formulaArray = []
    for (rate in ratings) {
      formulaArray.push("=(COUNTIF(dataCopy!" + cols[col] + "2:" + cols[col] + ", \"=" + ratings[rate] + "\")/COUNT(dataCopy!" + cols[col] + "2:" + cols[col] + "))");
    }
    formulas.push(formulaArray);
  }
  var ratingRange = sheet.getRange("B4:D23").setNumberFormat("0.#%");
  ratingRange.setFormulas(formulas);
  
}

function convertData(grade){
  var ss = SpreadsheetApp.getActive();
  ss.duplicateActiveSheet().setName("dataCopy");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dataCopy");
  //  get the current data range values as an array
  //  Less calls to access the sheet, lower overhead 
  var values = sheet.getDataRange().getValues(); 
  
  if (grade == 3) {
    
    // Replace 3YO ELA 1-4 (Verbal Planning)
    
    replaceInSheet(values,"Gesture","0");
    replaceInSheet(values,"Center","0");
    replaceInSheet(values,"Action","1");
    replaceInSheet(values,"Roles","2");
    replaceInSheet(values,"Props","2");
    replaceInSheet(values,"Role Interaction","2");
    replaceInSheet(values,"Joint Planning","2");
    
    //    Replace 3YO ELA 1-4 (Writing)
  
    replaceInSheet(values,"Plan","0");
    replaceInSheet(values,"Picture","0");
    replaceInSheet(values,"Message","1");
    replaceInSheet(values,"Lines","2");
    replaceInSheet(values,"IS","2");
    replaceInSheet(values,"ES","2");
    replaceInSheet(values,"MS","2");
    replaceInSheet(values,"AP","2");
    
    // Replace 3YO ELA 1-4 (Letter Recognition)
    
    replaceInSheet(values,"0","0");
    replaceInSheet(values,"1-6","1");
    replaceInSheet(values,"7-13","2");
    replaceInSheet(values,"14-20","2");
    replaceInSheet(values,"21-26","2");
    
    // Replace 3YO ELA 1-4 (Letter Sound Recognition)
    
    replaceInSheet(values,"0","0");
    replaceInSheet(values,"1-6","1");
    replaceInSheet(values,"7-13","2");
    replaceInSheet(values,"14-20","2");
    replaceInSheet(values,"21-26","2");
    
    // Replace 3YO ELA 5-8 (Language Acquisition)
    
    replaceInSheet(values,"Gesture","0");
    replaceInSheet(values,"\"I\" Statements","0");
    replaceInSheet(values,"Phrases","0");
    replaceInSheet(values,"Sentences","1");
    replaceInSheet(values,"Stories","2");
    
    // Replace 3YO ELA 5-8 (Vocabulary)
    
    replaceInSheet(values,"Shows understanding of everyday words","0");
    replaceInSheet(values,"Shows understanding of new word by acting it out","0");
    replaceInSheet(values,"Uses the word in story lab to describe","1");
    replaceInSheet(values,"Applies the word in new situation","2");
    replaceInSheet(values,"Uses synonyms for the word as well as examples to","2");
    
    // Replace 3YO ELA 5-8 (Listening Comprehension)
    
    replaceInSheet(values,"Attends to the book and may comment on the story","0");
    replaceInSheet(values,"Answers questions and responds to story (describes","0");
    replaceInSheet(values,"Makes text to self connections","1");
    replaceInSheet(values,"Makes text to text and text to world connections","2");
    replaceInSheet(values,"Retells stories accurately identifying beg/mid/end","2");
    
    // Replace 3YO ELA 5-8 (Phonological Awareness)
    
    replaceInSheet(values,"Participates in rhyming activities with group","0");
    replaceInSheet(values,"Completes the rhyming part of a song","1");
    replaceInSheet(values,"Identifies rhyming words","2");
    replaceInSheet(values,"Identifies and produces simple rhyming words","2");
    
    
    // Replace 3YO Math 1-3 (Counting Objects)
    
    replaceInSheet(values,"Counts objects with/without correct number order","0");
    replaceInSheet(values,"Counts 1-10 objects & compares with more/less/same","1");
    replaceInSheet(values,"Counts > 10 objects knowing last number is total","2");
    replaceInSheet(values,"Adds and subtracts in groups of 10 objects","2");
    
    // Replace 3YO Math 1-3 (Recognizing Numbers)
    
    replaceInSheet(values,"Recognizes numbers in the environment","0");
    replaceInSheet(values,"Recognizes numerals 1-10","1");
    replaceInSheet(values,"Recognizes numbers 1-10 and writes 1-10","2");
    replaceInSheet(values,"Recognizes numerals > 10 and writes many 1-20","2");
    
    // Replace 3YO Math 4-6 (Shapes)
    
    replaceInSheet(values,"Begins to identify basic shapes","0");
    replaceInSheet(values,"Identifies and can draw  basic shapes","1");
    replaceInSheet(values,"Recognizes and names 2D and 3D shapes","2");
    replaceInSheet(values,"Understands the connection b/w 3D and 2D shapes","2");
    
    // Replace 3YO Math 4-6 (Sorting)
    
    replaceInSheet(values,"Matches objects that are identical","0");
    replaceInSheet(values,"Sorts objects into small groups by 1 attribute","1");
    replaceInSheet(values,"Reclassifies already sorted objects by attribute","2");
    replaceInSheet(values,"Classifies/compares subgroups within larger groups","2");
    
    // Replace 3YO Math 4-6 (Sorting)
    
    replaceInSheet(values,"Begins to use concepts of measurement for puzzles","0");
    replaceInSheet(values,"Compares objects & uses comparative language","1");
    replaceInSheet(values,"Measures with non-standard and standard tools","2");
    replaceInSheet(values,"Measures using a common base describes attribute","2");
    
    // Replace 3YO Math 1-3 (Rote Counting)
    
    replaceInSheet(values,"Begins to verbally recite numbers partially correc","0");
    replaceInSheet(values,"Counts 1-10","1");
    replaceInSheet(values,"Counts 1-20","2");
    replaceInSheet(values,"Counts beyond 20 and can count backwards from 10","2");
   } else if (grade == 4) {
     
    // Replace 4YO ELA 1-4 (Verbal Planning)
    
    replaceInSheet(values,"Gesture","0");
    replaceInSheet(values,"Center","0");
    replaceInSheet(values,"Action","0");
    replaceInSheet(values,"Roles","0");
    replaceInSheet(values,"Props","0");
    replaceInSheet(values,"Role Interaction","1");
    replaceInSheet(values,"Joint Planning","2");
     
     //    Replace 4YO ELA 1-4 (Writing)
  
    replaceInSheet(values,"Plan","0");
    replaceInSheet(values,"Picture","0");
    replaceInSheet(values,"Message","0");
    replaceInSheet(values,"Lines","0");
    replaceInSheet(values,"IS","1");
    replaceInSheet(values,"ES","2");
    replaceInSheet(values,"MS","2");
    replaceInSheet(values,"AP","2");
    
    // Replace 4YO ELA 1-4 (Letter Recognition)
    
    replaceInSheet(values,"0","0");
    replaceInSheet(values,"1-6","0");
    replaceInSheet(values,"7-13","0");
    replaceInSheet(values,"14-20","1");
    replaceInSheet(values,"21-26","2");
    
    // Replace 4YO ELA 1-4 (Letter Sound Recognition)
    
    replaceInSheet(values,"0","0");
    replaceInSheet(values,"1-6","0");
    replaceInSheet(values,"7-13","0");
    replaceInSheet(values,"14-20","1");
    replaceInSheet(values,"21-26","2");
    
    // Replace 4YO ELA 5-8 (Language Acquisition)
    
    replaceInSheet(values,"Gesture","0");
    replaceInSheet(values,"\"I\" Statements","0");
    replaceInSheet(values,"Phrases","0");
    replaceInSheet(values,"Sentences","0");
    replaceInSheet(values,"Stories","1");
    
    // Replace 4YO ELA 5-8 (Vocabulary)
    
    replaceInSheet(values,"Shows understanding of everyday words","0");
    replaceInSheet(values,"Shows understanding of new word by acting it out","0");
    replaceInSheet(values,"Uses the word in story lab to describe","0");
    replaceInSheet(values,"Applies the word in new situation","1");
    replaceInSheet(values,"Uses synonyms for the word as well as examples to","2");
    
    // Replace 4YO ELA 5-8 (Listening Comprehension)
    
    replaceInSheet(values,"Attends to the book and may comment on the story","0");
    replaceInSheet(values,"Answers questions and responds to story (describes","0");
    replaceInSheet(values,"Makes text to self connections","0");
    replaceInSheet(values,"Makes text to text and text to world connections","1");
    replaceInSheet(values,"Retells stories accurately identifying beg/mid/end","2");
    
    // Replace 4YO ELA 5-8 (Phonological Awareness)
    
    replaceInSheet(values,"Participates in rhyming activities with group","0");
    replaceInSheet(values,"Completes the rhyming part of a song","0");
    replaceInSheet(values,"Identifies rhyming words","0");
    replaceInSheet(values,"Identifies and produces simple rhyming words","1");
    
    
    // Replace 4YO Math 1-3 (Counting Objects)
    
    replaceInSheet(values,"Counts objects with/without correct number order","0");
    replaceInSheet(values,"Counts 1-10 objects & compares with more/less/same","0");
    replaceInSheet(values,"Counts > 10 objects knowing last number is total","1");
    replaceInSheet(values,"Adds and subtracts in groups of 10 objects","2");
    
    // Replace 4YO Math 1-3 (Recognizing Numbers)
    
    replaceInSheet(values,"Recognizes numbers in the environment","0");
    replaceInSheet(values,"Recognizes numerals 1-10","0");
    replaceInSheet(values,"Recognizes numbers 1-10 and writes 1-10","1");
    replaceInSheet(values,"Recognizes numerals > 10 and writes many 1-20","2");
    
    // Replace 4YO Math 4-6 (Shapes)
    
    replaceInSheet(values,"Begins to identify basic shapes","0");
    replaceInSheet(values,"Identifies and can draw  basic shapes","0");
    replaceInSheet(values,"Recognizes and names 2D and 3D shapes","1");
    replaceInSheet(values,"Understands the connection b/w 3D and 2D shapes","2");
    
    // Replace 4YO Math 4-6 (Sorting)
    
    replaceInSheet(values,"Matches objects that are identical","0");
    replaceInSheet(values,"Sorts objects into small groups by 1 attribute","0");
    replaceInSheet(values,"Reclassifies already sorted objects by attribute","1");
    replaceInSheet(values,"Classifies/compares subgroups within larger groups","2");
    
    // Replace 4YO Math 4-6 (Measurement)
    
    replaceInSheet(values,"Begins to use concepts of measurement for puzzles","0");
    replaceInSheet(values,"Compares objects & uses comparative language","0");
    replaceInSheet(values,"Measures with non-standard and standard tools","1");
    replaceInSheet(values,"Measures using a common base describes attribute","2");
    
    // Replace 4YO Math 1-3 (Rote Counting)
    
    replaceInSheet(values,"Begins to verbally recite numbers partially correc","0");
    replaceInSheet(values,"Counts 1-10","0");
    replaceInSheet(values,"Counts 1-20","1");
    replaceInSheet(values,"Counts beyond 20 and can count backwards from 10","2");
   }
  
  
  
  //write the updated values to the sheet, again less call;less overhead
  sheet.getDataRange().setValues(values).setNumberFormat("0");
}

function replaceInSheet(values, to_replace, replace_with) {
    //loop over the rows in the array
      for(var row in values){
         //use Array.map to execute a replace call on each of the cells in the row.
         var replaced_values = values[row].map(function(original_value){
            return original_value.toString().replace(to_replace,replace_with);
          });
    //replace the original row values with the replaced values
    values[row] = replaced_values;
  }
}

