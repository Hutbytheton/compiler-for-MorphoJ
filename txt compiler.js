/************************************************
Name: Coordinate Data Compiler for use with MorphoJ (http://www.flywings.org.uk/morphoj_page.htm)
Purpose: This program uses 2D coordinate data entered in a Google Sheet to create .txt files 
         suitable for use with MorphoJ.
Authors: Hutton Brandon
************************************************/

////// Global Variables //////
// TODO: Find a different way to store these so they're not plain global variables
var entrySheetName = "Enter File Names and Classifiers";
var phases = ["Oral Phase Start", "Oral Phase End", "Pharyngeal Phase End"];
var phaseNames = ["Pre-Oral Phase", "Oral Phase", "Pharyngeal Phase", "Esophageal Phase"];




////// Function to test code in isolation //////
function testFunction() {
  //var sheet = selectEntrySheet();
  Logger.log(SpreadsheetApp.getActive().getSheets()[0].getRange('C44').getValue());
  
  
 
}

//
// Create menu for custom functions
//
function onOpen() {
  
  var menu = [],
      spreadsheet;
  
  // get the active spreadsheet
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  menu = [{name: "Import Txt File Names", functionName: "fileNameCopier"}, {name: 'Process Txt Files', functionName: 'txtProcessor'}];
  spreadsheet.addMenu('Compiler Tools', menu);
};

//                             //
//////// Master Function ////////
//                             //
function txtProcessor() {
  
  var ss,
      sheet,
      classifiersArray,
      filesArray,
      folderNames = {},
      inputFolder = "",
      outputFolder = "",
      filesObject,
      phasesObj,
      names,
      input,
      dateStamp,
      sh,
      file1,
      file2;
  
  sheet = selectEntrySheet();
  classifiersArray = getClassifiers(sheet);
  
  filesArray = getFilenames(sheet);
  
  folderNames = getFolderNames();
  inputFolder = folderNames["inputFolderName"];
  outputFolder = folderNames["outputFolderName"];
  
  filesObject = getFilesObject(classifiersArray, filesArray, inputFolder);
  phasesObj = getPhases(sheet);
  names = getName(filesObject);
  
  // way to avoid output if error encountered in filesObject
  if (filesObject) {
    ss = createSpreadsheet(names, outputFolder); 
  } else {
    Browser.msgBox("Error encountered: No filesObject. Please check that everything is set up correctly and try again.");
    return "Error encountered";
  }
  
  // way to avoid finishing output if error encountered while making spreadsheet
  if (!ss) { return "Error encountered with spreadsheet creation or move" };
  
  dateStamp = getDateStamp();
  populateSpreadsheet(filesObject, phasesObj, classifiersArray, ss);
  //Browser.msgBox("Check your output folder to find the generated spreadsheet.");
  
  sh = ss.getSheets()[0];
  file1 = createTxtFile("Coordinates " + dateStamp, sheetToString(sh), outputFolder);
  sh = ss.getSheets()[1];
  file2 = createTxtFile("Classifiers " + dateStamp, sheetToString(sh), outputFolder);
  Browser.msgBox("Processor finished. Check your output folder to find the files.");
  Logger.log("file1 size = " + file1.getSize() + "  files 2 size = " + file2.getSize());
}

// takes the name of a folder in drive and the sheet to be copied to. 
// converts all the names of the files of that folder as an array of form [[filename1],[filename2],[filename2],...]
// adds them to (A2:An) on the specified sheet and returns filenames array
function fileNameCopier() {
  
  var spreadsheet,  
      folderName = "",
      files,
      file,
      filenames = [[]],
      i = 0;
  
  // get the active spreadsheet
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  folderName = getFolderNames()["inputFolderName"];
  
  try {
    files = DriveApp.getFoldersByName(folderName).next().getFiles();
  } catch (e) {
    Browser.msgBox("[" + e + "]" + "  No folder of such name or folder is empty. Please check that you have the correct folder name and files, and try again.");
  }
  
  while (files.hasNext()) {
    file = files.next();
    filenames[i] = [file.getName()];
    i++;
    Logger.log(filenames);
  }
  
  spreadsheet.getSheetByName("Enter File Names and Classifiers").getRange(5,1,filenames.length).setValues(filenames);
  
  Browser.msgBox("File names imported successfully");
  return filenames;
}


////// Seleting entry sheets //////


// select the entry sheet. if name is not found, select the first sheet.
function selectEntrySheet() {
  
  var spreadsheet,
      entrySheet;
  
  // get the active spreadsheet
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  
  entrySheet = spreadsheet.getSheetByName(entrySheetName);
  
  if (entrySheet) {
    return entrySheet
  } else {
    Logger.log("Can't select entry sheet named " + entrySheetName);
  }
}
    
//////// Reading the entry sheet ////////



// get swallowing phase frames numbers
// returns nested object of form {filename1:{(name of phase):frame#, (name of phase2):frame#),...}, filename2:...}
function getPhases(entrySheet) {
  // TODO: remove entrysheet calls in loops and instead store as an array before loops
  var phasesObject = {},
      numRows = 0,
      filename = "",
      innerObject = {},
      phasesLength = 0,
      value = 0;
  
  numRows = entrySheet.getLastRow() - 1;
  
  // loop through rows of entry sheet
  for (var i=0; i<numRows; i++) {
    filename = entrySheet.getRange(i+5,1).getValue();
    innerObject = {};
    
    // loop through phases
    phasesLength = phases.length;
    for (var j=0; j<phasesLength; j++) {
      value = entrySheet.getRange(i+5,j+2).getValue()
      
      innerObject[phases[j]] = value;
    }
    phasesObject[filename] = innerObject;   
  }
  
  return phasesObject;
}
    

// get names of other characteristics in array of form [characteristic1, characteristic2, characteristic3, ... ]
function getClassifiers(entrySheet) {
  
  var numColumns = 0.0,
      classifiersArray = [];
  
  numColumns = entrySheet.getLastColumn() - 4;
  if (numColumns < 1.0) {
    Browser.msgBox("Please add an additional classifier. This script only works if at least cell E4 is occupied.");
    return null;
  }
  else {
    classifiersArray = entrySheet.getRange(4,5,1,numColumns).getValues()[0];
  
    return classifiersArray;
  }
}
  
  
// get filenames on entry sheet from A2 down 
// in 2d array of filenames (e.g. [ [Filename1, characteristic1, characteristic2, ...], [Filename2, characteristic1, characteristic2, ...] ])
function getFilenames(entrySheet) {
  
  var numEntries = 0.0,
      numColumns = 0.0,
      filenamesArray = [];
  
  numEntries = entrySheet.getLastRow() - 4;
  numColumns = entrySheet.getLastColumn();
  
  filenamesArray = entrySheet.getRange(5,1,numEntries,numColumns).getValues();
  
  return filenamesArray;
} 

// returns object of form {inputFolderName: "name1", outputFolderName: "name2"}
function getFolderNames() {
  
  var entrySheet,
      inputFolder = "",
      outputFolder = "",
      outputObj = {};
   
  entrySheet = selectEntrySheet();
  
  inputFolder = entrySheet.getRange(1,2).getValue();
  outputFolder = entrySheet.getRange(2,2).getValue();
  
  outputObj = {
    inputFolderName: inputFolder,
    outputFolderName: outputFolder
  };
  
  return outputObj;
}


/////// Processing Input ////////

// input of form getFilesObject(number of landmarks, array of characteristic names, 
// create an object with filenames as keys
// values of form {file:File, $othercharacteristic:value,...}
function getFilesObject(classifiersArray, filenamesArray, folderName) {
  
  var folders,
      folder,
      filesObject = {},
      filenamesArrayLength = 0.0,
      filename = "",
      fileItself,
      charactArrayLength = 0.0;
  
  // input folder naming
  try { 
    folders = DriveApp.getFoldersByName(folderName);
    folder = folders.next();
    if (folders.hasNext()) { throw "Input folder name " + folder + " not unique in your Drive." }
  } catch (e) {
    Browser.msgBox("[" + e + "]" + "  Something is wrong with your specified input folder. Please delete any output, check your input folder name, and try again.");
    return null;
  }
  // loop through rows in filenamesArray
  filenamesArrayLength = filenamesArray.length;
  for (var i=0; i<filenamesArrayLength; i++) {
    
    filename = filenamesArray[i][0];
    
    try {
      fileItself = folder.getFilesByName(filename).next();
    } 
    catch(e) {
      Browser.msgBox("[" + e + "]" + "  File: " + filename + " not found in " + folder + " folder. Please remove any output, make sure the files in the entry sheet match those in the " + folder + " folder, and try again.");
      return null;
    }
    filesObject[filename] = {
      file: fileItself
    };
    
    // loop through characteristics in classifiersArray and columns in filenamesArray and add to object
    charactArrayLength = classifiersArray.length;
    for (var j=0; j<charactArrayLength; j++) {
   
      filesObject[filename][classifiersArray[j]] = filenamesArray[i][j+4];
      
     }    
  }
  
  return filesObject;
}
  

// return names of first file to be compiled as a string
function getName(filesObject) {
  
  var name = "";
  
  for (var filename in filesObject) {
    names = filesObject[Object.keys(filesObject)[0]];
  }
  
  return name;
}


// move file by id to output folder. returns the folder
function moveToOutputFolder(fileId, folderName) {
  
  var folders,
      outputFolder,
      file;
  
  try { 
    folders = DriveApp.getFoldersByName(folderName);
    outputFolder = folders.next();
    if (folders.hasNext()) { throw "Output folder name " + outputFolder + " not unique in your Drive." }
  } catch (e) {
    Browser.msgBox("[" + e + "]" + "  Something is wrong with your output folder. Please delete any output, check your output folder name, and try again.");
    return null;
  }
 
  file = DriveApp.getFileById(fileId)
  outputFolder.addFile(file);
  return outputFolder;
}

// return dateStamp
function getDateStamp() {
  
  var date,
      dateStamp = "";
   
  // create date stamp
  date = new Date();
  dateStamp = date.getDate() + "-" + (date.getMonth() + 1) + "-" + date.getFullYear() + " " + date.getHours() + ":" + date.getMinutes();
 
  return dateStamp;
  
}   

////// Creating and populating spreadsheet //////

// create spreadsheet with concatenated text files
function createSpreadsheet(names, outputFolder) {
  
  var name = "",
      newSpreadsheet,
      moveValue;
      
  
  // base name of output files
  name = "Processer Output Sheet " + names + " " + getDateStamp();
  
  /* not sure I want this functionality for now
  
  // check for spreadsheets with same name. if found, delete them
  var ditto = DriveApp.getFilesByName(name);
  if (ditto.hasNext()) {
    ditto.next().setTrashed(true)
  } 
  */
  
  // create spreadsheet in Drive root
  newSpreadsheet = SpreadsheetApp.create(name);
  
  // move to output folder
  moveValue = moveToOutputFolder(newSpreadsheet.getId(), outputFolder);
  
  // end function if output folder not right
  if (!moveValue) { 
    Browser.msgBox("Something wrong with output folder selection. Please check your the main folder of your Drive for a file named \"" + name + "\" and remove it before trying again.");
    return null; 
  }
  
  // create 3 sheets: raw concatenated data, concatenated with ID, characteristics
  newSpreadsheet.getSheets()[0].setName("ID w/ Data");
  newSpreadsheet.insertSheet("ID w/ Classifiers");
  //newSpreadsheet.insertSheet("Raw Concatenated Data");
  
  return newSpreadsheet;
}

// bitwise OR float truncating function
function float2int (value) {
    return value | 0;
}

// accepts a 2d array of row of [filename, floats...] and rounds each to nearest integer and returns the modified array
//Not that the rows must be the same length, this was specified to improve performance
function roundArrayToInteger(array) {
  
  var rows = 0,
      columns = 0,
      i = 0,
      j = 0;
      
  rows = array.length;
  columns = array[0].length - 1;
  for (i=0; i<rows; i++) {
    // j=1 to skip the filename
    for (j=1; j<columns; j++) {
      array[i][j] = Math.round(array[i][j]);
    }
  }
  return array;                            
}

// replaces the next line characters of different operating systems with a single one
function canonicalizeNewlines (str) {
      return str.replace(/(\r\n|\r|\n)/g, '\n');
};


// returns object of form {dataArray: [array of data], headers: [headers]} from object containing files
function filesDataObj(filesObject) {
  
  var dataArray = [],
      headers = [],
      blob,
      blobString = "",
      fileArray = [],
      fileArrayLength = 0,
      returnedObject = {};
  
  for (var filename in filesObject) {
    
    try {
      blob = filesObject[filename]["file"].getBlob();
    }
    catch (e) {
      Browser.msgBox("[" + e + "]" + "  Not all files found. Please make sure the file in the entry sheet match those in the entry folder.")
    }
    
    blobString = blob.getDataAsString();
    blobString = canonicalizeNewlines(blobString);
    // create array of form [frame1  x1  y1  x2  y2  ..., frame2  x1... , ...] (values are separated by \t)
    fileArray = blobString.split("\n");
    //Logger.log("Filename: " + filename + " Array: " + fileArray.slice(0,4));
    // create 2d array of form [[frame1,x1,y1,x2,y2...],[frame2,x1,y1,x2,...],...]
    fileArrayLength = fileArray.length;
    for (var j=0; j<fileArrayLength; j++) {
      fileArray[j] = fileArray[j].split("\t");
      
      // add filename to row
      fileArray[j].unshift(filename);
    };
    Logger.log("Filename: " + filename + " Array: " + fileArray.slice(0,2));
    // remove and record the first element (the headers)
    headers = fileArray.shift();
    
    // add to dataArray
    fileArray.pop();
    //Logger.log("Filename: " + filename + " Array: " + fileArray[0].slice(0,4));
    dataArray = dataArray.concat(fileArray); 
  }
  
  // round all floats to integers
  roundArrayToInteger(dataArray);
  
  returnedObject = {dataArrayKey: dataArray, headersKey: headers};
  
  return returnedObject;
}

// returns array of data to go on "ID w/ Data" sheet
function createIdDataArray (dataObj) {

  var retrievedArray = [],
      retrievedArrayLength = 0,
      dataArray = [],
      headers =[],
      IdDataArray = [],
      dataArrayLength = 0;
  
  // get object of form {dataArrayKey: array of data, headersKey: headers}
  // array of data has form [[frame1,x1,y1,x2,y2...],[frame2,x1,y1,x2,...],...]
  // headers has form [header1, header2, ...]
  
  // have to loop through the rows to properly clone the array
  retrievedArray = dataObj["dataArrayKey"];
  retrievedArrayLength = retrievedArray.length;
  for (var i =0; i<retrievedArrayLength; i++) {
    dataArray[i] = retrievedArray[i].slice(0);
  }
    
  headers = dataObj["headersKey"];
  
  // remove filename and frame# headers, replace with "Swallow ID" 
  headers.splice(0,2,"Swallow ID");
  
  // add to first row of output array
  IdDataArray[0] = headers;
  
  // add rows of data to output array
  dataArrayLength = dataArray.length;
  for (var i=0; i<dataArrayLength; i++) {
    dataArray[i][0] = dataArray[i][0].slice(0,-4) + "_" + dataArray[i][1];
    dataArray[i].splice(1,1);
    IdDataArray.push(dataArray[i]);
  }
  Logger.log("IdDataArray length: " + IdDataArray[0].length);
  return IdDataArray;
}



// return array to go on "ID w/ Classifiers" sheet
function createIdClassifiersArray(filesObject, phasesObject, classifiersArray, dataObj) {
  
  var header = [],
      IdClassifiersArray = [],
      retrievedArray = [],
      retrievedArrayLength = 0,
      dataArray = [],
      dataArrayLength = 0,
      classifiersArrayLength = 0,
      nameLength = 0,
      phaseNums = {},
      Ostart =0,
      Oend = 0,
      Pend = 0,
      newRow = [],
      frame = 0;
  
  header = classifiersArray.slice(0);
  // add extra headers
  header.unshift("Swallow Phase");
  header.unshift("Swallow ID");
  // add header to first row of output array
  IdClassifiersArray[0] = header;
  
   // array of data has form [[frame1,x1,y1,x2,y2...],[frame2,x1,y1,x2,...],...]
  // have to loop through the rows to properly clone the array
  retrievedArray = dataObj["dataArrayKey"];
  retrievedArrayLength = retrievedArray.length;
  for (var i =0; i<retrievedArrayLength; i++) {
    dataArray[i] = retrievedArray[i].slice(0,4);
  }
  
  // create row with IDs and classifier values
  // set up values outside the huge loop
  dataArrayLength = dataArray.length;
  classifiersArrayLength = classifiersArray.length;
  
  // loop through filenames
  for (var filename in filesObject) {
    
    
    // values defined here to be used in next loop
    nameLength = filename.length - 4;
    // retrieve the innerObject of the phasesObject that corresponds to current filename and set frame numbers for phases
    phaseNums = phasesObject[filename];
    Ostart = phaseNums[phases[0]];
    Oend = phaseNums[phases[1]];
    Pend = phaseNums[phases[2]];
    
    // loop through rows of dataArray seeing if filenames match. if they do, create row of form [swallow ID, "phase", characteristic1, charact2...]
    for (var j=0; j<dataArrayLength; j++) {
      // for some reason, dataArray is still getting manipulated by the above functions although I tried to clone it and protect it. not sure what's
      //going on, but dataArray[j][0] returns a Swallow ID, so for now some awkward slicing is the way forward 
      if (filename.slice(0,-4) == dataArray[j][0].slice(0,-4)) {
        
        //create swallow ID and newRow
        dataArray[j][0] = dataArray[j][0].slice(0,-4) + "_" + dataArray[j][1];
        //// add phase to newRow
        
        // check to see which phase the current ID fits under by checking the number at end of ID, then assign proper phase name to newRow[1]
        frame = dataArray[j][0].slice(nameLength+1);
        
        if (Ostart <= frame && frame <= Oend) {
          newRow = [dataArray[j][0],phaseNames[1]];
        } else if (Oend < frame && frame <= Pend) {
          newRow = [dataArray[j][0],phaseNames[2]];
        } else if (Pend < frame) {
          newRow = [dataArray[j][0],phaseNames[3]];
        } else {
          newRow = [dataArray[j][0],phaseNames[0]];
        }
        
        // loop through classifiersArray so that correct order is preserved when pulling values from filesObject
        // again, the array keeps getting messed up, so the temporary fix is to ignore the first 2 values of classifiersArray //
        for (var k=0; k<classifiersArrayLength;k++) {
          newRow.push(filesObject[filename][classifiersArray[k]]);
        }
        
        
        
        IdClassifiersArray.push(newRow);
      } 
    }
  }  
  return IdClassifiersArray;  
}


// concatenate txt file in first tab of new spreadsheet. returns spreadsheet
function populateSpreadsheet(filesObject, phasesObject, classifiersArray, ss) {
  
  var dataObj = {},
      sheet,
      IdDataArray = [],
      IdClassifiersArray = [];
  
  // get object of form {dataArrayKey: array of data, headersKey: headers}
  // array of data has form [[frame1,x1,y1,x2,y2...],[frame2,x1,y1,x2,...],...]
  // headers has form [header1, header2, ...]
  dataObj = filesDataObj(filesObject);
  
  //// populate "ID w/ Data" sheet ////
  sheet = ss.getSheetByName("ID w/ Data");
  // use dataObj to create array to be inserted on "ID w/ Data" sheet
  IdDataArray = createIdDataArray (dataObj);
  // select range on sheet with corresponding width and height and add in the new array
  sheet.getRange(1,1,IdDataArray.length,IdDataArray[0].length).setValues(IdDataArray);
  
  //// populate "ID w/ Classifiers" sheet ////
  sheet = ss.getSheetByName("ID w/ Classifiers");
  // use inputs to create array to be inserted on "ID w/ Classifiers" sheet
  IdClassifiersArray = createIdClassifiersArray(filesObject, phasesObject, classifiersArray, dataObj);
  // select range on sheet with corresponding width and height and add in the new array
  sheet.getRange(1,1,IdClassifiersArray.length,IdClassifiersArray[0].length).setValues(IdClassifiersArray);
  
  return ss;
}

// returns string with column values separated by a \t and rows separated by \n
function sheetToString(sheet) {
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  
  var valuesArray = sheet.getSheetValues(1,1,numRows,numColumns);
  
  for (var i=0; i<valuesArray.length; i++) {
    valuesArray[i] = valuesArray[i].join("\t");
  }
  var valuesString = valuesArray.join("\n");
  
  return valuesString;
}

// string size can not exceed 10MB. so far, test files max out at around 10kB, so this should be able to handle around 1000 txt files 
function createTxtFile(name, string, outputFolder) {
  var newFile = DriveApp.createFile(name,string);
  moveToOutputFolder(newFile.getId(), outputFolder);
  
  return newFile;
}


/* code for raw conactenated data sheet. no need to do this in finished tool and it slows performance, but I wanted to keep it handy in case I wanted
to put it back in
  
  //// populate concatenated raw data sheet ////
  var sheet = ss.getSheetByName("Raw Concatenated Data");
  var sheet1Array = dataArray.slice(0);
  
  // add header
  headers.shift();
  headers.unshift("File Name");
  sheet.appendRow(headers);
  
  // add rows of data
  var sheet1ArrayLength = sheet1Array.length;  
  for (var i=0; i<sheet1ArrayLength; i++) {
    sheet1Array[i][0] = sheet1Array[i][0].slice(0,-4);
    sheet.appendRow(sheet1Array[i]);
  }
  
  */