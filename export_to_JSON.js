/*
  Exports the data from the specified Google Sheet to a JSON string.
  @return {string} A JSON string representing the sheet data, or null if an error occurs.
*/
function exportSheetToJson(spreadsheetId, sheetName, includeHeader = true, prettyJson = true){
  try{
    //Get the spreadsheet, spreadsheet name and sheet
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const ssName = ss.getName();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found.`);
      return null;
    }

    //Get the data in the sheet
    const data = sheet.getDataRange().getValues();
    Logger.log(data);

    const header = includeHeader?Array.from(data.shift(),item => item || null):null;  //Remove header row if needed
    //'filter', 'map', and 'forEach' only iterate over existing properties, 'Array.from' iterates over all properties => replace 'empty' by 'null' to avoid this problem

    //Convert data to JSON
    const jsonArray = [];

    data.forEach(row => {
      const jsonObject = {};
      if(header){
        let group;
        header.forEach((headerCell, index) => {
          if(headerCell){group = [row[index]]}
          else{group.push(row[index])}

          if(group.length===1){jsonObject[headerCell] = group.join()}
          else{jsonObject[header[index - group.length + 1]] = group.filter(item => item)}
        })
      }else{
        row.forEach((cell, index) => {jsonObject[index] = cell})   //If no header, use index as keys (0, 1, 2, etc.)
      }
      jsonArray.push(jsonObject);
    });

    const jsonString = prettyJson?JSON.stringify(jsonArray, null, 2):JSON.stringify(jsonArray);
    return [jsonString,ssName];

  }catch (error){
    Logger.log(`Error exporting sheet: ${error}`);
    return null;
  }
}

/*
  Saves the JSON string to a Google Drive file.
  @return {string} The ID of the created file, or null if an error occurred.
 */
function saveJsonToDrive(ssName, sheetName, jsonString, folderId){
  try{
    if(!jsonString){
      Logger.log("No JSON string to save.");
      return null;
    }

    const fileName = `Generated JSON file - ${ssName} - ${sheetName}.json`;
    const blob = Utilities.newBlob(jsonString, MimeType.JSON, fileName);

    let file;
    if(folderId){
       const folder = DriveApp.getFolderById(folderId);
       file = folder.createFile(blob);
    }else{
      file = DriveApp.createFile(blob); //Save to the root of Google Drive
    }
    
    Logger.log(`JSON saved to Drive: https://drive.google.com/file/d/${file.getId()}/view`);
    return file.getId();

  }catch (error){
    Logger.log(`Error saving JSON to Drive: ${error}`);
    return null;
  }
}

/*
  Steps:
  1. Set 'spreadsheetId', 'sheetName' and 'folderId' properly - Leave 'folderId' empty to save file to your Drive root
  2. Set 'includeHeader' and 'prettyJson' if desired
  3. Run function 'exportAndSave'
*/
function exportAndSaveJSON() {
  const spreadsheetId = 'YOUR_SPREADSHEET_ID'; //Replace with your Spreadsheet ID
  const sheetName = 'YOUR_SHEET_NAME'; //Replace with your sheet name
  const saveToFolder = 'YOUR_DRIVE_FOLDER_ID'; //(Optional) Replace with your Google Drive folder ID, or leave blank/null
  const includeHeader = true;   //Boolean, true by default, use 'true' if your sheet has a header in the first line
  const prettyJson = true;    //Boolean, true by default, use 'true' for pretty printing or 'false' for compact JSON

  const [jsonString,ssName] = exportSheetToJson(spreadsheetId, sheetName, includeHeader, prettyJson);
  if(jsonString){saveJsonToDrive(ssName, sheetName, jsonString, saveToFolder)}
}