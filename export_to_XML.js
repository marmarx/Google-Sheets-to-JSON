/*
  Exports the data from the specified Google Sheet to an XML string, handling joined columns.
  @return{string}An XML string representing the sheet data, or null if an error occurs.
*/
function exportSheetToXml(spreadsheetId, sheetName, includeHeader = true){
  try{
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);

    if(!sheet){
      Logger.log(`Sheet "${sheetName}" not found.`);
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const header = includeHeader?Array.from(data.shift(),item => item || null):null;  //Remove header row if needed
    //'filter', 'map', and 'forEach' only iterate over existing properties, 'Array.from' iterates over all properties => replace 'empty' by 'null' to avoid this problem

    let xmlString = '<?xml version="1.0" encoding="UTF-8"?>\n'; //XML declaration
    xmlString += `<${sheetName}>\n`; //Root element(using sheet name)

    data.forEach((row) => {
      xmlString += `  <row>\n`;

      if(header){
        let group;
        header.forEach((headerCell, index) => {
          if(headerCell){group = [row[index]]}
          else{group.push(row[index])}
          
          if(header[index+1]||index===header.length-1){
            if(group.length===1){
              xmlString += `    <${headerCell}>${group.join()}</${headerCell}>\n`
            }else{
              xmlString += `    <${header[index-group.length+1]}>\n`;
              group.forEach(item => xmlString += item?`      <value>${item}</value>\n`:'');
              xmlString += `    </${header[index-group.length+1]}>\n`;
            }
          }
        });

      }else{
        row.forEach((cell, index) => {
          xmlString += `    <column${index}>${cell}</column${index}>\n`; //Generic column names
        });
      }

      xmlString += `  </row>\n`;
    });

    xmlString += `</${sheetName}>\n`; //Close root element

    return xmlString;

  }catch(error){
    Logger.log(`Error exporting sheet to XML: ${error}`);
    return null;
  }
}

/*
  Saves the XML string to a Google Drive file
  @return{string}The ID of the created file, or null if an error occurred
*/
function saveXmlToDrive(spreadsheetId, sheetName, xmlString, folderId){
  try{
    if(!xmlString){
      Logger.log("No XML string to save.");
      return null;
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const spreadsheetName = ss.getName();

    const fileName = `${spreadsheetName}- ${sheetName}.xml`;
    const blob = Utilities.newBlob(xmlString, MimeType.XML, fileName);

    let file;
    if(folderId){
       const folder = DriveApp.getFolderById(folderId);
       file = folder.createFile(blob);
    }else{
      file = DriveApp.createFile(blob); //Save to the root of Google Drive
    }
    
    Logger.log(`XML saved to Drive: https://drive.google.com/file/d/${file.getId()}/view`);
    return file.getId();

  }catch(error){
    Logger.log(`Error saving XML to Drive: ${error}`);
    return null;
  }
}

function exportAndSaveXML(){
  const spreadsheetId = 'YOUR_SPREADSHEET_ID'; //Replace with your Spreadsheet ID
  const sheetName = 'YOUR_SHEET_NAME'; //Replace with your sheet name
  const saveToFolder = 'YOUR_DRIVE_FOLDER_ID'; //(Optional) Replace with your Google Drive folder ID, or leave blank/null
  const includeHeader = true;   //Boolean, true by default, use 'true' if your sheet has a header in the first line

  const xmlString = exportSheetToXml(spreadsheetId, sheetName, includeHeader);

  if(xmlString){
    saveXmlToDrive(spreadsheetId, sheetName, xmlString, saveToFolder);
  }
}
