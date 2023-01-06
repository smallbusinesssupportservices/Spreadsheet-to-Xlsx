/**
 * @author William Brewer <william@smallbusinessessupport.services>
 * 
 * This function will create a xlsx file link from a spreadsheet id or spreadsheet url
 * 
 * @param {Object} data 
 * 
 * @property {String} data.ssId      - Id of a spreadsheet
 * @property {String} data.ssUrl     - Url to spreadsheet
 * @property {String} data.sheetName - Name sheet to convert to xlsx
 * 
 * @returns {String} download link 
 */
function sheetToExcel(data) {

  let ss = getSpreadSheet(data)
  let sheet = ss.getSheetByName(data.sheetName)
  let downloadUrl = buildDownloadUrl(ss,sheet)

  return downloadUrl

}

function getSpreadSheet(data) {
  if (Object.keys(data).includes('ssId')) {
    ss = SpreadsheetApp.openById(data.ssId)
  } else if (Object.keys(data).includes('ssUrl')) {
    ss = SpreadsheetApp.openByUrl(data.ssUrl)
  } else {
    throw new Error('Enter a Spreadsheet id or url')
  }
  return ss

}

function buildDownloadUrl(ss,sheet){

  return "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx&gid=" + sheet.getSheetId().toString()

}
