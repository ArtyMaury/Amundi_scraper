const API_KEY="<TOKEN_FROM_BROWSER>"

const OPERATIONS_URL="https://epargnant.amundi-ee.com/api/individu/operations?flagFiltrageWebSalarie=true&flagInfoOC=Y&filtreStatutModeExclusion=false&flagRu=true&offset="

const DISPOSITIFS_URL="https://epargnant.amundi-ee.com/api/individu/dispositifs?flagUrlFicheFonds=true&codeLangueIso2=fr"

function main() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();

  const dispositifs = getDispositifs()
  const dispositifsSheet = resetSheet(doc, "DISPOSITIFS")
  insertRawData(dispositifsSheet, dispositifs)
  
  const operations = getOperations()  
  const operationsSheet = resetSheet(doc, "OPERATIONS")
  insertRawData(operationsSheet, operations)

  updateMainSheet(doc)
}

function getDispositifs() {
  const data = UrlFetchApp.fetch(DISPOSITIFS_URL, {headers: {"X-noee-authorization": API_KEY }})
  const jsonData = JSON.parse(data.getContentText())
  return jsonData.listPositionsSalarieDispositifsDto.filter(dispo => dispo.mtBrut != 0)
}

function getOperations() {
  const operations = []
  let nbOperations = 1
  let offset=0
  while(operations.length < nbOperations){
    const data = UrlFetchApp.fetch(OPERATIONS_URL + offset, {headers: {"X-noee-authorization": API_KEY }})
    const jsonData = JSON.parse(data.getContentText())
    operations.push(...jsonData.operationsIndividuelles);
    nbOperations = jsonData.nbOperationsIndividuelles
    offset ++;
  }
  return operations
}

/**
 *  @param {SpreadsheetApp.Spreadsheet} doc 
 *  @return {SpreadsheetApp.Sheet}
 */
function resetSheet(doc, sheetName){
  if (doc.getSheetByName(sheetName)){
    doc.deleteSheet(doc.getSheetByName(sheetName))
  }
  const rawSheet = doc.insertSheet()
  rawSheet.setName(sheetName)
  return rawSheet
}

function updateMainSheet(doc) {
  let mainSheet = doc.getSheetByName("main")
  if (!mainSheet) {
    mainSheet = doc.insertSheet(0)
    mainSheet.setName("main")
  }
  doc.setActiveSheet(mainSheet)
  const fullRange = mainSheet.getRange(1,1,500,50)
  recalculate(fullRange)
}

/**
 *  @param {SpreadsheetApp.Sheet} sheet 
 *  @param {any[]} data 
 */
function insertRawData(sheet, data){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (data.length==0){
    ss.toast("No data")
  } else {
    const headers = Object.keys(data[0]).sort((a,b) => a.localeCompare(b))
    insertHeaders(sheet, headers);
    insertRows(sheet, data, headers);
  }

}

/**
 *  @param {SpreadsheetApp.Sheet} sheet 
 *  @param {string[]} headers
 */
function insertHeaders(sheet, headers) {
  let i = 1;
  headers.forEach(header => {
    sheet.getRange(1,i).setValue(header)
    i++;
  })
}

/**
 *  @param {SpreadsheetApp.Sheet} sheet 
 *  @param {any[]} data
 *  @param {string[]} headers
 */
function insertRows(sheet, data, headers) {
  const destinationRange = sheet.getRange(2, 1, data.length, headers.length);
  const values = data.map(row => headers.map(h => row[h]))
  destinationRange.setValues(values);
}

/**
 * @param {SpreadsheetApp.Range} range
 */
function recalculate(range){
  var originalFormulas = range.getFormulas();
  var originalValues = range.getValues();
  
  var valuesToEraseFormula = [];
  var valuesToRestoreFormula = [];
  
  originalFormulas.forEach((rowValues,rowNum) => {
    valuesToEraseFormula[rowNum] = [];
    valuesToRestoreFormula[rowNum] = [];
    rowValues.forEach((cellValue, columnNum) => {
      const a = originalValues[rowNum][columnNum]
      if('' === cellValue){
        //The cell doesn't have formula
        valuesToEraseFormula[rowNum][columnNum] = originalValues[rowNum][columnNum];
        valuesToRestoreFormula[rowNum][columnNum] = originalValues[rowNum][columnNum];
      }else{
        //The cell has a formula.
        valuesToEraseFormula[rowNum][columnNum] = null;
        valuesToRestoreFormula[rowNum][columnNum] = originalFormulas[rowNum][columnNum];
      }
    })
  })
  
  range.setValues(valuesToEraseFormula);
  range.setValues(valuesToRestoreFormula);
}
