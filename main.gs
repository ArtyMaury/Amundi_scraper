const API_KEY ='<TOKEN_FROM_BROWSER>'

const OPERATIONS_URL =
  'https://epargnant.amundi-ee.com/api/individu/operations?flagFiltrageWebSalarie=true&flagInfoOC=Y&filtreStatutModeExclusion=false&flagRu=true&offset='

const DISPOSITIFS_URL =
  'https://epargnant.amundi-ee.com/api/individu/dispositifs?flagUrlFicheFonds=true&codeLangueIso2=fr'
const DISPOSITIF_URL_PREFIX = 'https://epargnant.amundi-ee.com/api/individu/graphe/historiquePositionsParDispositif/'
const DISPOSITIF_URL_SUFFIX = '?format=json&duree=60'

function main() {
  const dispositifs = getDispositifs()
  insertRawData('DISPOSITIFS', dispositifs)

  dispositifs //
    .map(dispositif => getDispositif(dispositif)) //
    .forEach(dispositif => insertRawData(dispositif.nom, dispositif.data))

  const operations = getOperations()
  insertRawData('OPERATIONS', operations)

  updateMainSheet()
}

function getDispositifs() {
  const data = UrlFetchApp.fetch(DISPOSITIFS_URL, { headers: { 'X-noee-authorization': API_KEY } })
  const jsonData = JSON.parse(data.getContentText())
  return jsonData.listPositionsSalarieDispositifsDto.filter(dispo => dispo.mtBrut != 0)
}

function getDispositif(dispositif) {
  const data = UrlFetchApp.fetch(DISPOSITIF_URL_PREFIX + dispositif.idDispositif + DISPOSITIF_URL_SUFFIX, {
    headers: { 'X-noee-authorization': API_KEY },
  })
  const jsonData = JSON.parse(data.getContentText())
  return { nom: dispositif.libelleDispositif, data: jsonData.listegrapheHistoriqueMvDetailDto }
}

function getOperations() {
  const operations = []
  let nbOperations = 1
  let offset = 0
  while (operations.length < nbOperations) {
    const data = UrlFetchApp.fetch(OPERATIONS_URL + offset, { headers: { 'X-noee-authorization': API_KEY } })
    const jsonData = JSON.parse(data.getContentText())
    operations.push(...jsonData.operationsIndividuelles)
    nbOperations = jsonData.nbOperationsIndividuelles
    offset++
  }
  return operations.filter(ope => ope.montantNet != 0)
}

/**
 *  @param {SpreadsheetApp.Spreadsheet} doc
 *  @return {SpreadsheetApp.Sheet}
 */
function resetSheet(doc, sheetName) {
  if (doc.getSheetByName(sheetName)) {
    doc.deleteSheet(doc.getSheetByName(sheetName))
  }
  const rawSheet = doc.insertSheet()
  rawSheet.setName(sheetName)
  return rawSheet
}

function updateMainSheet() {
  const doc = SpreadsheetApp.getActiveSpreadsheet()
  let mainSheet = doc.getSheetByName('main')
  if (!mainSheet) {
    mainSheet = doc.insertSheet(0)
    mainSheet.setName('main')
  }
  doc.setActiveSheet(mainSheet)
  const fullRange = mainSheet.getRange(1, 1, 500, 50)
  recalculate(fullRange)
}

/**
 *  @param {String} sheetName
 *  @param {any[]} data
 */
function insertRawData(sheetName, data) {
  const doc = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = resetSheet(doc, sheetName)
  
  const flatData = data.map(d => flatten(d))
  const headers = Object.keys(flatData[0]).sort((a, b) => a.localeCompare(b))
  insertHeaders(sheet, headers)
  insertRows(sheet, flatData, headers)
}

function flatten(object, prefix='') {
  if (object == null || Array.isArray(object)){
    return object
  }
  let finalObject = {}
  Object.keys(object).forEach(key => {
    if (typeof object[key] == 'object' && object[key] != null && !Array.isArray(object[key])){
      finalObject = {...finalObject, ...flatten(object[key], key+'__')}
    }
    else {
      finalObject[prefix+key] = object[key]
    }
  })
  return finalObject
}

/**
 *  @param {SpreadsheetApp.Sheet} sheet
 *  @param {string[]} headers
 */
function insertHeaders(sheet, headers) {
  sheet.getRange(1,1,1,headers.length).setValues([headers])
}

/**
 *  @param {SpreadsheetApp.Sheet} sheet
 *  @param {any[]} data
 *  @param {string[]} headers
 */
function insertRows(sheet, data, headers) {
  const destinationRange = sheet.getRange(2, 1, data.length, headers.length)
  const values = data.map(row =>
    headers.map(h => {
      if (h.includes('date')) {
        return new Date(row[h])
      } else {
        return row[h]
      }
    })
  )
  destinationRange.setValues(values)
}

/**
 * @param {SpreadsheetApp.Range} range
 */
function recalculate(range) {
  var originalFormulas = range.getFormulas()
  var originalValues = range.getValues()

  var valuesToEraseFormula = []
  var valuesToRestoreFormula = []

  originalFormulas.forEach((rowValues, rowNum) => {
    valuesToEraseFormula[rowNum] = []
    valuesToRestoreFormula[rowNum] = []
    rowValues.forEach((cellValue, columnNum) => {
      const a = originalValues[rowNum][columnNum]
      if ('' === cellValue) {
        //The cell doesn't have formula
        valuesToEraseFormula[rowNum][columnNum] = originalValues[rowNum][columnNum]
        valuesToRestoreFormula[rowNum][columnNum] = originalValues[rowNum][columnNum]
      } else {
        //The cell has a formula.
        valuesToEraseFormula[rowNum][columnNum] = null
        valuesToRestoreFormula[rowNum][columnNum] = originalFormulas[rowNum][columnNum]
      }
    })
  })

  range.setValues(valuesToEraseFormula)
  range.setValues(valuesToRestoreFormula)
}
