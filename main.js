const MORed = [245, 117, 117];
const MOGray = [189, 189, 189];

// Finds the sheet by name and returns the sheet object
function getSheet(sheetName) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  //console.log(sheet);
  return sheet;
}

// Gets all rows in the active data range
function getRows(sheet) {
  let rows = sheet.getDataRange().getValues();;

  // Add row index at the beginning of each element
  let i = 1
  rows.forEach( row => {
    row.unshift(i++);
  });

  return rows
}

// Gets all rows with operations in them
function getOperations(sheet) {
  let rows = getRows(sheet);
  rows.shift(); //Remove heading row

  let fRows = [];
  rows.forEach( row => {
    if (row[3] != "") fRows.push(row); //Remove operations without a Zeus
 });
  return fRows;
}

// Changes the background of the given row index. STARTS FOR 1 AND NOT 0
function changeRowBackground(sheet, index, color) {
  //let activeRange = sheet.getDataRange();
  let changeRange = sheet.getRange(index,1,1,sheet.getLastColumn());
  changeRange.setBackgroundRGB(color[0], color[1], color[2]);

  return true;
}

// Refreshes the filter that hides past operations
function refreshFilters(sheetName) {
  console.log("Refreshing filters...");

  let col = 2;
  let filter = getSheet(sheetName).getFilter();
  let criteria = filter.getColumnFilterCriteria(col); //Criteria is the condition for the filter
  //console.log(criteria.getCriteriaValues());
  filter.setColumnFilterCriteria(col, criteria);

  console.log("Filters refreshed");
  return true;
}

// Checks whether the given date is in the past. Accuracy is down to one day.
// Input: Date object
function isInThePast(date) {
  const now = new Date();

  //console.log(date.toString() + " | " + now.toString())
  if (date.getUTCFullYear() < now.getUTCFullYear()) return true
  if (date.getUTCMonth() < now.getUTCMonth()) return true
  if (date.getUTCDate() + 1 < now.getUTCDate()) return true
  return false;
};

// Updates and changes the color of the past and upcoming operation
function updateHighlight(sheetName) {
  console.log("Updating row colors...");
  const sheet = getSheet(sheetName);
  //console.log(sheet);

  let rows = getOperations(sheet);
  rows.every( (row) => {
    if (isInThePast(row[2])) {
      changeRowBackground(sheet, row[0], MOGray);
      return true;
    };
    console.log(`NEXT OPERATION: ${row[4]} by ${row[3]}`);
    changeRowBackground(sheet, row[0], MORed);
    return false;
 });

  console.log("Row colors updated");
  return true;
}

function main() {
  /*
  refreshFilters("test");
  updateHighlight("test");
  */
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach( sheet => {
    console.log("---------------------------");
    console.log(`Updating sheet ${sheet.getSheetName()}`)
    refreshFilters(sheet.getSheetName());
    updateHighlight(sheet.getSheetName());
  });

  console.log("ALL SHEETS UPDATED");
  return true;
}

main();

