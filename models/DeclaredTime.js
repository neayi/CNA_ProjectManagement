
let allDeclaredTimes = [];
function getDeclaredTimes() {
  if (allDeclaredTimes.length > 0) {
    return allDeclaredTimes;
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés');

  if (!sheet) {
    throw new Error("La feuille 'Temps déclarés' n'existe pas dans le classeur.");
  }

  let data = sheet.getDataRange().getValues();
  let headers = data.shift();

  allDeclaredTimes = data.map(row => new DeclaredTime(row, headers));

  return allDeclaredTimes;
}

class DeclaredTime {
  constructor(row, headers) {

    this.wp = getValue(row, headers, 'Work package');
    this.employee = getValue(row, headers, 'Collaborateur');
    this.month = getValue(row, headers, 'Mois');
    this.declaredTime = getValue(row, headers, 'Temps (PM)');
    this.project = getValue(row, headers, 'Projet');
  }
}
