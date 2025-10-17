/**
 * Model for the "Temps déclarés" sheet.
 */

class DeclaredTime {
  constructor(row, headers) {

    this.wp = getValue(row, headers, 'Work package');
    this.employee = getValue(row, headers, 'Collaborateur');
    this.month = getDateValue(row, headers, 'Mois');
    this.year = getValue(row, headers, 'Année');
    this.declaredTime = getValue(row, headers, 'Temps (PM)');
    this.project = getValue(row, headers, 'Projet');
  }

    
  static getDeclaredTimes() {
    if (DeclaredTime.allDeclaredTimes != undefined && DeclaredTime.allDeclaredTimes.length > 0) {
      return DeclaredTime.allDeclaredTimes;
    }

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés');

    if (!sheet) {
      throw new Error("La feuille 'Temps déclarés' n'existe pas dans le classeur.");
    }

    let data = sheet.getDataRange().getValues();
    let headers = data.shift();

    DeclaredTime.allDeclaredTimes = data.map(row => new DeclaredTime(row, headers));

    return DeclaredTime.allDeclaredTimes;
  }


}
