/**
 * This class models the "Reporting periods" sheet.
 */

class ReportingPeriod {
  constructor(row, headers) {

    this.project = getValue(row, headers, 'Projet');
    this.start = getDateValue(row, headers, 'Date début');
    this.end = getDateValue(row, headers, 'Date de fin');
    this.name = getValue(row, headers, 'Reporting Period').trim();
  }

  static getReportingPeriods() {
    if (ReportingPeriod.allReportingPeriods != undefined && ReportingPeriod.allReportingPeriods.length > 0) {
      return ReportingPeriod.allReportingPeriods;
    }

    let rpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reporting periods');

    if (!rpSheet) {
      throw new Error("La feuille 'Reporting periods' n'existe pas dans le classeur.");
    }

    let data = rpSheet.getDataRange().getValues();
    let headers = data.shift();

    ReportingPeriod.allReportingPeriods = data.map(row => new ReportingPeriod(row, headers)).filter(row => row.name != '');

    return ReportingPeriod.allReportingPeriods;
  }
}
