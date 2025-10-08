/**
 * Model for the "Import des ordres de missions" sheet.
 */

class Mission {
  constructor(row, headers) {

    this.employee = getValue(row, headers, 'Nom complet');
    this.project = getValue(row, headers, 'Nom du projet');
    this.wpCode = getValue(row, headers, 'Work package');

    this.dateStart = getDateValue(row, headers, 'Date de début déplacement');
    this.dateEnd = getDateValue(row, headers, 'Date de fin déplacement');

    if (this.dateStart != null) {
      this.month = new Date(this.dateStart.getTime());
      this.month.setDate(1); // Set to the first day of the month
    }
    else {
      this.month = null; // If dateStart is null, set month to null
    }

    this.workPackage = WorkPackage.getWorkPackageForProjectAndCode(this.project, this.wpCode);
  }

  static getMissions() {
    if (Mission.allMissions != undefined && Mission.allMissions.length > 0) {
      return Mission.allMissions;
    }

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Import des ordres de missions');

    if (!sheet) {
      throw new Error("La feuille 'Import des ordres de missions' n'existe pas dans le classeur.");
    }

    let data = sheet.getDataRange().getValues();
    let headers = data.shift();

    Mission.allMissions = data.map(row => new Mission(row, headers));

    return Mission.allMissions;
  }

  static getMissionsForEmployee(employeeName, workPackageName, year) {
    return Mission.getMissions().filter(mission => {
      return mission.employee.toLowerCase() === employeeName.toLowerCase() &&
             mission.workPackage != null &&
             mission.workPackage.name.toLowerCase() === workPackageName.toLowerCase() &&
             mission.month != null &&
             mission.month.getFullYear() === year;
    });
  }

}
