/**
 * This class models the "Work packages" sheet.
 */

class WorkPackage {
  constructor(row, headers) {

    this.name = getValue(row, headers, 'Work package');
    this.employee = getValue(row, headers, 'Nom');
    this.project = getValue(row, headers, 'Projet');
    this.employeesNames = getValue(row, headers, 'Personnes affectées').split(',').map(e => e.trim());
    this.wpCode = getValue(row, headers, 'Numero du workpackage');

    this.budgetedTimes = new Map();

    headers.forEach((header, index) => {
      const found = header.match(/^PM([0-9]{4}) prévu/);
      if (found != null) {
        let year = found[1];
        this.budgetedTimes.set(year, row[index]);
      }
    });
  }

  static getWorkPackages() {
    if (WorkPackage.allWorkPackages != undefined && WorkPackage.allWorkPackages.length > 0) {
      return WorkPackage.allWorkPackages;
    }

    let wpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Work packages');

    if (!wpSheet) {
      throw new Error("La feuille 'Work packages' n'existe pas dans le classeur.");
    }

    let data = wpSheet.getDataRange().getValues();
    data.shift(); // helper comments
    let headers = data.shift();

    WorkPackage.allWorkPackages = data.map(row => new WorkPackage(row, headers));

    return WorkPackage.allWorkPackages;
  }

  static getDeclaredTimeForYear(year) {
    return DeclaredTime.getDeclaredTimes().filter(declaredTime => {
      return declaredTime.wp.toLowerCase() === this.name.toLowerCase() && declaredTime.year == year;
    }).reduce((total, declaredTime) => total + declaredTime.declaredTime, 0);
  }

  static getWorkPackagesForProject(projectName) {
    return WorkPackage.getWorkPackages().filter(workPackage => {
      return workPackage.project.toLowerCase() === projectName.toLowerCase();
    });
  }

  static getWorkPackageForProjectAndCode(projectName, wpCode) {
    return WorkPackage.getWorkPackages().find(workPackage => {
      return workPackage.project.toLowerCase() === projectName.toLowerCase() && workPackage.wpCode.toLowerCase() === wpCode.toLowerCase();
    });
  }

  getBudgetedTimeForYear(year) {
    return this.budgetedTimes.get(year.toString()) || 0; // Return 0 if no budgeted time for the year
  }

  /**
   * Returns the budgeted time for this work package for the current year and substract the times already accounted for by each employee
   */
  getRemainingBudgetedTime(year) {
    let budgetedTime = this.getBudgetedTimeForYear(year);

    this.wpPersons.forEach(wpPerson => {
      budgetedTime -= wpPerson.budgetedTime;
    });

    return Math.max(budgetedTime, 0);
  }
}
