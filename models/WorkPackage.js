class WorkPackage {
  constructor(row, headers) {

    this.name = getValue(row, headers, 'Work package');
    this.employee = getValue(row, headers, 'Nom');
    this.project = getValue(row, headers, 'Projet');

    this.budgetedTimes = new Map();

    headers.forEach((header, index) => {
      const found = header.match(/^PM([0-9]{4}) prévu/);
      if (found != null) {
        let year = found[1];
        this.budgetedTimes.set(year, row[index]);
      }
    });
  }
    
  static allWorkPackages = [];
  static getWorkPackages() {
    if (WorkPackage.allWorkPackages.length > 0) {
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


  getBudgetedTimeForYear(year) {
    return this.budgetedTimes.get(year.toString()) || 0; // Return 0 if no budgeted time for the year
  }

  getDeclaredTimeForYear(year) {
    return DeclaredTime.getDeclaredTimes().filter(declaredTime => {
      return declaredTime.wp === this.name && declaredTime.month.getFullYear() === year;
    }).reduce((total, declaredTime) => total + declaredTime.declaredTime, 0);
  }
}
