/**
 * Model for the "Temps déclarés" sheet.
 */

class Project {
  constructor(row, headers) {

    this.project = getValue(row, headers, 'Projet');
  }

    
  static getProjects() {
    if (Project.allProjects != undefined && Project.allProjects.length > 0) {
      return Project.allProjects;
    }

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projets');

    if (!sheet) {
      throw new Error("La feuille 'Projets' n'existe pas dans le classeur.");
    }

    let data = sheet.getDataRange().getValues();
    data.shift(); // helper comments
    let headers = data.shift();

    Project.allProjects = data.map(row => new Project(row, headers));

    return Project.allProjects;
  }


}
