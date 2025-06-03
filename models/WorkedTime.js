let allWorkedTimes = [];
function getWorkedTimes() {
    if (allWorkedTimes.length > 0) {
        return allWorkedTimes;
    }

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Salaires collaborateurs');

    if (!sheet) {
        throw new Error("La feuille 'Salaires collaborateurs' n'existe pas dans le classeur.");
    }

    let data = sheet.getDataRange().getValues();
    let headers = data.shift();

    allWorkedTimes = data.map(row => new WorkedTime(row, headers));

    return allWorkedTimes;
}

class WorkedTime {
    constructor(row, headers) {
        this.employee = getValue(row, headers, 'Collaborateur');
        this.month = getDateValue(row, headers, 'Mois');
        this.salary = getValue(row, headers, 'Salaire chargé réel mensuel');
        this.percentWorked = getValue(row, headers, 'Temps de travail dans le mois');
        this.workedDays = getValue(row, headers, 'Nombre de jours travaillés');
        this.salaryPerPM = getValue(row, headers, 'Salaire chargé 1 PM (ETP)');
        this.year = getValue(row, headers, 'Année');
        this.status = getValue(row, headers, 'Statut');
    }
}
