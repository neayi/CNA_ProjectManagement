class WorkedTime {
    constructor(row, headers) {
        this.employee = getValue(row, headers, 'Collaborateur');
        this.month = getDateValue(row, headers, 'Mois');
        this.salary = getValue(row, headers, 'Salaire chargé réel mensuel');
        
        this.workedDays = getValue(row, headers, 'Nombre de jours travaillés');
        this.salaryPerPM = getValue(row, headers, 'Salaire chargé 1 PM (ETP)');
        this.year = getValue(row, headers, 'Année');
        this.status = getValue(row, headers, 'Statut');
        
        this.percentWorked = getValue(row, headers, 'PM Effectif');
    }

    static allWorkedTimes = [];
    static getWorkedTimes() {
        if (WorkedTime.allWorkedTimes.length > 0) {
            return WorkedTime.allWorkedTimes;
        }

        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Salaires collaborateurs');

        if (!sheet) {
            throw new Error("La feuille 'Salaires collaborateurs' n'existe pas dans le classeur.");
        }

        let data = sheet.getDataRange().getValues();
        let headers = data.shift();

        WorkedTime.allWorkedTimes = data.map(row => new WorkedTime(row, headers));

        return WorkedTime.allWorkedTimes;
    }


}
