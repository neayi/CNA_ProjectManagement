/**
 * This class models the "Budget par projet et par personne" sheet.
 */

class BudgetedTime {
    constructor(row, headers) {

        this.name = getValue(row, headers, 'Name');
        this.employee = getValue(row, headers, 'Collaborateurs');
        this.project = getValue(row, headers, 'Projet');

        this.budgetedTimes = new Map();

        headers.forEach((header, index) => {
            const found = header.match(/^P([0-9]{4})/);

            if (found != null) {
                let year = found[1]; // Remove the 'P' prefix to get the year
                this.budgetedTimes.set(year, row[index]);
            }
        });
    }

    static getBudgetedTimes() {
        if (BudgetedTime.allBudgetedTimes != undefined && BudgetedTime.allBudgetedTimes.length > 0) {
            return BudgetedTime.allBudgetedTimes;
        }

        let budgetedTimesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Budget par projet et par personne');

        if (!budgetedTimesSheet) {
            throw new Error("La feuille 'Budget par projet et par personne' n'existe pas dans le classeur.");
        }

        let data = budgetedTimesSheet.getDataRange().getValues();
        data.shift(); // helper comments
        let headers = data.shift();

        BudgetedTime.years = headers.filter(header => header.match(/P[0-9]+/)).map(header => header.substring(1));

        BudgetedTime.allBudgetedTimes = data.map(row => new BudgetedTime(row, headers));

        return BudgetedTime.allBudgetedTimes;
    }

    static getYears() {
        if (BudgetedTime.years == undefined) {
            BudgetedTime.getBudgetedTimes(); // This will populate the years
        }
        return BudgetedTime.years;
    }
    
    getBudgetedTimeForYear(year) {
        return this.budgetedTimes.get(year.toString()) || 0; // Return 0 if no budgeted time for the year
    }

    getWorkPackages() {
        return WorkPackage.getWorkPackages().filter(workPackage => {
            return workPackage.project.toLowerCase() == this.project.toLowerCase();
        });
    }

    static getEmployeesWithBudgetedTimes(projectName) {
        return BudgetedTime.getBudgetedTimes().filter(budgetedTime => {
            return budgetedTime.project.toLowerCase() === projectName.toLowerCase();
        }).map(budgetedTime => budgetedTime.employee);
    }

    static getBudgetForWPPerson(workPackageName, year) {
        let budgetedTimes = BudgetedTime.getBudgetedTimes().filter(budgetedTime => {
            return budgetedTime.name.toLowerCase() === workPackageName.toLowerCase();
        });

        if (budgetedTimes.length === 0) {
            return 0; // No budgeted time found for this work package
        }

        return budgetedTimes[0].getBudgetedTimeForYear(year);
    }
}
