let allEmployees = [];
function getEmployees() {
    if (allEmployees.length > 0) {
        return allEmployees;
    }

    // This function should return an array of Employee objects that have worked on the projects between the two dates.
    // For now, we will return an empty array.
    let employeesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Collaborateurs');
    if (!employeesSheet) {
        throw new Error("La feuille 'Collaborateurs' n'existe pas dans le classeur.");
    }

    let data = employeesSheet.getDataRange().getValues();
    data.shift(); // helper comments
    let headers = data.shift();

    allEmployees = data.map(row => new Employee(row, headers));

    return allEmployees;
}


class Employee {
    constructor(row, headers) {

        this.name = getValue(row, headers, 'Collaborateur');
        this.startDate = getDateValue(row, headers, 'Entrée');
        this.endDate = getDateValue(row, headers, 'Sortie');

        this.declaredTimes = new Map();
    }

    hasWorkedBetween(startDate, endDate) {
        if (this.startDate && this.startDate > endDate)
            return false;

        if (this.endDate && this.endDate < startDate)
            return false;

        return true;
    }

    getBudgetedTimesOnProjects(startDate, endDate) {
        let budgetedTimes = getBudgetedTimes();

        return budgetedTimes.filter(budgetedTime => {
            if (budgetedTime.employee == this.name) {
                for (let year = startDate.getFullYear(); year <= endDate.getFullYear(); year++) {
                    if (budgetedTime.budgetedTimes.has(year.toString())) {
                        return true;
                    }
                }
            }

            return false;
        });
    }

    /**
     * 
     * @param {*} month in human notation (1-12)
     * @param {*} year 
     * @returns 
     */
    getWorkedTime(month, year) {
        let workedTime = getWorkedTimes().filter(workedTime => {
            return workedTime.employee === this.name && workedTime.month.getMonth() === (month - 1) && workedTime.month.getFullYear() === year;
        }).at(0);

        if (workedTime !== undefined) {
            return workedTime.percentWorked || 0;
        }

        return 0;
    }

    getDeclaredTimeForYearAndProject(year, project) {
        return getDeclaredTimes().filter(declaredTime => {
                return declaredTime.employee === this.name && declaredTime.month.getFullYear() === year && declaredTime.project === project;
            }).reduce((total, declaredTime) => total + declaredTime.declaredTime, 0);
    }

    /**
     * 
     * @param {*} month in human notation (1-12)
     * @param {*} year 
     * @returns 
     */
    getDeclaredTimeForMonth(month, year) {
        return getDeclaredTimes().filter(declaredTime => {
                return declaredTime.employee === this.name && 
                       declaredTime.month.getMonth() === (month - 1) &&
                       declaredTime.month.getFullYear() === year;
            }).reduce((total, declaredTime) => total + declaredTime.declaredTime, 0);
    }
}

