class Employee {
    constructor(row, headers) {

        this.name = getValue(row, headers, 'Collaborateur');
        this.startDate = getDateValue(row, headers, 'Entrée');
        this.endDate = getDateValue(row, headers, 'Sortie');

        this.declaredTimes = new Map();
    }

    static getEmployees() {
        if (Employee.allEmployees != undefined && Employee.allEmployees.length > 0) {
            return Employee.allEmployees;
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

        Employee.allEmployees = data.map(row => new Employee(row, headers));

        return Employee.allEmployees;
    }

    hasWorkedBetween(startDate, endDate) {
        if (this.startDate && this.startDate > endDate)
            return false;

        if (this.endDate && this.endDate < startDate)
            return false;

        return true;
    }

    getBudgetedTimesOnProjects(startDate, endDate) {
        let budgetedTimes = BudgetedTime.getBudgetedTimes();

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
        let workedTime = WorkedTime.getWorkedTimes().filter(workedTime => {
            return workedTime.employee === this.name && workedTime.month.getMonth() === (month - 1) && workedTime.month.getFullYear() === year;
        }).at(0);

        if (workedTime !== undefined) {
            return workedTime.percentWorked || 0;
        }

        return 0;
    }

    getDeclaredTimeForYearAndProject(year, project) {
        return DeclaredTime.getDeclaredTimes().filter(declaredTime => {
            return declaredTime.employee === this.name && declaredTime.month.getFullYear() === year && declaredTime.project === project;
        }).reduce((total, declaredTime) => total + declaredTime.declaredTime, 0);
    }

    getDeclaredTimeForYearAndWorkPackage(year, workPackage) {
        return DeclaredTime.getDeclaredTimes().filter(declaredTime => {
            return declaredTime.employee === this.name && declaredTime.month.getFullYear() === year && declaredTime.wp === workPackage;
        }).reduce((total, declaredTime) => total + declaredTime.declaredTime, 0);
    }

    /**
     * 
     * @param {*} month in human notation (1-12)
     * @param {*} year 
     * @returns 
     */
    getDeclaredTimeForMonth(month, year) {
        return DeclaredTime.getDeclaredTimes().filter(declaredTime => {
            return declaredTime.employee === this.name &&
                declaredTime.month.getMonth() === (month - 1) &&
                declaredTime.month.getFullYear() === year;
        }).reduce((total, declaredTime) => total + declaredTime.declaredTime, 0);
    }

    /**
     * Returns the budgeted time for this work package for the current year and substract the times already accounted for by each WP
     */
    getRemainingBudgetedTime() {
        let budgetedTime = this.budgetTimeForYear;

        this.wpPersons.forEach(wpPerson => {
            budgetedTime -= wpPerson.budgetedTime;
        });

        return Math.max(budgetedTime, 0);
    }

    /**
     * Same as getRemainingBudgetedTime, but we divide the remaining budgeted time by the number of work packages that have times left to fill up
     */
    getAverageRemainingBudgetedTime() {
        let remainingWPCount = this.wpPersons.filter(wpPerson => (wpPerson.budgetedTime == 0 && wpPerson.isTarget)).length;
        if (remainingWPCount === 0) return 0;

        return this.getRemainingBudgetedTime() / remainingWPCount;
    }
}

