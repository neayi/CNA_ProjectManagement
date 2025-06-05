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

    static getWorkedTimes() {
        if (WorkedTime.allWorkedTimes != undefined && WorkedTime.allWorkedTimes.length > 0) {
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

    /**
     * Pour tous les salaires passés, présents dans l'onglet "Import Salaires", on met à jour le salaire (colonne C). 
     * On ne modifie pas le temps de travail dans le mois qui a pu être corrigé dans le cadre des apprentis.
     * 
     * Pour les salaires futurs, on créé les salaires en se basant sur les données de "Salaires à venir"
     */
    static UpdateEmployeeSalaries() {
        // Start by getting the list of all salaries in "Import Salaires"
        let importSalariesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Import Salaires');
        if (!importSalariesSheet) {
            throw new Error("La feuille 'Import Salaires' n'existe pas dans le classeur.");
        }

        let importSalariesData = importSalariesSheet.getDataRange().getValues();       
        let importSalariesHeaders = importSalariesData.shift();

        // Create a map with the employee name and the month as key, and the salary as value
        let salaryMap = new Map();
        importSalariesData.forEach(row => {
            let employee = getValue(row, importSalariesHeaders, 'Nom du salarié');
            let month = getDateValue(row, importSalariesHeaders, 'Période');
            let salary = getValue(row, importSalariesHeaders, 'Coût global CNA');

            if (salary > 0 && month) {
                let key = `${employee}-${getDateKey(month)}`;
                salaryMap.set(key, salary);
            }
        });

        // Then update the column C (Salaire chargé réel mensuel) with the new salary values from the "Salaires collaborateurs" sheet
        let salariesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Salaires collaborateurs');
        if (!salariesSheet) {
            throw new Error("La feuille 'Salaires collaborateurs' n'existe pas dans le classeur.");
        }
            
        // Remove any existing filter on the salaries sheet
        const filter = salariesSheet.getFilter();
        if (filter) {
            filter.remove();
        }

        // Reset the style for the entire sheet
        salariesSheet.getRange(2, 1, salariesSheet.getLastRow() - 1, salariesSheet.getLastColumn())
            .setFontStyle("normal");

        let salariesData = salariesSheet.getDataRange().getValues();
        let salariesHeaders = salariesData.shift();
        let updatedSalaries = [];
        salariesData.forEach(row => {
            let employee = getValue(row, salariesHeaders, 'Collaborateur');
            let month = getDateValue(row, salariesHeaders, 'Mois');
            if (!month)
                return;

            let key = `${employee}-${getDateKey(month)}`;

            if (salaryMap.has(key)) {
                updatedSalaries.push([salaryMap.get(key)]);

                // Remove the key from the map to avoid duplicates
                salaryMap.delete(key);
            } else {
                updatedSalaries.push([getValue(row, salariesHeaders, 'Salaire chargé réel mensuel')]);
            }
        });

        // Update the salaries in the sheet
        let salaryColumnIndex = salariesHeaders.indexOf('Salaire chargé réel mensuel') + 1; // +1 because getRange is 1-based
        salariesSheet.getRange(2, salaryColumnIndex, updatedSalaries.length, 1).setValues(updatedSalaries);

        // Now add the rows that were in "Import Salaires" but not in "Salaires collaborateurs"
        const lastRow = salariesSheet.getLastRow();
        let newRows = [];
        importSalariesData.forEach(row => {
            let employee = getValue(row, importSalariesHeaders, 'Nom du salarié');
            let month = getDateValue(row, importSalariesHeaders, 'Période');
            let salary = getValue(row, importSalariesHeaders, 'Coût global CNA');
            let time = getValue(row, importSalariesHeaders, 'Temps de travail');

            if (salary > 0 && month) {
                let key = `${employee}-${getDateKey(month)}`;

                // Check if this key is still in the salaryMap, meaning it was not found in the existing salaries
                // and should be added as a new row
                if (salaryMap.has(key)) {
                    let newRow = WorkedTime.getWorkedRow(employee, month, salary, time, 'A saisir !!', lastRow + newRows.length + 1);

                    newRows.push(newRow);
                }
            }
        });

        // If there are new rows, add them to the sheet
        if (newRows.length > 0) {
            // Add the new rows starting from the next empty row
            salariesSheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        }

        // Flush the changes to the sheet
        SpreadsheetApp.flush();

        // Clear the cached worked times
        WorkedTime.allWorkedTimes = [];

        // Clear the content all salaries that are in the future
        salariesData.forEach((row, rowIndex) => {
            let month = getDateValue(row, salariesHeaders, 'Mois');
            let employee = getValue(row, salariesHeaders, 'Collaborateur');

            if ((month && month > new Date()) || employee == 'A embaucher') {
                // If the month is in the future or the employee is "A embaucher", clear the content for this row
                salariesSheet.getRange(rowIndex + 2, 1, 1, salariesHeaders.length).clearContent();
            }
        });

        // Sort the sheet by month in ascending order
        salariesSheet.sort(salariesHeaders.indexOf('Mois') + 1);
        SpreadsheetApp.flush();

        // Now create the future salaries based on the "Salaires à venir" sheet
        WorkedTime.processFutureSalaries(new Date());

        SpreadsheetApp.getUi().alert("Les salaires ont été mis à jour.");
    }



    /**
     * Remplir pour tous les collaborateurs dans le premier tableau et pour tous les mois entre date début et un mois avant la date 
     * changement (ou date de fin si date de changement est vide) les lignes du tableau à modifier avec le salaire chargé et le 
     * temps de travail spécifié dans le temps de travail par mois. 
     * 
     * Attention à chaque date anniversaire spécifique au collaborateur, il faut pour chaque collaborateur procéder à l'augmentation
     * fixe précisée dans le premier tableau. 
     */
    static processFutureSalaries(startDate) {
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Salaires à venir");
        let destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Salaires collaborateurs");

        let data = sheet.getDataRange();
        let values = data.getValues();

        let headers = values.shift(); // get the titles row

        let employees = new Map();

        // Start by getting the list of employees
        values.forEach((row, rowIndex) => {
            let employeeName = getValue(row, headers, "Nom");
            let employeeData = employees.get(employeeName) || {
                'name': employeeName,
                'salaries': new Map()
            };

            let salaryRow = {};
            salaryRow.startDate = getDateValue(row, headers, "Date début");
            salaryRow.endDate = getDateValue(row, headers, "Date fin");
            salaryRow.salary = getValue(row, headers, "Salaire réel");
            salaryRow.time = getValue(row, headers, "Temps de travail");
            salaryRow.status = getValue(row, headers, "Statut");

            employeeData.salaries.set(getDateKey(salaryRow.startDate), salaryRow);
            console.log("Adding a salary row for employee " + employeeName + " (initial): ");
            console.log(salaryRow);

            // Now add a salaries item if change date is not empty:
            let changedSalaryRow = {};
            changedSalaryRow.startDate = getDateValue(row, headers, "Date nouveau salaire");
            changedSalaryRow.salary = getValue(row, headers, "Nouveau salaire");
            changedSalaryRow.time = getValue(row, headers, "Nouveau temps de travail");
            changedSalaryRow.status = salaryRow.status;

            if (!changedSalaryRow.time)
                changedSalaryRow.time = salaryRow.time;

            if (changedSalaryRow.salary > 0 && changedSalaryRow.startDate) {
                employeeData.salaries.set(getDateKey(changedSalaryRow.startDate), changedSalaryRow);
                console.log("Adding a salary row for employee " + employeeName + " (update): ");
                console.log(changedSalaryRow);
            }

            employeeData.anniversaryDate = getDateValue(row, headers, "Date anniversaire");
            employeeData.annualRaise = getValue(row, headers, "Augmentation annuelle automatique");

            employees.set(employeeName, employeeData);
        });

        // Generate new "virtual" rows for annual raises
        employees.forEach((employee, key) => {
            let lastStartDate = null;
            let lastSalary = 0;
            let lastAverageTime = 0;
            let endDate = null;

            employee.salaries.forEach(salary => {
                if (lastStartDate == null || lastStartDate > salary.startDate) {
                    lastStartDate = new Date(salary.startDate);
                    lastSalary = salary.salary;
                    lastAverageTime = salary.time;
                }

                if (endDate == null || endDate < salary.endDate)
                    endDate = new Date(salary.endDate);
            });

            for (let d = new Date(lastStartDate.getTime()); d <= endDate; d.setMonth(d.getMonth() + 1)) {
                let key = getDateKey(d);

                if (employee.salaries.has(key)) {
                    console.log("Found a custom raise for employee " + employee.number + " on date " + key);

                    let salary = employee.salaries.get(key);
                    lastSalary = salary.salary;
                    lastAverageTime = salary.time;
                    continue; // Don't generate an anniversary salary if we already have a salary change for this month
                }

                if (employee.anniversaryDate && employee.annualRaise && d.getMonth() == employee.anniversaryDate.getMonth()) {
                    // Add a virtual raise:
                    let salaryRow = {};

                    salaryRow.startDate = new Date(d.getTime());

                    // Set the start date on the 1st of the next month
                    if (employee.anniversaryDate.getDate() > 1) {
                        salaryRow.startDate.setMonth(salaryRow.startDate.getMonth() + 1);
                        salaryRow.startDate.setDate(1);
                    }

                    salaryRow.endDate = new Date(endDate.getTime());
                    lastSalary = lastSalary + employee.annualRaise;
                    salaryRow.salary = lastSalary;
                    salaryRow.time = lastAverageTime;

                    employee.salaries.set(key, salaryRow);
                    console.log("Adding a salary row for employee " + employee.number + " (anniversary): ");
                    console.log(salaryRow);
                }
            }

            console.log(employee);

            let newValues = [];

            // Find the last row in the sheet
            const lastRow = destSheet.getLastRow();

            let currentSalary = employee.salaries.values().next().value;

            // Now generate rows for each month of the employee
            for (let d = new Date(startDate.getTime()); d <= endDate; d.setMonth(d.getMonth() + 1)) {
                let key = getDateKey(d);

                if (employee.salaries.has(key))
                    currentSalary = employee.salaries.get(key);

                if (currentSalary.startDate > d)
                    continue;

                let newRow = WorkedTime.getWorkedRow(employee.name, d, currentSalary.salary, currentSalary.time, currentSalary.status, lastRow + newValues.length + 1);

                newValues.push(newRow);
            }

            // Add rows starting from the next empty row
            if (newValues.length > 0)
                destSheet.getRange(lastRow + 1, 1, newValues.length, newValues[0].length)
                    .setValues(newValues)
                    .setFontStyle("italic");

            SpreadsheetApp.flush();
        });

    }

    static getWorkedRow(name, month, salary, time, status, rowIndex) {
        let newRow = [];

        newRow.push(name);
        newRow.push(getMonthStringForDate(month));
        newRow.push(salary);
        newRow.push(time);
        newRow.push(`=C${rowIndex}/D${rowIndex}`);
        newRow.push(`=year(B${rowIndex})`);
        newRow.push(status);
        newRow.push(`=sumifs('Import Salaires'!K:K; 'Import Salaires'!C:C; A${rowIndex}; 'Import Salaires'!A:A; B${rowIndex})`);
        newRow.push(`=(vlookup(B${rowIndex}; 'Import Salaires'!M:N; 2; false) - H${rowIndex}) * D${rowIndex}`);
        newRow.push(`=I${rowIndex}/vlookup(B${rowIndex}; 'Import Salaires'!M:N; 2; false)`);
        newRow.push(`=sumifs('Import Salaires'!F:F; 'Import Salaires'!C:C; A${rowIndex}; 'Import Salaires'!A:A; B${rowIndex})`);
        newRow.push(`=sumifs('Import Salaires'!H:H; 'Import Salaires'!C:C; A${rowIndex}; 'Import Salaires'!A:A; B${rowIndex})`);
            
        return newRow;
    }
}
