/**
 * Model for the "Salaires collaborateurs" sheet.
 */
class WorkedTime {
    constructor(row, headers, mode) {
        if (mode === 'CNA') {
            this.employee = getValue(row, headers, 'Collaborateur');
        } else {
            this.employee = getValue(row, headers, 'Collaborateur nom projets');
        }
            
        this.salaryPerPM = getValue(row, headers, 'Salaire chargé 1 PM (ETP)');
        this.year = getValue(row, headers, 'Année');
        this.status = getValue(row, headers, 'Statut');

        this.month = getDateValue(row, headers, 'Mois');
        this.salary = getValue(row, headers, 'Salaire chargé réel mensuel');
        this.percentWorked = getValue(row, headers, 'PM Effectif');

        this.rtt = getValue(row, headers, 'RTT');
        this.workDays = getValue(row, headers, "Nb jours moyen à l'année travaillés");
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

        let mode = 'CNA';
        let index = headers.findIndex(item => 'n° cegid' === item.toLowerCase());
        if (index >= 0) {
            mode = 'VER DE TERRE';
        }

        WorkedTime.allWorkedTimes = data.map(row => new WorkedTime(row, headers, mode));

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

        let mode = 'CNA';
        let index = importSalariesHeaders.findIndex(item => 'code cegid salarié' === item.toLowerCase());
        if (index >= 0) {
            mode = 'VER DE TERRE';
        }

        Logger.log("Mise à jour de la feuille 'Salaires collaborateurs' avec le mode : " + mode + ' ' + importSalariesHeaders.join(', '));

        // Create a map with the employee name and the month as key, and the salary as value
        let salaryMap = new Map();
        importSalariesData.forEach(row => {
            let employee = '';
            let month = null;
            let salary = 0;
            let time = 0;

            if (mode === 'CNA') {
                // CNA
                employee = getValue(row, importSalariesHeaders, 'Nom du salarié');
                month = getDateValue(row, importSalariesHeaders, 'Période');
                salary = getValue(row, importSalariesHeaders, 'Coût global CNA');
                time = getValue(row, importSalariesHeaders, 'Temps de travail');
            } else {
                // Ver de terre
                employee = getValue(row, importSalariesHeaders, 'Code CEGID Salarié');
                month = getDateValue(row, importSalariesHeaders, 'Date');
                salary = getValue(row, importSalariesHeaders, 'Cout global (attention comprends des frais déplacements)');
                time = getValue(row, importSalariesHeaders, '%Temps travaillé');
            }

            if (salary > 0 && month) {
                let key = `${employee}-${getDateKey(month)}`;
                salaryMap.set(key, [salary, time]);
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
            let employee = '';

            if (mode === 'CNA') {
                employee = getValue(row, salariesHeaders, 'Collaborateur');
            } else {
                employee = getValue(row, salariesHeaders, 'N° CEGID');
            }

            let month = getDateValue(row, salariesHeaders, 'Mois');

            if (!month)
                return;

            let key = `${employee}-${getDateKey(month)}`;

            if (salaryMap.has(key)) {
                updatedSalaries.push(salaryMap.get(key));

                // Remove the key from the map to avoid duplicates
                salaryMap.delete(key);
            } else {
                // No update, keep the existing salary
                updatedSalaries.push([getValue(row, salariesHeaders, 'Salaire chargé réel mensuel'), getValue(row, salariesHeaders, 'Temps de travail dans le mois')]);
            }
        });

        // Update the salaries and time worked in the sheet
        let salaryColumnIndex = salariesHeaders.findIndex(item => 'salaire chargé réel mensuel' === item.toLowerCase()) + 1; // +1 because getRange is 1-based
        salariesSheet.getRange(2, salaryColumnIndex, updatedSalaries.length, 2).setValues(updatedSalaries);

        // Now add the rows that were in "Import Salaires" but not in "Salaires collaborateurs"
        const lastRow = salariesSheet.getLastRow();
        let newRows = [];
        importSalariesData.forEach(row => {
            let employee = '';
            let cegid = '';
            let month = null;
            let salary = 0;
            let time = null;

            if (mode === 'CNA') {
                employee = getValue(row, importSalariesHeaders, 'Nom du salarié');
                month = getDateValue(row, importSalariesHeaders, 'Période');
                salary = getValue(row, importSalariesHeaders, 'Coût global CNA');
                time = getValue(row, importSalariesHeaders, 'Temps de travail');
            } else {
                cegid = getValue(row, importSalariesHeaders, 'Code CEGID Salarié');
                let nom = getValue(row, importSalariesHeaders, 'Nom');
                let prenom = getValue(row, importSalariesHeaders, 'Prénom');

                employee = prenom + ' ' + nom; // Collaborateur nom projets

                month = getDateValue(row, importSalariesHeaders, 'Date');
                salary = getValue(row, importSalariesHeaders, 'Cout global (attention comprends des frais déplacements)');
                time = getValue(row, importSalariesHeaders, '%Temps travaillé');
            }
        
            if (salary > 0 && month) {
                let key = `${employee}-${getDateKey(month)}`;

                // Check if this key is still in the salaryMap, meaning it was not found in the existing salaries
                // and should be added as a new row
                if (salaryMap.has(key)) {
                    let newRow = [];

                    if (mode === 'CNA') {
                        newRow = WorkedTime.createSpreadsheetRowCNA(employee, month, salary, time, 'A saisir !!', 'RTT A saisir !!', lastRow + newRows.length + 1);
                    } else {
                        newRow = WorkedTime.createSpreadsheetRowVDT(employee, cegid, month, salary, time, 'A saisir !!', 'RTT A saisir !!', lastRow + newRows.length + 1);
                    }

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
        
        let debutDuMois = new Date();
        debutDuMois.setHours(0,0,0,0);
        debutDuMois.setDate(1);

        // Clear the content all salaries that are in the future
        salariesData.forEach((row, rowIndex) => {
            let month = getDateValue(row, salariesHeaders, 'Mois');
            let employee = getValue(row, salariesHeaders, 'Collaborateur');

            if ((month && month >= debutDuMois) ||
                 employee.toLowerCase().match(/embaucher/) ||
                 employee.toLowerCase().match(/recruter/)) {
                // If the month is in the future or the employee is "A embaucher/A recruter", clear the content for this row
                salariesSheet.getRange(rowIndex + 2, 1, 1, salariesHeaders.length).clearContent();
            }
        });

        // Sort the sheet by month in ascending order
        salariesSheet.sort(salariesHeaders.findIndex(item => 'mois' === item.toLowerCase()) + 1);
        SpreadsheetApp.flush();

        // Now create the future salaries based on the "Salaires à venir" sheet
        WorkedTime.processFutureSalaries(debutDuMois);

        // Sort the sheet by month in ascending order
        salariesSheet.sort(salariesHeaders.findIndex(item => 'mois' === item.toLowerCase()) + 1);
        SpreadsheetApp.flush();

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

        let mode = 'CNA';
        let index = headers.findIndex(item => 'n° salarié' === item.toLowerCase());
        if (index >= 0) {
            mode = 'VER DE TERRE';
        }

        let employees = new Map();

        // Start by getting the list of employees
        values.forEach((row, rowIndex) => {
            let employeeName = '';

            if (mode == 'VER DE TERRE')
                employeeName = getValue(row, headers, "Nom CEGID");
            else
                employeeName = getValue(row, headers, "Nom");

            let employeeData = employees.get(employeeName) || {
                'name': employeeName,
                'salaries': new Map()
            };

            let salaryRow = {};
            salaryRow.startDate = getDateValue(row, headers, "Date début");
            salaryRow.endDate = getDateValue(row, headers, "Date fin");
            salaryRow.salary = getValue(row, headers, "Salaire réel");
            salaryRow.time = getValue(row, headers, "Temps de travail");
            salaryRow.rtt = getValue(row, headers, "RTT");
            salaryRow.status = getValue(row, headers, "Statut");

            employeeData.salaries.set(getDateKey(salaryRow.startDate), salaryRow);
            console.log("Adding a salary row for employee " + employeeName + " (initial): ");
            console.log(salaryRow);

            // Now add a salaries item if change date is not empty:
            let changedSalaryRow = {};
            changedSalaryRow.startDate = getDateValue(row, headers, "Date nouveau salaire");
            changedSalaryRow.salary = getValue(row, headers, "Nouveau salaire");
            changedSalaryRow.time = getValue(row, headers, "Nouveau temps de travail");
            changedSalaryRow.rtt = salaryRow.rtt;
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

            if (mode === 'VER DE TERRE') {
                employeeData.cegid = getValue(row, headers, "N° salarié");
            }

            employees.set(employeeName, employeeData);
        });

        Logger.log("Employees: ");
        Logger.log([...employees]);

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

                    let salary = employee.salaries.get(key);
                    lastSalary = salary.salary;
                    lastAverageTime = salary.time;

                    console.log("Found a custom raise for employee " + employee.name + " on date " + key + " - " + salary.salary + " with time " + salary.time);

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
                    console.log("Adding a salary row for employee " + employee.name + " (anniversary): ");
                    console.log(salaryRow);
                }
            }

            console.log(employee);
            console.log([...employee.salaries]);

            // Now generate the rows in the destination sheet
            let newValues = [];

            // Find the last row in the sheet
            const lastRow = destSheet.getLastRow();

            // Get the last row in the past:
            let currentSalary = employee.salaries.values().toArray().filter(item => item.startDate <= startDate).sort((a, b) => b.startDate - a.startDate).at(0);

            console.log("Starting salary for employee " + employee.name + " on date " + currentSalary.startDate + ": ");
            console.log(currentSalary);

            // Now generate rows for each month of the employee
            for (let d = new Date(startDate.getTime()); d <= endDate; d.setMonth(d.getMonth() + 1)) {
                let key = getDateKey(d);

                if (employee.salaries.has(key))
                    currentSalary = employee.salaries.get(key);

                if (currentSalary.startDate > d)
                    continue;

                let newRow = [];
                if (mode === 'CNA') {
                    newRow = WorkedTime.createSpreadsheetRowCNA(employee.name, d, currentSalary.salary, currentSalary.time, currentSalary.status, currentSalary.rtt, lastRow + newValues.length + 1, true);
                } else {
                    newRow = WorkedTime.createSpreadsheetRowVDT(employee.name, employee.cegid, d, currentSalary.salary, currentSalary.time, currentSalary.status, currentSalary.rtt, lastRow + newValues.length + 1, true);
                }

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

    static createSpreadsheetRowCNA(name, month, salary, time, status, rtt, rowIndex, isFuture = false) {
        let newRow = [];

        newRow.push(name);
        newRow.push(getMonthStringForDate(month));
        newRow.push(salary);
        newRow.push(time);
        newRow.push(`=C${rowIndex}/D${rowIndex}`);
        newRow.push(`=year(B${rowIndex})`);
        newRow.push(status);

        if (isFuture) {
            newRow.push('');
            newRow.push('');
            newRow.push(`=D${rowIndex}`); // PM Effectif
            
            newRow.push('');
        } else {
            newRow.push(`=sumifs('Import Salaires'!K:K; 'Import Salaires'!C:C; A${rowIndex}; 'Import Salaires'!A:A; B${rowIndex})`); // Jours d'absence (RTT, Vacances, Maladies)
            newRow.push(`=(vlookup(B${rowIndex}; 'Import Salaires'!M:N; 2; false) - H${rowIndex}) * D${rowIndex}`); // Nb de jours de présence
            newRow.push(`=I${rowIndex}/N${rowIndex}`); // PM Effectif
            
            newRow.push(`=sumifs('Import Salaires'!F:F; 'Import Salaires'!C:C; A${rowIndex}; 'Import Salaires'!A:A; B${rowIndex})`); // Salaire effectif
        }

        newRow.push(rtt); // RTT
        newRow.push(`=(filter('Import Salaires'!$T$2:$T;'Import Salaires'!$R$2:$R=F${rowIndex})-25-L${rowIndex})/12`); // Nb jours moyen à l'année travaillés

        return newRow;
    }

    static createSpreadsheetRowVDT(name, cegid, month, salary, time, status, rtt, rowIndex, isFuture = false) {
        let newRow = [];
      
        newRow.push(titleCase(name));        // Collaborateur nom projets
        newRow.push(cegid);   // N° CEGID
        newRow.push(name);   // Collaborateur
        newRow.push(getMonthStringForDate(month));       // Mois
        newRow.push(salary);    // Salaire chargé réel mensuel
        newRow.push(time);        // Temps de travail dans le mois
        newRow.push(`=IF(F${rowIndex} = 0; 0; E${rowIndex}/F${rowIndex})`);        // Salaire chargé 1 PM (ETP)

        if (isFuture) {
            newRow.push(`=IF(D${rowIndex}<=DATE(2023;8;31); YEAR(D${rowIndex}); IF(MONTH(D${rowIndex}) <= 8; RIGHT(YEAR(D${rowIndex})-1; 2) & RIGHT(YEAR(D${rowIndex}); 2); RIGHT(YEAR(D${rowIndex}); 2) & RIGHT(YEAR(D${rowIndex})+1; 2)))`); // Année
            newRow.push(status);    // Statut
            newRow.push('');        // Nb jours travaillés
            newRow.push(`=F${rowIndex}`); //  PM Effectif
        } else {
            newRow.push(`=IF(D${rowIndex}<=DATE(2023;8;31); YEAR(D${rowIndex}); IF(MONTH(D${rowIndex}) <= 8; RIGHT(YEAR(D${rowIndex})-1; 2) & RIGHT(YEAR(D${rowIndex}); 2); RIGHT(YEAR(D${rowIndex}); 2) & RIGHT(YEAR(D${rowIndex})+1; 2)))`); // Année
            newRow.push(status);        // Statut
            newRow.push(`=FILTER('Import salaires'!S:S;'Import salaires'!A:A=D${rowIndex}*1;'Import salaires'!B:B=B${rowIndex})`);  // Nb jours travaillés
            newRow.push(`=FILTER('Import salaires'!P:P;'Import salaires'!A:A=D${rowIndex}*1;'Import salaires'!B:B=B${rowIndex})`);  // PM Effectif
        }

        newRow.push(rtt); // RTT
        newRow.push(`=(filter('Jours ouvrés par mois'!K:K;'Jours ouvrés par mois'!G:G=H${rowIndex}*1)-25-L${rowIndex})/12`); // Nb jours moyen à l'année travaillés

        // Heures travaillées dans le mois réel qui sert à remplir les temps
        // Heures travaillées officiels CEGID qui sert à calculer les coûts FEADER
        // Vérif salaire brut : ok !
        // Vérifs
        // Données venant de suivi salaires
        // Nb jours moyen à l'année travaillés

        return newRow;
    }

}
