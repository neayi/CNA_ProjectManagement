function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Génération des temps')
        .addItem('Générer les temps déclarés', 'CreateDeclaredTimes')
        .addItem('Mettre à jour les salaires collaborateurs', 'UpdateEmployeeSalaries')
        .addToUi();
}


/**
 * Fill the "Temps déclarés" sheet with rows that match the work packages and the people who worked on them.
 */
function CreateDeclaredTimes() {
    const html = HtmlService.createHtmlOutputFromFile('interfaces/SelectDates.html')
        .setWidth(500)
        .setHeight(550);

    SpreadsheetApp.getUi().showModalDialog(html, 'Choisir les dates de début et de fin');
}

function generateTimesForDates(startDate, endDate, deleteExistingTimes, projectName) {
    startDate = new Date(startDate);
    endDate = new Date(endDate);

    const currentUser = getCurrentUser();
    const currentDate = new Date().toLocaleDateString('fr-FR');

    Logger.log("Génération des temps déclarés pour les dates entre " + startDate + " et " + endDate + " (deleteExistingTimes = " + deleteExistingTimes + ") par " + currentUser);

    let declaredTimesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés test');

    // On garde la dernière ligne générée pour récupérer les formules calculées
    const lastGoodFormulasRowIndex = declaredTimesSheet.getLastRow(); // On pourrait améliorer en vérifiant que les formules ne sont effectivement pas vides

    // On copie les formules calculées pour les utiliser plus tard au moment de générer de nouvelles lignes
    let formulasFJ = declaredTimesSheet.getRange(lastGoodFormulasRowIndex, 6, 1, 5).getFormulasR1C1();
    let formulasLU = declaredTimesSheet.getRange(lastGoodFormulasRowIndex, 12, 1, 10).getFormulasR1C1();

    // If deleteExistingTimes is true, we will delete the existing times in the "Temps déclarés" sheet.
    if (deleteExistingTimes) {
        let declaredTimesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés test');
        if (declaredTimesSheet) {

            let existingDeclaredTimes = DeclaredTime.getDeclaredTimes();

            for (let i = existingDeclaredTimes.length - 1; i >= 1; i--) {
                if (isDateInRange(existingDeclaredTimes[i].month, startDate, endDate) &&
                    existingDeclaredTimes[i].project.toLowerCase() === projectName.toLowerCase()) {
                    declaredTimesSheet.deleteRow(i + 2); // +2 because the first row is the header and the index is 0-based
                }
            }

            Logger.log("Les temps déclarés existants ont été supprimés.");
        }
    }

    // Start by getting the budgeted times for each work package and each person, per year.
    let employeesNames = BudgetedTime.getEmployeesWithBudgetedTimes(projectName);

    let employees = new Map();

    Employee.getEmployees().forEach(employee => {
        if (employeesNames.findIndex(item => employee.name.toLowerCase() === item.toLowerCase()) !== -1 &&
            employee.hasWorkedBetween(startDate, endDate)
            // && employee.name == "Laurence Fontaine";
        ) {
            employees.set(employee.name, employee);
    }});

    // On trie les employés par salaires (décroissant) pour charger avant tout ceux avec un plus gros salaire et éviter de déclarer des temps sur des stagiaires par exemple
    employees = new Map([...employees.entries()].sort((a, b) => b[1].salary - a[1].salary));

    // On créé une liste de workpackages, qui nous seront nécessaires pour manipuler les données de wpPersons au niveau de chaque workpackage
    let workPackages = WorkPackage.getWorkPackagesForProject(projectName);

    Logger.log("Found " + employees.size + " employees who have worked between " + startDate + " and " + endDate);
    Logger.log([...employees]);

    BudgetedTime.getYears().filter(yearString => isDateInRange(startDate, endDate, yearString)).forEach(yearString => {

        // On travaille année par année. 
        flushSpreadsheetAndCache();

        // On créé une matrice d'objets WorkPackage/employee qui contiendra les temps déclarés pour chaque work package et chaque employé.
        let wpPersons = new Map();

        workPackages.forEach(workPackage => {
            workPackage.wpPersons = new Array();
        });

        employees.forEach(employee => {
            employee.wpPersons = new Array();

            // On récupère les temps prévus pour ce projet pour chaque employé (Budget par projet et par personne)
            let budgetKey = employee.name + ' - ' + projectName;
            employee.budgetTimeForYear = BudgetedTime.getBudgetForWPPerson(budgetKey, yearString);
        });

        // On commence par initialiser wpPersons:
        employees.forEach(employee => {
            workPackages.forEach(workPackage => {
                let wpPerson = new DeclaredTimePerPerson(employee.name, workPackage.name, projectName, yearString);
                wpPersons.set(wpPerson.getKey(), wpPerson);
                employee.wpPersons.push(wpPerson);
                workPackage.wpPersons.push(wpPerson);

                // Ensuite, si le temps prévu pour ce projet et cet employé est zéro, alors on met le temps déclaré à zéro pour chacun des work packages/employé
                if (employee.budgetTimeForYear === 0) {
                    wpPerson.setAsNotWorked();
                }
            });
        });

        // Pour chaque work package, on regarde les personnes qui bossent dessus. On enlève les personnes qui n'ont pas de temps prévu sur le projet.
        workPackages.forEach(workPackage => {
            workPackage.employees = new Map();
            workPackage.employeesNames.forEach(employeeName => {
                let employee = employees.get(employeeName);
                if (employee && employee.budgetTimeForYear > 0) {
                    workPackage.employees.set(employeeName, employee);

                    let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, yearString));
                    if (wpPerson)
                        wpPerson.isTarget = true; // This is the target person for the work package
                }
            });

            // Si il n'y a qu'une personne pour ce WP, alors on lui affecte le temps prévu pour le WP sur l'année.
            if (workPackage.employees.size === 1) {
                let employeeName = workPackage.employees.keys().next().value;
                let employee = employees.get(employeeName);
                if (employee) {
                    let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, yearString));
                    if (wpPerson) {
                        wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(yearString), employee.getRemainingBudgetedTime()));
                    }
                }
            }
        });

        // Ensuite, pour chaque personne, on ajoute des temps à chaque WP, dans la limite du temps prévu pour cette personne et pour ce WP
        employees.forEach(employee => {
            const averageRemainingBudgetedTime = employee.getAverageRemainingBudgetedTime();

            workPackages.forEach(workPackage => {
                if (!workPackage.employees.has(employee.name))
                    return; // Si l'employé n'est pas prévu sur ce WP, on passe au suivant

                let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, yearString));
                if (wpPerson && wpPerson.budgetedTime == 0) {
                    wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(yearString), averageRemainingBudgetedTime));
                }
            });
        });

        // Deuxième passe, cette fois on complète les temps pour chaque personne, en fonction du temps restant à déclarer pour chaque WP.
        employees.forEach(employee => {
            workPackages.forEach(workPackage => {
                if (!workPackage.employees.has(employee.name))
                    return; // Si l'employé n'est pas prévu sur ce WP, on passe au suivant

                let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, yearString));
                if (wpPerson) {
                    wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(yearString), employee.getRemainingBudgetedTime()));
                }
            });
        });

        // S'il reste du temps à déclarer pour une personne, on l'affecte aux WP qui ne sont pas encore remplis, même si la personne n'était pas prévue au départ
        employees.forEach(employee => {
            workPackages.forEach(workPackage => {
                let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, yearString));
                if (wpPerson) {
                    wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(yearString), employee.getRemainingBudgetedTime()));
                }
            });
        });

        // On affiche les résultats dans la console pour debug
        debugLog(employees, workPackages, wpPersons, projectName, yearString);

        // Maintenant, on va ajouter les temps déclarés dans la feuille "Temps déclarés"
        employees.forEach(employee => {

            workPackages.forEach(workPackage => {
                const key = DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, yearString);
                let wpPerson = wpPersons.get(key);

                if (wpPerson == false || wpPerson.budgetedTime == 0)
                    return;

                console.log("Adding declared time for " + employee.name + " on " + workPackage.name + " for year " + yearString + ": " + debugRound(wpPerson.budgetedTime));

                let WPYearlyMissingTime = wpPerson.budgetedTime - employee.getDeclaredTimeForYearAndWorkPackage(yearString, workPackage.name);

                if (WPYearlyMissingTime <= 0) {
                    console.log("Skipping " + employee.name + " on " + workPackage.name + " for year " + yearString + " because no missing time. Declared time for year and work package: " + debugRound(employee.getDeclaredTimeForYearAndWorkPackage(yearString, workPackage.name)));
                    return; // Si le temps déclaré est déjà supérieur ou égal au temps prévu, on ne fait rien
                }

                let rows = [];

                let months = new Set(); // de 1 à 12

                // Commencer par identifier les ordres de mission pour ce salarié sur ce projet dans cette année
                let missions = Mission.getMissionsForEmployee(employee.name, workPackage.name, yearString);
                missions.forEach(mission => {
                    const m = mission.month.getMonth() + 1;

                    if (m >= getStartMonth(yearString, startDate) && m <= getEndMonth(yearString, endDate)) {
                        months.add(m);
                    }
                });

                // Ajouter les autres mois de l'année
                for (let m = getStartMonth(yearString, startDate); m <= getEndMonth(yearString, endDate); m++) {
                    months.add(m);
                }

                for (const m of months) {
                    const year = getYearForMonth(m, yearString);

                    if (notInRange(m, year, startDate, endDate)) {
                        continue;
                    }

                    let workedTime = employee.getWorkedTime(m, yearString); // Temps travaillé par le salarié
                    let declaredTimeForMonth = employee.getDeclaredTimeForMonth(m, yearString); // temps déjà déclaré pour ce mois
                    let remainingTimeForMonth = workedTime - declaredTimeForMonth; // Temps restant à déclarer pour ce mois

                    if (remainingTimeForMonth <= 0) {
                        console.log("Skipping " + employee.name + " for month " + m + " of year " + yearString + " because no remaining time. Worked time: " + debugRound(workedTime) + ", already declared time: " + debugRound(declaredTimeForMonth));
                        continue; // Si le temps travaillé est inférieur ou égal au temps déjà déclaré, on ne fait rien
                    }

                    let newDeclaredTime = Math.min(WPYearlyMissingTime, remainingTimeForMonth); // On ne peut pas déclarer plus que le temps travaillé moins le temps déjà déclaré

                    console.log("Declaring time for " + employee.name + " for month " + m + ", declared time for month: " + debugRound(newDeclaredTime));

                    rows.push([
                        workPackage.name,
                        employee.name,
                        "01/" + m + "/" + year,
                        0, // This is in hours - ignored for the moment
                        newDeclaredTime,
                        '', // Planifié dans l'année - calculated
                        '', // 	Total déclaré dans l'année - calculated
                        '', // 	Reste à déclarer dans l'année - calculated
                        '', // 	Dispo pour ce collaborateur dans le mois - calculated
                        '', // 	Reste dispo pour ce collaborateur dans le mois - calculated
                        projectName, // Projet
                        '', // Reporting period
                        '', // Année
                        '', // Salaire
                        '', // Salaire ETP
                        '', // Coût réel
                        '', // FTE
                        '', // Statut
                        '', // Temps travaillé en heure sur la période sur  tous les projets
                        '', // Daily rate
                        '', // Coût
                        currentDate, // Date de génération
                        currentUser  // Acteur de la génération
                    ]);

                    WPYearlyMissingTime -= newDeclaredTime;

                    if (WPYearlyMissingTime <= 0) {
                        break; // No more time to declare for this work package or this project
                    }
                }

                if (rows.length > 0) {
                    // If we have rows to add, we need to add them to the "Temps déclarés" sheet.

                    const newRowIndex = declaredTimesSheet.getLastRow() + 1;
                    // Add the new rows
                    declaredTimesSheet.getRange(newRowIndex, 1, rows.length, rows[0].length).setValues(rows);

                    rows.forEach((row, index) => {
                        declaredTimesSheet.getRange(newRowIndex + index, 6, 1, 5).setFormulasR1C1(formulasFJ);
                        declaredTimesSheet.getRange(newRowIndex + index, 12, 1, 10).setFormulasR1C1(formulasLU);
                    });

                    // Flush the changes to the sheet
                    flushSpreadsheetAndCache(true);
                    Logger.log("Ajout de " + rows.length + " lignes dans la feuille 'Temps déclarés' pour l'employé " + employee.name + " et le projet " + projectName);
                }
            });
        });
    });    // For each year

    SpreadsheetApp.getUi().alert("Les temps ont été générés.");
}

function debugLog(employees, workPackages, wpPersons, projectName, yearString) {

    let row = ["Tps travaillé".padEnd(25, " ")];
    workPackages.forEach(workPackage => {
        
        let name = workPackage.name;

        let wp = name.match(/wp[^ ]+/i);
        if (wp) {
            name = wp[0];
        }

        row.push(name);
    });

    row.push(' for year ' + yearString);

    console.log(row.join("\t"));

    employees.forEach(employee => {
        let row = [employee.name.padEnd(25, ".")];

        workPackages.forEach(workPackage => {
            let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, yearString));
            if (wpPerson) {
                row.push(debugRound(wpPerson.budgetedTime));
            } else {
                row.push(0);
            }
        });

        row.push('');
        row.push(debugRound(employee.budgetTimeForYear));

        console.log(row.join("\t"));
    });
}

function debugRound(value) {
    value = Math.round(value * 100) / 100;

    return String(value).replace('.', ',');
}

// utilities


/**
 * Takes a year (in both format 2023 or 2334) and a date. If the year matches the year of the date, returns the month of the date (1-12). Otherwise returns 1.
 * For dates that are in the format 2324, in order to return the month of startDate, it must be between September of the first year and August of the next year.
 * 
 * @param {*} yearString 
 * @param {*} startDate 
 * @returns Integer between 1 and 12
 */
function getStartMonth(yearString, startDate) {
    if (yearString > 2300) {
        // yearString is in the format 2324
        yearString = 2000 + (yearString % 100) - 1;

        if (startDate.getMonth() + 1 >= 9) { // if startDate is in September or later, it belongs to the first year
            if (yearString === startDate.getFullYear())
                return startDate.getMonth() + 1;
            else
                return 1;
        } else { // if startDate is before September, it belongs to the second yearString
            if (yearString + 1 === startDate.getFullYear())
                return startDate.getMonth() + 1;
            else
                return 1;
        }
    }

    return (yearString === startDate.getFullYear()) ? startDate.getMonth() + 1 : 1; // Months are 0-indexed in JavaScript
}

/**
 * Takes a year (in both format 2023 or 2334) and a date. If the year matches the year of the date, returns the month of the date (1-12). Otherwise returns 12.
 * For dates that are in the format 2324, in order to return the month of startDate, it must be between September of the first year and August of the next year.
 * 
 * @param {*} yearString 
 * @param {*} startDate 
 * @returns Integer between 1 and 12
 */
function getEndMonth(yearString, endDate) {
    if (yearString > 2300) {
        // yearString is in the format 2324
        yearString = 2000 + (yearString % 100) - 1;

        if (endDate.getMonth() + 1 >= 9) { // if endDate is in September or later, it belongs to the first year
            if (yearString === endDate.getFullYear())
                return endDate.getMonth() + 1;
            else
                return 12;
        } else { // if endDate is before September, it belongs to the second year
            if (yearString + 1 === endDate.getFullYear())
                return endDate.getMonth() + 1;
            else
                return 12;
        }
    }

    return (yearString === endDate.getFullYear()) ? endDate.getMonth() + 1 : 12; // Months are 0-indexed in JavaScript
}

function getYearForMonth(m, yearString) {
    if (yearString > 2300) {
        // yearString is in the format 2324
        let firstYear = 2000 + (yearString % 100) - 1;
        let secondYear = firstYear + 1;

        if (m >= 9) {
            return firstYear;
        } else {
            return secondYear;
        }
    }

    return yearString;
}

function getValue(row, headers, field) {
    let index = headers.findIndex(item => field.toLowerCase() === item.toLowerCase());

    if (index === -1) {
        throw new Error(`Le champ '${field}' n'existe pas dans les en-têtes.`);
    }
    return row[index];
}

function getDateValue(row, headers, field) {
    let index = headers.findIndex(item => field.toLowerCase() === item.toLowerCase());
    if (index === -1) {
        throw new Error(`Le champ '${field}' n'existe pas dans les en-têtes.`);
    }
    
    let date = row[index] == '' ? null : new Date(row[index]);

    // if the date is not null but invalid, try to parse it from a string in dd/mm/yyyy format
    if (date != null && isNaN(date.getTime())) {
        let parts = row[index].split('/');
        if (parts.length === 3) {
            let day = parseInt(parts[0], 10);
            let month = parseInt(parts[1], 10) - 1; // Months are 0-based in JS
            let year = parseInt(parts[2], 10);
            if (year < 100) { // If year is in yy format, convert to yyyy
                year += 2000;
            }
            date = new Date(year, month, day);
        } else {
            date = null; // If we can't parse it, set it to null
        }
    }

    return date;
}

function titleCase(str) {
  return str
    .toLowerCase()                 // Tout en minuscules d’abord
    .split(' ')                    // Séparer les mots
    .map(word => {
      if(word.length === 0) return ''; // éviter les mots vides
      return word[0].toUpperCase() + word.slice(1);
    })
    .join(' ');                    // Rejoindre les mots
}

function flushSpreadsheetAndCache(ClearDeclaredTimesOnly = false) {
    SpreadsheetApp.flush();

    DeclaredTime.allDeclaredTimes = [];

    if (!ClearDeclaredTimesOnly) {
        WorkPackage.allWorkPackages = [];
        BudgetedTime.allBudgetedTimes = [];
        Employee.allEmployees = [];
        WorkedTime.allWorkedTimes = [];
        Project.allProjects = [];
        // Missions.allMissions = []; // Since this data is imported, it won't change during our process
    }
}

function isDateInRange(date, startDate, endDate) {
    if (!(date instanceof Date)) {
        return false; // If date is not a Date object, return false
    }

    date.setHours(0, 0, 0, 0);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    return date >= startDate && date <= endDate;
}

/**
 * Fill the "Temps déclarés" sheet with rows that match the work packages and the people who worked on them.
 */
function UpdateEmployeeSalaries() {
    const html = HtmlService.createHtmlOutputFromFile('interfaces/WarningBeforeUpdateSalaries.html')
        .setWidth(500)
        .setHeight(550);

    SpreadsheetApp.getUi().showModalDialog(html, 'Générer les salaires des collaborateurs');
}

function ConfirmUpdateEmployeeSalaries() {
    WorkedTime.UpdateEmployeeSalaries();
}

function getDateKey(date) {
    return date.getMonth() + '-' + date.getFullYear();
}

function getMonthStringForDate(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), '01/MM/yyyy');
}

/**
 * Returns a list of projects for the picker.
 * Each project should have an id and a name.
 */
function getProjects() {
    // We extract the list of projects from the "Projects" sheet, using the Projects model
    let existingProjects = Project.getProjects();

    return existingProjects.map(project => ({
        id: project.project,
        name: project.project
    }));
}

/**
 * Returns a list of reporting periods for the picker.
 * Each period should have a project, name, start, and end date.
 */
function getReportingPeriodsForPeriodPicker() {
    let id = 1;
    const ret = ReportingPeriod.getReportingPeriods().map(period => ({
        id: 'rp' + (id++),
        projectId: period.project,
        name: period.name,
        start: period.start.valueOf(),
        end: period.end.valueOf()
    }));

    return ret;
}

function getCurrentUser() {
    const email = Session.getActiveUser().getEmail();
    if (!email) return "Utilisateur inconnu";

    const localPart = email.split("@")[0]; // avant le @
    const parts = localPart.split(".");
    const firstName = parts[0].charAt(0).toUpperCase() + parts[0].slice(1);

    return firstName;
}

function isDateInRange(startDate, endDate, yearString) {
    let minYear = yearString;
    let maxYear = yearString;
    
    if (yearString > 2300) {
        // year is in the format 2324
        minYear = 2000 + (yearString % 100) - 1;
        maxYear = minYear + 1;
    }

    return startDate.getFullYear() <= maxYear && minYear <= endDate.getFullYear();
}

function notInRange(m, year, startDate, endDate) {
    let monthDate = new Date(year, m - 1, 1);

    return monthDate < startDate || monthDate > endDate;
}
