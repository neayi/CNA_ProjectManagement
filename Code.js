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

    let declaredTimesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés');

    // On garde la dernière ligne générée pour récupérer les formules calculées
    const lastGoodFormulasRowIndex = declaredTimesSheet.getLastRow(); // On pourrait améliorer en vérifiant que les formules ne sont effectivement pas vides
    
    // On copie les formules calculées pour les utiliser plus tard au moment de générer de nouvelles lignes
    let formulasFJ = declaredTimesSheet.getRange(lastGoodFormulasRowIndex, 6, 1, 5).getFormulasR1C1();
    let formulasLR = declaredTimesSheet.getRange(lastGoodFormulasRowIndex, 12, 1, 7).getFormulasR1C1();

    // If deleteExistingTimes is true, we will delete the existing times in the "Temps déclarés" sheet.
    if (deleteExistingTimes) {
        let declaredTimesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés');
        if (declaredTimesSheet) {

            let existingDeclaredTimes = DeclaredTime.getDeclaredTimes();

            for (let i = existingDeclaredTimes.length - 1; i >= 1; i--) {
                if (isDateInRange(existingDeclaredTimes[i].month, startDate, endDate) && existingDeclaredTimes[i].project === projectName) {
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
        if (employeesNames.indexOf(employee.name) !== -1 && 
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

    const startYear = startDate.getFullYear();
    const endYear = endDate.getFullYear();
    for (let year = startYear; year <= endYear; year++) {
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
            employee.budgetTimeForYear = BudgetedTime.getBudgetForWPPerson(budgetKey, year);
        });

        // On commence par initialiser wpPersons:
        employees.forEach(employee => {
            workPackages.forEach(workPackage => {
                let wpPerson = new DeclaredTimePerPerson(employee.name, workPackage.name, projectName, year);
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

                    let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, year));
                    if (wpPerson)
                        wpPerson.isTarget = true; // This is the target person for the work package
                }
            });

            // Si il n'y a qu'une personne pour ce WP, alors on lui affecte le temps prévu pour le WP sur l'année.
            if (workPackage.employees.size === 1) {
                let employeeName = workPackage.employees.keys().next().value;
                let employee = employees.get(employeeName);
                if (employee) {
                    let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, year));
                    if (wpPerson) {
                        wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(year), employee.getRemainingBudgetedTime()));
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

                let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, year));
                if (wpPerson && wpPerson.budgetedTime == 0) {
                    wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(year), averageRemainingBudgetedTime));
                }
            });
        });

        // Deuxième passe, cette fois on complète les temps pour chaque personne, en fonction du temps restant à déclarer pour chaque WP.
        employees.forEach(employee => {
            workPackages.forEach(workPackage => {
                if (!workPackage.employees.has(employee.name))
                    return; // Si l'employé n'est pas prévu sur ce WP, on passe au suivant

                let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, year));
                if (wpPerson) {
                    wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(year), employee.getRemainingBudgetedTime()));
                }
            });
        });

        // S'il reste du temps à déclarer pour une personne, on l'affecte aux WP qui ne sont pas encore remplis, même si la personne n'était pas prévue au départ
        employees.forEach(employee => {
            workPackages.forEach(workPackage => {
                let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, year));
                if (wpPerson) {
                    wpPerson.addDeclaredTime(Math.min(workPackage.getRemainingBudgetedTime(year), employee.getRemainingBudgetedTime()));
                }
            });
        });

        // On affiche les résultats dans la console pour debug
        debugLog(employees, workPackages, wpPersons, projectName, year);

        // Maintenant, on va ajouter les temps déclarés dans la feuille "Temps déclarés"
        employees.forEach(employee => {

            workPackages.forEach(workPackage => {
                const key = DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, year);
                let wpPerson = wpPersons.get(key);

                if (wpPerson == false || wpPerson.budgetedTime == 0)
                    return;

                console.log("Adding declared time for " + employee.name + " on " + workPackage.name + " for year " + year + ": " + wpPerson.budgetedTime);

                let WPYearlyMissingTime = wpPerson.budgetedTime - employee.getDeclaredTimeForYearAndWorkPackage(year, workPackage.name);

                if (WPYearlyMissingTime <= 0)
                    return; // Si le temps déclaré est déjà supérieur ou égal au temps prévu, on ne fait rien

                let rows = [];

                let months = new Set(); // de 1 à 12

                // Commencer par identifier les ordres de mission pour ce salarié sur ce projet dans cette année 
                let missions = Mission.getMissionsForEmployee(employee.name, workPackage.name, year);
                missions.forEach(mission => {
                    const m = mission.month.getMonth() + 1;

                    if (m >= getStartMonth(year, startDate) && m <= getEndMonth(year, endDate)) {
                        months.add(m);
                    }
                });

                // Ajouter les autres mois de l'année
                for (let m = getStartMonth(year, startDate); m <= getEndMonth(year, endDate); m++) {
                    months.add(m);
                }

                console.log(Array.from(months));

                for (const m of months) {
                    let workedTime = employee.getWorkedTime(m, year); // Temps travaillé par le salarié
                    let declaredTimeForMonth = employee.getDeclaredTimeForMonth(m, year); // temps déjà déclaré pour ce mois
                    let remainingTimeForMonth = workedTime - declaredTimeForMonth; // Temps restant à déclarer pour ce mois

                    if (remainingTimeForMonth <= 0)
                        continue; // Si le temps travaillé est inférieur ou égal au temps déjà déclaré, on ne fait rien
                    
                    let newDeclaredTime = Math.min(WPYearlyMissingTime, remainingTimeForMonth); // On ne peut pas déclarer plus que le temps travaillé moins le temps déjà déclaré

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
                        declaredTimesSheet.getRange(newRowIndex + index, 12, 1, 7).setFormulasR1C1(formulasLR);
                    });

                    // Flush the changes to the sheet
                    flushSpreadsheetAndCache(true);
                    Logger.log("Ajout de " + rows.length + " lignes dans la feuille 'Temps déclarés' pour l'employé " + employee.name + " et le projet " + projectName);
                }
            });
        });
    }    // For each year

    SpreadsheetApp.getUi().alert("Les temps ont été générés.");
}

function debugLog(employees, workPackages, wpPersons, projectName, year) {

    let row = ["Tps travaillé"];
    workPackages.forEach(workPackage => {
        row.push(workPackage.name);
    });
    console.log(row.join("\t"));

    employees.forEach(employee => {
        let row = [employee.name];

        workPackages.forEach(workPackage => {
            let wpPerson = wpPersons.get(DeclaredTimePerPerson.makeKey(employee.name, workPackage.name, projectName, year));
            if (wpPerson) {
                row.push(String(wpPerson.budgetedTime).replace('.', ',')); // Replace dots with commas for decimal values
            } else {
                row.push(0);
            }
        });

        row.push('');
        row.push(String(employee.budgetTimeForYear).replace('.', ','));

        console.log(row.join("\t"));
    });

}

// utilities
function getStartMonth(year, startDate) {
    return (year === startDate.getFullYear()) ? startDate.getMonth() + 1 : 1; // Months are 0-indexed in JavaScript
}

function getEndMonth(year, endDate) {
    return (year === endDate.getFullYear()) ? endDate.getMonth() + 1 : 12; // Months are 0-indexed in JavaScript
}

function getValue(row, headers, field) {
    let index = headers.indexOf(field);
    if (index === -1) {
        throw new Error(`Le champ '${field}' n'existe pas dans les en-têtes.`);
    }
    return row[index];
}

function getDateValue(row, headers, field) {
    let index = headers.indexOf(field);
    if (index === -1) {
        throw new Error(`Le champ '${field}' n'existe pas dans les en-têtes.`);
    }
    return row[index] == '' ? null : new Date(row[index]);
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

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return "Utilisateur inconnu";

  const localPart = email.split("@")[0]; // avant le @
  const parts = localPart.split(".");
  const firstName = parts[0].charAt(0).toUpperCase() + parts[0].slice(1);
  
  return firstName;
}