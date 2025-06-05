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

function generatedTimesForDates(startDate, endDate, deleteExistingTimes) {
    startDate = new Date(startDate);
    endDate = new Date(endDate);

    Logger.log("Génération des temps déclarés pour les dates entre " + startDate + " et " + endDate + " (deleteExistingTimes = " + deleteExistingTimes + ")");

    // If deleteExistingTimes is true, we will delete the existing times in the "Temps déclarés" sheet.
    if (deleteExistingTimes) {
        let declaredTimesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés');
        if (declaredTimesSheet) {
            
            let existingDeclaredTimes = DeclaredTime.getDeclaredTimes();

            for (let i = existingDeclaredTimes.length - 1; i >= 1; i--) {
                if (isDateInRange(existingDeclaredTimes[i].month, startDate, endDate)) {
                   declaredTimesSheet.deleteRow(i + 2); // +2 because the first row is the header and the index is 0-based
                }
            }

            Logger.log("Les temps déclarés existants ont été supprimés.");
        }
    }

    // Don't forget to flush the spreadsheet and cache before starting the generation.
    flushSpreadsheetAndCache();

    // TODO: We could add a step here to select the projects on which we want to work.


    // Start by getting the budgeted times for each work package and each person, per year.
    let employees = Employee.getEmployees().filter(employee => {
        return employee.hasWorkedBetween(startDate, endDate); // && employee.name == "Laurence Fontaine";
    });

    Logger.log("Found " + employees.length + " employees who have worked between " + startDate + " and " + endDate);
    Logger.log(employees);


    employees.forEach(employee => {
        let budgetedTimes = employee.BudgetedTime.getBudgetedTimesOnProjects(startDate, endDate); // returns an array of projects that the employee has budgeted times on, between the two dates

        budgetedTimes.forEach(budgetedTime => {
            for (let year = startDate.getFullYear(); year <= endDate.getFullYear(); year++) {
                let declaredTimeOnProject = employee.getDeclaredTimeForYearAndProject(year, budgetedTime.project); // returns the declared time for the employee for the given year 

                if (declaredTimeOnProject < budgetedTime.getBudgetedTimeForYear(year)) {
                    // If the declared time is less than the budgeted time, we need to generate the missing times.

                    let missingTime = budgetedTime.getBudgetedTimeForYear(year) - declaredTimeOnProject;

                    let workPackages = budgetedTime.WorkPackage.getWorkPackages();

                    console.log("Pour l'employé " + employee.name + " et le projet " + budgetedTime.project + ", il reste " + missingTime + " PM à déclarer pour l'année " + year);

                    workPackages.forEach(workPackage => {
                        if (workPackage.getBudgetedTimeForYear(year) > 0 && 
                            workPackage.getDeclaredTimeForYear(year) < workPackage.getBudgetedTimeForYear(year)) {
                            
                            if (missingTime <= 0) {
                                return; // No more time to declare for this project
                            }

                            let rows = [];

                            let WPMissingTime = workPackage.getBudgetedTimeForYear(year) - workPackage.getDeclaredTimeForYear(year);

                            let timeToDeclare = Math.min(missingTime, WPMissingTime);

                            for (let m = getStartMonth(year, startDate); m <= getEndMonth(year, endDate); m++) {
                                let workedTime = employee.getWorkedTime(m, year); // in man month
                                let declaredTimeForMonth = employee.getDeclaredTimeForMonth(m, year); // in man month

                                if (workedTime > declaredTimeForMonth) {
                                    let newDeclaredTime = Math.min(timeToDeclare, workedTime - declaredTimeForMonth);

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
                                        budgetedTime.project // Projet
                                    ]);

                                    timeToDeclare -= newDeclaredTime;
                                    missingTime -= newDeclaredTime;

                                    if (timeToDeclare <= 0 || missingTime <= 0) {
                                        break; // No more time to declare for this work package or this project
                                    }
                                }
                            }

                            if (rows.length > 0) {
                                // If we have rows to add, we need to add them to the "Temps déclarés" sheet.
                                let declaredTimesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temps déclarés');

                                // Add the new rows
                                declaredTimesSheet.getRange(declaredTimesSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

                                // Flush the changes to the sheet
                                flushSpreadsheetAndCache(true);
                                Logger.log("Ajout de " + rows.length + " lignes dans la feuille 'Temps déclarés' pour l'employé " + employee.name + " et le projet " + budgetedTime.name);
                            }
                        }
                    });
                }

            }

        });

    });

    SpreadsheetApp.getUi().alert("Les temps ont été générés.");
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

    if (!ClearDeclaredTimesOnly){
        WorkPackage.allWorkPackages = [];
        BudgetedTime.allBudgetedTimes = [];
        Employee.allEmployees = [];
        WorkedTime.allWorkedTimes = [];
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