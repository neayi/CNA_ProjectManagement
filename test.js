function GenerationTest() {
    flushSpreadsheetAndCache(); // Ensure the spreadsheet and cache are flushed before starting the test

    // Test the different models
    let startDate = new Date('2025-01-01');
    let endDate = new Date('2025-12-01');
    
    generatedTimesForDates(startDate, endDate, true);
}

function UnitTest() {
    flushSpreadsheetAndCache(); // Ensure the spreadsheet and cache are flushed before starting the test

    // Test the different models
    let startDate = new Date('2024-01-01');
    let endDate = new Date('2024-12-31');
    let employees = Employee.getEmployees();
    let workedTimes = WorkedTime.getWorkedTimes();
    let declaredTimes = DeclaredTime.getDeclaredTimes();
    let budgetedTimes = BudgetedTime.getBudgetedTimes();
    let workPackages = WorkPackage.getWorkPackages();

    employees = employees.filter(employee => {
        return employee.hasWorkedBetween(startDate, endDate);
    });

    let martin = employees.filter(employee => {
        return employee.name == "Martin Rollet";
    }).at(0);

    let budgetedProjects = martin.BudgetedTime.getBudgetedTimesOnProjects(startDate, endDate);

    let declaredTimesForMartinIn2024 = martin.getDeclaredTimeForYear(2024); // 2.2
    let declaredTimesForMartinInJanuary2024 = martin.getDeclaredTimeForMonth(0, 2024); // 0.2
    let workedTimeForMartin = martin.getWorkedTime(0, 2025); // January 2025 : 0.5

    if (workedTimeForMartin != 0.5) {
        throw new Error("Martin should not have worked 50% in January 2025, but got " + workedTimeForMartin);
    }
    if (declaredTimesForMartinIn2024 != 2.2) {
        throw new Error("Martin should have declared 2.2 in 2024, but got " + declaredTimesForMartinIn2024);
    }
    if (declaredTimesForMartinInJanuary2024 != 0.2) {
        throw new Error("Martin should have declared 0.2 in January 2024, but got " + declaredTimesForMartinInJanuary2024);
    }

    let wp = budgetedTimes.filter(wp => wp.project == 'CONSERWA').at(0).WorkPackage.getWorkPackages();
    let wp1 = wp.at(0);
    
    let wp1budgetFor2023 = wp1.getBudgetedTimeForYear(2023); // 0.2
    let wp1declaredFor2023 = wp1.getDeclaredTimeForYear(2023); // 0.2

    if (wp1budgetFor2023 != 0.2) {
        throw new Error("wp1 budget for 2023 should be 0.2, but got " + wp1budgetFor2023);
    }
    
    if (wp1declaredFor2023 != 0.2) {
        throw new Error("wp1 declared for 2023 should be 0.2, but got " + wp1declaredFor2023);
    }

    Logger.log("Test completed successfully!");
}