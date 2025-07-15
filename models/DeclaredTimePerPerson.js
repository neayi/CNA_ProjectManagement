/**
 * This class allows to manage the time we need to declare for each person on each work package each year.
 * It will contain a number of methods to fill up the time up to the constraints that we have:
 * - Budget for each work package
 * - Budget for each person per year
 * - Time already declared for each work package
 * - Time already declared for each person
 */

class DeclaredTimePerPerson {
    constructor(personName, workPackage, projectName, year) {
        this.personName = personName;
        this.workPackage = workPackage;
        this.projectName = projectName;

        this.isTarget = false; // If this is the target person for the work package
        
        this.declaredTimes = new Map(); // Map of times declared per month - key is month in format 'MM/YYYY' and value is the declared time in PM
        this.year = year;

        this.budgetedTime = 0; // Budgeted time for this work package and person for the year
        this.notWorked = false;
    }

    /**
     * Returns a key for this instance, which is a string combining the person name, work package, project name, and year.
     * This key can be used to uniquely identify this instance in a collection or map.
     * @returns {string} A unique key for this instance.
     */
    getKey() {
        return DeclaredTimePerPerson.makeKey(this.personName, this.workPackage, this.projectName, this.year);
    }

    setAsNotWorked() {
        this.budgetedTime = 0;
        this.notWorked = true;
    }

    /**
     * Attempt to set a given amount of time (man months), but do not exceed the budgeted time for this work package and person.
     */
    addDeclaredTime(declaredTime) {
        if (this.notWorked) {
            return; // If the person is marked as not worked, we do not set any declared time
        }


        this.budgetedTime += declaredTime;
    }

    static makeKey(personName, workPackage, projectName, year) {
        return `${personName}:${workPackage}:${projectName}:${year}`;
    }
}