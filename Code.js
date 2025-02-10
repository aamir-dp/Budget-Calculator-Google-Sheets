// Replace with your actual Google Sheet ID
var SPREADSHEET_ID = '12I3LwYSykIJXjlEILn772IvyLT9bK-5ZQwSiNsxt9g4';

/**
 * onOpen – Adds a custom menu when the spreadsheet is opened.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Accounting')
        .addItem('Add Payment Record', 'showPaymentForm')
        .addItem('Add Employee Record', 'showEmployeeForm')
        .addItem('Deduct Salaries & Notify Owner', 'deductSalariesAndNotifyOwner')
        .addItem('Add Expense Record', 'showExpenseForm')
        .addItem('Pay Unpaid Expenses', 'payUnpaidExpenses')
        .addItem('Update Payment Record', 'showUpdatePaymentForm')
        .addSeparator()
        .addItem('Add Month Headings', 'addMonthHeadings') // ✅ New function in UI
        .addItem('Generate Monthly Summary', 'showFinancialSummaryForm') // ✅ New function in UI
        .addToUi();
}

// function onEdit(e) {
//     var sheet = e.source.getSheetByName('Payments'); // Adjust to your sheet
//     if (sheet) {
//         addMonthHeadings();
//     }
// }


function showFinancialSummaryForm() {
    var html = HtmlService.createHtmlOutputFromFile('FinancialSummaryForm')
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, "Generate Monthly Summary");
}


function showSalaryPaymentForm() {
    var html = HtmlService.createHtmlOutputFromFile('SalaryPaymentForm')
        .setTitle('Salary Payment')
        .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
}


function showUpdatePaymentForm() {
    var html = HtmlService.createHtmlOutputFromFile('UpdatePaymentForm')
        .setTitle('Update Payment Record');
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * showPaymentForm – Opens the Payment Form in a sidebar.
 */
function showPaymentForm() {
    var html = HtmlService.createHtmlOutputFromFile('PaymentForm')
        .setTitle('Add Payment Record');
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * showEmployeeForm – Opens the Employee Form in a sidebar.
 */
function showEmployeeForm() {
    var html = HtmlService.createHtmlOutputFromFile('EmployeeForm')
        .setTitle('Add Employee Record');
    SpreadsheetApp.getUi().showSidebar(html);
}

// function addPaymentRecord(record) {
//     try {
//         var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//         var sheet = ss.getSheetByName('Payments');
//         if (!sheet) {
//             throw new Error("Sheet 'Payments' not found.");
//         }

//         // Get the foreign currency amount and selected currency.
//         var totalPaymentFC = parseFloat(record.totalPaymentFC) || 0;
//         var currency = record.currency.toUpperCase();

//         // Build the formula for "Total Payment Received PKR" (Column J).
//         var totalPaymentPKRFormula = "";
//         if (currency === "PKR") {
//             totalPaymentPKRFormula = "=" + totalPaymentFC + "*1";
//         } else {
//             totalPaymentPKRFormula = "=" + totalPaymentFC + "*GOOGLEFINANCE(\"CURRENCY:" + currency + "PKR\")";
//         }

//         // Build the new row array matching the sheet columns:
//         // 1. Invoice Number  
//         // 2. Date of Invoice  
//         // 3. Date of Payment  
//         // 4. Client Name  
//         // 5. Job Description  
//         // 6. Payment Account  
//         // 7. Payment Status  
//         // 8. Currency  
//         // 9. Total Payment Received F.C  
//         // 10. Total Payment Received PKR (formula)
//         var newRow = [
//             record.invoiceNumber,
//             record.invoiceDate,
//             record.paymentDate,
//             record.clientName,
//             record.jobDescription,
//             record.paymentAccount,
//             record.paymentStatus,
//             record.currency,
//             totalPaymentFC,
//             totalPaymentPKRFormula
//         ];

//         sheet.appendRow(newRow);

//         return { status: 'success', message: 'Payment record added successfully.' };
//     } catch (e) {
//         return { status: 'error', message: e.toString() };
//     }
// }

function addPaymentRecord(record) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Payments');
        var dashboardSheet = ss.getSheetByName('Dashboard');

        if (!sheet || !dashboardSheet) {
            throw new Error("One or more required sheets are missing.");
        }

        var totalPaymentFC = parseFloat(record.totalPaymentFC) || 0;
        var totalPaymentPKR = parseFloat(record.totalPaymentPKR) || 0; // ✅ Manually inputted by user

        // New row for payments
        var newRow = [
            record.invoiceNumber,
            record.invoiceDate,
            record.paymentDate,
            record.clientName,
            record.jobDescription,
            record.paymentAccount,
            record.paymentStatus,
            record.currency,
            totalPaymentFC,
            totalPaymentPKR // ✅ No more formula-based conversion
        ];

        sheet.appendRow(newRow);

        // ✅ If Payment Status is "Received", add to Running Total
        if (record.paymentStatus.toLowerCase() === "received") {
            var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
            var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;
            var newRunningTotal = currentRunningTotal + totalPaymentPKR;
            runningTotalCell.setValue(newRunningTotal);
        }

        return { status: 'success', message: 'Payment record added successfully.' };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}



/**
 * addEmployeeRecord – Appends a new employee record into the "Employees" sheet.
 *
 * Expected record properties:
 *   - employeeName, salary, designation, dateJoined, status
 */
function addEmployeeRecord(record) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Employees');
        if (!sheet) {
            throw new Error("Sheet 'Employees' not found.");
        }

        // Build the new row array for employee data
        var newRow = [
            record.employeeName,
            record.salary,
            record.designation,
            record.dateJoined,
            record.status,
            record.salaryReceived // ✅ New column for Salary Received
        ];

        sheet.appendRow(newRow);
        return { status: 'success', message: 'Employee record added successfully.' };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}

/**
 * addExpenseRecord – Appends a new expense record into the "Expenses" sheet.
 *
 * Expected record properties:
 *   - date, description, category, amount
 */
function addExpenseRecord(record) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Expenses');
        var dashboardSheet = ss.getSheetByName('Dashboard');

        if (!sheet || !dashboardSheet) {
            throw new Error("One or more required sheets are missing.");
        }

        var expenseAmount = parseFloat(record.amount) || 0;
        var expenseStatus = String(record.status).trim().toLowerCase();

        // New row for expenses
        var newRow = [
            record.expenseDate,
            record.description,
            record.category,
            expenseAmount,
            record.status // Paid or Unpaid
        ];

        sheet.appendRow(newRow);

        // ✅ If the expense is already "Paid", deduct it from Running Total immediately
        if (expenseStatus === "paid") {
            var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
            var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;

            if (currentRunningTotal >= expenseAmount) {
                var newRunningTotal = currentRunningTotal - expenseAmount;
                runningTotalCell.setValue(newRunningTotal); // ✅ Deduct immediately
            } else {
                throw new Error("Not enough balance in Running Total to cover this expense.");
            }

            // ✅ Update Dashboard
            updateDashboard();
        }

        return { status: 'success', message: 'Expense record added successfully.' };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}


/**
 * showExpenseForm – Opens the Expense Form in a sidebar.
 */
function showExpenseForm() {
    var html = HtmlService.createHtmlOutputFromFile('ExpenseForm')
        .setTitle('Add Expense Record');
    SpreadsheetApp.getUi().showSidebar(html);
}

// If your script is container-bound (attached to the spreadsheet), use getActiveSpreadsheet()
// Otherwise, use openById() with a valid spreadsheet ID.
// For a container-bound script, you can remove the SPREADSHEET_ID variable and use getActiveSpreadsheet().

/**
 * showDashboard – Opens the Dashboard in a sidebar.
 */
function showDashboard() {
    var html = HtmlService.createHtmlOutputFromFile('Dashboard')
        .setTitle('Dashboard');
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * getDashboardData – Aggregates dashboard metrics.
 *
 * This function reads:
 *   - Total revenue from the "Payments" sheet (sum of Total Payment Received PKR, assumed to be in column 10)
 *   - Total salaries from the "Employees" sheet (assumed to be in column 2)
 *   - Total expenses from the "Expenses" sheet (assumed to be in column 2)
 *
 * It then computes:
 *   - Profit = Revenue - Salaries - Expenses
 *   - Partner Share = 50% of Profit (as an example)
 */
function getDashboardData() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var paymentsSheet = ss.getSheetByName('Payments');
        var dashboardSheet = ss.getSheetByName('Dashboard');
        var expensesSheet = ss.getSheetByName('Expenses');

        if (!paymentsSheet || !dashboardSheet || !expensesSheet) {
            throw new Error("One or more required sheets are missing.");
        }

        var paymentsData = paymentsSheet.getDataRange().getValues();
        var expensesData = expensesSheet.getDataRange().getValues();
        var runningTotal = 0;
        var totalPaidExpenses = 0;

        // Loop through Payments data (excluding header)
        for (var i = 1; i < paymentsData.length; i++) {
            var status = String(paymentsData[i][6]).trim().toLowerCase(); // Column G: Payment Status
            var totalReceivedPKR = parseFloat(paymentsData[i][9]) || 0; // Column J: Total Payment Received PKR

            if (status === "received") {
                runningTotal += totalReceivedPKR;
            }
        }

        // Loop through Expenses data (only Paid expenses)
        for (var i = 1; i < expensesData.length; i++) {
            if (String(expensesData[i][4]).trim().toLowerCase() === "paid") { // Column E: Status
                totalPaidExpenses += parseFloat(expensesData[i][3]) || 0; // Column D: Amount
            }
        }

        // Deduct expenses from Running Total
        runningTotal -= totalPaidExpenses;

        var partnerShare = 50; // Fixed 50% Share
        var partnerShareAmount = (runningTotal * partnerShare) / 100; // 50% of Running Total

        // **Update the Dashboard Sheet Table**
        dashboardSheet.getRange("A1:C1").setValues([["Metric", "Value", "Percentage"]]); // Headers
        dashboardSheet.getRange("A2:C2").setValues([["Running Total", runningTotal, "-"]]);
        dashboardSheet.getRange("A3:C3").setValues([["Partner Share (%)", partnerShare, "50%"]]);
        dashboardSheet.getRange("A4:C4").setValues([["Partner Share Amount", partnerShareAmount, "-"]]);

        return {
            status: 'success',
            runningTotal: runningTotal.toFixed(2),
            partnerShare: partnerShare,
            partnerShareAmount: partnerShareAmount.toFixed(2)
        };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}



/**
 * getPendingInvoices
 *
 * Retrieves the invoice numbers for records with Payment Status "Pending".
 */
function getPendingInvoices() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Payments');
        var data = sheet.getDataRange().getValues(); // includes header row
        var pending = [];
        // Loop through rows (starting at row 2)
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][6]).toLowerCase() === "pending") { // Column G (index 6)
                pending.push(data[i][0]); // Column A (index 0) holds Invoice Number
            }
        }
        return { status: 'success', invoices: pending };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}

/**
 * getPaymentRecord
 *
 * Retrieves the details of a payment record by invoice number.
 */
function getPaymentRecord(invoiceNumber) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Payments');
        var data = sheet.getDataRange().getValues();

        console.log("Looking for invoice: " + invoiceNumber); // Debug

        for (var i = 1; i < data.length; i++) {
            console.log("Checking row " + i + " Invoice: " + data[i][0]); // Debug
            if (String(data[i][0]).trim() === String(invoiceNumber).trim()) {
                var record = {
                    invoiceNumber: data[i][0],
                    invoiceDate: data[i][1],
                    paymentDate: data[i][2],
                    clientName: data[i][3],
                    jobDescription: data[i][4],
                    paymentAccount: data[i][5],
                    paymentStatus: data[i][6],
                    currency: data[i][7],
                    totalPaymentFC: data[i][8]
                };
                console.log("Record Found:", record); // Debug
                return { status: 'success', record: record };
            }
        }
        console.warn("Invoice not found:", invoiceNumber);
        return { status: 'error', message: 'Invoice not found.' };
    } catch (e) {
        console.error("Error in getPaymentRecord:", e);
        return { status: 'error', message: e.toString() };
    }
}

function getPaymentsSheetData() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Payments');
        if (!sheet) {
            throw new Error("Sheet 'Payments' not found.");
        }
        var data = sheet.getDataRange().getValues(); // Get entire sheet data
        return data;
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}



/**
 * updatePaymentRecord
 *
 * Updates an existing payment record (identified by invoice number) with new values.
 * This function also sets the Payment Status to "Received" and recalculates the running totals.
 *
 * Expected properties in the updated record:
 *   - invoiceNumber, invoiceDate, paymentDate, clientName,
 *     jobDescription, paymentAccount, currency, totalPaymentFC
 *
 * Payment Status will be forced to "Received" when updating.
 */
// function updatePaymentRecord(record) {
//     try {
//         var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//         var sheet = ss.getSheetByName('Payments');
//         var dashboardSheet = ss.getSheetByName('Dashboard');

//         if (!sheet || !dashboardSheet) {
//             throw new Error("One or more required sheets are missing.");
//         }

//         // Read all rows to locate the record by invoice number.
//         var data = sheet.getDataRange().getValues();
//         var rowToUpdate = -1;
//         var totalPaymentPKR = 0;

//         for (var i = 1; i < data.length; i++) {
//             if (String(data[i][0]).trim() === String(record.invoiceNumber).trim()) {
//                 rowToUpdate = i + 1; // Convert array index to sheet row number
//                 totalPaymentPKR = parseFloat(data[i][9]) || 0; // Column J: Total Payment Received PKR
//                 break;
//             }
//         }

//         if (rowToUpdate === -1) {
//             throw new Error("Invoice number " + record.invoiceNumber + " not found.");
//         }

//         // Update Payment Status to "Received" (Column G)
//         sheet.getRange(rowToUpdate, 7).setValue("Received");

//         // Update Date of Payment (Column C)
//         sheet.getRange(rowToUpdate, 3).setValue(record.paymentDate);

//         // ✅ Update Running Total in Dashboard
//         var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
//         var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;
//         var newRunningTotal = currentRunningTotal + totalPaymentPKR;

//         runningTotalCell.setValue(newRunningTotal); // ✅ Add received payment

//         // ✅ Refresh Dashboard Data
//         updateDashboard();

//         return { status: 'success', message: 'Payment record updated successfully.' };
//     } catch (e) {
//         return { status: 'error', message: e.toString() };
//     }
// }

function updatePaymentRecord(record) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Payments');
        var dashboardSheet = ss.getSheetByName('Dashboard');

        if (!sheet || !dashboardSheet) {
            throw new Error("One or more required sheets are missing.");
        }

        var data = sheet.getDataRange().getValues();
        var rowToUpdate = -1;
        var totalPaymentPKR = parseFloat(record.totalPaymentPKR) || 0; // ✅ Manually inputted

        for (var i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === String(record.invoiceNumber).trim()) {
                rowToUpdate = i + 1; // Convert array index to sheet row number
                break;
            }
        }

        if (rowToUpdate === -1) {
            throw new Error("Invoice number " + record.invoiceNumber + " not found.");
        }

        // ✅ Update Payment Status and Date of Payment
        sheet.getRange(rowToUpdate, 7).setValue("Received"); // Column G
        sheet.getRange(rowToUpdate, 3).setValue(record.paymentDate); // Column C

        // ✅ Update Total Payment Received (PKR) manually
        sheet.getRange(rowToUpdate, 10).setValue(totalPaymentPKR); // Column J

        // ✅ Add to Running Total
        var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
        var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;
        var newRunningTotal = currentRunningTotal + totalPaymentPKR;
        runningTotalCell.setValue(newRunningTotal);

        // ✅ Refresh Dashboard
        updateDashboard();

        return { status: 'success', message: 'Payment record updated successfully.' };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}



/**
 * recalcRunningTotals
 *
 * Recalculates the Running Total (Column K) for all rows in the Payments sheet.
 * Only rows with Payment Status "Received" contribute their Total Payment Received PKR value.
 * For row 2 (first data row): =IF(G2="Received", J2, 0)
 * For subsequent rows: =K(previous row) + IF(G[current row]="Received", J[current row], 0)
 */
function recalcRunningTotals(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // No data rows

    for (var i = 2; i <= lastRow; i++) {
        var formula = "";
        if (i === 2) {
            formula = '=IF(G2="Received", J2, 0)'; // First row formula
        } else {
            formula = '=K' + (i - 1) + ' + IF(G' + i + '="Received", J' + i + ', 0)';
        }
        sheet.getRange(i, 11).setFormula(formula); // Column K
    }
}


function markPaymentAsReceived(invoiceNumber, paymentDate) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Payments');
        var dashboardSheet = ss.getSheetByName('Dashboard');

        if (!sheet || !dashboardSheet) {
            throw new Error("One or more required sheets are missing.");
        }

        var data = sheet.getDataRange().getValues();
        var rowToUpdate = -1;
        var totalPaymentPKR = 0;

        for (var i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === String(invoiceNumber).trim()) {
                rowToUpdate = i + 1; // Convert array index to sheet row number
                totalPaymentPKR = parseFloat(data[i][9]) || 0; // Column J: Total Payment Received PKR
                break;
            }
        }

        if (rowToUpdate === -1) {
            throw new Error("Invoice number " + invoiceNumber + " not found.");
        }

        // ✅ Update Payment Status to "Received" (Column G)
        sheet.getRange(rowToUpdate, 7).setValue("Received");

        // ✅ Update Date of Payment (Column C)
        sheet.getRange(rowToUpdate, 3).setValue(paymentDate);

        // ✅ Add amount to Running Total
        var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
        var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;
        var newRunningTotal = currentRunningTotal + totalPaymentPKR;
        runningTotalCell.setValue(newRunningTotal);

        // ✅ Refresh Dashboard
        updateDashboard();

        return { status: 'success', message: 'Payment updated to Received and Running Total updated.' };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}




// function getActiveEmployees() {
//     try {
//         var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//         var employeeSheet = ss.getSheetByName('Employees');

//         if (!employeeSheet) {
//             throw new Error("Sheet 'Employees' not found.");
//         }

//         var employeeData = employeeSheet.getDataRange().getValues();
//         var activeEmployees = [];

//         Logger.log("Total Rows in Employees Sheet: " + employeeData.length); // Debugging

//         for (var i = 1; i < employeeData.length; i++) {
//             if (!employeeData[i] || employeeData[i].length < 5) continue; // Ensure row exists and has data

//             var employeeName = employeeData[i][0] || "Unknown"; // Column A: Employee Name
//             var salary = parseFloat(employeeData[i][1]) || 0; // Column B: Salary
//             var designation = employeeData[i][2] || "Unknown"; // Column C: Designation
//             var dateJoined = employeeData[i][3] || "Unknown"; // Column D: Date Joined
//             var status = String(employeeData[i][4] || "").trim().toLowerCase(); // Column E: Status

//             Logger.log(`Row ${i}: ${employeeName}, Status: ${status}`); // Debugging

//             if (status === "active") {
//                 activeEmployees.push({
//                     name: employeeName,
//                     salary: salary,
//                     designation: designation,
//                     dateJoined: dateJoined
//                 });
//             }
//         }

//         Logger.log("Active Employees Found: " + activeEmployees.length); // Debugging
//         return activeEmployees.length > 0 ? activeEmployees : [];
//     } catch (e) {
//         Logger.log("Error in getActiveEmployees: " + e.toString());
//         return [];
//     }
// }



// function processSelectedSalaries(selectedEmployeeNames) {
//     try {
//         var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//         var employeeSheet = ss.getSheetByName('Employees');
//         var dashboardSheet = ss.getSheetByName('Dashboard');

//         if (!employeeSheet || !dashboardSheet) {
//             throw new Error("One or more required sheets are missing.");
//         }

//         var employeeData = employeeSheet.getDataRange().getValues();
//         var today = new Date();
//         var currentMonth = today.getMonth() + 1; // JS months are 0-based
//         var currentYear = today.getFullYear();

//         var totalSalaries = 0;
//         var salaryDetails = [];

//         for (var i = 1; i < employeeData.length; i++) {
//             var employeeName = employeeData[i][0];
//             if (!selectedEmployeeNames.includes(employeeName)) continue;

//             var salary = parseFloat(employeeData[i][1]) || 0; // Column B: Salary
//             var dateJoined = new Date(employeeData[i][3]); // Column D: Date Joined
//             var joiningMonth = dateJoined.getMonth() + 1;
//             var joiningYear = dateJoined.getFullYear();

//             var finalSalary = salary;

//             // If employee joined this month, calculate working days & salary
//             if (joiningYear === currentYear && joiningMonth === currentMonth) {
//                 var totalDaysInMonth = new Date(currentYear, currentMonth, 0).getDate();
//                 var workingDays = totalDaysInMonth - dateJoined.getDate() + 1;
//                 finalSalary = (salary / totalDaysInMonth) * workingDays;
//             }

//             totalSalaries += finalSalary;
//             salaryDetails.push(`${employeeName}: PKR ${finalSalary.toFixed(2)}`);
//         }

//         // Deduct from Running Total in Dashboard
//         var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
//         var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;

//         if (currentRunningTotal >= totalSalaries) {
//             var newRunningTotal = currentRunningTotal - totalSalaries;
//             runningTotalCell.setValue(newRunningTotal);
//         } else {
//             throw new Error("Not enough balance in Running Total to pay salaries.");
//         }

//         // Send email notification
//         var ownerEmail = Session.getEffectiveUser().getEmail();
//         var subject = "Salary Payment Processed";
//         var body = `Dear Owner,\n\nThe selected salaries have been processed successfully.\n\nTotal Salaries Paid: PKR ${totalSalaries.toFixed(2)}\n\nBreakdown:\n${salaryDetails.join("\n")}\n\nNew Running Total: PKR ${newRunningTotal.toFixed(2)}\n\nBest Regards,\nAccounting System`;

//         MailApp.sendEmail(ownerEmail, subject, body);

//         return { status: 'success', message: `Salaries paid and email sent to ${ownerEmail}` };
//     } catch (e) {
//         return { status: 'error', message: e.toString() };
//     }
// }

// function refreshDashboard() {
//     try {
//         var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//         var dashboardSheet = ss.getSheetByName('Dashboard');
//         var paymentsSheet = ss.getSheetByName('Payments');
//         var expensesSheet = ss.getSheetByName('Expenses');

//         if (!dashboardSheet || !paymentsSheet || !expensesSheet) {
//             throw new Error("One or more required sheets are missing.");
//         }

//         var paymentsData = paymentsSheet.getDataRange().getValues();
//         var expensesData = expensesSheet.getDataRange().getValues();
//         var runningTotal = 0;
//         var totalPaidExpenses = 0;

//         // ✅ Sum only "Received" payments from the Payments sheet
//         for (var i = 1; i < paymentsData.length; i++) {
//             var status = String(paymentsData[i][6] || "").trim().toLowerCase(); // Column G: Payment Status
//             var totalReceivedPKR = parseFloat(paymentsData[i][9]) || 0; // Column J: Total Payment Received PKR

//             if (status === "received") {
//                 runningTotal += totalReceivedPKR;
//             }
//         }

//         // ✅ Deduct all "Paid" expenses from the Expenses sheet
//         for (var i = 1; i < expensesData.length; i++) {
//             var expenseStatus = String(expensesData[i][4] || "").trim().toLowerCase(); // Column E: Status
//             if (expenseStatus === "paid") {
//                 totalPaidExpenses += parseFloat(expensesData[i][3]) || 0; // Column D: Amount
//             }
//         }

//         // ✅ Update Running Total by subtracting paid expenses
//         runningTotal -= totalPaidExpenses;

//         var partnerShare = 50; // 50% Share
//         var partnerShareAmount = (runningTotal * partnerShare) / 100;

//         // ✅ Update the Dashboard Sheet
//         dashboardSheet.getRange("A1:C1").setValues([["Metric", "Value", "Percentage"]]);
//         dashboardSheet.getRange("A2:C4").setValues([
//             ["Running Total", runningTotal, "-"],
//             ["Partner Share (%)", partnerShare, "50%"],
//             ["Partner Share Amount", partnerShareAmount, "-"]
//         ]);

//         Logger.log("Dashboard Updated: Running Total = " + runningTotal);

//         // ✅ Return updated values for UI refresh
//         return {
//             status: 'success',
//             runningTotal: runningTotal.toFixed(2),
//             partnerShare: partnerShare,
//             partnerShareAmount: partnerShareAmount.toFixed(2)
//         };
//     } catch (e) {
//         Logger.log("Error in refreshDashboard: " + e.toString());
//         return { status: 'error', message: e.toString() };
//     }
// }

/**
 * deductSalariesAndNotifyOwner - Deduct salaries and update the dashboard.
 */
function deductSalariesAndNotifyOwner() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var employeeSheet = ss.getSheetByName('Employees');
        var dashboardSheet = ss.getSheetByName('Dashboard');

        if (!employeeSheet || !dashboardSheet) {
            throw new Error("One or more required sheets are missing.");
        }

        var employeeData = employeeSheet.getDataRange().getValues();
        var today = new Date();
        var currentMonth = today.getMonth() + 1;
        var currentYear = today.getFullYear();

        var totalSalaries = 0;
        var salaryDetails = [];
        var employeesToUpdate = [];

        for (var i = 1; i < employeeData.length; i++) {
            if (!employeeData[i] || employeeData[i].length < 6) continue; // Skip empty rows

            var status = String(employeeData[i][4] || "").trim().toLowerCase(); // Column E: Status
            var salaryReceivedStatus = String(employeeData[i][5] || "").trim().toLowerCase(); // Column F: Salary Received

            if (status !== "active" || salaryReceivedStatus === "received") continue; // Skip inactive or already paid employees

            var employeeName = employeeData[i][0] || "Unknown"; // Column A: Employee Name
            var salary = parseFloat(employeeData[i][1]) || 0; // Column B: Salary
            var dateJoined = new Date(employeeData[i][3]); // Column D: Date Joined
            var joiningMonth = dateJoined.getMonth() + 1;
            var joiningYear = dateJoined.getFullYear();

            var finalSalary = salary;

            // If employee joined this month, calculate pro-rata salary
            if (joiningYear === currentYear && joiningMonth === currentMonth) {
                var totalDaysInMonth = new Date(currentYear, currentMonth, 0).getDate();
                var workingDays = totalDaysInMonth - dateJoined.getDate() + 1;
                finalSalary = (salary / totalDaysInMonth) * workingDays;
            }

            totalSalaries += finalSalary;
            salaryDetails.push(`${employeeName}: PKR ${finalSalary.toFixed(2)}`);

            // ✅ Store employee row to update Salary Received status
            employeesToUpdate.push(i + 1); // +1 because sheet rows start at 1
        }

        // ✅ Get the Updated Running Total from Dashboard Sheet
        var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
        var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;

        if (currentRunningTotal >= totalSalaries) {
            var newRunningTotal = currentRunningTotal - totalSalaries;
            runningTotalCell.setValue(newRunningTotal); // ✅ Update Dashboard Sheet
        } else {
            throw new Error("Not enough balance in Running Total to pay salaries.");
        }

        // ✅ Mark Salary as "Received" for the employees we just paid
        employeesToUpdate.forEach(row => {
            employeeSheet.getRange(row, 6).setValue("Received"); // Column F
        });

        // ✅ Update Dashboard Sheet
        updateDashboard();

        // Send email notification
        var ownerEmail = Session.getEffectiveUser().getEmail();
        var subject = "Salary Payment Processed";
        var body = `Dear Owner,\n\nThe salaries for this month have been processed successfully.\n\nTotal Salaries Paid: PKR ${totalSalaries.toFixed(2)}\n\nBreakdown:\n${salaryDetails.join("\n")}\n\nNew Running Total: PKR ${newRunningTotal.toFixed(2)}\n\nBest Regards,\nAccounting System`;

        MailApp.sendEmail(ownerEmail, subject, body);

        return { status: 'success', message: `Salaries paid and email sent to ${ownerEmail}` };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}

/**
 * payUnpaidExpenses - Pay all unpaid expenses and update the dashboard.
 */
function payUnpaidExpenses() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var expensesSheet = ss.getSheetByName('Expenses');
        var dashboardSheet = ss.getSheetByName('Dashboard');

        if (!expensesSheet || !dashboardSheet) {
            throw new Error("One or more required sheets are missing.");
        }

        var expensesData = expensesSheet.getDataRange().getValues();
        var totalUnpaidAmount = 0;
        var expensesToUpdate = [];

        for (var i = 1; i < expensesData.length; i++) {
            var expenseStatus = String(expensesData[i][4] || "").trim().toLowerCase();
            var expenseAmount = parseFloat(expensesData[i][3]) || 0;

            if (expenseStatus === "unpaid") { // Only process unpaid expenses
                totalUnpaidAmount += expenseAmount;
                expensesToUpdate.push(i + 1);
            }
        }

        var runningTotalCell = dashboardSheet.getRange("B2"); // Running Total (B2)
        var currentRunningTotal = parseFloat(runningTotalCell.getValue()) || 0;

        if (currentRunningTotal >= totalUnpaidAmount) {
            runningTotalCell.setValue(currentRunningTotal - totalUnpaidAmount);
        } else {
            throw new Error("Not enough balance in Running Total to pay expenses.");
        }

        // ✅ Mark Unpaid Expenses as "Paid"
        expensesToUpdate.forEach(row => {
            expensesSheet.getRange(row, 5).setValue("Paid"); // Column E
        });

        // ✅ Update Dashboard
        updateDashboard();

        return { status: 'success', message: "Expenses paid successfully." };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}


/**
 * updateDashboard - Recalculates and updates the dashboard only when needed.
 */
function updateDashboard() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dashboardSheet = ss.getSheetByName('Dashboard');

    if (!dashboardSheet) {
        throw new Error("Dashboard sheet not found.");
    }

    var runningTotal = parseFloat(dashboardSheet.getRange("B2").getValue()) || 0;
    var partnerShare = 50; // Fixed Partner Share Percentage
    var partnerShareAmount = (runningTotal * partnerShare) / 100;

    // ✅ Update the Dashboard Sheet
    dashboardSheet.getRange("A2:C4").setValues([
        ["Running Total", runningTotal, "-"],
        ["Partner Share (%)", partnerShare, "50%"],
        ["Partner Share Amount", partnerShareAmount, "-"]
    ]);
}

function addMonthHeadings() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Payments'); // Adjust if needed

    if (!sheet) {
        SpreadsheetApp.getUi().alert("Sheet 'Payments' not found.");
        return;
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < 2) {
        SpreadsheetApp.getUi().alert("No data found in the 'Payments' sheet.");
        return;
    }

    var data = sheet.getDataRange().getValues();
    var lastMonth = "";
    var rowsToInsert = [];
    var monthTracker = {}; // Track first occurrence of each month

    // **Step 1: Detect Where to Insert New Month Headers**
    for (var i = 1; i < data.length; i++) { // Skip the header row
        var dateValue = data[i][1]; // Column B (Index 1)
        if (dateValue instanceof Date) {
            var monthYear = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "MMMM yyyy");

            // **Only store the first occurrence of each month**
            if (!monthTracker[monthYear]) {
                monthTracker[monthYear] = i + 1; // Store row index
            }
        }
    }

    // **Step 2: Insert Month Headers Before First Entry of Each Month**
    Object.entries(monthTracker).reverse().forEach(([month, rowIndex]) => {
        sheet.insertRowBefore(rowIndex);
        var range = sheet.getRange(rowIndex, 1, 1, lastColumn);
        range.merge();
        range.setValue(month);
        range.setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
        range.setBackground("#D9EAD3"); // Light green color for visibility
    });

    SpreadsheetApp.getUi().alert("Month headings added successfully!");
}

function generateMonthlySummary(selectedMonth, openingBalance) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var paymentsSheet = ss.getSheetByName('Payments');
    var expensesSheet = ss.getSheetByName('Expenses');
    var employeesSheet = ss.getSheetByName('Employees');
    var dashboardSheet = ss.getSheetByName('Dashboard');

    if (!paymentsSheet || !expensesSheet || !employeesSheet || !dashboardSheet) {
        return { status: 'error', message: "One or more required sheets are missing." };
    }

    var paymentsData = paymentsSheet.getDataRange().getValues();
    var expensesData = expensesSheet.getDataRange().getValues();
    var employeesData = employeesSheet.getDataRange().getValues();
    var closingBalance = dashboardSheet.getRange("B2").getValue(); // Running Total from Dashboard

    var totalPayments = 0;
    var totalSalaries = 0;
    var expensesByCategory = {};

    // **Step 1: Process Payments**
    for (var i = 1; i < paymentsData.length; i++) {
        var date = paymentsData[i][1]; // Date of Invoice (Column B)
        var status = paymentsData[i][6]; // Payment Status (Column G)
        var amount = parseFloat(paymentsData[i][9]) || 0; // Total Payment Received PKR (Column J)

        if (date instanceof Date && status === "Received") {
            var monthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy");
            if (monthYear === selectedMonth) {
                totalPayments += amount;
            }
        }
    }

    // **Step 2: Process Expenses**
    for (var i = 1; i < expensesData.length; i++) {
        var date = expensesData[i][0]; // Date (Column A)
        var category = expensesData[i][2]; // Category (Column C)
        var amount = parseFloat(expensesData[i][3]) || 0; // Amount (Column D)
        var status = expensesData[i][4]; // Status (Column E)

        if (date instanceof Date && status === "Paid") {
            var monthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy");
            if (monthYear === selectedMonth) {
                if (!expensesByCategory[category]) {
                    expensesByCategory[category] = 0;
                }
                expensesByCategory[category] += amount;
            }
        }
    }

    // **Step 3: Process Salaries**
    for (var i = 1; i < employeesData.length; i++) {
        var dateJoined = employeesData[i][3]; // Date Joined (Column D)
        var salary = parseFloat(employeesData[i][1]) || 0; // Salary (Column B)
        var status = employeesData[i][5]; // Salary Received (Column F)

        if (dateJoined instanceof Date && status === "Received") {
            var monthYear = Utilities.formatDate(dateJoined, Session.getScriptTimeZone(), "MMMM yyyy");
            if (monthYear === selectedMonth) {
                totalSalaries += salary;
            }
        }
    }

    var totalExpenses = Object.values(expensesByCategory).reduce((a, b) => a + b, 0);
    var columnOffset = 5; // Column E (Index 5)

    // **Step 4: Append the Monthly Summary to Dashboard**
    var lastDashboardRow = dashboardSheet.getLastRow() + 1;

    // **Beautification: Merge & Style the Month Header**
    var monthTitleRange = dashboardSheet.getRange(lastDashboardRow, columnOffset, 1, 2);
    monthTitleRange.merge();
    monthTitleRange.setValue(selectedMonth);
    monthTitleRange.setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
    monthTitleRange.setBackground("#4CAF50").setFontColor("white"); // Green Background

    var row = lastDashboardRow + 1;

    // **Beautify Each Row**
    function addRow(label, value, backgroundColor = "#F1F1F1", bold = false) {
        dashboardSheet.getRange(row, columnOffset).setValue(label);
        dashboardSheet.getRange(row, columnOffset + 1).setValue(value);
        dashboardSheet.getRange(row, columnOffset, 1, 2).setBackground(backgroundColor);
        if (bold) {
            dashboardSheet.getRange(row, columnOffset, 1, 2).setFontWeight("bold");
        }
        row++;
    }

    addRow("Opening Balance", openingBalance, "#D9EAD3", true); // Light Green
    addRow("Total Payments Received", totalPayments, "#FFF2CC", true); // Light Yellow
    addRow("Total Salaries Paid", totalSalaries, "#FFEBEE", true); // Light Red

    // **Add Expenses with Categorization**
    if (Object.keys(expensesByCategory).length > 0) {
        addRow("Expenses", "", "#BBDEFB", true); // Light Blue Header for Expenses
        for (var category in expensesByCategory) {
            addRow(" - " + category, expensesByCategory[category], "#E3F2FD"); // Light Blue Shade for Expenses
        }
    }

    addRow("Closing Balance (Running Total)", closingBalance, "#C8E6C9", true); // Light Green
    addRow("", "", "white"); // Empty Row for Spacing

    return { status: 'success', message: "Summary for " + selectedMonth + " added to Dashboard." };
}




function getAvailableMonths() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var paymentsSheet = ss.getSheetByName('Payments');
    var expensesSheet = ss.getSheetByName('Expenses');

    if (!paymentsSheet || !expensesSheet) {
        return [];
    }

    var paymentsData = paymentsSheet.getDataRange().getValues();
    var expensesData = expensesSheet.getDataRange().getValues();
    var months = new Set();

    for (var i = 1; i < paymentsData.length; i++) {
        var date = paymentsData[i][1];
        if (date instanceof Date) {
            months.add(Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy"));
        }
    }

    for (var i = 1; i < expensesData.length; i++) {
        var date = expensesData[i][0];
        if (date instanceof Date) {
            months.add(Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy"));
        }
    }

    return Array.from(months).sort();
}

