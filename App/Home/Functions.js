/// <reference path="../App.js" />

var loggingFlag = 99;
var logRange;

(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // If not using Excel 2016, return
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
                app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions (1.3+) of Excel.");
                return;
            }
        });
    }
})();

// Create & Switch to the Welcome sheet 
function viewWelcomeSheet(args) {
     // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        // Check App logging
        Checklog(ctx);
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                // Queue commands to select and activate Welcome sheet
                var sheet = ctx.workbook.worksheets.getItem("Welcome");
                sheet.activate();
                var range = sheet.getRange("A1");
                range.select();
                logEntry(ctx, "Welcome", "Welcome sheet is active.");
                //Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync();
            })
    })
    .catch(function (error) {
        handleError(error);
    });
    args.completed();
}

// Switch to the Dashboard sheet
function viewDashboard(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        // Check App logging
        Checklog(ctx);
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                // Queue commands to activate Dashboard sheet
                var sheet = ctx.workbook.worksheets.getItem("Dashboard");
                sheet.activate();
                var range = sheet.getRange("A1");
                range.select();
                logEntry(ctx, "Dashboard", "Dashboard sheet is active.");
                //Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync();
            })
    })
    .catch(function (error) {
        handleError(error);
    });
    args.completed();
}

// Switch to the Transactions sheet
function viewTransactions(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        // Check App logging
        Checklog(ctx);
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                // Queue commands to activate Dashboard sheet
                var sheet = ctx.workbook.worksheets.getItem("Actuals");
                sheet.activate();
                var range = sheet.getRange("A1");
                range.select();
                logEntry(ctx, "Actuals", "Actuals sheet is active.");
                //Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync();
            })
    })
    .catch(function (error) {
        handleError(error);
    });
    args.completed();
}

// Categories Setup on Budget sheet
function categoriesSetup(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        Checklog(ctx);
        // Queue a command to find name of active sheet
        var activesheet = ctx.workbook.worksheets.getActiveWorksheet();
        activesheet.load("name");
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                // Activate Budget worksheet
                var sheet = ctx.workbook.worksheets.getItem("Budget");
                sheet.activate();
                var range = sheet.getRange("F2");
                if (activesheet.name != "Budget") {
                    range.values = [["Next time, select BUDGET before clicking Categories."]];
                    range.select();
                    return ctx.sync();
                } 
                range.values= [[""]];
                logEntry(ctx, "Categories", "Budget sheet is active.");
                logEntry(ctx, "Categories", "Ready to setup CategoryTable.");
                // Delete Category table first.
                var CategoryTable = ctx.workbook.tables.getItem("CategoryTable");
                var CategoryRange = CategoryTable.getDataBodyRange();
                CategoryRange.delete("Up");
                var CategoryTableRows = CategoryTable.rows;
                logEntry(ctx, 'Categories', 'Cleared CategoryTable.');
                // Get the Income Table categories
                var incomeTable = ctx.workbook.tables.getItem("IncomeTable");
                var incomeColumnRange = incomeTable.columns.getItem("Category").getDataBodyRange().load("values");
                logEntry(ctx, 'Categories', 'Getting IncomeTable Categories.');
                // Get the Transfer Table categories
                var transferTable = ctx.workbook.tables.getItem("TransferTable");
                var transferColumnRange = transferTable.columns.getItem("Category").getDataBodyRange().load("values");
                logEntry(ctx, 'Categories', 'Getting TransferTable Categories.');
                // Get the Expense Table categories
                var expenseTable = ctx.workbook.tables.getItem("ExpenseTable");
                var expenseColumnRange = expenseTable.columns.getItem("Category").getDataBodyRange().load("values");
                logEntry(ctx, 'Categories', 'Getting ExpenseTable Categories.');
                return ctx.sync()
                    .then(function () {
                        // Process merge the Income, Transfer and Expense Category Column into Cleared Category Table
                        CategoryTableRows.add(null, incomeColumnRange.values);
                        CategoryTableRows.add(null, transferColumnRange.values);
                        CategoryTableRows.add(null, expenseColumnRange.values);
                        logEntry(ctx, 'Categories', 'Merging Income, Transfer & Expense categories into cleared CategoryTable.');
                        return ctx.sync();
                    })
            })
    })
    .catch(function (error) {
         handleError(error);
    });
    args.completed();
}

// Preview imported transactions
function previewTransactions(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        Checklog(ctx);
        // Queue a command to find name of active sheet
        var activesheet = ctx.workbook.worksheets.getActiveWorksheet();
        activesheet.load("name");
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                var sheet = ctx.workbook.worksheets.getItem("Import");
                sheet.activate();
                var range = sheet.getRange("V2");
                if (activesheet.name != "Import") {
                    range.values = [["Next time, select Import sheet before clicking Preview."]];
                    range.select();
                    return ctx.sync();
                };
                // We're on the Import sheet and ready to go
                range.values = [[""]];  // Clear prior error message, if any
                logEntry(ctx, "Preview", "Import sheet is active.");
                logEntry(ctx, "Preview", "Ready to process RAW data.");
                //Load import rule cell; hopefully it's the selected range
                var ruleCell = ctx.workbook.getSelectedRange().load("values, address");
                // load 9 columns after to get all rule details
                var ruleRow = ruleCell.getColumnsAfter(9); 
                ruleRow.load("values");  
                // Check to see if there's also RAW data to process
                // Drop/paste area for import trx is range B20:I1020
                // Only process up to 1,000 rows 
                var rangeAddress = "B20:I1020";  
                // Get the Address of the RAW Data Used Range in worksheet
                // Find area that actually contains data from user paste operation
                var rangeRAW = sheet.getRange(rangeAddress).getUsedRange();
                // Queue Load the data command for next sync.
                rangeRAW.load("address,columnCount,rowCount, values");
                // Queue commands to get DataBodyrange of results table
                var importResultsTable = ctx.workbook.tables.getItem("ImportResults");
                var importResultsRange = importResultsTable.getDataBodyRange();
                importResultsRange.load("rowCount");
                //Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync()
                    .then(function () {
                        //Check to see if an import rule is selected
                        logEntry(ctx, 'Preview', 'Import Range: ' + rangeRAW.address + ' Columns:' + rangeRAW.columnCount + ' Rows:' + rangeRAW.rowCount);
                        logEntry(ctx, 'Preview', 'Data Row 1: ' + rangeRAW.values[1][0] +', '+ rangeRAW.values[1][1] +', '+ rangeRAW.values[1][2] +','+ rangeRAW.values[1][3] +','+ rangeRAW.values[1][4]);
                        var RuleType = ruleRow.values[0][0];
                        var StartRow = ruleRow.values[0][1];
                        var PayeeCol = ruleRow.values[0][2];
                        var DateCol = ruleRow.values[0][3];
                        var AmtCol = ruleRow.values[0][4];
                        var ReverseFlag = ruleRow.values[0][6];
                        var ReverseAmt = 1;
                        if (ReverseFlag == 'Y') {
                            ReverseAmt = -1;
                        }
                        var ImportZeroTrxFlag = ruleRow.values[0][8];
                        if ((RuleType !='CSV') && (RuleType != 'QIF')) {
                                logEntry(ctx, 'Preview', 'Invalid Rule Type: ' + RuleType);
                                range.values = [["Invalid Rule:" + ruleCell.address + ", Please select a Rule by name in column B."]];
                                range.select();
                                return ctx.sync();
                        }
                        logEntry(ctx, 'Preview', 'Rule Selected: ' + ruleCell.address + ', ' + ruleCell.values + ', ' + ruleRow.values);
                        range.values = [["Processing with Rule: " + ruleCell.address + " Name: " + ruleCell.values]];
                        // Delete Results table and process transformed data into it
                        if (importResultsRange.rowCount > 1) {
                            importResultsRange.delete("Up");
                            logEntry(ctx, 'Preview', 'Cleared Preview Results table.');
                        }
                        for (var i = StartRow - 1; i < rangeRAW.rowCount; i++) {
                            if ((ImportZeroTrxFlag == 'N') && (rangeRAW.values[i][AmtCol - 1] == 0)) {
                                // skip zero transactions
                            } else { 
                            importResultsTable.rows.add(null, [[rangeRAW.values[i][PayeeCol - 1], rangeRAW.values[i][DateCol - 1], ruleRow.values[0][7], "Expense", rangeRAW.values[i][AmtCol - 1] * ReverseAmt, ruleRow.values[0][5], "", "Imported!"]]);
                            }
                        }
                        logEntry(ctx, 'Preview', 'Processed and added each row of import data to Preview Results table.');
                        // Queue a command to sort data by the fourth column of the table (descending)
                        var sortRange = importResultsTable.getDataBodyRange();
                        sortRange.sort.apply([{ key: 1, ascending: true, },]);
                        logEntry(ctx, 'Preview', 'Sorted Preview Results table ascending by date.');
                        return ctx.sync();
                    })

            })
    })
    .catch(function (error) {
        console.log("Error: " + error);
        handleError(error);
    });
    args.completed();
}

// Accept imported transactions 
function acceptTransactions(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        Checklog(ctx);
        // Queue a command to find name of active sheet
        var activesheet = ctx.workbook.worksheets.getActiveWorksheet();
        activesheet.load("name");
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                var sheet = ctx.workbook.worksheets.getItem("Import");
                sheet.activate();
                var range = sheet.getRange("V2");
                if (activesheet.name != "Import") {
                    range.values = [["Next time, select Import sheet before clicking Accept."]];
                    range.select();
                    return ctx.sync();
                }
                // We're on the Import sheet and ready to go
                range.values = [["Accepting Formatted RESULT data."]];  // Clear prior error message, if any
                logEntry(ctx, "Preview", "Import sheet is active.");
                logEntry(ctx, "Preview", "Ready to Accept Formatted RESULT data into Actuals table.");
                // Get the Preview Table results
                var importResultsTable = ctx.workbook.tables.getItem("ImportResults");
                var importResultsRange = importResultsTable.getDataBodyRange();
                importResultsRange.load("formulas");
                var transactionTable = ctx.workbook.tables.getItem("Trx_Table");
                var tableRows = transactionTable.rows;
                return ctx.sync()
                    .then(function () {
                        // Add all rows of importTable to end of Trx_Table
                        tableRows.add(null, importResultsRange.formulas);
                        logEntry(ctx, 'Accept', 'Formatted RESULT data copied to Trx_Table.');
                        // Delete Results table and process transformed data into it
                        importResultsRange.delete("Up");
                        logEntry(ctx, 'Accept', 'Formatted RESULT table cleared');
                        range.values = [["Done. Formatted RESULT data copied."]];
                        // Sync to run the queued command in Excel
                        return ctx.sync();
                    })
            })
    })
    .catch(function (error) {
            handleError(error);
    });
    args.completed();
}

// Year-End Reset on Reconcile and Transaction sheets
function myFinanceReset(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        Checklog(ctx);
        // Queue a command to find name of active sheet
        var activesheet = ctx.workbook.worksheets.getActiveWorksheet();
        activesheet.load("name");
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                var sheet = ctx.workbook.worksheets.getItem("Reconcile");
                sheet.activate();
                var range = sheet.getRange("I2");
                if (activesheet.name != "Reconcile") {
                    range.values = [["Next time, select Reconcile sheet before clicking Reset."]];
                    range.select();
                    return ctx.sync();
                }
                // Clear any prior error messages and make log entries
                range.values = [[""]];
                logEntry(ctx, "Reset", "Reconcile sheet is active.");
                logEntry(ctx, "Reset", "Ready to reset workbook for new year.");
                // Check to make sure user has entered "CLEAR" in the Reconcile Table Entry
                resetRange = sheet.getRange("Reset_Flag");
                resetRange.load("values");
                endBalanceRange = sheet.getRange("End_Balance");
                endBalanceRange.load("values");
                beginBalanceRange = sheet.getRange("Begin_Balance");
                beginBalanceRange.load("values, address");
                return ctx.sync()
                    .then(function () {
                        if (resetRange.values[0] == "CLEAR") {
                            //Display user message that we're processing end of year reset
                            range.values = [["Processing End of Year RESET"]];  
                            // First erase CLEAR so function will not run again by accident if Reset menu option is chosen
                            resetRange.values = [[""]];
                            // Copy each accounts year-ending balance to year-begin balance, and clear monthly reconciliation entries
                            logEntry(ctx, "Reset", "Copying " + (endBalanceRange.values.length + 1) + " balances and clearing reconcilation entries");
                            var reconcileRanges = []; // Array to hold address of ranges to be cleared
                            var recRow = beginBalanceRange.address.substr(beginBalanceRange.address.length - 1, 1); //Last digit of address is Row #
                            var recCol = beginBalanceRange.address.substr(beginBalanceRange.address.length - 5, 1); //5th digit from end of address is starting Column
                            for (var i = 0; i < endBalanceRange.values.length + 1; i++) {
                                var cellToChange = beginBalanceRange.getCell(0, i * 4); // Each Begin balance entry is offset by 4 columns from first
                                cellToChange.values = endBalanceRange.values[0][i];       // Copy Ending balance to Beginning balance
                                //  Now clear the corresponding reconciliation entries for this(i) account's beginning balance
                                // RecCol needs to be offset by 3 columns on each pass- FIX this
                                reconcileRanges[i] = "Reconcile!" + recCol + (Number(recRow) + 4) + ":" + recCol + (Number(recRow) + 15); 
                                range = sheet.getRange(reconcileRanges[i]);
                                range.values = 0;  // set previously entered end of month institution balance to zero
                            }
                            return ctx.sync()
                                .then(function () {
                                    logEntry(ctx, 'Reset', 'Cleared reconciliation entries.');
                                    // Now, Delete the transactions table Data Body
                                    var trxTable = ctx.workbook.tables.getItem("Trx_Table");
                                    var trxRange = trxTable.getDataBodyRange();
                                    trxRange.delete("Up");
                                    logEntry(ctx, 'Reset', 'Cleared entries in Actuals table Trx_Table.');
                                    var pivotTable = ctx.workbook.pivotTables.getItem("ReconcilePivot");
                                    pivotTable.refresh();
                                    var pivotTable = ctx.workbook.pivotTables.getItem("IncomeExpensePivot");
                                    pivotTable.refresh();
                                    logEntry(ctx, 'Reset', 'Refreshed Income/Expense & Reconcile PivotTables.');
                                    return ctx.sync();
                                })
                        } else {
                            range.values = [['You must enter "CLEAR" for this function to proceed. READ NOTES! ']];
                            range.select();
                            logEntry(ctx, "Reset", "User hasn't entered CLEAR, function aborted.");
                            return ctx.sync();
                        }

                    })
            })
    })
    .catch(function (error) {
        handleError(error);
    });
    args.completed();
}

function toggleLogging(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        // Check App logging
        Checklog(ctx);
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                var activesheet = ctx.workbook.worksheets.getActiveWorksheet();
                activesheet.load("name");
                var sheet = ctx.workbook.worksheets.getItem("Log");
                sheet.activate();
                //Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync()
                    .then(function () {
                        var range = sheet.getRange("G2");
                        if (activesheet.name != "Log") {
                            range.values = [["Next time, select Log sheet before clicking Toggle."]];
                            range.select();
                            return ctx.sync();
                        }
                        range.values = [[""]];  // Clear error message if any
                        logEntry(ctx, "Toggle", "Log sheet is active.");
                        logEntry(ctx, "Toggle", "Log entries are placed here by myFinance v0.8 commands");
                        logEntry(ctx, "Toggle", "Clear this list whenever you want!");
                        // Toggle Flag
                        if (loggingFlag == 0) {
                            logRange.values = [["On"]];
                            loggingFlag = 1;
                            logEntry(ctx, "Toggle", "Logging is turned: " + logRange.values[0]);
                        } else {
                            logRange.values = [["Off"]];
                            logEntry(ctx, "Toggle", "Logging is turned: " + logRange.values[0]);
                            loggingFlag = 0;
                        }
                        return ctx.sync();
                    })
            })
    })
    .catch(function (error) {
            handleError(error);
    });
    args.completed();
}

function clearLogging(args) {
    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {
        // Check App logging
        Checklog(ctx);
        return ctx.sync()
            .then(function () {
                if (logRange.values[0] == "Off") {
                    loggingFlag = 0;
                } else {
                    loggingFlag = 1;
                }
                var activesheet = ctx.workbook.worksheets.getActiveWorksheet();
                activesheet.load("name");
                var sheet = ctx.workbook.worksheets.getItem("Log");
                sheet.activate();
                //Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync()
                    .then(function () {
                        var range = sheet.getRange("G2");
                        if (activesheet.name != "Log") {
                            range.values = [["Next time, select Log sheet before clicking Toggle."]];
                            range.select();
                            return ctx.sync();
                        }
                        range.values = [[""]];  // Clear error message if any
                        logEntry(ctx, "ClearLog", "Log sheet is active.");

                        // Delete Logging Table Data Body
                        var logTable = ctx.workbook.tables.getItem("logTable");
                        var logRange = logTable.getDataBodyRange();
                        logRange.delete("Up");
                        logEntry(ctx, 'ClearLog', 'Cleared LogTable.');
                        return ctx.sync();
                    })
            })
    })
    .catch(function (error) {
        handleError(error);
    });
    args.completed();
}

// Handle errors
function handleError(error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
}

// Log Application Events
function logEntry(ctx, section, entry) {
    if (loggingFlag == 1) {
        var logTable = ctx.workbook.tables.getItem("logTable");
        var d = new Date();
        var logdate = d.getMonth() + 1 + '/' + d.getDate() + '/' + d.getFullYear() + '-' + d.getHours() + ':' + d.getMinutes() + ':' + d.getSeconds();
        logTable.rows.add(null /* add rows to end */, [[logdate, section, entry]]);
    }
}

function Checklog(ctx) {
    // Queue commands to check the automation Log table
    if (loggingFlag == 99) {
        var sheet = ctx.workbook.worksheets.getItem("Log");
        logRange = sheet.getRange("LogState");
        logRange.load("values");
    }
}

/**
 * Courtesy of Chris West's Blog 
 * Takes a positive integer and returns the corresponding column name.
 * @param {number} num  The positive integer to convert to a column name.
 * @return {string}  The column name.
 */
function toColumnName(num) {
    for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
        ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
    }
    return ret;
}