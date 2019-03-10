'use strict';

(function () {
    Office.onReady(function () {
        // Office is ready
        console.log('Office is ready');
        if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
            console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
        }
        $(document).ready(function () {
            // The document is ready
            console.log('Document is ready');

            var today = new Date();
            var mm = today.getMonth(); // January = 0
            var yyyy = today.getFullYear();
            var nextYear = yyyy + 1;
            console.log('Current year is: ' + yyyy + ' and current month is: ' + mm);
            console.log('Next year is: ' + nextYear);
            //Year drop down list
            var yearSelect = document.getElementById("yearSelect");
            var currentYearOption = document.createElement("OPTION");
            currentYearOption.text = yyyy;
            var nextYearOption = document.createElement("OPTION");
            nextYearOption.text = nextYear;
            yearSelect.add(currentYearOption);
            yearSelect.add(nextYearOption);
            if (mm == 11) {
                yearSelect.value = yyyy + 1; // If it's December, then next month is January of the next year
            } else { yearSelect.value = yyyy; }
            // Month drop down list
            var monthSelect = document.getElementById("monthSelect");
            var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            for (var i = 0; i < monthNames.length; i++) {
                var monthOption = document.createElement("OPTION");
                monthOption.text = monthNames[i];
                monthSelect.add(monthOption);
            }
            var nextMM = (mm + 1) % 12;
            console.log('nextMM is: ' + nextMM);
            var nextMonth = monthNames[nextMM];
            console.log('nextMonth is: ' + nextMonth);
            monthSelect.value = nextMonth;
            console.log('Selected month is: ' + nextMonth);
            Excel.run(function (context) {
                var wb = context.workbook;
                var functionResult = wb.functions.year(wb.functions.today());
                functionResult.load('value');
                return context.sync().then(function () {
                    console.log('Result of the function: ' + functionResult.value);
                }).catch(function (error) {
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                })
            });
        });
    });


}
)();

var button = $('#generate');
console.log(button);
$('#generate').click(createData);

function createData() {
    Excel.run(function (context) {
        var currentWorksheet = context.workbook.worksheets.getItemOrNullObject('Staff Data');
        console.log("current worksheet is: " + JSON.stringify(currentWorksheet, null, 4));
        if (currentWorksheet === undefined) {
            // This doesn't work!
            document.getElementById('errorText').innerHTML = "Was looking for a sheet called Staff Data but couldn't find one.";
        } else {
            var staffListRange = currentWorksheet.getUsedRangeOrNullObject();
            staffListRange.load("values");
    
            return context.sync()
                .then(function () {
                console.log(JSON.stringify(staffListRange.values, null, 4));
                console.log(staffListRange.values.length);
                var rosterTableHeaderRow = [];

                for (var i = 0; i < staffListRange.values.length; i++) {

                    for (var j = 0; j < staffListRange.values[i].length; j++) {
                        console.log('Next value is: ' + JSON.stringify(staffListRange.values[i][j], null, 4));

                        if (staffListRange.values[i][j] != '') {
                            rosterTableHeaderRow.push(staffListRange.values[i][j]);
                            //rosterTable.getHeaderRowRange().values[i] = rosterTableHeaderRow;
                        }
                    }
                }

                //for var (i = 0; i < )
                rosterTableHeaderRow.shift(); // Remove the header value (i.e. 'Staff List')
                rosterTableHeaderRow.unshift('', '');// Two columns for the day of the week and the date
                console.log(rosterTableHeaderRow);

                
                            });        
        
        

        }
    }
        ).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}
    
function createRoster(rosterTableHeaderRow) {
    Excel.run(function (context) {
        var rosterTable = currentWorksheet.tables.add("A2:AZ2", true /*hasHeaders*/);
        var rosterTableRange = rosterTable.getHeaderRowRange();
        rosterTable.getHeaderRowRange().values = rosterTableHeaderRow;
        console.log(JSON.stringify(rosterTableRange.values, null,4));
        rosterTable.name = "RosterTable";
        console.log('Made a roster table');
        rosterTable.getRange().format.autofitColumns();
        rosterTable.getRange().format.autofitRows();

        return context.sync();
// expensesTable.getHeaderRowRange().values =
        //     [["Date", "Merchant", "Category", "Amount"]];

        // expensesTable.rows.add(null /*add at the end*/, [
        //     ["1/1/2017", "The Phone Company", "Communications", "120"],
        //     ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        //     ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        //     ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        //     ["1/11/2017", "Bellows College", "Education", "350.1"],
        //     ["1/15/2017", "Trey Research", "Other", "135"],
        //     ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
        // ]);
        // expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];        
        //return context.sync().then(function() {
            //console.log(worksheetTables);
    }        
    ).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}