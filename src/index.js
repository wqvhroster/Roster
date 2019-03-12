'use strict';
import moment from 'moment';

var unavailableStaff = [[], [], [], [], [], [], []];
var datesByDaysOfTheWeek = [[], [], [], [], [], [], []];
var pubHols = [];

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
            /*Excel.run(function (context) {
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
            });*/
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
        var staffListRange = currentWorksheet.getUsedRangeOrNullObject();
        staffListRange.load("values");
        return context.sync()
            .then(function () {

                if (staffListRange.values === undefined) {
                    document.getElementById('errorText').innerHTML = "Error! Was looking for a sheet called Staff Data but couldn't find one.";
                    return;
                } else {
                    document.getElementById('errorText').innerHTML = "Happy days!";

                }
                console.log(JSON.stringify(staffListRange.values, null, 4));
                console.log(staffListRange.values.length);
                var rosterTableHeaderRow = [];

                // Start at line 1 to ignore the header row
                for (var i = 1; i < staffListRange.values.length; i++) {

                    if (staffListRange.values[i][0] != '') {
                        rosterTableHeaderRow.push(staffListRange.values[i][0]);
                    }

                    for (var j = 1; j < staffListRange.values[i].length; j++) {
                        if (staffListRange.values[i][j] != "") {
                            unavailableStaff[j - 1].push(i - 1);
                        }
                    }

                }

                var calendar = generateCalendar(rosterTableHeaderRow.length);
                rosterTableHeaderRow.unshift('', '');// Two columns for the day of the week and the date
                console.log(rosterTableHeaderRow);

                calendar.unshift(rosterTableHeaderRow); // Add the header row to the top of the calendar
                console.log("Here's the calendar: " + calendar);
                // Figure out what the public holidays are:
                var pubHolSheet = context.workbook.worksheets.getItemOrNullObject('Public Holidays');
                var pubHolRange = pubHolSheet.getUsedRangeOrNullObject();
                populatePublicHolidayData(context, pubHolRange);

                // Create the roster
//                var lastCell = convertToNumberingScheme(rosterTableHeaderRow.length);
//                lastCell += calendar.length;
//                console.log("the last cell for the roster range is: " + lastCell);
//                context.workbook.worksheets.add("Roster");
//                var rosterRange = context.workbook.worksheets.getItemOrNullObject('Roster').getRange("A1:" + lastCell);
//                rosterRange.values = calendar;

//                return context.sync().then(function(context) {
//                    formatRoster(context, rosterRange);

//                });

                return context.sync();

            });




    }
    ).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

function formatRoster(context, rosterRange) {
    rosterRange.load("values");
    return context.sync().then(function () {
        //console.log("range values size is: " + rosterRange.values.length);
        for (var i = 0; i < rosterRange.values.length; i++) {
            // Find the weekends and set the fill to grey
            var row = rosterRange.values[i];
            //console.log ("Current row is: " + JSON.stringify(row, null, 4));
            //console.log ("First column in the current row is: " + row[0]);
            if (row[0] === "Sat" || row[0] === "Sun") {
                var rangeString = "A" + (i + 1) + ":" + convertToNumberingScheme(row.length) + (i + 1);
                //console.log(rangeString);
                var weekendRange = context.workbook.worksheets.getItemOrNullObject('Roster').getRange(rangeString);
                weekendRange.format.fill.color = "#C8C8C8";
            }

        }
        for (var i = 0; i < datesByDaysOfTheWeek.length; i++) {
            var staff = unavailableStaff[i];
            var days = datesByDaysOfTheWeek[i];
            //console.log("For dates: " + datesByDaysOfTheWeek[i] + ", these staff are unavailable: " + unavailableStaff[i]);
            for (var j = 0; j < staff.length; j++) {
                for (var k = 0; k < days.length; k++) {
                    var column = staff[j] + 3;
                    var row = days[k] + 1;
                    //console.log("I wish to grey out cell: " + " j: " + column + ", k: " + row);
                    var rangeString = convertToNumberingScheme(column) + row;
                    var dayOffRange = context.workbook.worksheets.getItemOrNullObject('Roster').getRange(rangeString);
                    dayOffRange.format.fill.color = "#C8C8C8";
                }
            }
        }
        rosterRange.format.autofitColumns();
        return context.sync();

    });

}

function generateCalendar(numColumns) {
    console.log(monthSelect.value + ' ' + yearSelect.value);
    var daysOfWeek = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    var startDate = new Date(yearSelect.value, monthSelect.selectedIndex, 1);
    var endDate = new Date(yearSelect.value, monthSelect.selectedIndex + 1, 0);
    var endDay = endDate.getDate();
    console.log("Start date is: " + startDate.toDateString() + "; end date is: " + endDate.toDateString());
    var calendar = [];
    for (var i = 1; i < endDay + 1; i++) {
        var row = [];
        var currentDate = new Date(yearSelect.value, monthSelect.selectedIndex, i);
        datesByDaysOfTheWeek[currentDate.getDay()].push(currentDate.getDate());
        row.push(daysOfWeek[currentDate.getDay()]);
        row.push(currentDate.getDate());

        for (var j = 0; j < numColumns; j++) {
            row.push(""); //Everyone starts off working every day
        }
        calendar[i - 1] = row;
    }
    return calendar;
}

function populatePublicHolidayData(context, pubHolRange) {
    console.log("Getting public holiday values");
    pubHolRange.load("values");
    return context.sync().then(function() {
        if (pubHolRange.values === undefined) {
            document.getElementById('errorText').innerHTML = "Warning! Was looking for a sheet called Public Holidays but couldn't find one. Public holidays will not be applied! Continuing anyway...";
        } else {
            console.log("Found public holiday values: " + JSON.stringify(pubHolRange.values, null, 4));
            console.log("From this moment... " + moment().format());

        }
    
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    

}

function convertToNumberingScheme(number) {
    var baseChar = ("A").charCodeAt(0),
        letters = "";

    do {
        number -= 1;
        letters = String.fromCharCode(baseChar + (number % 26)) + letters;
        number = (number / 26) >> 0; // quick `floor`
    } while (number > 0);

    return letters;
}